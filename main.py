import os
import sys
import json
import time
import shutil
import subprocess
from xml.sax.saxutils import escape as _xml_escape
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

# --- 全局常量 ---
CONFIG_FILE = os.path.expanduser('~/.knock_sync_gui_config.json')
PLIST_PATH = os.path.expanduser('~/Library/LaunchAgents/com.user.knock.sync.plist')
PLIST_LABEL = "com.user.knock.sync"
WIN_TASK_NAME = "KnockSyncBackupDaemon"
WIN_STARTUP_CMD_NAME = "KnockSyncBackup-daemon.cmd"

_IS_WIN = sys.platform == "win32"
_IS_MAC = sys.platform == "darwin"


def _daemon_log_paths():
    """后台任务日志路径（与界面提示一致）。"""
    if _IS_WIN:
        base = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")), "knock_sync_logs")
        os.makedirs(base, exist_ok=True)
        return (
            os.path.join(base, "knock_sync_out.log"),
            os.path.join(base, "knock_sync_err.log"),
        )
    return "/tmp/knock_sync_out.log", "/tmp/knock_sync_err.log"


def _redirect_daemon_stdio_if_needed():
    """Windows 下无控制台时，将 print 写入临时目录日志。"""
    if not _IS_WIN or not getattr(sys, "frozen", False):
        return
    try:
        out_p, err_p = _daemon_log_paths()
        sys.stdout = open(out_p, "a", encoding="utf-8", buffering=1)
        sys.stderr = open(err_p, "a", encoding="utf-8", buffering=1)
    except Exception:
        pass


def _win_daemon_tr_string(exe_path):
    """供 schtasks /tr 使用的命令行（处理路径中的空格）。"""
    exe_path = os.path.normpath(exe_path)
    if " " in exe_path:
        return f'"{exe_path}" --daemon'
    return f"{exe_path} --daemon"


def _windows_task_command_line():
    """当前进程对应的「登录时启动后台」完整命令行。"""
    if getattr(sys, "frozen", False):
        return _win_daemon_tr_string(os.path.realpath(sys.executable))
    py = os.path.normpath(sys.executable)
    script = os.path.normpath(os.path.abspath(sys.argv[0]))
    return f'"{py}" "{script}" --daemon'


def _win_startup_cmd_path():
    """当前用户「启动」文件夹中的本程序启动脚本路径。"""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    return os.path.join(
        appdata,
        "Microsoft",
        "Windows",
        "Start Menu",
        "Programs",
        "Startup",
        WIN_STARTUP_CMD_NAME,
    )


def _remove_windows_startup_cmd():
    p = _win_startup_cmd_path()
    if p and os.path.isfile(p):
        try:
            os.remove(p)
        except OSError:
            pass


def _write_windows_startup_cmd():
    """
    在「启动」文件夹写入 .cmd，登录后自动运行 --daemon（无需计划任务权限）。
    使用 UTF-8 BOM，避免中文路径乱码。
    """
    path = _win_startup_cmd_path()
    if not path:
        raise RuntimeError("无法获取 %APPDATA%，无法写入启动文件夹。")
    os.makedirs(os.path.dirname(path), exist_ok=True)

    if getattr(sys, "frozen", False):
        exe = os.path.normpath(os.path.realpath(sys.executable))
        line = f'start "" /min "{exe}" --daemon'
    else:
        py = os.path.normpath(sys.executable)
        script = os.path.normpath(os.path.abspath(sys.argv[0]))
        line = f'start "" /min "{py}" "{script}" --daemon'

    content = (
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        f"{line}\r\n"
    )
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write(content)
    return path


def _install_windows_startup_task():
    """
    优先创建计划任务（登录时运行）；若遇「拒绝访问」等权限问题，
    则回退到当前用户的「启动」文件夹中的 .cmd（效果同为登录后自启）。
    返回 ("schtasks", None) 或 ("startup_cmd", 脚本路径)。
    """
    sub_kw = {}
    if _IS_WIN and hasattr(subprocess, "CREATE_NO_WINDOW"):
        sub_kw["creationflags"] = subprocess.CREATE_NO_WINDOW

    subprocess.run(
        ["schtasks", "/delete", "/tn", WIN_TASK_NAME, "/f"],
        capture_output=True,
        **sub_kw,
    )

    tr = _windows_task_command_line()
    # 多种参数组合，提高在家庭版/企业策略下的成功率
    attempts = [
        ["schtasks", "/create", "/tn", WIN_TASK_NAME, "/tr", tr, "/sc", "onlogon", "/f"],
        ["schtasks", "/create", "/tn", WIN_TASK_NAME, "/tr", tr, "/sc", "onlogon", "/rl", "LIMITED", "/f"],
    ]
    last_err = ""
    for args in attempts:
        r = subprocess.run(
            args,
            capture_output=True,
            text=True,
            **sub_kw,
        )
        if r.returncode == 0:
            _remove_windows_startup_cmd()
            return "schtasks", None
        last_err = ((r.stderr or "") + (r.stdout or "")).strip() or f"退出码 {r.returncode}"

    try:
        cmd_path = _write_windows_startup_cmd()
        return "startup_cmd", cmd_path
    except Exception as e:
        raise RuntimeError(
            "计划任务创建失败（可能被组策略限制或需管理员权限）：\n"
            f"{last_err}\n\n"
            f"启动文件夹回退也失败：{e}"
        ) from e


def _plist_xml_string(s):
    """plist 内 <string> 需转义 &、<、>。"""
    return _xml_escape(str(s), {'"': "&quot;", "'": "&apos;"})


def _sync_done_hint(morning=True):
    if getattr(sys, "frozen", False):
        out_p, err_p = _daemon_log_paths()
        return (
            "同步已完成。打包版无终端窗口，请用「诊断源目录」查看；"
            f"后台日志：{out_p}、{err_p}"
        )
    verb = "抽存" if morning else "灌回"
    return f"{verb}同步完毕！详情请看终端。"


# ==========================================
# 核心引擎：改用 pathlib.rglob（与旧脚本一致）
# ==========================================

def normalize_path_value(path_value):
    """
    兼容 GUI 输入异常（如多行重复路径），返回可用的单一路径字符串。
    """
    if path_value is None:
        return ""
    text = str(path_value).replace("\r", "\n").strip()
    if not text:
        return ""
    parts = [line.strip() for line in text.split("\n") if line.strip()]
    return parts[0] if parts else ""

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def _count_files_under(root: Path):
    """递归统计目录下文件总数（用于诊断）；失败时返回 None。"""
    try:
        n = 0
        for child in root.rglob("*"):
            if child.is_file():
                n += 1
                if n > 200000:
                    return n
        return n
    except (PermissionError, OSError):
        return None


def _mtime_ns(path: Path) -> int:
    """文件修改时间（纳秒），便于跨平台比较。"""
    st = path.stat()
    ns = getattr(st, "st_mtime_ns", None)
    if ns is not None:
        return ns
    return int(st.st_mtime * 1_000_000_000)


def _merge_config_patch(patch: dict):
    cfg = load_config()
    cfg.update(patch)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)


def _ask_overwrite_newer_dialog(master, sample_rel: str):
    """源较新需覆盖时弹窗。返回 (是否覆盖本次全部较新项, 是否记住长期策略)。"""
    result = {'ok': False, 'remember': False}
    win = tk.Toplevel(master)
    win.title("确认覆盖较新的源文件")
    win.transient(master)
    win.grab_set()
    msg = (
        "检测到源文件修改时间较新，将覆盖目标中的同名文件。\n\n"
        f"示例：{sample_rel}\n\n"
        "• 确定：本次同步中，所有「源较新」的情况均覆盖目标。\n"
        "• 取消：本次同步中，均不覆盖（仅复制新增文件）。\n"
    )
    tk.Label(win, text=msg, justify="left", wraplength=460).pack(padx=12, pady=10)
    remember_var = tk.BooleanVar(value=False)
    tk.Checkbutton(
        win,
        text="记住本次选择（确定=以后始终覆盖较新；取消=以后始终跳过较新、仅新增）",
        variable=remember_var,
    ).pack(anchor="w", padx=12)

    bf = tk.Frame(win)
    bf.pack(pady=10)

    def on_ok():
        result['ok'] = True
        result['remember'] = remember_var.get()
        win.destroy()

    def on_cancel():
        result['ok'] = False
        result['remember'] = remember_var.get()
        win.destroy()

    tk.Button(bf, text="确定", command=on_ok, width=10).pack(side="left", padx=5)
    tk.Button(bf, text="取消", command=on_cancel, width=10).pack(side="left", padx=5)
    win.wait_window()
    return result['ok'], result['remember']


def _resolve_overwrite_newer(
    gui_master,
    gui_interactive,
    overwrite_session,
    on_policy_saved,
    sample_relative,
):
    """
    返回 True：执行覆盖；False：跳过本次覆盖。
    overwrite_session 为可变 dict，键 'decision'：None | True | False，供同一轮多次 copy 共用。
    """
    cfg = load_config()
    mode = cfg.get('newer_overwrite_mode', 'ask')
    if mode == 'always':
        return True
    if mode == 'skip_newer':
        return False
    # mode == ask
    if not gui_interactive:
        return True
    if gui_master is None:
        return True
    if overwrite_session is None:
        overwrite_session = {}
    if overwrite_session.get('decision') is not None:
        return overwrite_session['decision']

    ok, remember = _ask_overwrite_newer_dialog(gui_master, str(sample_relative))
    if remember:
        if ok:
            _merge_config_patch({'newer_overwrite_mode': 'always'})
            if on_policy_saved:
                on_policy_saved('always')
        else:
            _merge_config_patch({'newer_overwrite_mode': 'skip_newer'})
            if on_policy_saved:
                on_policy_saved('skip_newer')
    overwrite_session['decision'] = bool(ok)
    return ok


def copy_without_overwrite(
    src_dir, tgt_dir, log_prefix="",
    *,
    gui_master=None,
    gui_interactive=False,
    overwrite_session=None,
    on_policy_saved=None,
):
    """
    用 pathlib.rglob 深度遍历（与你旧脚本一致的方式）
    + iCloud 占位符自动检测下载
    + 目标已存在同名文件时：仅当源文件修改时间更新才覆盖（以较新者为准）
    """
    if not src_dir or not tgt_dir:
        print(f"[{datetime.now()}] {log_prefix} ❌ 路径未设置 (src={src_dir}, tgt={tgt_dir})")
        return

    src_dir = normalize_path_value(src_dir)
    tgt_dir = normalize_path_value(tgt_dir)
    src = Path(src_dir).expanduser().resolve()
    tgt = Path(tgt_dir).expanduser().resolve()

    if not src.exists():
        print(f"[{datetime.now()}] {log_prefix} ❌ 源文件夹不存在: {src}")
        return

    tgt.mkdir(parents=True, exist_ok=True)
    print(f"[{datetime.now()}] {log_prefix} 🚀 开始扫描: {src}")

    if overwrite_session is None:
        overwrite_session = {'decision': None}

    # ===== 诊断输出：看看 Python 到底看到了什么 =====
    try:
        raw_items = sorted(src.iterdir(), key=lambda p: p.name)
        print(f"   (诊断) 👀 顶层共 {len(raw_items)} 个项目:")
        for item in raw_items:
            if item.name == '.DS_Store':
                continue
            if item.is_dir():
                tag = "📁 文件夹"
            elif item.name.startswith('.') and item.name.endswith('.icloud'):
                tag = "☁️  iCloud云端"
            else:
                tag = "📄 本地文件"
            print(f"   (诊断)   {tag}  {item.name}")
    except Exception as e:
        print(f"   (诊断) ⛔ 无法读取目录: {e}")

    total_under = _count_files_under(src)
    if total_under is not None:
        print(f"   (诊断) 📊 递归统计：整棵目录下共 {total_under} 个文件（含子文件夹内）")
    else:
        print(f"   (诊断) 📊 递归统计：无法遍历子目录（权限或路径限制）")

    count = 0
    skipped = 0
    cloud_downloaded = 0
    cloud_failed = 0

    try:
        # ===== 核心：用 rglob 递归遍历（和旧脚本完全一致） =====
        for src_file in src.rglob("*"):
            if not src_file.is_file():
                continue

            name = src_file.name

            # 跳过 macOS 系统垃圾
            if name in ('.DS_Store', '.localized', 'Thumbs.db', 'desktop.ini'):
                continue

            # ===== iCloud 占位符处理（仅 macOS） =====
            if _IS_MAC and name.startswith('.') and name.endswith('.icloud'):
                real_name = name[1:-7]  # .脚本号.docx.icloud → 脚本号.docx
                real_file = src_file.parent / real_name

                # 如果真实文件已经在本地了，直接用它
                if real_file.exists() and real_file.is_file():
                    src_file = real_file
                    name = real_name
                else:
                    # 触发 iCloud 下载
                    print(f"  ☁️  云端文件: {real_name}，正在触发下载...")
                    try:
                        subprocess.run(
                            ['brctl', 'download', str(src_file)],
                            capture_output=True, timeout=15
                        )
                    except FileNotFoundError:
                        # brctl 不存在时尝试 open 命令
                        try:
                            subprocess.run(
                                ['open', str(src_file)],
                                capture_output=True, timeout=15
                            )
                        except Exception:
                            pass
                    except Exception:
                        pass

                    # 等待下载完成（最多 120 秒）
                    downloaded = False
                    for wait_round in range(24):
                        if real_file.exists() and real_file.is_file():
                            downloaded = True
                            break
                        time.sleep(5)
                        if wait_round % 4 == 3:
                            print(f"      ⏳ 仍在等待下载: {real_name} ({(wait_round+1)*5}秒)")

                    if downloaded:
                        src_file = real_file
                        name = real_name
                        cloud_downloaded += 1
                        print(f"  ☁️→✅ 下载完成: {real_name}")
                    else:
                        cloud_failed += 1
                        print(f"  ☁️→❌ 下载超时: {real_name}")
                        continue

            # 跳过其余隐藏文件（.DS_Store 之类的，注意 iCloud 的已在上面处理）
            if name.startswith('.'):
                continue

            # ===== 计算目标路径（保持子目录结构） =====
            relative = src_file.relative_to(src)
            dst_file = tgt / relative

            # 同名：目标已存在且修改时间不早于源 → 跳过；否则覆盖（双向备份以较新的一侧为准）
            replacing = False
            if dst_file.exists():
                try:
                    if _mtime_ns(src_file) <= _mtime_ns(dst_file):
                        skipped += 1
                        continue
                    replacing = True
                except OSError:
                    print(f"  ! 无法比较修改时间，跳过: {relative}")
                    skipped += 1
                    continue

            if replacing:
                if not _resolve_overwrite_newer(
                    gui_master,
                    gui_interactive,
                    overwrite_session,
                    on_policy_saved,
                    relative,
                ):
                    skipped += 1
                    continue

            # 创建目标子目录 & 拷贝
            dst_file.parent.mkdir(parents=True, exist_ok=True)
            try:
                shutil.copy2(str(src_file), str(dst_file))
                count += 1
                if replacing:
                    print(f"  ↑ 覆盖(源较新): {relative}")
                else:
                    print(f"  + 新增: {relative}")
            except Exception as e:
                print(f"  ! 搬运失败: {relative} → {e}")

    except PermissionError as e:
        print(f"[{datetime.now()}] {log_prefix} ⛔️ 权限被拦截: {e}")
        if _IS_MAC:
            print(f"   请到【系统设置 → 隐私与安全 → 完全磁盘访问权限】中授权！")
        else:
            print(f"   请检查该文件夹的 NTFS 权限，或以管理员身份运行后再试。")
    except Exception as e:
        print(f"[{datetime.now()}] {log_prefix} ⛔️ 未知错误: {e}")
        import traceback
        traceback.print_exc()

    summary = (f"新增/覆盖 {count}, 跳过 {skipped}"
               + (f", 云端下载 {cloud_downloaded}" if cloud_downloaded else "")
               + (f", 云端超时 {cloud_failed}" if cloud_failed else ""))
    print(f"[{datetime.now()}] {log_prefix} ✅ 完成: {summary}\n")


def run_morning_task(
    gui_master=None,
    gui_interactive=False,
    on_policy_saved=None,
    overwrite_session=None,
):
    cfg = load_config()
    if overwrite_session is None:
        overwrite_session = {'decision': None}
    print("\n" + "=" * 60)
    print("=== 执行早晨同步规则 (抽出保存) ===")
    print("=" * 60)
    copy_without_overwrite(
        cfg.get('files_src'), cfg.get('files_tgt'), "[文档]",
        gui_master=gui_master,
        gui_interactive=gui_interactive,
        overwrite_session=overwrite_session,
        on_policy_saved=on_policy_saved,
    )
    copy_without_overwrite(
        cfg.get('images_src'), cfg.get('images_tgt'), "[图片]",
        gui_master=gui_master,
        gui_interactive=gui_interactive,
        overwrite_session=overwrite_session,
        on_policy_saved=on_policy_saved,
    )


def run_evening_task(
    gui_master=None,
    gui_interactive=False,
    on_policy_saved=None,
    overwrite_session=None,
):
    cfg = load_config()
    if overwrite_session is None:
        overwrite_session = {'decision': None}
    print("\n" + "=" * 60)
    print("=== 执行晚间同步规则 (灌回防删) ===")
    print("=" * 60)
    copy_without_overwrite(
        cfg.get('files_tgt'), cfg.get('files_src'), "[文档]",
        gui_master=gui_master,
        gui_interactive=gui_interactive,
        overwrite_session=overwrite_session,
        on_policy_saved=on_policy_saved,
    )
    copy_without_overwrite(
        cfg.get('images_tgt'), cfg.get('images_src'), "[图片]",
        gui_master=gui_master,
        gui_interactive=gui_interactive,
        overwrite_session=overwrite_session,
        on_policy_saved=on_policy_saved,
    )


def start_daemon_loop():
    _redirect_daemon_stdio_if_needed()
    cfg = load_config()
    morning = cfg.get('morning_time', "10:10")
    evening = cfg.get('evening_time', "19:00")
    print(f"[后台服务] 已启动！早: {morning}  晚: {evening}")

    try:
        import schedule
        schedule.every().day.at(morning).do(run_morning_task)
        schedule.every().day.at(evening).do(run_evening_task)
        while True:
            schedule.run_pending()
            time.sleep(30)
    except Exception as e:
        # 兜底：即使没有 schedule 依赖，也能按分钟轮询执行。
        print(f"[后台服务] schedule 不可用，启用内置调度: {e}")
        last_morning_date = None
        last_evening_date = None
        while True:
            now = datetime.now()
            now_hm = now.strftime("%H:%M")
            today = now.date()
            if now_hm == morning and last_morning_date != today:
                run_morning_task()
                last_morning_date = today
            if now_hm == evening and last_evening_date != today:
                run_evening_task()
                last_evening_date = today
            time.sleep(20)


# ==========================================
# GUI 面板
# ==========================================

class AppGUI:
    def __init__(self, root):
        self.root = root
        title = "全自动防丢失备份工厂"
        if _IS_MAC:
            title = "Mac " + title
        elif _IS_WIN:
            title = "Windows " + title
        self.root.title(title)
        self.root.geometry("620x680")

        self.vars = {
            'files_src': tk.StringVar(), 'files_tgt': tk.StringVar(),
            'images_src': tk.StringVar(), 'images_tgt': tk.StringVar(),
            'morning_time': tk.StringVar(value="10:10"),
            'evening_time': tk.StringVar(value="19:00"),
            'newer_overwrite_mode': tk.StringVar(value="ask"),
        }

        self.load_history()
        self.build_ui()

    def build_ui(self):
        lf1 = tk.LabelFrame(self.root, text=" 📂 文档备份路径 (Files) ", padx=10, pady=10)
        lf1.pack(fill="x", padx=15, pady=5)
        self.create_path_row(lf1, "原始地址 (源):", 'files_src')
        self.create_path_row(lf1, "目标地址 (备):", 'files_tgt')

        lf2 = tk.LabelFrame(self.root, text=" 🖼️ 图片备份路径 (Images) ", padx=10, pady=10)
        lf2.pack(fill="x", padx=15, pady=5)
        self.create_path_row(lf2, "原始地址 (源):", 'images_src')
        self.create_path_row(lf2, "目标地址 (备):", 'images_tgt')

        lf3 = tk.LabelFrame(self.root, text=" ⏰ 自动化双向规则 ", padx=10, pady=10)
        lf3.pack(fill="x", padx=15, pady=5)
        tk.Label(lf3, text="🌞 早晨 (抽出保存):").grid(row=0, column=0, sticky="e")
        tk.Entry(lf3, textvariable=self.vars['morning_time'], width=10).grid(row=0, column=1, padx=5)
        tk.Label(lf3, text="🌙 晚间 (灌回防删):").grid(row=1, column=0, sticky="e", pady=5)
        tk.Entry(lf3, textvariable=self.vars['evening_time'], width=10).grid(row=1, column=1, padx=5)

        lf4 = tk.LabelFrame(self.root, text=" 📌 同名文件源较新时（覆盖目标） ", padx=10, pady=10)
        lf4.pack(fill="x", padx=15, pady=5)
        fr_mode = tk.Frame(lf4)
        fr_mode.pack(anchor="w")
        tk.Radiobutton(
            fr_mode,
            text="每次询问（弹窗内可记住：覆盖较新 / 跳过较新）",
            variable=self.vars['newer_overwrite_mode'],
            value="ask",
            command=self.save_current_config,
        ).pack(anchor="w")
        tk.Radiobutton(
            fr_mode,
            text="始终按较新的覆盖（不再弹窗）",
            variable=self.vars['newer_overwrite_mode'],
            value="always",
            command=self.save_current_config,
        ).pack(anchor="w")
        tk.Radiobutton(
            fr_mode,
            text="始终跳过较新（仅复制新增，不再弹窗）",
            variable=self.vars['newer_overwrite_mode'],
            value="skip_newer",
            command=self.save_current_config,
        ).pack(anchor="w")

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill="x", padx=15, pady=10)
        tk.Button(btn_frame, text="▶️ 立刻: 抽存",
                  command=self.force_morning, bg="lightgreen").pack(side="left", padx=5)
        tk.Button(btn_frame, text="◀️ 立刻: 灌回",
                  command=self.force_evening, bg="lightblue").pack(side="left", padx=5)
        tk.Button(btn_frame, text="🔍 诊断源目录",
                  command=self.diagnose, bg="lightyellow").pack(side="left", padx=5)

        tk.Button(self.root, text="✅ 保存 & 打入静默后台",
                  command=self.save_and_enable_daemon,
                  fg="red", font=("", 13, "bold")).pack(pady=10)

    def create_path_row(self, parent, label_text, var_name):
        frame = tk.Frame(parent)
        frame.pack(fill="x", pady=2)
        tk.Label(frame, text=label_text, width=12, anchor="e").pack(side="left")
        tk.Entry(frame, textvariable=self.vars[var_name], width=40).pack(side="left", padx=5)
        tk.Button(frame, text="选择...", command=lambda: self.select_dir(var_name)).pack(side="left")

    def select_dir(self, var_name):
        path = filedialog.askdirectory()
        if path:
            self.vars[var_name].set(path)
            self.save_current_config()

    def load_history(self):
        cfg = load_config()
        for k, v in self.vars.items():
            if k == 'newer_overwrite_mode':
                m = cfg.get(k) or 'ask'
                if m not in ('ask', 'always', 'skip_newer'):
                    m = 'ask'
                v.set(m)
                continue
            if k in cfg and cfg[k]:
                if k.endswith('_src') or k.endswith('_tgt'):
                    v.set(normalize_path_value(cfg[k]))
                else:
                    v.set(cfg[k])

    def save_current_config(self):
        cfg = {}
        for k, v in self.vars.items():
            if k == 'newer_overwrite_mode':
                m = v.get().strip() or 'ask'
                if m not in ('ask', 'always', 'skip_newer'):
                    m = 'ask'
                cfg[k] = m
                continue
            value = v.get().strip()
            if k.endswith('_src') or k.endswith('_tgt'):
                value = normalize_path_value(value)
                self.vars[k].set(value)
            cfg[k] = value
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
        return cfg

    def check_source_permissions(self):
        """检测源目录是否可读；macOS 下额外处理 Containers 静默拦截。"""
        paths_to_check = [
            ('files_src', '文档源'),
            ('images_src', '图片源'),
        ]
        for key, label in paths_to_check:
            path = normalize_path_value(self.vars[key].get().strip())
            if not path:
                continue

            p = Path(path)
            if not p.exists():
                continue

            try:
                items = list(p.iterdir())
            except PermissionError:
                items = None  # 明确报错

            # ===== macOS：Containers 目录 + 0 个文件 = 被静默拦截 =====
            if _IS_MAC and items is not None and len(items) == 0 and '/Containers/' in path:
                msg = (
                    f"⚠️ 检测到【{label}】被 macOS 静默拦截！\n\n"
                    f"路径在应用沙盒容器内：\n{path}\n\n"
                    f"macOS 不报错但返回空列表（0个文件）。\n\n"
                    f"解决方法：\n"
                    f"1. 点击【去设置】打开系统权限页面\n"
                    f"2. 把 Terminal.app 和 Python 都加入「完全磁盘访问权限」\n"
                    f"3. 重启终端和本软件"
                )
                if messagebox.askokcancel("沙盒权限拦截", msg):
                    subprocess.run(["open",
                                    "x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles"])
                return False

            if items is None:
                if _IS_MAC:
                    msg = ("⚠️ 权限被明确拒绝！\n\n请到系统设置授予「完全磁盘访问权限」。")
                    if messagebox.askokcancel("权限不足", msg):
                        subprocess.run(["open",
                                        "x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles"])
                else:
                    msg = (
                        f"⚠️ 无法读取【{label}】：\n{path}\n\n"
                        f"请检查该文件夹权限，或以管理员身份运行本程序。"
                    )
                    messagebox.showerror("权限不足", msg)
                return False

        return True

    # ===== 新增：诊断按钮 =====
    def diagnose(self):
        """一键诊断：打印 Python 在源目录里到底看到了什么"""
        self.save_current_config()
        cfg = load_config()

        report_lines = []
        for label, key in [("文档源", 'files_src'), ("图片源", 'images_src')]:
            path = normalize_path_value(cfg.get(key, ''))
            report_lines.append(f"{'='*50}")
            report_lines.append(f"🔍 {label}: {path}")

            if not path:
                report_lines.append("   ❌ 路径为空！请先设置")
                continue

            p = Path(path)
            if not p.exists():
                report_lines.append(f"   ❌ 路径不存在！")
                continue

            try:
                items = sorted(p.iterdir(), key=lambda x: x.name)
                local_files = []
                icloud_files = []
                dirs = []
                hidden = []

                for item in items:
                    if item.is_dir():
                        dirs.append(item.name)
                    elif item.name.startswith('.') and item.name.endswith('.icloud'):
                        real_name = item.name[1:-7]
                        icloud_files.append(real_name)
                    elif item.name.startswith('.'):
                        hidden.append(item.name)
                    else:
                        local_files.append(item.name)

                report_lines.append(
                    "   ℹ️ 说明：「本地文件」仅统计顶层、且名称不以 . 开头的文件；"
                    "子目录内的文件见下方「递归文件总数」。"
                )
                report_lines.append(f"   📄 顶层本地文件 ({len(local_files)}):")
                for f in local_files:
                    report_lines.append(f"      ✅ {f}")

                report_lines.append(f"   👁️ 顶层隐藏项 ({len(hidden)})（名称以 . 开头，如 .git、.env）:")
                for f in hidden:
                    report_lines.append(f"      · {f}")

                report_lines.append(f"   ☁️  iCloud云端 ({len(icloud_files)}):")
                for f in icloud_files:
                    report_lines.append(f"      ☁️ {f}")

                report_lines.append(f"   📁 顶层子文件夹 ({len(dirs)}):")
                for d in dirs:
                    report_lines.append(f"      📁 {d}")

                total_under = _count_files_under(p)
                if total_under is not None:
                    report_lines.append(f"   📊 递归文件总数（含所有子文件夹）: {total_under}")
                    if not local_files and total_under > 0:
                        report_lines.append(
                            "   💡 顶层没有「非隐藏文件」属正常：文件都在子文件夹里，抽存仍会递归备份。"
                        )
                else:
                    report_lines.append("   📊 递归统计：无法遍历子目录（权限或拒绝访问）。")

                if icloud_files:
                    report_lines.append(f"\n   ⚠️ 有 {len(icloud_files)} 个文件在iCloud云端！")
                    report_lines.append(f"   备份时会自动触发下载，请确保网络畅通。")

            except PermissionError:
                report_lines.append("   ⛔ 权限被拦截！")
            except Exception as e:
                report_lines.append(f"   ⛔ 错误: {e}")

        report = "\n".join(report_lines)
        print(report)

        # 同时弹窗显示
        diag_win = tk.Toplevel(self.root)
        diag_win.title("源目录诊断报告")
        diag_win.geometry("600x450")
        text = tk.Text(diag_win, wrap="word", font=("Menlo", 11))
        text.pack(fill="both", expand=True, padx=10, pady=10)
        text.insert("1.0", report)
        text.config(state="disabled")

    def force_morning(self):
        if not self.check_source_permissions():
            return
        self.save_current_config()
        session = {'decision': None}

        def on_saved(mode):
            self.vars['newer_overwrite_mode'].set(mode)

        run_morning_task(
            gui_master=self.root,
            gui_interactive=True,
            on_policy_saved=on_saved,
            overwrite_session=session,
        )
        messagebox.showinfo("完成", _sync_done_hint(morning=True))

    def force_evening(self):
        if not self.check_source_permissions():
            return
        self.save_current_config()
        session = {'decision': None}

        def on_saved(mode):
            self.vars['newer_overwrite_mode'].set(mode)

        run_evening_task(
            gui_master=self.root,
            gui_interactive=True,
            on_policy_saved=on_saved,
            overwrite_session=session,
        )
        messagebox.showinfo("完成", _sync_done_hint(morning=False))

    def save_and_enable_daemon(self):
        if not self.check_source_permissions():
            return
        self.save_current_config()

        try:
            if _IS_WIN:
                method, extra = _install_windows_startup_task()
                log_dir = os.path.join(
                    os.environ.get("TEMP", os.path.expanduser("~")), "knock_sync_logs"
                )
                if method == "schtasks":
                    msg = (
                        f"已创建计划任务（{WIN_TASK_NAME}），登录后自动运行后台。\n"
                        f"请注销并重新登录后生效。\n\n后台日志目录：{log_dir}"
                    )
                else:
                    msg = (
                        "当前系统不允许创建计划任务（常见于组策略/权限限制），"
                        "已改为「启动」文件夹方式，效果同为登录后自动运行后台。\n\n"
                        f"启动脚本：\n{extra}\n\n"
                        "请注销并重新登录后生效；若不需要可删除上述 .cmd 文件。\n\n"
                        f"后台日志目录：{log_dir}"
                    )
                messagebox.showinfo("部署成功！", msg)
            elif _IS_MAC:
                current_script = os.path.abspath(sys.argv[0])
                if getattr(sys, "frozen", False):
                    exe_path = os.path.realpath(sys.executable)
                    working_dir = os.path.dirname(exe_path) or os.path.expanduser("~")
                    arg_lines = (
                        f"        <string>{_plist_xml_string(exe_path)}</string>\n"
                        f"        <string>--daemon</string>"
                    )
                else:
                    python_exe = sys.executable or "/usr/bin/python3"
                    working_dir = os.path.dirname(current_script) or os.path.expanduser("~")
                    arg_lines = (
                        f"        <string>{_plist_xml_string(python_exe)}</string>\n"
                        f"        <string>{_plist_xml_string(current_script)}</string>\n"
                        f"        <string>--daemon</string>"
                    )
                plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" \
"http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>{PLIST_LABEL}</string>
    <key>ProgramArguments</key>
    <array>
{arg_lines}
    </array>
    <key>WorkingDirectory</key>
    <string>{_plist_xml_string(working_dir)}</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>/tmp/knock_sync_out.log</string>
    <key>StandardErrorPath</key>
    <string>/tmp/knock_sync_err.log</string>
</dict>
</plist>"""
                subprocess.run(["launchctl", "unload", PLIST_PATH], capture_output=True)
                with open(PLIST_PATH, 'w', encoding='utf-8') as f:
                    f.write(plist_content)
                subprocess.run(["launchctl", "load", PLIST_PATH], capture_output=True)
                messagebox.showinfo("部署成功！", "开机自启已就绪！可安全关闭此窗口。")
            else:
                messagebox.showinfo(
                    "提示",
                    "当前系统仅支持在 macOS / Windows 上一键写入自启。\n"
                    "其他系统请自行用计划任务/cron 运行：python backup.py --daemon",
                )
        except Exception as e:
            messagebox.showerror("错误", f"配置后台失败: {str(e)}")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--daemon":
        start_daemon_loop()
    else:
        root = tk.Tk()
        app = AppGUI(root)
        root.mainloop()

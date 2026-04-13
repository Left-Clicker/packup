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


def _install_windows_startup_task():
    sub_kw = {}
    if _IS_WIN and hasattr(subprocess, "CREATE_NO_WINDOW"):
        sub_kw["creationflags"] = subprocess.CREATE_NO_WINDOW
    subprocess.run(
        ["schtasks", "/delete", "/tn", WIN_TASK_NAME, "/f"],
        capture_output=True,
        **sub_kw,
    )
    tr = _windows_task_command_line()
    r = subprocess.run(
        [
            "schtasks",
            "/create",
            "/tn",
            WIN_TASK_NAME,
            "/tr",
            tr,
            "/sc",
            "onlogon",
            "/f",
        ],
        capture_output=True,
        text=True,
        **sub_kw,
    )
    if r.returncode != 0:
        err = (r.stderr or "") + (r.stdout or "")
        raise RuntimeError(err.strip() or f"schtasks 退出码 {r.returncode}")


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


def copy_without_overwrite(src_dir, tgt_dir, log_prefix=""):
    """
    用 pathlib.rglob 深度遍历（与你旧脚本一致的方式）
    + iCloud 占位符自动检测下载
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

            # 增量逻辑：目标已有则跳过
            if dst_file.exists():
                skipped += 1
                continue

            # 创建目标子目录 & 拷贝
            dst_file.parent.mkdir(parents=True, exist_ok=True)
            try:
                shutil.copy2(str(src_file), str(dst_file))
                count += 1
                print(f"  + 搬运成功: {relative}")
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

    summary = (f"新增 {count}, 跳过 {skipped}"
               + (f", 云端下载 {cloud_downloaded}" if cloud_downloaded else "")
               + (f", 云端超时 {cloud_failed}" if cloud_failed else ""))
    print(f"[{datetime.now()}] {log_prefix} ✅ 完成: {summary}\n")


def run_morning_task():
    cfg = load_config()
    print("\n" + "=" * 60)
    print("=== 执行早晨同步规则 (抽出保存) ===")
    print("=" * 60)
    copy_without_overwrite(cfg.get('files_src'), cfg.get('files_tgt'), "[文档]")
    copy_without_overwrite(cfg.get('images_src'), cfg.get('images_tgt'), "[图片]")


def run_evening_task():
    cfg = load_config()
    print("\n" + "=" * 60)
    print("=== 执行晚间同步规则 (灌回防删) ===")
    print("=" * 60)
    copy_without_overwrite(cfg.get('files_tgt'), cfg.get('files_src'), "[文档]")
    copy_without_overwrite(cfg.get('images_tgt'), cfg.get('images_src'), "[图片]")


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
        self.root.geometry("620x580")

        self.vars = {
            'files_src': tk.StringVar(), 'files_tgt': tk.StringVar(),
            'images_src': tk.StringVar(), 'images_tgt': tk.StringVar(),
            'morning_time': tk.StringVar(value="10:10"),
            'evening_time': tk.StringVar(value="19:00"),
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
            if k in cfg and cfg[k]:
                if k.endswith('_src') or k.endswith('_tgt'):
                    v.set(normalize_path_value(cfg[k]))
                else:
                    v.set(cfg[k])

    def save_current_config(self):
        cfg = {}
        for k, v in self.vars.items():
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

                report_lines.append(f"   📄 本地文件 ({len(local_files)}):")
                for f in local_files:
                    report_lines.append(f"      ✅ {f}")

                report_lines.append(f"   ☁️  iCloud云端 ({len(icloud_files)}):")
                for f in icloud_files:
                    report_lines.append(f"      ☁️ {f}")

                report_lines.append(f"   📁 子文件夹 ({len(dirs)}):")
                for d in dirs:
                    report_lines.append(f"      📁 {d}")

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
        run_morning_task()
        messagebox.showinfo("完成", _sync_done_hint(morning=True))

    def force_evening(self):
        if not self.check_source_permissions():
            return
        self.save_current_config()
        run_evening_task()
        messagebox.showinfo("完成", _sync_done_hint(morning=False))

    def save_and_enable_daemon(self):
        if not self.check_source_permissions():
            return
        self.save_current_config()

        try:
            if _IS_WIN:
                _install_windows_startup_task()
                log_dir = os.path.join(
                    os.environ.get("TEMP", os.path.expanduser("~")), "knock_sync_logs"
                )
                messagebox.showinfo(
                    "部署成功！",
                    f"已创建登录时启动的后台任务（{WIN_TASK_NAME}）。\n"
                    f"注销并重新登录后生效。\n后台日志目录：{log_dir}",
                )
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

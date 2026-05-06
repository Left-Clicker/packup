import sys
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageChops, ImageFilter
import os
import platform

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError:
    messagebox.showerror("缺少组件", "请在终端执行: pip install tkinterdnd2")
    exit()

TEMPLATE_CONFIGS = {
    "template_E6.png": {
        "style": "E6_Classic",
        "photo": {"size": 405, "x": 92, "y": 122},
        # 主景横光辉（⑨）与昵称坐标解耦；默认与旧版「昵称默认位置」视觉对齐
        "glow": {"x": 0, "y": 514},
        "fields": {
            "name": {"label": "✎ 玩家名称 (Name):", "size": 28, "x": 0, "y": 502,
                     "color": "#FAD355", "stroke": "#111111"}
        }
    },
    "template_E7.png": {
        "style": "E7_Italic",
        "photo": {"size": 240, "x": 417, "y": 90},
        "fields": {
            "name": {"label": "✎ 玩家昵称 (Name):", "size": 26, "x": -6, "y": 311,
                     "color": "#FFFFFF", "stroke": "#111111"},
            "city": {"label": "⌂ 玩家城市 (City):", "size": 23, "x": -6, "y": 351,
                     "color": "#FAD355", "stroke": "#111111"}
        }
    }
}

C0 = "#0D0D0D"
C1 = "#161616"
C2 = "#222222"
C3 = "#333333"
CG = "#D4AF37"
CT = "#E6C27A"
CW = "#E0E0E0"


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("宣传图制作工作台 V2.2")
        self.root.geometry("1400x850")
        self.root.configure(bg=C0)
        self.app_dir = self._ap()
        self.queue = []
        self.cur = None
        self.final = None
        self.fc = {}
        self.sfonts = self._sf()
        self.fld = {}
        self._rp = None
        self._ri = None
        self._tn = None
        self.ti = None
        self.pi = None
        self.sc = 1.0
        self.vx = self.vy = 0
        self.dsx = self.dsy = self.dbx = self.dby = 0
        self.psx = self.psy = self.bvx = self.bvy = 0
        self.tki = None
        self.cid = None
        self._f1 = True
        self._rj = None
        self.drag_mode = None
        self.handle_size = 14
        self.resize_handle = None

        st = ttk.Style()
        st.theme_use('clam')
        st.configure("TCombobox", fieldbackground=C3, background=C2,
                     foreground=CG, arrowcolor=CG, bordercolor=C2)
        st.map("TCombobox", fieldbackground=[("readonly", C3)],
               selectbackground=[("readonly", CG)], selectforeground=[("readonly", "black")],
               foreground=[("readonly", CG)])
        for k, v in [('background', C3), ('foreground', CG),
                     ('selectBackground', CG), ('selectForeground', 'black')]:
            self.root.option_add(f'*TCombobox*Listbox.{k}', v)

        self._ui()
        self._st()
        self._dnd()

    @staticmethod
    def _ap():
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            internal_dir = os.path.join(exe_dir, "_internal")
            if os.path.isdir(internal_dir):
                return internal_dir
            return exe_dir
        return os.path.dirname(os.path.abspath(__file__))

    def _ui(self):
        L = tk.Frame(self.root, bg=C1, width=420)
        L.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        L.pack_propagate(False)
        tk.Label(L, text="⚜️ 核心工作台", fg=CG, bg=C1,
                 font=("Microsoft YaHei UI", 18, "bold")).pack(pady=(15, 10))

        b1 = tk.LabelFrame(L, text=" 模板与录入 ", bg=C2, fg=CG,
                           font=("Microsoft YaHei UI", 11, "bold"), bd=1)
        b1.pack(fill=tk.X, padx=10, pady=5)
        self.cmb = ttk.Combobox(b1, state="readonly", font=("Arial", 12, "bold"))
        self.cmb.pack(fill=tk.X, padx=15, pady=(15, 10))
        self.cmb.bind("<<ComboboxSelected>>", self._tc)
        tk.Button(b1, text="[ 点击多选 / 拖入图片 ]\n建立批量排队", command=self._sel,
                  bg="#2B2B2B", fg=CT, activebackground=CG, activeforeground="black",
                  font=("Microsoft YaHei UI", 10), height=2, relief="groove"
                  ).pack(fill=tk.X, padx=15, pady=(5, 10))
        self.lq = tk.Label(b1, text="[队列空闲]", fg="#888", bg=C2,
                           font=("Microsoft YaHei UI", 9))
        self.lq.pack(pady=(0, 10))

        b2 = tk.LabelFrame(L, text=" 空间参数 ", bg=C2, fg=CG,
                           font=("Microsoft YaHei UI", 10), bd=1)
        b2.pack(fill=tk.X, padx=10, pady=5)
        fp = tk.Frame(b2, bg=C2)
        fp.pack(pady=10)

        def mks(p, t, lo, hi, c):
            tk.Label(p, text=t, bg=C2, fg=CW,
                     font=("Microsoft YaHei", 9)).grid(row=0, column=c * 2, padx=(5, 0))
            s = tk.Spinbox(p, from_=lo, to=hi, width=5, bg=C3, fg=CG,
                           insertbackground=CG, buttonbackground=C2)
            s.grid(row=0, column=c * 2 + 1, padx=2)
            return s

        self.ssz = mks(fp, "宽:", 10, 8000, 0)
        self.sx = mks(fp, "X:", -4000, 4000, 1)
        self.sy = mks(fp, "Y:", -4000, 4000, 2)
        self.ssz.config(command=self._full)
        self.ssz.bind("<Return>", lambda e: self._full())
        self.sx.config(command=self._fast)
        self.sx.bind("<Return>", lambda e: self._fast())
        self.sy.config(command=self._fast)
        self.sy.bind("<Return>", lambda e: self._fast())

        b3 = tk.LabelFrame(L, text=" 🌟 光效 ", bg=C2, fg=CG,
                           font=("Microsoft YaHei UI", 10), bd=1)
        b3.pack(fill=tk.X, padx=10, pady=5)

        def mksc(par, txt, lo, hi, val):
            f = tk.Frame(par, bg=C2)
            f.pack(fill=tk.X, padx=10, pady=4)
            tk.Label(f, text=txt, bg=C2, fg=CW,
                     font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
            s = tk.Scale(f, from_=lo, to=hi, orient=tk.HORIZONTAL, bg=C2, fg=CG,
                         bd=0, highlightthickness=0, troughcolor=C3,
                         activebackground=CG, command=lambda v: self._fast())
            s.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
            s.set(val)
            return s

        self.sli = mksc(b3, "光线强度:", 0, 100, 80)
        self.sfl = mksc(b3, "横光偏移:", -200, 200, 23)

        # E6 专用：主景横光辉位置（与昵称「左右/上下」独立）
        self.fglow = tk.Frame(b3, bg=C2)

        def mkg(p, t, lo, hi):
            tk.Label(p, text=t, bg=C2, fg=CW,
                     font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
            s = tk.Spinbox(p, from_=lo, to=hi, width=6, bg=C3, fg=CG,
                           insertbackground=CG, buttonbackground=C2,
                           command=self._fast)
            s.pack(side=tk.LEFT, padx=8)
            s.bind("<Return>", lambda e: self._fast())
            return s

        gf = tk.Frame(self.fglow, bg=C2)
        gf.pack(fill=tk.X, padx=10, pady=4)
        self.sgx = mkg(gf, "横光左右:", -2000, 2000)
        self.sgy = mkg(gf, "横光上下:", -500, 2500)

        self.btxt = tk.LabelFrame(L, text=" 文字编辑 ", bg=C2, fg=CG,
                                  font=("Microsoft YaHei UI", 10), bd=1)
        self.btxt.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        bf = tk.Frame(L, bg=C1)
        bf.pack(fill=tk.X, padx=10, pady=15, side=tk.BOTTOM)
        tk.Button(bf, text="⟳ 刷新", command=self._full, bg="#222", fg=CW,
                  relief="groove").pack(fill=tk.X, pady=(0, 10))
        self.bsv = tk.Button(bf, text="✦ 渲染输出 ✦", command=self._save, bg=CG,
                             fg="black", activebackground="#FFE47A", activeforeground="black",
                             font=("Microsoft YaHei UI", 13, "bold"), height=2,
                             state="disabled", cursor="hand2")
        self.bsv.pack(fill=tk.X)

        R = tk.LabelFrame(self.root,
                          text=" 🖥️ [滚轮缩放视图] [Ctrl+滚轮细调大小] [拖边角缩放] [左键拖位置] [右键平移] ",
                          bg=C0, fg=CG, font=("Microsoft YaHei UI", 11, "bold"), bd=1)
        R.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 10), pady=10)
        self.cv = tk.Canvas(R, bg="#080808", highlightthickness=0, cursor="tcross")
        self.cv.pack(fill=tk.BOTH, expand=True)
        self.cv.create_text(400, 300, text="等待图像…", fill="#444",
                            font=("Microsoft YaHei", 16))
        for ev, fn in [("<MouseWheel>", self._zm), ("<Button-4>", self._zm),
                       ("<Button-5>", self._zm), ("<Control-MouseWheel>", self._czm),
                       ("<Control-Button-4>", self._czm), ("<Control-Button-5>", self._czm),
                       ("<ButtonPress-1>", self._ds), ("<B1-Motion>", self._dm),
                       ("<ButtonRelease-1>", self._de),
                       ("<ButtonPress-3>", self._ps), ("<B3-Motion>", self._pm),
                       ("<ButtonPress-2>", self._ps), ("<B2-Motion>", self._pm),
                       ("<Control-ButtonPress-1>", self._ps), ("<Control-B1-Motion>", self._pm)]:
            self.cv.bind(ev, fn)

    def _dnd(self):
        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>',
                               lambda e: self._aq(self.root.tk.splitlist(e.data)))
        except:
            pass

    def _sel(self):
        p = filedialog.askopenfilenames(
            title="选择照片", filetypes=[("Images", "*.png *.jpg *.jpeg *.webp")])
        if p:
            self._aq(p)

    def _aq(self, paths):
        v = [p for p in paths
             if str(p).lower().endswith(('.png', '.jpg', '.jpeg', '.webp'))]
        if not v:
            return
        self.queue.extend(v)
        if not self.cur:
            self._nx()
        else:
            self._uq()

    def _nx(self):
        if not self.queue:
            self.cur = None
            self.lq.config(text="✅ 完成", fg="#2ECC71")
            self.bsv.config(state="disabled", text="✦ 输出 ✦")
            return
        self.cur = self.queue.pop(0)
        self._uq()
        self._full()

    def _uq(self):
        nm = os.path.basename(self.cur)
        r = len(self.queue)
        if r > 0:
            self.lq.config(text=f"▶ {nm} | ⏳{r}张", fg=CT)
            self.bsv.config(text=f"✦ 保存并下一张({r}) ✦")
        else:
            self.lq.config(text=f"▶ {nm} | 🏁最后", fg=CG)
            self.bsv.config(text="✦ 保存终图 ✦")

    def _save(self):
        if not self.final or not self.cur:
            return
        try:
            d = os.path.dirname(self.cur)
            n = os.path.splitext(os.path.basename(self.cur))[0]
            tpl = self.cmb.get().lower()
            pfx = "E7出图" if "e7" in tpl else "E6出图"
            self.final.save(
                os.path.join(d, f"{pfx}-{n}.png"),
                format="PNG",
                optimize=False,
                compress_level=0
            )
            self._nx()
            if not self.queue and not self.cur:
                messagebox.showinfo("收工", "🎉 全部完成！")
        except Exception as e:
            messagebox.showerror("异常", str(e))

    def _st(self):
        vl = []
        try:
            for f in os.listdir(self.app_dir):
                lo = f.lower()
                if "template_e6" in lo and lo.endswith(('.png', '.jpg')):
                    vl.append(f)
                    TEMPLATE_CONFIGS[f] = TEMPLATE_CONFIGS["template_E6.png"]
                elif "template_e7" in lo and lo.endswith(('.png', '.jpg')):
                    vl.append(f)
                    TEMPLATE_CONFIGS[f] = TEMPLATE_CONFIGS["template_E7.png"]
        except:
            pass
        if vl:
            self.cmb['values'] = vl
            self.cmb.current(0)
            self._tc()
        else:
            messagebox.showwarning("警告", "没找到模板！")

    def _tc(self, ev=None):
        cfg = TEMPLATE_CONFIGS.get(self.cmb.get())
        if not cfg:
            return
        for s, k in [(self.ssz, "size"), (self.sx, "x"), (self.sy, "y")]:
            s.delete(0, 'end')
            s.insert(0, cfg["photo"][k])
        gv = cfg.get("glow", {"x": 0, "y": 514})
        self.sgx.delete(0, 'end')
        self.sgx.insert(0, str(gv.get("x", 0)))
        self.sgy.delete(0, 'end')
        self.sgy.insert(0, str(gv.get("y", 514)))
        if cfg.get("style") == "E6_Classic":
            self.fglow.pack(fill=tk.X, padx=0, pady=0)
        else:
            self.fglow.pack_forget()
        self.vx = self.vy = 0
        for w in self.btxt.winfo_children():
            w.destroy()
        self.fld.clear()
        for fk, fv in cfg["fields"].items():
            g = tk.Frame(self.btxt, bg=C2)
            g.pack(fill=tk.X, padx=10, pady=(8, 0))
            tk.Label(g, text=fv["label"], fg=CT, bg=C2,
                     font=("Microsoft YaHei UI", 9, "bold")).pack(anchor="w")
            ent = tk.Entry(g, font=("Trebuchet MS", 12), justify="center",
                           bg=C3, fg=CG, insertbackground=CG, bd=0)
            ent.pack(fill=tk.X, ipady=4, pady=2)
            ent.bind("<KeyRelease>", lambda e: self._fast())
            fc = tk.Frame(g, bg=C2)
            fc.pack(fill=tk.X, pady=2)

            def mk(p, t, a, b):
                tk.Label(p, text=t, bg=C2, fg=CW,
                         font=("Microsoft YaHei", 8)).pack(side=tk.LEFT, padx=(5, 0))
                s = tk.Spinbox(p, from_=a, to=b, width=4, bg=C3, fg=CG,
                               buttonbackground=C2, command=self._fast)
                s.pack(side=tk.LEFT, padx=2)
                s.bind("<Return>", lambda e: self._fast())
                return s

            ss = mk(fc, "字号:", 10, 400)
            ss.delete(0, 'end')
            ss.insert(0, fv["size"])
            sxx = mk(fc, "左右:", -1000, 1000)
            sxx.delete(0, 'end')
            sxx.insert(0, fv["x"])
            syy = mk(fc, "上下:", -500, 2000)
            syy.delete(0, 'end')
            syy.insert(0, fv["y"])
            self.fld[fk] = {"e": ent, "ss": ss, "sx": sxx, "sy": syy,
                            "c": fv["color"], "s": fv["stroke"]}
        if self.cur:
            self._full()

            # ═══ 画布交互 ═══

    def _zm(self, e):
        if not self.final or e.state & 0x0004:
            return
        self.sc *= 1.15 if (e.delta > 0 or e.num == 4) else 1 / 1.15
        self._disp()

    def _czm(self, e):
        if not self.final:
            return
        try:
            c = int(self.ssz.get())
            fine = max(1, min(8, int(round(c * 0.01))))
            coarse = max(5, min(30, int(round(c * 0.03))))
            step = coarse if e.state & 0x0001 else fine
            d = getattr(e, 'delta', 0)
            if e.num == 4 or d > 0:
                n = c + step
            elif e.num == 5 or d < 0:
                n = c - step
            else:
                return
            self.ssz.delete(0, 'end')
            self.ssz.insert(0, str(max(10, min(n, 8000))))
            self._full()
        except:
            pass

    def _ds(self, e):
        if not self.final:
            return
        self.drag_mode = "move"
        self.dsx = e.x
        self.dsy = e.y
        try:
            self.dbx = int(self.sx.get())
            self.dby = int(self.sy.get())
            self.resize_handle = self._hit_resize_handle(e.x, e.y)
            if self.resize_handle:
                self.drag_mode = "resize"
                self.dbw = int(self.ssz.get())
        except:
            pass

    def _dm(self, e):
        if not self.final:
            return
        if self.drag_mode == "resize":
            self._drag_resize(e)
            return
        self.sx.delete(0, 'end')
        self.sx.insert(0, int(self.dbx + (e.x - self.dsx) / self.sc))
        self.sy.delete(0, 'end')
        self.sy.insert(0, int(self.dby + (e.y - self.dsy) / self.sc))
        self._fast()

    def _ps(self, e):
        if not self.final:
            return
        self.psx = e.x
        self.psy = e.y
        self.bvx = self.vx
        self.bvy = self.vy

    def _de(self, e):
        self.drag_mode = None
        self.resize_handle = None

    def _pm(self, e):
        if not self.final or not self.cid:
            return
        self.vx = self.bvx + (e.x - self.psx)
        self.vy = self.bvy + (e.y - self.psy)
        self._disp()

    def _photo_bounds(self):
        if not self.final or not self.pi:
            return None
        w, h = self.final.size
        nw = max(10, int(w * self.sc))
        nh = max(10, int(h * self.sc))
        cx = self.cv.winfo_width() / 2 + self.vx
        cy = self.cv.winfo_height() / 2 + self.vy
        bx = cx - nw / 2 + int(self.sx.get()) * self.sc
        by = cy - nh / 2 + int(self.sy.get()) * self.sc
        bw = self.pi.size[0] * self.sc
        bh = self.pi.size[1] * self.sc
        return bx, by, bw, bh

    def _hit_resize_handle(self, x, y):
        bounds = self._photo_bounds()
        if not bounds:
            return None
        bx, by, bw, bh = bounds
        hs = self.handle_size
        mx = bx + bw / 2
        my = by + bh / 2
        hit_boxes = {
            "nw": (bx - hs, by - hs, bx + hs, by + hs),
            "n": (mx - hs, by - hs, mx + hs, by + hs),
            "ne": (bx + bw - hs, by - hs, bx + bw + hs, by + hs),
            "e": (bx + bw - hs, my - hs, bx + bw + hs, my + hs),
            "se": (bx + bw - hs, by + bh - hs, bx + bw + hs, by + bh + hs),
            "s": (mx - hs, by + bh - hs, mx + hs, by + bh + hs),
            "sw": (bx - hs, by + bh - hs, bx + hs, by + bh + hs),
            "w": (bx - hs, my - hs, bx + hs, my + hs),
        }
        for name, (x1, y1, x2, y2) in hit_boxes.items():
            if x1 <= x <= x2 and y1 <= y <= y2:
                return name
        return None

    def _drag_resize(self, e):
        bounds = self._photo_bounds()
        if not bounds or not self._ri:
            return
        bx, by, bw, bh = bounds
        ow, oh = self._ri.size
        if ow <= 0 or oh <= 0 or self.sc == 0 or not self.resize_handle:
            return
        ratio = oh / float(ow)
        right = bx + bw
        bottom = by + bh
        width_candidates = []
        if "e" in self.resize_handle:
            width_candidates.append((e.x - bx) / self.sc)
        if "w" in self.resize_handle:
            width_candidates.append((right - e.x) / self.sc)
        if "n" in self.resize_handle:
            width_candidates.append(((bottom - e.y) / self.sc) / ratio)
        if "s" in self.resize_handle:
            width_candidates.append(((e.y - by) / self.sc) / ratio)
        if not width_candidates:
            return
        new_width = max(10, min(8000, int(round(max(width_candidates)))))
        old_width = max(1, self.dbw)
        old_height = max(1, int(round(old_width * ratio)))
        new_height = max(1, int(round(new_width * ratio)))
        new_x = self.dbx
        new_y = self.dby
        if "w" in self.resize_handle:
            new_x = self.dbx + (old_width - new_width)
        if "n" in self.resize_handle:
            new_y = self.dby + (old_height - new_height)
        self.ssz.delete(0, 'end')
        self.ssz.insert(0, str(new_width))
        self.sx.delete(0, 'end')
        self.sx.insert(0, str(new_x))
        self.sy.delete(0, 'end')
        self.sy.insert(0, str(new_y))
        self._full()

        # ═══ 字体 ═══

    def _sf(self):
        sys_name = platform.system().lower()
        if sys_name == "windows":
            fd = "C:\\Windows\\Fonts"
            names = ["seguiemj.ttf", "seguisym.ttf", "impact.ttf", "arialbd.ttf",
                     "msyhbd.ttc", "tahomabd.ttf", "tahoma.ttf", "arial.ttf",
                     "msyh.ttc", "simsun.ttc"]
        elif sys_name == "darwin":
            fd = "/System/Library/Fonts"
            names = [
                "PingFang.ttc", "Helvetica.ttc", "HelveticaNeue.ttc",
                "Arial.ttf", "Arial Bold.ttf", "Hiragino Sans GB.ttc",
                "STHeiti Medium.ttc", "Apple Color Emoji.ttc"
            ]
        else:
            fd = "/usr/share/fonts"
            names = [
                "NotoSansCJK-Regular.ttc", "NotoSansCJK-Bold.ttc",
                "DejaVuSans.ttf", "LiberationSans-Regular.ttf"
            ]
        pri = []
        for n in names:
            full = os.path.join(fd, n)
            pri.append(full if os.path.exists(full) else n)
        try:
            extra_dirs = [fd]
            if sys_name == "darwin":
                extra_dirs.extend([
                    "/Library/Fonts",
                    os.path.expanduser("~/Library/Fonts")
                ])
            elif sys_name == "linux":
                extra_dirs.extend([
                    "/usr/local/share/fonts",
                    os.path.expanduser("~/.fonts")
                ])
            ex = set(os.path.basename(p).lower() for p in pri)
            ext = []
            for d in extra_dirs:
                if not os.path.isdir(d):
                    continue
                for f in os.listdir(d):
                    lo = f.lower()
                    if lo.endswith(('.ttf', '.ttc', '.otf')) and lo not in ex:
                        ext.append(os.path.join(d, f))
            return pri + ext
        except:
            return pri

    def _gf(self, char, size):
        ck = f"{char}_{size}"
        if ck in self.fc:
            return self.fc[ck]
        fpk = f"__fp_{ord(char)}"
        if fpk in self.fc:
            try:
                f = ImageFont.truetype(self.fc[fpk], size)
                self.fc[ck] = (f, char)
                return f, char
            except:
                del self.fc[fpk]
        for fp in self.sfonts:
            try:
                f = ImageFont.truetype(fp, size)
                if f.getlength(char) <= 0:
                    continue
                if not self._fok(fp, char):
                    continue
                self.fc[ck] = (f, char)
                self.fc[fpk] = fp
                return f, char
            except:
                continue
        return ImageFont.load_default(), char

    def _fok(self, fp, char):
        k = f"__ok_{fp}_{ord(char)}"
        if k in self.fc:
            return self.fc[k]
        ok = False
        try:
            tf = ImageFont.truetype(fp, 20)
            m1 = tf.getmask(char)
            m2 = tf.getmask('\ufffe')
            if m1.size != m2.size:
                ok = m1.getbbox() is not None
            else:
                ok = (m1.tobytes() != m2.tobytes()) and (m1.getbbox() is not None)
        except:
            pass
        self.fc[k] = ok
        return ok

        # ═══ 文字渲染引擎 ═══

    def _dtxt(self, tgt, text, tw, th, sz, ox, yp, mc, sc_, sty):

        # 极限 5 倍超采样抗锯齿 (SSAA)
        SSA = 5
        sz_s = sz * SSA

        cd = []
        totw_s = 0
        for ch in text:
            if ch.isspace():
                w_s = sz_s * 0.4
                cd.append((" ", None, w_s))
                totw_s += w_s
                continue
            fn, fc = self._gf(ch, sz_s)
            try:
                w_s = fn.getlength(fc)
            except:
                w_s = sz_s * 0.8
            cd.append((fc, fn, w_s))
            totw_s += w_s

        pad_s = int(sz_s * 2.5)
        cw_s = int(totw_s + pad_s * 2)
        cvh_s = int(sz_s * 4.5)
        tx_s = pad_s
        ty_s = sz_s
        bw_s = max(1, int(sz_s * 0.045))

        mt = Image.new("L", (cw_s, cvh_s), 0)
        dt = ImageDraw.Draw(mt)
        sw_s = max(2, int(sz_s * 0.07))
        mo = Image.new("L", (cw_s, cvh_s), 0)
        do = ImageDraw.Draw(mo)
        gw_s = max(6, int(sz_s * 0.22))
        mg = Image.new("L", (cw_s, cvh_s), 0)
        dg_ = ImageDraw.Draw(mg)

        # 极薄厚底设置 (使用超采样尺度)
        dep_s = max(1, int(sz_s * 0.025))

        ra = 0
        for _, fn, _ in cd:
            if fn:
                ra = fn.getmetrics()[0]
                break

        cx_s = tx_s
        for c, fn, w_s in cd:
            if fn:
                dy = ra - fn.getmetrics()[0]
                dt.text((cx_s, ty_s + dy), c, font=fn, fill=255,
                        stroke_width=bw_s, stroke_fill=255)
                do.text((cx_s, ty_s + dy), c, font=fn, fill=255,
                        stroke_width=sw_s + bw_s, stroke_fill=255)
                dg_.text((cx_s, ty_s + dy), c, font=fn, fill=255,
                         stroke_width=gw_s + bw_s, stroke_fill=255)
            cx_s += w_s

        gold = mc != "#FFFFFF"
        up = Image.new("RGBA", (cw_s, cvh_s), (0, 0, 0, 0))
        is7 = sty == "E7_Italic"

        if gold and not is7:
            gr_s = max(3, int(sz_s * 0.18))
            gb = mg.filter(ImageFilter.GaussianBlur(radius=gr_s))
            gt = Image.new("RGBA", (cw_s, cvh_s), (255, 190, 60, 0))
            ga = np.array(gb, dtype=np.float32) / 255.0
            yt0 = max(0, ty_s - int(sz_s * 0.3))
            yb0 = min(cvh_s - 1, ty_s + int(sz_s * 1.6))
            vt = np.zeros(cvh_s, dtype=np.float32)
            for y in range(cvh_s):
                if y <= yt0:
                    vt[y] = 0.02
                elif y <= yb0:
                    vt[y] = 0.02 + 0.98 * ((y - yt0) / max(1, yb0 - yt0))
                else:
                    vt[y] = 1.0
            al = np.clip(ga * vt.reshape(-1, 1) * 0.8 * 255, 0, 255).astype(np.uint8)
            gt.putalpha(Image.fromarray(al))
            up.alpha_composite(gt)
        elif not gold:
            gr_s = max(3, int(sz_s * 0.18))
            gb = mg.filter(ImageFilter.GaussianBlur(radius=gr_s))
            gt = Image.new("RGBA", (cw_s, cvh_s), (180, 200, 255, 0))
            gt.putalpha(gb.point(lambda p: min(255, int(p * 0.35))))
            up.alpha_composite(gt)

            # 👇 --- 修改点 7：为 E6 和 E7 重启统一的顶光阴影系统 --- 👇
        m_full = mt.copy()
        for i in range(1, dep_s + 1):
            shifted = Image.new("L", (cw_s, cvh_s), 0)
            shifted.paste(mt, (0, i))
            m_full = ImageChops.lighter(m_full, shifted)

        olw_s = max(1, int(sz_s * 0.025))
        expanded = m_full.filter(ImageFilter.MaxFilter(olw_s * 2 + 1))

        shifted_expanded = Image.new("L", (cw_s, cvh_s), 0)
        shifted_expanded.paste(expanded, (0, olw_s + 1))

        ring = ImageChops.subtract(shifted_expanded, m_full)
        ring_soft = ring.filter(ImageFilter.GaussianBlur(2.5))

        # 为金字和白字分别配制最合适的阴影颜色和透明度
        if gold:
            ring_soft = ring_soft.point(lambda p: int(p * 0.65))  # 金字用65%透明度
            shadow_color = (20, 10, 0, 255)  # 暖暗褐色
        else:
            ring_soft = ring_soft.point(lambda p: int(p * 0.80))  # 白字需要更高透明度(80%)分离背景
            shadow_color = (15, 15, 20, 255)  # 冷暗灰偏青色

        ol_layer = Image.new("RGBA", (cw_s, cvh_s), (0, 0, 0, 0))
        ol_layer.paste(Image.new("RGBA", (cw_s, cvh_s), shadow_color), mask=ring_soft)
        up.alpha_composite(ol_layer)
        # 👆 ----------------------------------------------------------- 👆

        # ③ 斜面厚度特效
        bev_dep_s = max(2, dep_s // 2) if is7 else dep_s
        for i in range(bev_dep_s, 0, -1):
            t = i / max(1, bev_dep_s)
            if gold:
                r = int(230 * t + 185 * (1 - t))
                g = int(180 * t + 145 * (1 - t))
                b = int(45 * t + 15 * (1 - t))
            else:
                v = int(90 * t + 55 * (1 - t))
                r, g, b = v, v, int(v * 1.1)
            up.paste(Image.new("RGBA", (cw_s, cvh_s), (r, g, b, 255)),
                     (0, i), mask=mt)

            # ④ 渐变补色
        if gold:
            stops = [
                (0.00, (255, 255, 245)),
                (0.25, (255, 235, 120)),
                (0.60, (220, 160, 20)),
                (1.00, (110, 60, 0)),
            ]
        else:
            stops = [
                (0.00, (255, 255, 255)), (0.25, (238, 240, 248)),
                (0.45, (175, 180, 200)), (0.55, (148, 152, 172)),
                (0.75, (215, 220, 235)), (1.00, (248, 250, 255)),
            ]

        grd = Image.new("RGBA", (cw_s, cvh_s), (0, 0, 0, 0))
        gd = ImageDraw.Draw(grd)

        bbox = mt.getbbox()
        if bbox:
            yt_s = bbox[1]
            yb_s = bbox[3]
        else:
            yt_s = ty_s - int(sz_s * 0.08)
            yb_s = ty_s + int(sz_s * 1.08)

        sp_s = max(1, yb_s - yt_s)
        for y in range(yt_s, yb_s + 1):
            t = max(0.0, min(1.0, (y - yt_s) / sp_s))
            lo, hi = stops[0], stops[-1]
            for j in range(len(stops) - 1):
                if stops[j][0] <= t <= stops[j + 1][0]:
                    lo, hi = stops[j], stops[j + 1]
                    break
            f = (t - lo[0]) / max(0.0001, hi[0] - lo[0])
            f = max(0.0, min(1.0, f))
            f = f * f * (3 - 2 * f)
            rgb = tuple(int(lo[1][c] + (hi[1][c] - lo[1][c]) * f) for c in range(3))
            gd.line([(0, y), (cw_s, y)], fill=(*rgb, 255))

        fc2 = Image.new("RGBA", (cw_s, cvh_s), (0, 0, 0, 0))
        fc2.paste(grd, mask=mt)
        up.alpha_composite(fc2)

        # ⑤ 高光点缀
        si = Image.new("RGBA", (cw_s, cvh_s), (0, 0, 0, 0))
        sd = ImageDraw.Draw(si)
        if gold:
            cy2_s = ty_s + int(sz_s * 0.18)
            cr_s = int(sz_s * 0.22)
            for y in range(cy2_s - cr_s, cy2_s + cr_s):
                d = abs(y - cy2_s) / max(1, cr_s)
                a = int(140 * max(0, 1 - d ** 1.4))
                if a > 0:
                    sd.line([(0, y), (cw_s, y)], fill=(255, 255, 240, a))
        else:
            cy2_s = ty_s + int(sz_s * 0.2)
            cr_s = int(sz_s * 0.18)
            for y in range(cy2_s - cr_s, cy2_s + cr_s):
                d = abs(y - cy2_s) / max(1, cr_s)
                a = int(120 * max(0, 1 - d ** 1.5))
                if a > 0:
                    sd.line([(0, y), (cw_s, y)], fill=(255, 255, 255, a))
        sf = Image.new("RGBA", (cw_s, cvh_s), (0, 0, 0, 0))
        sf.paste(si, mask=mt)
        up.alpha_composite(sf)

        # ⑥ 顶部微射反光
        if gold:
            eh = Image.new("RGBA", (cw_s, cvh_s), (0, 0, 0, 0))
            ed = ImageDraw.Draw(eh)
            rim_s = max(2, int(sz_s * 0.1))
            for y in range(yt_s, yt_s + rim_s):
                a = int(130 * (1 - (y - yt_s) / max(1, rim_s)))
                ed.line([(0, y), (cw_s, y)], fill=(255, 255, 248, a))
            ef = Image.new("RGBA", (cw_s, cvh_s), (0, 0, 0, 0))
            ef.paste(eh, mask=mt)
            up.alpha_composite(ef)

            # ⑧ 斜体变形 (完美抗锯齿变形，阴影也会一起自然横斜！)
        if is7:
            stamp_s = up.transform((cw_s, cvh_s), Image.AFFINE,
                                   (1, 0.22, -0.22 * (cvh_s / 2), 0, 1, 0),
                                   resample=Image.Resampling.BICUBIC)
        else:
            stamp_s = up

            # --- 降采样（消除所有锯齿并缩回到目标尺寸） ---
        target_cw = int(cw_s / SSA)
        target_cvh = int(cvh_s / SSA)
        stamp = stamp_s.resize((target_cw, target_cvh), Image.Resampling.LANCZOS)

        # 等比还原物理坐标系统
        totw = totw_s / SSA
        tx_ = tx_s / SSA
        ty_ = ty_s / SSA

        # ⑨ 主景横光辉（仅E6）（保持独立背景层）
        if gold and not is7:
            foff = int(self.sfl.get())
            bwf, bhf = 520, 260
            xg = np.linspace(-1, 1, bwf)
            yg = np.linspace(-1, 1, bhf)
            xx, yy = np.meshgrid(xg, yg)
            core = np.exp(-(xx ** 2 / 0.01 + yy ** 2 / 0.008))
            vs = np.maximum(0.0005, 0.014 * (1 - np.abs(xx) ** 1.6 * 0.88))
            streak = np.exp(-(xx ** 2 / 3.5)) * np.exp(-(yy ** 2 / vs))
            mid = np.exp(-(xx ** 2 / 0.25 + yy ** 2 / 0.018))
            amb = np.exp(-(xx ** 2 / 0.08 + yy ** 2 / 0.035))
            li = np.clip(core + streak * 0.65 + mid * 0.28 + amb * 0.18,
                         0, 1).astype(np.float32)
            rgba = np.zeros((bhf, bwf, 4), dtype=np.uint8)
            rgba[..., 0] = np.minimum(255, 255 * li).astype(np.uint8)
            rgba[..., 1] = np.minimum(255, 225 * li).astype(np.uint8)
            rgba[..., 2] = np.minimum(255, 100 * li).astype(np.uint8)
            rgba[..., 3] = np.minimum(255, 120 * li).astype(np.uint8)
            fo = Image.fromarray(rgba, "RGBA")
            ow2 = min(tw, int(totw * 5))
            oh2 = int(sz * 3)
            fo = fo.resize((ow2, oh2), Image.Resampling.BICUBIC)
            try:
                gx = int(self.sgx.get())
                gy = int(self.sgy.get())
            except Exception:
                gx, gy = 0, 514
            cxd = int(tw / 2 + gx) - ow2 // 2
            cyd = gy + foff - oh2 // 2
            bl = Image.new("RGBA", (tw, th), (0, 0, 0, 0))
            bl.paste(fo, (cxd, cyd))
            tgt.alpha_composite(bl)

        tgt.alpha_composite(stamp,
                            dest=(int((tw - totw) / 2 + ox - tx_), int(yp - ty_)))

        # ═══ 渲染管线 ═══

    def _full(self):
        if not self.cur:
            return
        try:
            self.root.update_idletasks()
            tn = self.cmb.get()
            rw = int(self.ssz.get())
            if self._tn != tn:
                self.ti = Image.open(os.path.join(self.app_dir, tn)).convert("RGBA")
                self._tn = tn
            if self._rp != self.cur:
                self._ri = Image.open(self.cur).convert("RGBA")
                self._rp = self.cur
            ow, oh = self._ri.size
            r = rw / float(ow) if ow > 0 else 1
            self.pi = self._ri.resize((rw, max(1, int(oh * r))),
                                      Image.Resampling.LANCZOS)
            if self._f1:
                self.sc = 0.6 if self.ti.size[0] > 800 else 1.0
                self._f1 = False
            self._do_render()
            self.bsv.config(state="normal")
        except:
            pass

    def _fast(self):
        if self._rj:
            self.root.after_cancel(self._rj)
        self._rj = self.root.after(30, self._do_render)

    def _do_render(self):
        self._rj = None
        if not self.ti or not self.pi:
            return
        try:
            px = int(self.sx.get())
            py = int(self.sy.get())
            tn = self.cmb.get()
            ts = TEMPLATE_CONFIGS.get(
                next((k for k in TEMPLATE_CONFIGS if k == tn), None), {}
            ).get("style", "E6_Classic")
            tw, th = self.ti.size

            pl = Image.new("RGBA", (tw, th), (0, 0, 0, 0))
            pl.paste(self.pi, (px, py))

            lv = int(self.sli.get())
            if lv > 0:
                ins = lv / 100.0
                rv, gv, bv, pa = pl.split()
                rgb = Image.merge("RGB", (rv, gv, bv))
                gt = Image.new("RGB", (tw, th), (255, 185, 50))
                cg = Image.blend(rgb, ImageChops.multiply(rgb, gt),
                                 0.35 * ins).convert("RGBA")

                lw2, lh2 = 160, 160
                md = (lw2 ** 2 + lh2 ** 2) ** 0.5
                ms = ((lw2 / 2) ** 2 + lh2 ** 2) ** 0.5
                yy, xx = np.mgrid[0:lh2, 0:lw2]
                xx_f = xx.astype(np.float64)
                yy_f = yy.astype(np.float64)

                dL = np.sqrt(xx_f ** 2 + yy_f ** 2)
                pL = np.clip(1 - dL / (md * 0.85), 0, 1) ** 1.6
                dR = np.sqrt((lw2 - xx_f) ** 2 + yy_f ** 2)
                pR = np.clip(1 - dR / (md * 0.85), 0, 1) ** 1.6
                tp = np.clip(pL + pR, 0, 1)
                la = np.zeros((lh2, lw2, 4), dtype=np.uint8)
                la[..., 0] = 255
                la[..., 1] = 235
                la[..., 2] = 150
                la[..., 3] = np.minimum(255, (190 * tp * ins)).astype(np.uint8)
                lb = Image.fromarray(la, "RGBA")

                ds = np.sqrt((lw2 / 2 - xx_f) ** 2 + (lh2 - yy_f) ** 2)
                sr = np.clip(1 - ds / (ms * 0.9), 0, 1)
                sa = np.zeros((lh2, lw2, 4), dtype=np.uint8)
                sa[..., 0] = 25
                sa[..., 1] = 12
                sa[..., 3] = np.minimum(255, (145 * sr ** 1.2 * ins)).astype(np.uint8)
                sb = Image.fromarray(sa, "RGBA")

                lm = lb.resize((tw, th), Image.Resampling.BICUBIC)
                sm = sb.resize((tw, th), Image.Resampling.BICUBIC)
                lit = Image.alpha_composite(Image.alpha_composite(cg, sm), lm)
                mem = Image.merge("RGBA", (*lit.split()[:3], pa))
            else:
                mem = pl

            mem = Image.alpha_composite(mem, self.ti)

            for _, ct in self.fld.items():
                tv = ct["e"].get().strip()
                if not tv:
                    continue
                self._dtxt(mem, tv, tw, th, int(ct["ss"].get()),
                           int(ct["sx"].get()), int(ct["sy"].get()),
                           ct["c"], ct["s"], ts)

            self.final = mem.copy()
            self._disp()
        except:
            pass

    def _disp(self):
        if not self.final:
            return
        self.cv.delete("all")
        w, h = self.final.size
        nw = max(10, int(w * self.sc))
        nh = max(10, int(h * self.sc))
        self.tki = ImageTk.PhotoImage(
            self.final.resize((nw, nh), Image.Resampling.LANCZOS))
        cx = self.cv.winfo_width() / 2 + self.vx
        cy = self.cv.winfo_height() / 2 + self.vy
        self.cid = self.cv.create_image(cx, cy, anchor=tk.CENTER, image=self.tki)
        try:
            if self.pi:
                bx = cx - nw / 2 + int(self.sx.get()) * self.sc
                by = cy - nh / 2 + int(self.sy.get()) * self.sc
                self.cv.create_rectangle(
                    bx, by,
                    bx + self.pi.size[0] * self.sc,
                    by + self.pi.size[1] * self.sc,
                    outline="#00FFCC", width=1, dash=(5, 3))
                hs = self.handle_size
                left = bx
                top = by
                right = bx + self.pi.size[0] * self.sc
                bottom = by + self.pi.size[1] * self.sc
                mx = (left + right) / 2
                my = (top + bottom) / 2
                for hx, hy in [
                    (left, top), (mx, top), (right, top),
                    (right, my), (right, bottom), (mx, bottom),
                    (left, bottom), (left, my)
                ]:
                    self.cv.create_rectangle(
                        hx - hs / 2, hy - hs / 2,
                        hx + hs / 2, hy + hs / 2,
                        fill="#00FFCC", outline="#003B33", width=1)
        except:
            pass
        self.cv.create_line(self.cv.winfo_width() / 2, 0,
                            self.cv.winfo_width() / 2, self.cv.winfo_height(),
                            fill="#7A5C12", dash=(6, 4), width=1)
        self.cv.create_line(0, self.cv.winfo_height() / 2,
                            self.cv.winfo_width(), self.cv.winfo_height() / 2,
                            fill="#7A5C12", dash=(6, 4), width=1)


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()

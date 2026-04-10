import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont, ImageTk
import os

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError:
    messagebox.showerror("缺少组件", "请在终端执行: pip install tkinterdnd2")
    exit()

# ==========================================
# ⚙️ 预设据数据库
# ==========================================
TEMPLATE_CONFIGS = {
    "template_E6.png": {
        "style": "E6_Classic",
        "photo": {"size": 405, "x": 92, "y": 122},
        "fields": {
            "name": {"label": "✎ 玩家名称 (Name):", "size": 28, "x": 0, "y": 505, "color": "#FAD355",
                     "stroke": "#111111"}
        }
    },
    "template_E7.png": {
        "style": "E7_Italic",
        "photo": {"size": 240, "x": 417, "y": 90},
        "fields": {
            "name": {"label": "✎ 玩家昵称 (Name):", "size": 26, "x": -6, "y": 317, "color": "#FFFFFF",
                     "stroke": "#111111"},
            "city": {"label": "⌂ 玩家城市 (City):", "size": 23, "x": -6, "y": 347, "color": "#FAD355",
                     "stroke": "#111111"}
        }
    }
}

# 🎨 尊贵黑金色卡
C_BG_BASE = "#0D0D0D"
C_BG_PANEL = "#161616"
C_BG_BLOCK = "#222222"
C_BG_INPUT = "#333333"
C_GOLD_MAIN = "#D4AF37"
C_GOLD_TXT = "#E6C27A"
C_WHITE = "#E0E0E0"


class UltimateImageComposerV37:
    def __init__(self, root):
        self.root = root
        self.root.title("宣传图制作工作台 - V3.7 终极防乱码对齐版")
        self.root.geometry("1250x850")
        self.root.configure(bg=C_BG_BASE)

        self.app_dir = self._get_application_path()
        self.image_queue = []
        self.current_image_path = None
        self.final_generated_img = None

        self.field_controls = {}
        self.font_cache = {}
        self.tofu_cache = {}
        self.system_fonts = self._init_system_fonts()

        self._current_opened_path = None
        self._current_raw_img = None
        self._current_tpl_name = None

        self.cache_tmpl_img = None
        self.cache_player_img = None
        self.canvas_scale = 1.0
        self.view_offset_x = 0
        self.view_offset_y = 0
        self.canvas_img_id = None

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TCombobox", fieldbackground=C_BG_INPUT, background=C_BG_BLOCK, foreground=C_GOLD_MAIN,
                        arrowcolor=C_GOLD_MAIN, bordercolor=C_BG_BLOCK)
        style.map("TCombobox", fieldbackground=[("readonly", C_BG_INPUT)], selectbackground=[("readonly", C_GOLD_MAIN)],
                  selectforeground=[("readonly", "black")], foreground=[("readonly", C_GOLD_MAIN)])
        self.root.option_add('*TCombobox*Listbox.background', C_BG_INPUT)
        self.root.option_add('*TCombobox*Listbox.foreground', C_GOLD_MAIN)
        self.root.option_add('*TCombobox*Listbox.selectBackground', C_GOLD_MAIN)
        self.root.option_add('*TCombobox*Listbox.selectForeground', 'black')

        self.setup_ui()
        self.scan_templates()
        self.setup_drag_and_drop()

    def _get_application_path(self):
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        elif __file__:
            return os.path.dirname(os.path.abspath(__file__))
        return os.getcwd()

    def setup_ui(self):
        self.panel_left = tk.Frame(self.root, bg=C_BG_PANEL, width=420)
        self.panel_left.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        self.panel_left.pack_propagate(False)

        tk.Label(self.panel_left, text="⚜️ 工作台", fg=C_GOLD_MAIN, bg=C_BG_PANEL,
                 font=("Microsoft YaHei UI", 18, "bold")).pack(pady=(15, 20))

        box_top = tk.LabelFrame(self.panel_left, text=" 模板确认和图片录入 ", bg=C_BG_BLOCK, fg=C_GOLD_MAIN,
                                font=("Microsoft YaHei UI", 11, "bold"), bd=1)
        box_top.pack(fill=tk.X, padx=10, pady=5)

        self.combo_tpl = ttk.Combobox(box_top, state="readonly", font=("Arial", 12, "bold"))
        self.combo_tpl.pack(fill=tk.X, padx=15, pady=(15, 10))
        self.combo_tpl.bind("<<ComboboxSelected>>", self.on_template_change)

        self.btn_img = tk.Button(box_top, text="[ 点击 / 拖入多张图片 ]\n建立自动流水线队列",
                                 command=self.select_image_batch, bg="#2B2B2B", fg=C_GOLD_TXT,
                                 activebackground=C_GOLD_MAIN, activeforeground="black",
                                 font=("Microsoft YaHei UI", 10), height=2, relief="groove")
        self.btn_img.pack(fill=tk.X, padx=15, pady=(5, 10))

        self.lbl_queue = tk.Label(box_top, text="[当前队列空闲]", fg="#888888", bg=C_BG_BLOCK,
                                  font=("Microsoft YaHei UI", 9))
        self.lbl_queue.pack(pady=(0, 10))

        self.box_photo = tk.LabelFrame(self.panel_left, text=" 图片位置参数 ", bg=C_BG_BLOCK, fg=C_GOLD_MAIN,
                                       font=("Microsoft YaHei UI", 10), bd=1)
        self.box_photo.pack(fill=tk.X, padx=10, pady=10)
        f_p = tk.Frame(self.box_photo, bg=C_BG_BLOCK)
        f_p.pack(pady=10)

        def _make_spinbox(parent, text, from_, to, col):
            tk.Label(parent, text=text, bg=C_BG_BLOCK, fg=C_WHITE, font=("Microsoft YaHei", 9)).grid(row=0,
                                                                                                     column=col * 2,
                                                                                                     padx=(5, 0))
            sp = tk.Spinbox(parent, from_=from_, to=to, width=5, bg=C_BG_INPUT, fg=C_GOLD_MAIN,
                            insertbackground=C_GOLD_MAIN, buttonbackground=C_BG_BLOCK)
            sp.grid(row=0, column=col * 2 + 1, padx=2)
            return sp

        self.spin_p_size = _make_spinbox(f_p, "大小宽:", 10, 8000, 0)
        self.spin_p_x = _make_spinbox(f_p, "X轴:", -4000, 4000, 1)
        self.spin_p_y = _make_spinbox(f_p, "Y轴:", -4000, 4000, 2)

        self.spin_p_size.config(command=self.trigger_full_render)
        self.spin_p_size.bind("<Return>", lambda e: self.trigger_full_render())
        self.spin_p_x.config(command=self.trigger_fast_render)
        self.spin_p_x.bind("<Return>", lambda e: self.trigger_fast_render())
        self.spin_p_y.config(command=self.trigger_fast_render)
        self.spin_p_y.bind("<Return>", lambda e: self.trigger_fast_render())

        self.box_texts = tk.LabelFrame(self.panel_left, text=" 文字操作工作台 ", bg=C_BG_BLOCK, fg=C_GOLD_MAIN,
                                       font=("Microsoft YaHei UI", 10), bd=1)
        self.box_texts.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        f_btns = tk.Frame(self.panel_left, bg=C_BG_PANEL)
        f_btns.pack(fill=tk.X, padx=10, pady=15, side=tk.BOTTOM)

        self.btn_preview = tk.Button(f_btns, text="⟳ 强制刷新画面", command=self.trigger_full_render, bg="#222222",
                                     fg=C_WHITE, relief="groove")
        self.btn_preview.pack(fill=tk.X, pady=(0, 10))

        self.btn_save = tk.Button(f_btns, text="✦ 渲染输出并切下张 (Next) ✦", command=self.save_and_next,
                                  bg=C_GOLD_MAIN, fg="black", activebackground="#FFE47A", activeforeground="black",
                                  font=("Microsoft YaHei UI", 13, "bold"), height=2, state="disabled", cursor="hand2")
        self.btn_save.pack(fill=tk.X)

        title_text = " 🖥️ 原画视口：[滚轮:缩全图] | [Ctrl+滚轮:缩放照片] | [绿框:辅助雷达] "
        self.panel_right = tk.LabelFrame(self.root, text=title_text, bg=C_BG_BASE, fg=C_GOLD_MAIN,
                                         font=("Microsoft YaHei UI", 11, "bold"), bd=1)
        self.panel_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 10), pady=10)
        self.canvas = tk.Canvas(self.panel_right, bg="#080808", highlightthickness=0, cursor="tcross")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.txt_hint = self.canvas.create_text(400, 300, text="⚔️\n等待列队...", fill="#444444",
                                                font=("Microsoft YaHei", 16), justify="center")

        self.canvas.bind("<MouseWheel>", self.master_mouse_wheel)
        self.canvas.bind("<Button-4>", self.master_mouse_wheel)
        self.canvas.bind("<Button-5>", self.master_mouse_wheel)
        self.canvas.bind("<ButtonPress-1>", self.on_photo_drag_start)
        self.canvas.bind("<B1-Motion>", self.on_photo_drag_motion)
        self.canvas.bind("<ButtonPress-3>", self.on_canvas_pan_start)
        self.canvas.bind("<B3-Motion>", self.on_canvas_pan_motion)

    def setup_drag_and_drop(self):
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', lambda e: self._handle_drag_data(e.data))

    def select_image_batch(self):
        paths = filedialog.askopenfilenames(title="选择多张照片建立队列",
                                            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp")])
        if paths:
            self._add_to_queue(paths)

    def _handle_drag_data(self, data):
        paths = self.root.tk.splitlist(data)
        self._add_to_queue(paths)

    def _add_to_queue(self, paths):
        valid_paths = [p for p in paths if str(p).lower().endswith(('.png', '.jpg', '.jpeg', '.webp'))]
        if not valid_paths:
            return
        self.image_queue.extend(valid_paths)
        if not self.current_image_path:
            self.load_next_in_queue()
        else:
            self._update_queue_ui()

    def load_next_in_queue(self):
        if not self.image_queue:
            self.current_image_path = None
            self.lbl_queue.config(text="✅ 队列全部完成！等待新任务", fg="#2ECC71")
            self.btn_save.config(state="disabled", text="✦ 渲染输出并继续下一张 (Next) ✦")
            return
        self.current_image_path = self.image_queue.pop(0)
        self._update_queue_ui()
        self.trigger_full_render()

    def _update_queue_ui(self):
        cur_name = os.path.basename(self.current_image_path)
        rem_count = len(self.image_queue)
        if rem_count > 0:
            self.lbl_queue.config(text=f"▶ 当前: {cur_name} | ⏳ 队列剩余: {rem_count}张", fg=C_GOLD_TXT)
            self.btn_save.config(text=f"✦ 渲染保存 >> 自动切下一张 ({rem_count}) ✦")
        else:
            self.lbl_queue.config(text=f"▶ 当前: {cur_name} | 🏁 队列最后一张", fg=C_GOLD_MAIN)
            self.btn_save.config(text="✦ 渲染保存终图 ✦")

    def save_and_next(self):
        if not self.final_generated_img or not self.current_image_path:
            return
        try:
            o_dir = os.path.dirname(self.current_image_path)
            o_name = os.path.splitext(os.path.basename(self.current_image_path))[0]
            save_dest = os.path.join(o_dir, f"出图_{o_name}.png")
            self.final_generated_img.save(save_dest)
            if self.image_queue:
                self.load_next_in_queue()
            else:
                self.load_next_in_queue()
                messagebox.showinfo("收工", "🎉 撒花！所有排队图片已全部处理输出完毕！")
        except Exception as e:
            messagebox.showerror("写入异常", f"保存失败:\n{str(e)}")

    def scan_templates(self):
        valid = []
        try:
            all_files = os.listdir(self.app_dir)
            for file in all_files:
                name_lower = file.lower()
                if "template_e6" in name_lower and name_lower.endswith(('.png', '.jpg', '.jpeg')):
                    valid.append(file)
                    TEMPLATE_CONFIGS[file] = TEMPLATE_CONFIGS["template_E6.png"]
                elif "template_e7" in name_lower and name_lower.endswith(('.png', '.jpg', '.jpeg')):
                    valid.append(file)
                    TEMPLATE_CONFIGS[file] = TEMPLATE_CONFIGS["template_E7.png"]
        except Exception:
            pass

        if valid:
            self.combo_tpl['values'] = valid
            self.combo_tpl.current(0)
            self.on_template_change()
        else:
            messagebox.showwarning("❌ 警告：未找到模板！",
                                   f"在以下路径扫描不到模板：\n{self.app_dir}\n请放入 template_E6 / E7")

    def on_template_change(self, event=None):
        tpl_name = self.combo_tpl.get()
        config = TEMPLATE_CONFIGS.get(tpl_name)
        if not config:
            return

        self.spin_p_size.delete(0, 'end')
        self.spin_p_size.insert(0, config["photo"]["size"])
        self.spin_p_x.delete(0, 'end')
        self.spin_p_x.insert(0, config["photo"]["x"])
        self.spin_p_y.delete(0, 'end')
        self.spin_p_y.insert(0, config["photo"]["y"])

        self.view_offset_x = 0
        self.view_offset_y = 0
        for widget in self.box_texts.winfo_children():
            widget.destroy()
        self.field_controls.clear()

        for field_k, field_v in config["fields"].items():
            f_group = tk.Frame(self.box_texts, bg=C_BG_BLOCK)
            f_group.pack(fill=tk.X, padx=10, pady=(8, 0))
            tk.Label(f_group, text=field_v["label"], fg=C_GOLD_TXT, bg=C_BG_BLOCK,
                     font=("Microsoft YaHei UI", 9, "bold")).pack(anchor="w")

            entry = tk.Entry(f_group, font=("Trebuchet MS", 12), justify="center", bg=C_BG_INPUT, fg=C_GOLD_MAIN,
                             insertbackground=C_GOLD_MAIN, bd=0)
            entry.pack(fill=tk.X, ipady=4, pady=2)
            entry.bind("<KeyRelease>", lambda e: self.trigger_fast_render())

            f_t_ctrl = tk.Frame(f_group, bg=C_BG_BLOCK)
            f_t_ctrl.pack(fill=tk.X, pady=2)

            def _txt_spin(parent, text, from_, to):
                tk.Label(parent, text=text, bg=C_BG_BLOCK, fg=C_WHITE, font=("Microsoft YaHei", 8)).pack(side=tk.LEFT,
                                                                                                         padx=(
                                                                                                             5 if from_ < 10 else 0,
                                                                                                             0))
                sp = tk.Spinbox(parent, from_=from_, to=to, width=4, bg=C_BG_INPUT, fg=C_GOLD_MAIN,
                                buttonbackground=C_BG_BLOCK, command=self.trigger_fast_render)
                sp.pack(side=tk.LEFT, padx=2)
                sp.bind("<Return>", lambda e: self.trigger_fast_render())
                return sp

            s_size = _txt_spin(f_t_ctrl, "字号:", 10, 400)
            s_size.delete(0, 'end')
            s_size.insert(0, field_v["size"])
            s_x = _txt_spin(f_t_ctrl, "左右:", -1000, 1000)
            s_x.delete(0, 'end')
            s_x.insert(0, field_v["x"])
            s_y = _txt_spin(f_t_ctrl, "上下:", -500, 2000)
            s_y.delete(0, 'end')
            s_y.insert(0, field_v["y"])

            self.field_controls[field_k] = {
                "entry": entry,
                "spin_size": s_size,
                "spin_x": s_x,
                "spin_y": s_y,
                "color": field_v["color"],
                "stroke": field_v["stroke"]
            }

        if self.current_image_path:
            self.trigger_full_render()

    def master_mouse_wheel(self, event):
        if not self.final_generated_img:
            return
        delta = getattr(event, 'delta', 0)
        is_zoom_in = (event.num == 4 or delta > 0)
        is_zoom_out = (event.num == 5 or delta < 0)
        is_ctrl_pressed = (event.state & 0x0004) != 0

        if is_ctrl_pressed:
            try:
                current_size = int(self.spin_p_size.get())
                if is_zoom_in:
                    new_size = int(current_size * 1.05) + 2
                elif is_zoom_out:
                    new_size = int(current_size * 0.95) - 2
                else:
                    return
                new_size = max(10, min(new_size, 8000))
                self.spin_p_size.delete(0, 'end')
                self.spin_p_size.insert(0, str(new_size))
                self.trigger_full_render()
            except:
                pass
        else:
            if is_zoom_in:
                self.canvas_scale *= 1.15
            elif is_zoom_out:
                self.canvas_scale /= 1.15
            self.update_canvas_display()

    def on_photo_drag_start(self, event):
        if not self.final_generated_img:
            return
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        try:
            self.drag_base_p_x = int(self.spin_p_x.get())
            self.drag_base_p_y = int(self.spin_p_y.get())
        except:
            pass

    def on_photo_drag_motion(self, event):
        if not self.final_generated_img:
            return
        dx = (event.x - self.drag_start_x) / self.canvas_scale
        dy = (event.y - self.drag_start_y) / self.canvas_scale
        self.spin_p_x.delete(0, 'end')
        self.spin_p_x.insert(0, int(self.drag_base_p_x + dx))
        self.spin_p_y.delete(0, 'end')
        self.spin_p_y.insert(0, int(self.drag_base_p_y + dy))
        self.trigger_fast_render()

    def on_canvas_pan_start(self, event):
        if not self.final_generated_img:
            return
        self.pan_start_x = event.x
        self.pan_start_y = event.y
        self.base_view_x = self.view_offset_x
        self.base_view_y = self.view_offset_y

    def on_canvas_pan_motion(self, event):
        if not self.final_generated_img or not self.canvas_img_id:
            return
        self.view_offset_x = self.base_view_x + (event.x - self.pan_start_x)
        self.view_offset_y = self.base_view_y + (event.y - self.pan_start_y)
        c_w = self.canvas.winfo_width() / 2
        c_h = self.canvas.winfo_height() / 2
        self.canvas.coords(self.canvas_img_id, c_w + self.view_offset_x, c_h + self.view_offset_y)

        # ==========================================

    # 🚀 超级字库引擎
    # ==========================================
    def _init_system_fonts(self):
        font_queue = [
            "segoeui.ttf", "seguisym.ttf", "msyh.ttc", "msyhbd.ttc",
            "arial.ttf", "arialbd.ttf", "tahoma.ttf", "simhei.ttf"
        ]
        valid_paths = []

        all_sys = []
        if os.path.exists("C:\\Windows\\Fonts"):
            for f in os.listdir("C:\\Windows\\Fonts"):
                if f.lower().endswith(('.ttf', '.ttc')):
                    all_sys.append(os.path.join("C:\\Windows\\Fonts", f))

        for target in font_queue:
            for p in all_sys:
                if os.path.basename(p).lower() == target:
                    valid_paths.append(p)
                    break

        for p in all_sys:
            if p not in valid_paths:
                valid_paths.append(p)

        return valid_paths

    # ==========================================
    # ✅ 修复核心：豆腐块检测 - 缓存键包含字号
    # ==========================================
    def is_char_supported(self, font, font_path, char, size):
        """
        用绘图法检测字体是否真正支持某个字符。
        修复：tofu_cache 的键从 font_path 改为 (font_path, size)，
        确保不同字号下的豆腐块像素签名尺寸一致，比对有效。
        """
        canvas_dim = size * 2
        img = Image.new("L", (canvas_dim, canvas_dim), 0)
        try:
            ImageDraw.Draw(img).text((size, size), char, font=font, fill=255, anchor="md")
            sig = img.tobytes()
        except:
            return False

        # 画出来全空白 = 字体没这个字
        if not any(sig):
            return False

        # ✅ 关键修复：缓存键必须包含字号
        # 旧代码用 font_path 做键，首次在 size=56 建立的豆腐签名是 112×112 bytes，
        # 后续 size=54 时签名是 108×108 bytes，长度不同永远匹配不上，
        # 导致所有字体都被误判为"支持"，第一个字体(segoeui)直接中选但实际不支持 → □□
        cache_key = (font_path, size)
        if cache_key not in self.tofu_cache:
            tofu_sigs = []
            for mc in ["\uFFFF", "\uFFFD"]:
                t_img = Image.new("L", (canvas_dim, canvas_dim), 0)
                try:
                    ImageDraw.Draw(t_img).text((size, size), mc, font=font, fill=255, anchor="md")
                    tofu_sigs.append(t_img.tobytes())
                except:
                    pass
            self.tofu_cache[cache_key] = tofu_sigs

        return sig not in self.tofu_cache[cache_key]

    def get_font_for_char(self, char, size):
        cache_key = f"{char}_{size}"
        if cache_key in self.font_cache:
            return self.font_cache[cache_key]

        # 空格秒过
        if char.isspace():
            f = ImageFont.truetype(self.system_fonts[0], size) if self.system_fonts else ImageFont.load_default()
            self.font_cache[cache_key] = (f, char)
            return f, char

        for fp in self.system_fonts:
            try:
                font = ImageFont.truetype(fp, size)
                if self.is_char_supported(font, fp, char, size):
                    self.font_cache[cache_key] = (font, char)
                    return font, char
            except:
                continue

        # 全部阵亡兜底
        fallback = ImageFont.truetype(self.system_fonts[0], size) if self.system_fonts else ImageFont.load_default()
        self.font_cache[cache_key] = (fallback, "囗")
        return fallback, "囗"

    # ==========================================
    # ✨ 3D金属质感文字渲染
    # ==========================================
    def draw_unified_3d_metal_text(self, target_img, text, tmpl_w, tmpl_h, size, offset_x, y_pos, main_color,
                                   stroke_color, style):
        SS = 2
        large_size = size * SS

        # 测宽定位
        char_data_list = []
        total_width = 0
        for char in text:
            font, final_char = self.get_font_for_char(char, large_size)
            try:
                w = font.getlength(final_char)
            except:
                w = large_size * 0.8
            char_data_list.append((final_char, font, w))
            total_width += w

        pad = int(large_size * 2)
        canv_w = int(total_width + pad * 2)
        canv_h = int(large_size * 4)

        temp_cx = pad
        baseline_y = int(large_size * 2.5)

        mask_text = Image.new("L", (canv_w, canv_h), 0)
        draw_text = ImageDraw.Draw(mask_text)
        mask_stroke = Image.new("L", (canv_w, canv_h), 0)
        draw_stroke = ImageDraw.Draw(mask_stroke)

        stroke_w = max(1, int(large_size * 0.025))
        depth = max(SS, int(large_size * 0.1))

        cx = temp_cx
        for char, font, w in char_data_list:
            try:
                draw_text.text((cx, baseline_y), char, font=font, fill=255, anchor="ls")
                draw_stroke.text((cx, baseline_y), char, font=font, fill=255, stroke_width=stroke_w, stroke_fill=255,
                                 anchor="ls")
            except TypeError:
                draw_text.text((cx, baseline_y - int(large_size)), char, font=font, fill=255)
                draw_stroke.text((cx, baseline_y - int(large_size)), char, font=font, fill=255, stroke_width=stroke_w,
                                 stroke_fill=255)
            cx += w

        upright_canvas = Image.new("RGBA", (canv_w, canv_h), (0, 0, 0, 0))
        color_black = (10, 10, 10, 255)
        if main_color == "#FFFFFF":
            color_bevel = (90, 100, 115, 255)
            c1 = (255, 255, 255)
            c2 = (230, 235, 240)
            c3 = (160, 165, 175)
            c4 = (220, 225, 235)
        else:
            color_bevel = (145, 80, 10, 255)
            c1 = (255, 255, 230)
            c2 = (255, 215, 60)
            c3 = (245, 170, 10)
            c4 = (255, 230, 80)

        img_black = Image.new("RGBA", (canv_w, canv_h), color_black)
        img_bevel = Image.new("RGBA", (canv_w, canv_h), color_bevel)

        for dy in range(depth + 1):
            upright_canvas.paste(img_black, (0, dy), mask=mask_stroke)
        for dy in range(1, depth + 1):
            upright_canvas.paste(img_bevel, (0, dy), mask=mask_text)

        grad_layer = Image.new("RGBA", (canv_w, canv_h), (0, 0, 0, 0))
        grad_draw = ImageDraw.Draw(grad_layer)

        y_top = baseline_y - int(large_size * 0.9)
        y_max = baseline_y + int(large_size * 0.1)
        scan_h = max(1, y_max - y_top)

        for y in range(y_top, y_max + 1):
            ratio = (y - y_top) / scan_h
            if ratio <= 0.45:
                r_i = ratio / 0.45
                c_up = c1
                c_dn = c2
            elif ratio <= 0.55:
                r_i = (ratio - 0.45) / 0.10
                c_up = c2
                c_dn = c3
            else:
                r_i = (ratio - 0.55) / 0.45
                c_up = c3
                c_dn = c4
            r = int(c_up[0] * (1 - r_i) + c_dn[0] * r_i)
            g = int(c_up[1] * (1 - r_i) + c_dn[1] * r_i)
            b = int(c_up[2] * (1 - r_i) + c_dn[2] * r_i)
            grad_draw.line([(0, y), (canv_w, y)], fill=(r, g, b, 255))

        face_layer = Image.new("RGBA", (canv_w, canv_h), (0, 0, 0, 0))
        face_layer.paste(grad_layer, (0, 0), mask=mask_text)
        upright_canvas.alpha_composite(face_layer)

        if style == "E7_Italic":
            shear_angle = 0.22
            shear_offset = -shear_angle * (canv_h / 2)
            matrix = (1, shear_angle, shear_offset, 0, 1, 0)
            final_stamp = upright_canvas.transform((canv_w, canv_h), Image.AFFINE, matrix,
                                                   resample=Image.Resampling.BICUBIC)
        else:
            final_stamp = upright_canvas

        final_w = canv_w // SS
        final_h = canv_h // SS
        final_stamp = final_stamp.resize((final_w, final_h), Image.Resampling.LANCZOS)

        orig_pad = pad // SS
        orig_tot_w = total_width // SS

        paste_x = int((tmpl_w - orig_tot_w) / 2 + offset_x - orig_pad)
        paste_y = int(y_pos - (baseline_y // SS) + size)
        target_img.alpha_composite(final_stamp, dest=(paste_x, paste_y))

        # ==========================================

    # 渲染流程
    # ==========================================
    def trigger_full_render(self):
        if not self.current_image_path:
            return

            # ✅ 字号变化时清除字体选择缓存，强制重新匹配
        self.font_cache.clear()

        try:
            self.root.update()
            tpl_name = self.combo_tpl.get()
            req_width = int(self.spin_p_size.get())

            if getattr(self, '_current_tpl_name', None) != tpl_name:
                actual_conf_key = None
                for key, val in TEMPLATE_CONFIGS.items():
                    if key == tpl_name:
                        actual_conf_key = key
                        break
                tpl_abs_path = os.path.join(self.app_dir, actual_conf_key if actual_conf_key else tpl_name)
                self.cache_tmpl_img = Image.open(tpl_abs_path).convert("RGBA")
                self._current_tpl_name = tpl_name

            if getattr(self, '_current_opened_path', None) != self.current_image_path:
                self._current_raw_img = Image.open(self.current_image_path).convert("RGBA")
                self._current_opened_path = self.current_image_path

            raw_player = self._current_raw_img
            orig_w, orig_h = raw_player.size
            ratio = req_width / float(orig_w) if orig_w > 0 else 1
            new_h = max(1, int(orig_h * ratio))

            self.cache_player_img = raw_player.resize((req_width, new_h), Image.Resampling.LANCZOS)

            if getattr(self, '_is_first_load', True):
                if self.cache_tmpl_img.size[0] > 800:
                    self.canvas_scale = 0.6
                else:
                    self.canvas_scale = 1.0
                self._is_first_load = False

            self.trigger_fast_render()
            self.btn_save.config(state="normal")
        except Exception as e:
            pass

    def trigger_fast_render(self):
        if not self.cache_tmpl_img or not self.cache_player_img:
            return
        try:
            p_x = int(self.spin_p_x.get())
            p_y = int(self.spin_p_y.get())
            tpl_name = self.combo_tpl.get()

            actual_conf_key = None
            for key, val in TEMPLATE_CONFIGS.items():
                if key == tpl_name:
                    actual_conf_key = key
                    break

            tmpl_style = TEMPLATE_CONFIGS.get(actual_conf_key, {}).get("style", "E6_Classic")

            tmpl_w, tmpl_h = self.cache_tmpl_img.size

            mem_img = Image.new("RGBA", (tmpl_w, tmpl_h), (0, 0, 0, 0))
            mem_img.paste(self.cache_player_img, (p_x, p_y))
            mem_img = Image.alpha_composite(mem_img, self.cache_tmpl_img)

            for key, ctrl in self.field_controls.items():
                text_val = ctrl["entry"].get().strip()
                if not text_val:
                    continue
                t_size = int(ctrl["spin_size"].get())
                t_x_off = int(ctrl["spin_x"].get())
                t_y = int(ctrl["spin_y"].get())

                self.draw_unified_3d_metal_text(
                    mem_img, text_val, tmpl_w, tmpl_h,
                    t_size, t_x_off, t_y,
                    ctrl["color"], ctrl["stroke"], tmpl_style
                )

            self.final_generated_img = mem_img.copy()
            self.update_canvas_display()
        except:
            pass

    def update_canvas_display(self):
        if not self.final_generated_img:
            return
        self.canvas.delete("all")
        w, h = self.final_generated_img.size
        new_w = max(10, int(w * self.canvas_scale))
        new_h = max(10, int(h * self.canvas_scale))
        preview = self.final_generated_img.resize((new_w, new_h), Image.Resampling.NEAREST)
        self.tk_canvas_img = ImageTk.PhotoImage(preview)

        c_w = self.canvas.winfo_width() / 2 + self.view_offset_x
        c_h = self.canvas.winfo_height() / 2 + self.view_offset_y
        self.canvas_img_id = self.canvas.create_image(c_w, c_h, anchor=tk.CENTER, image=self.tk_canvas_img)

        if self.cache_player_img:
            p_x = int(self.spin_p_x.get())
            p_y = int(self.spin_p_y.get())
            p_w, p_h = self.cache_player_img.size

            canvas_start_x = c_w - new_w / 2
            canvas_start_y = c_h - new_h / 2

            bbox_x1 = canvas_start_x + p_x * self.canvas_scale
            bbox_y1 = canvas_start_y + p_y * self.canvas_scale
            bbox_x2 = bbox_x1 + p_w * self.canvas_scale
            bbox_y2 = bbox_y1 + p_h * self.canvas_scale

            self.canvas.create_rectangle(bbox_x1, bbox_y1, bbox_x2, bbox_y2, outline="#00FFCC", width=1, dash=(5, 3))


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = UltimateImageComposerV37(root)
    root.mainloop()

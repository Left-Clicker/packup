import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.toast import ToastNotification
from ttkbootstrap.scrolled import ScrolledText
from tkinter import filedialog, messagebox, colorchooser
import tkinter as tk
import pandas as pd
import math
import re
import json
import os
import sys
import copy
import time
import threading
from datetime import datetime, timedelta
from itertools import groupby
from operator import itemgetter

os.environ['TK_SILENCE_DEPRECATION'] = '1'


# ================= 0. 模板管理器 =================

class TemplateManager:
    FILE_NAME = "templates.json"

    DEFAULT_TEMPLATES = {
        "contact_sum": (
            "Hi dear, your account will reach {bonus_name} bonus and get free {reward}*$99.99 packs, "
            "if you purchase {miss} more $99.99 packs before {deadline} City Time."
        ),
        "contact_count": (
            "Hi dear, your account will reach {bonus_name} bonus and get free {reward}*$99.99 packs, "
            "if you purchase {miss} more $99.99 non-Diamonds packs before {deadline} City Time."
        )
    }

    def __init__(self):
        self.templates = self.load_templates()

    def load_templates(self):
        base_path = get_app_path()
        full_path = os.path.join(base_path, self.FILE_NAME)
        if not os.path.exists(full_path): return self.DEFAULT_TEMPLATES.copy()
        try:
            with open(full_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k, v in self.DEFAULT_TEMPLATES.items():
                    if k not in data: data[k] = v
                return data
        except:
            return self.DEFAULT_TEMPLATES.copy()

    def save_templates(self, new_data):
        self.templates = new_data
        base_path = get_app_path()
        full_path = os.path.join(base_path, self.FILE_NAME)
        try:
            with open(full_path, "w", encoding="utf-8") as f:
                json.dump(new_data, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            return False

    def get(self, key):
        return self.templates.get(key, self.DEFAULT_TEMPLATES[key])

    def render(self, key, **kwargs):
        tmpl = self.get(key)
        try:
            return tmpl.format(**kwargs)
        except KeyError as e:
            return f"Error: Template missing variable {{{e.args[0]}}}. Raw: {tmpl}"


def get_app_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


tmpl_mgr = TemplateManager()

# ================= 1. 配置区域 =================

RULE_COLORS = {
    "1000+180": "#FF6B6B",
    "500+85": "#FF9F43",
    "168+25": "#Feca57",
    "68+8": "#2ECC71",
    "34+4": "#54A0FF",
    "10+1": "#D980FA",
    "DEFAULT": "#333333"
}

COLUMN_MAPPING_CONFIG = {
    'oid': '订单号', 'pid': '玩家昵称', 'real_id': '玩家Player_id', 'amt': '礼包价格',
    'time': '订单发放日期', 'server': '服务器昵称', 'status': '是否发放元宝',
    'is_marked': '是否已标记绩效', 'perf_owner': '业绩归属人', 'free_event': '是否计入免费包活动'
}

SUM_RULES = [
    {"name": "1000+180", "hours": 72, "type": "sum", "target": 99990.00, "priority": 10, "default_min_reached": 700,
     "reward": 180},
    {"name": "500+85", "hours": 48, "type": "sum", "target": 49995.00, "priority": 20, "default_min_reached": 300,
     "reward": 85},
    {"name": "168+25", "hours": 48, "type": "sum", "target": 16798.32, "priority": 30, "default_min_reached": 100,
     "reward": 25},
    {"name": "68+8", "hours": 48, "type": "sum", "target": 6799.32, "priority": 40, "default_min_reached": 54,
     "reward": 8},
    {"name": "34+4", "hours": 24, "type": "sum", "target": 3399.66, "priority": 50, "default_min_reached": 20,
     "reward": 4},
]

COUNT_RULE = {"name": "10+1", "hours": 24, "type": "count", "target": 10, "unit_price": 99.99, "priority": 60,
              "default_min_reached": 5, "reward": 1}

# 必须按优先级排序 (1000 -> 10)
ALL_RULES = sorted(SUM_RULES + [COUNT_RULE], key=lambda x: x['priority'])


def is_marked_performance(row): return str(row.get('is_marked', '')).strip() == '是'


def is_free_event_pack(row): return str(row.get('free_event', '否')).strip() != '否'


def format_event_text_full(e):
    oids = ", ".join([str(o['oid']) for o in e['data']['orders']])
    return f"Rule: {e['rule']['name']}\nTotal: ${e['data']['total']:.2f}\nCount: {e['data']['count']}\nOrders: {oids}"


def show_copy_toast(title, message):
    try:
        toast = ToastNotification(
            title=title,
            message=message,
            duration=2000,
            bootstyle="success",
            position=(50, 50, 'se')
        )
        toast.show_toast()
    except:
        pass

        # Display Logic


def generate_summary_grouped(events, calc_mode="normal", ref_time=None, current_oid=None, suppress_map=None):
    if not events: return ""

    # 【修复调整】彻底废除了原本用来隐藏 miss<=1 的 cutoff_priority 拦截逻辑
    # 确保上方的 Treeview 列表与底部卡片读取数据源保持绝对一致，如实显示差1。

    grouped = {}
    for e in events:
        r_name = e['rule']['name']
        r_priority = e['rule']['priority']

        if r_name not in grouped: grouped[r_name] = {'priority': r_priority, 'achieved_sets': 0, 'contacts': set()}

        if e['type'] == 'achieved':
            grouped[r_name]['achieved_sets'] += e['sets']
        elif e['type'] == 'contact':
            # 过滤已超时的催单
            if ref_time and e['data']['deadline'] <= ref_time: continue

            # 【被删除的代码】：不再拦截强行剔除 e['miss'] <= 1 的数据

            grouped[r_name]['contacts'].add(e['miss'])

    sorted_groups = sorted(grouped.items(), key=lambda x: x[1]['priority'])
    parts = []
    for r_name, data in sorted_groups:
        sub_parts = []
        if data['achieved_sets'] > 0: sub_parts.append(f"{r_name}达成x{data['achieved_sets']}")
        if data['contacts']:
            min_miss = min(data['contacts'])
            sub_parts.append(f"{r_name}差{min_miss}")
        if sub_parts: parts.append(" ".join(sub_parts))
    return " | ".join(parts)


# --- 核心计算引擎 (Standard - Waterfall) ---
# 逻辑：优先达成大额，达成即消耗并停止检查(Break)；未达成则检查催单，并继续检查低档位(Continue)。
def calculate_achievements_normal(orders, all_rules, time_ext, limit_configs={}, ref_time=None):
    if ref_time is None: ref_time = datetime.utcnow()
    n = len(orders)
    events = []
    used_indices = set()

    i = 0
    while i < n:
        if i in used_indices:
            i += 1
            continue

        consumed_in_this_pass = False

        for rule in all_rules:
            # 1. 确定起点
            target_price = rule.get('unit_price', 99.99)
            actual_start_idx = i

            if rule['type'] == 'count':
                temp_idx = i
                found_valid_start = False
                while temp_idx < n:
                    if temp_idx in used_indices:
                        temp_idx += 1;
                        continue
                    if abs(orders[temp_idx]['amt'] - target_price) < 0.05:
                        actual_start_idx = temp_idx
                        found_valid_start = True
                        break
                    temp_idx += 1
                if not found_valid_start: continue

            st = orders[actual_start_idx]['time_obj']
            t_ext_val = time_ext if rule['type'] == 'sum' else 0
            window_s = (rule['hours'] + t_ext_val) * 3600
            deadline = st + timedelta(seconds=window_s)

            scan_indices = []
            acc_val = 0.0
            valid_cnt = 0

            for j in range(actual_start_idx, n):
                if j in used_indices: continue
                row = orders[j]
                if row['time_obj'] > deadline: break

                is_valid = False
                if rule['type'] == 'sum':
                    acc_val += row['amt'];
                    is_valid = True
                elif rule['type'] == 'count':
                    if abs(row['amt'] - target_price) < 0.05:
                        valid_cnt += 1;
                        is_valid = True

                if is_valid: scan_indices.append(j)

                # 2. 判断状态
            is_achieved = False
            # 【修复1：防浮点不达成】增加 0.05 的容错，防止本该达成却判定失败
            if rule['type'] == 'sum' and acc_val >= (rule['target'] - 0.05):
                is_achieved = True
            elif rule['type'] == 'count' and valid_cnt >= int(rule['target']):
                is_achieved = True

            if is_achieved:
                # === 达成 ===
                final_orders = []
                temp_sum = 0.0;
                temp_cnt = 0
                for idx in scan_indices:
                    o = orders[idx]
                    final_orders.append(o)
                    used_indices.add(idx)
                    if rule['type'] == 'sum':
                        temp_sum += o['amt']
                        if temp_sum >= (rule['target'] - 0.05): break
                    elif rule['type'] == 'count':
                        temp_cnt += 1
                        if temp_cnt >= int(rule['target']): break

                events.append({
                    'type': 'achieved', 'rule': rule, 'sets': 1,
                    'data': {'orders': final_orders, 'total': sum(x['amt'] for x in final_orders),
                             'count': len(final_orders), 'start_t': final_orders[0]['time_obj'], 'deadline': deadline}
                })
                consumed_in_this_pass = True
                break

            else:
                # === 催单 ===
                limit = limit_configs.get(rule['name'], 0)
                if limit > 0:
                    is_contact = False;
                    miss = 0
                    if rule['type'] == 'sum':
                        if acc_val >= (limit * 99.99 - 0.05):
                            is_contact = True
                            # 【修复2：彻底解决连续13和最后多算1包的问题】
                            # 利用 round 先将结果逼近规范小数位，再 ceil 抹去末尾脏数据进行整数进位
                            miss = math.ceil(round((rule['target'] - acc_val) / 99.99, 4))
                    elif rule['type'] == 'count':
                        if valid_cnt >= limit:
                            is_contact = True
                            miss = int(rule['target']) - valid_cnt

                    if is_contact:
                        events.append({
                            'type': 'contact', 'rule': rule, 'miss': miss,
                            'data': {'orders': [orders[k] for k in scan_indices], 'total': acc_val,
                                     'count': len(scan_indices), 'start_t': st, 'deadline': deadline}
                        })

        if not consumed_in_this_pass:
            i += 1

    return events, used_indices


# Kernel logic with Pure Surplus Check (Safest Mode)
def calculate_achievements_complex(orders, all_rules, time_ext, limit_configs, ref_time):
    sum_rules = [r for r in all_rules if r['type'] == 'sum']
    normal_events, _ = calculate_achievements_normal(orders, sum_rules, time_ext, limit_configs, ref_time)
    achieved_sum_events = [e for e in normal_events if e['type'] == 'achieved']
    order_to_sum_event_idx = {}
    for idx, evt in enumerate(achieved_sum_events):
        for o in evt['data']['orders']: order_to_sum_event_idx[str(o['oid'])] = idx

    orders_99 = [o for o in orders if abs(o['amt'] - 99.99) < 0.05]
    n99 = len(orders_99)
    count_rule = next((r for r in all_rules if r['type'] == 'count'), None)
    if not count_rule: return normal_events, set()

    achieved_count_events = [];
    ids_consumed_by_count = set()
    i = 0
    while i <= n99 - 10:
        if str(orders_99[i]['oid']) in ids_consumed_by_count: i += 1; continue
        st = orders_99[i]['time_obj'];
        ddl = st + timedelta(hours=24)
        grp = [];
        temp_j = i
        while temp_j < n99:
            o = orders_99[temp_j]
            if str(o['oid']) in ids_consumed_by_count: temp_j += 1; continue
            if o['time_obj'] > ddl: break
            grp.append(o);
            if len(grp) == 10: break
            temp_j += 1
        if len(grp) < 10: i += 1; continue

        bounds = {}
        for o in grp:
            oid = str(o['oid'])
            if oid in order_to_sum_event_idx:
                idx = order_to_sum_event_idx[oid]
                if idx not in bounds: bounds[idx] = []
                bounds[idx].append(o)

        possible = True;
        plans = {}
        reserved_refill_oids = set()
        for e_idx, stolen in bounds.items():
            evt = achieved_sum_events[e_idx]
            cur = evt['data']['total'];
            tgt = evt['rule']['target']
            need = tgt - (cur - sum(o['amt'] for o in stolen))
            if need <= 0.05: plans[e_idx] = []; continue
            est = evt['data']['start_t'];
            eed = evt['data']['deadline']
            pot = []
            c_ids = {str(o['oid']) for o in grp}
            for o in orders:
                oid = str(o['oid'])
                if o['time_obj'] < est or o['time_obj'] > eed: continue
                # 特殊模式补单去重：同一轮中已分配给其它事件的订单不可重复使用
                if oid not in order_to_sum_event_idx and oid not in c_ids and oid not in reserved_refill_oids:
                    pot.append(o)
            refill = [];
            r_tot = 0.0
            for sp in pot:
                refill.append(sp);
                r_tot += sp['amt']
                if r_tot >= need - 0.05: break
            if r_tot < need - 0.05: possible = False; break
            plans[e_idx] = refill
            reserved_refill_oids.update(str(x['oid']) for x in refill)

        if possible:
            for e_idx, stolen in bounds.items():
                evt = achieved_sum_events[e_idx];
                sps = plans[e_idx]
                s_ids = [str(o['oid']) for o in stolen]
                for s_oid in s_ids:
                    # 被借走的订单不再属于原累充事件，必须清理映射
                    if order_to_sum_event_idx.get(s_oid) == e_idx:
                        del order_to_sum_event_idx[s_oid]
                new_list = [o for o in evt['data']['orders'] if str(o['oid']) not in s_ids] + sps
                evt['data']['orders'] = new_list
                evt['data']['total'] = sum(o['amt'] for o in new_list)
                evt['data']['count'] = len(new_list)
                for o in sps: order_to_sum_event_idx[str(o['oid'])] = e_idx
            achieved_count_events.append({'type': 'achieved', 'rule': count_rule, 'sets': 1,
                                          'data': {'orders': grp, 'total': sum(o['amt'] for o in grp), 'count': 10,
                                                   'start_t': grp[0]['time_obj'], 'deadline': ddl}
                                          })
            for o in grp: ids_consumed_by_count.add(str(o['oid']))
        i += 1

    final = achieved_sum_events + achieved_count_events
    min_r = limit_configs.get(count_rule['name'], 0)

    def is_borrowable_safely(oid):
        if oid in ids_consumed_by_count: return False
        if oid not in order_to_sum_event_idx: return True
        e_idx = order_to_sum_event_idx[oid]
        evt = achieved_sum_events[e_idx]
        if (evt['data']['total'] - 99.99) >= (evt['rule']['target'] - 0.05):
            return True
        return False

    if min_r > 0:
        i = 0
        while i < n99:
            o_start = orders_99[i]
            oid_start = str(o_start['oid'])
            if not is_borrowable_safely(oid_start):
                i += 1;
                continue
            st = o_start['time_obj'];
            ddl = st + timedelta(hours=24)
            current_chain = [o_start]
            for j in range(i + 1, n99):
                o_curr = orders_99[j]
                if o_curr['time_obj'] > ddl: break
                oid_curr = str(o_curr['oid'])
                if is_borrowable_safely(oid_curr):
                    current_chain.append(o_curr)
            cnt = len(current_chain)
            if cnt >= min_r and cnt < 10 and ref_time < ddl:
                final.append({'type': 'contact', 'rule': count_rule, 'miss': 10 - cnt,
                              'data': {'orders': current_chain, 'total': sum(o['amt'] for o in current_chain),
                                       'count': cnt, 'start_t': st, 'deadline': ddl}
                              })
            i += 1

    final += [e for e in normal_events if e['type'] == 'contact']
    return final, set()


# 全局辅助函数：寻找最佳起点 (Best Scenario)
def find_best_scenario(orders, all_rules, time_ext, limit_map, ref_time):
    if not orders: return []
    best_score = -1
    best_events = []

    scan_limit = min(len(orders), 100)

    for i in range(scan_limit):
        sub_orders = orders[i:]
        evts, _ = calculate_achievements_normal(sub_orders, all_rules, time_ext, limit_map, ref_time)
        score = sum(e['rule'].get('reward', 0) * e['sets'] for e in evts if e['type'] == 'achieved')

        if score > best_score:
            best_score = score
            best_events = evts

    return best_events


# ================= 3. 编辑器 UI =================

class TemplateEditorWindow(ttk.Toplevel):
    def __init__(self, parent):
        super().__init__(title="Edit Templates", master=parent)
        self.geometry("800x600")
        self.data = tmpl_mgr.load_templates()
        self.text_widgets = {}

        ttk.Label(self, text="模板编辑器", font=("Helvetica", 14, "bold"), bootstyle="primary").pack(pady=10)
        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        f1 = ttk.Labelframe(container, text="累充规则模板 (Sum Rules)", bootstyle="primary")
        f1.pack(fill=tk.BOTH, expand=True, pady=5)
        t1 = ScrolledText(f1, height=5, font=("Consolas", 10))
        t1.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        t1.insert("1.0", self.data.get("contact_sum", ""))
        self.text_widgets["contact_sum"] = t1

        f2 = ttk.Labelframe(container, text="计数规则模板 (Count Rules)", bootstyle="success")
        f2.pack(fill=tk.BOTH, expand=True, pady=5)
        t2 = ScrolledText(f2, height=5, font=("Consolas", 10))
        t2.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        t2.insert("1.0", self.data.get("contact_count", ""))
        self.text_widgets["contact_count"] = t2

        ttk.Button(self, text="💾 保存并生效", bootstyle="success", command=self.save).pack(pady=10, fill=tk.X, padx=50)

    def save(self):
        new_data = {}
        for k, w in self.text_widgets.items(): new_data[k] = w.get("1.0", "end-1c").strip()
        if tmpl_mgr.save_templates(new_data): self.destroy()

        # ================= 4. UI 组件 (SafeScrollableFrame) =================


class SafeScrollableFrame(ttk.Frame):
    def __init__(self, parent, columns=2, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.columns = columns
        style = ttk.Style()
        bg_color = style.lookup("TFrame", "background")
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg=bg_color)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind('<Configure>', self._configure_window_width)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.bind_all("<MouseWheel>", self._on_mousewheel)

    def _configure_window_width(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        try:
            x, y = self.winfo_pointerxy()
            widget = self.winfo_containing(x, y)
            if widget and str(self) in str(widget):
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except:
            pass

    def clear(self):
        for widget in self.scrollable_frame.winfo_children(): widget.destroy()

    def add_card(self, title, content, bootstyle="secondary", oids=None, template_text=None):
        count = len(self.scrollable_frame.winfo_children())
        row, col = count // self.columns, count % self.columns
        card = ttk.Labelframe(self.scrollable_frame, text=f" {title} ", bootstyle=bootstyle)
        card.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

        txt = tk.Text(card, height=5, width=40, font=("Consolas", 9), bg="#2b2b2b", fg="#eee", relief="flat",
                      insertbackground="white")
        txt.insert("1.0", content)
        txt.config(state=tk.DISABLED)
        txt.pack(fill="both", expand=True, padx=5, pady=5)

        def _enter_txt(e):
            self.unbind_all("<MouseWheel>")

        def _leave_txt(e):
            self.bind_all("<MouseWheel>", self._on_mousewheel)

        txt.bind("<Enter>", _enter_txt);
        txt.bind("<Leave>", _leave_txt)

        btn_frame = ttk.Frame(card);
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        def copy_str(s, msg):
            self.winfo_toplevel().clipboard_clear()
            self.winfo_toplevel().clipboard_append(s)
            show_copy_toast("COPIED", msg)

        if oids:
            if isinstance(oids, list):
                o_str = "\n".join(oids)
            else:
                o_str = str(oids)
            ttk.Button(btn_frame, text="🆔 Copy Orders", bootstyle="info-outline", cursor="hand2",
                       command=lambda: copy_str(o_str, f"{len(oids)} Orders Copied")).pack(side=tk.LEFT, padx=5,
                                                                                           fill=tk.X, expand=True)
        if template_text:
            ttk.Button(btn_frame, text="📋 Copy Msg", bootstyle="success-outline", cursor="hand2",
                       command=lambda: copy_str(template_text, "Template Text Copied")).pack(side=tk.LEFT, padx=5,
                                                                                             fill=tk.X, expand=True)
        self.scrollable_frame.grid_columnconfigure(col, weight=1)

        # ================= 5. 单玩家检查窗口类 =================


class SinglePlayerCheck(ttk.Toplevel):
    def __init__(self, parent, player_data, time_ext, limit_configs, ref_time):
        t_str = ref_time.strftime("%Y-%m-%d %H:%M:%S") if ref_time else "Unknown"
        super().__init__(title=f"单玩家检查 (ID:{player_data.get('real_id', '?')}) @ Ref: {t_str}", master=parent)
        self.geometry("1550x850")
        self.p_data = player_data
        self.time_ext = time_ext
        self.limit_configs = limit_configs
        self.ref_time = ref_time

        self.calc_mode = "normal"
        self.opt_diff_val = 0
        self.opt_10_info = ""
        self.frm_manual_stats = None

        self.precalc_data = {'normal': [], 'special': []}
        self.cache_view = {}
        self.iid_to_oid = {}
        self.oid_to_iid = {}

        self._init_ui()
        self._setup_tags()
        self._precompute_all()
        self._refresh_view()

    def _init_ui(self):
        frame_top = ttk.Frame(self);
        frame_top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        frame_legend = ttk.LabelFrame(frame_top, text=" Color Legend ")
        frame_legend.pack(side=tk.TOP, fill=tk.X, pady=(0, 10), ipadx=5, ipady=5)
        for rule_name, color_hex in RULE_COLORS.items():
            if rule_name == "DEFAULT": continue
            canvas = tk.Canvas(frame_legend, width=90, height=26, bg="#444444", highlightthickness=0)
            canvas.pack(side=tk.LEFT, padx=5, pady=5)
            canvas.create_text(45, 13, text=rule_name, fill=color_hex, font=("Helvetica", 10, "bold"))

        frame_controls = ttk.Frame(frame_top);
        frame_controls.pack(side=tk.TOP, fill=tk.X)
        self.frm_manual_stats = ttk.Frame(frame_controls);
        self.frm_manual_stats.pack(side=tk.LEFT, padx=(0, 20))

        # ================== 【新增区域：下拉菜单按钮组】 ==================
        filter_frame = ttk.Frame(frame_controls)
        filter_frame.pack(side=tk.LEFT, padx=10)

        # 1. 剔除归属人下拉菜单
        mb_owner = ttk.Menubutton(filter_frame, text="⛔ 剔除归属人订单", bootstyle="warning-outline")
        mb_owner.pack(side=tk.LEFT, padx=2)
        menu_owner = tk.Menu(mb_owner, tearoff=0)
        menu_owner.add_command(label="剔除【所有】归属人/标记订单", command=lambda: self.filter_owner(keep_10_1=False))
        menu_owner.add_command(label="剔除归属人 (但保留 '24小时内买10送1')",
                               command=lambda: self.filter_owner(keep_10_1=True))
        mb_owner['menu'] = menu_owner

        # 2. 剔除免费包下拉菜单
        mb_free = ttk.Menubutton(filter_frame, text="🆓 剔除免费包订单", bootstyle="warning-outline")
        mb_free.pack(side=tk.LEFT, padx=2)
        menu_free = tk.Menu(mb_free, tearoff=0)
        menu_free.add_command(label="剔除【所有】免费包订单", command=lambda: self.filter_free_event(keep_10_1=False))
        menu_free.add_command(label="剔除免费包 (但保留 '24小时内买10送1')",
                              command=lambda: self.filter_free_event(keep_10_1=True))
        mb_free['menu'] = menu_free

        # 3. 重置按钮 (保持普通按钮)
        ttk.Button(filter_frame, text="🔄 重置列表", bootstyle="info-outline", command=self.reset_filter).pack(
            side=tk.LEFT, padx=2)
        # ==========================================================

        rt_frame = ttk.Frame(frame_controls);

        rt_frame.pack(side=tk.RIGHT)
        ttk.Button(rt_frame, text="🧹 Clear Marks", bootstyle="danger-outline", command=self.clear_manual_marks).pack(
            side=tk.LEFT, padx=5)
        self.btn_mode = ttk.Button(rt_frame, text="计算中...", bootstyle="secondary", state=tk.DISABLED,
                                   command=self.toggle_mode)
        self.btn_mode.pack(side=tk.LEFT, padx=5)

        ttk.Label(frame_controls, text="Shift+Click选范围 -> 右键手动标记", font=("Helvetica", 9),
                  bootstyle="secondary").pack(side=tk.LEFT)

        self.paned = ttk.Panedwindow(self, orient=tk.VERTICAL);
        self.paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        frame_list = ttk.Frame(self.paned);
        self.paned.add(frame_list, weight=3)
        cols = ("time", "oid", "result", "score")
        self.tree = ttk.Treeview(frame_list, columns=cols, show="headings", height=10)
        self.tree.heading("time", text="Start Time (GMT+0)");
        self.tree.column("time", width=150)
        self.tree.heading("oid", text="Order | Price");
        self.tree.column("oid", width=220)
        self.tree.heading("result", text="Analysis Result");
        self.tree.column("result", width=850)
        self.tree.heading("score", text="Reward");
        self.tree.column("score", width=80, anchor="center")
        sb = ttk.Scrollbar(frame_list, command=self.tree.yview);
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True);
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        def _copy_id(e):
            sel = self.tree.selection()
            if sel: val = self.tree.item(sel[0], 'values'); self.clipboard_clear(); self.clipboard_append(
                val[1].split('|')[0].strip())
            show_copy_toast("COPIED", "Copied")

        self.tree.bind("<Control-c>", _copy_id)

        self.cm = tk.Menu(self.tree, tearoff=0)
        self.cm.add_command(label="Mark Selection (Mixed)", command=lambda: self.mark_selection_logic("all"))
        self.cm.add_command(label="Mark Selection (Only 99.99)", command=lambda: self.mark_selection_logic("99"))
        self.tree.bind("<Button-3>", lambda e: self.cm.post(e.x_root, e.y_root))

        frame_bottom = ttk.Frame(self.paned);
        self.paned.add(frame_bottom, weight=2)
        self.nb_cards = ttk.Notebook(frame_bottom, bootstyle="warning");
        self.nb_cards.pack(fill=tk.BOTH, expand=True)
        self.f_ach_cards = SafeScrollableFrame(self.nb_cards);
        self.nb_cards.add(self.f_ach_cards, text=" ACHIEVED ")
        self.f_con_cards = SafeScrollableFrame(self.nb_cards);
        self.nb_cards.add(self.f_con_cards, text=" CONTACT ")
        f_tmpl = ttk.Frame(self.nb_cards);
        self.nb_cards.add(f_tmpl, text=" LOGS ")
        self.txt = ScrolledText(f_tmpl, height=8, font=("Consolas", 10), bootstyle="secondary");
        self.txt.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def _setup_tags(self):
        for name, color in RULE_COLORS.items():
            if name != "DEFAULT": self.tree.tag_configure(name, background=color, foreground="white")
        self.tree.tag_configure('separator', background='#444444', foreground='#dddddd')
        self.tree.tag_configure('normal_row', background=RULE_COLORS["DEFAULT"], foreground="#cccccc")

        # ================== 【基础抽取改写区域】提取带参数据生成器 ==================

    def _precompute_data(self, orders):
        t_strict = self.ref_time - timedelta(hours=24)
        t_loose = self.ref_time - timedelta(hours=24 + self.time_ext)

        def compute_mode(mode):
            max_score = -1
            scan_limit = min(len(orders), 100)

            best_evts_for_suppress = []

            for k in range(scan_limit):
                sub = orders[k:]
                if mode == "normal":
                    evts_k, _ = calculate_achievements_normal(sub, ALL_RULES, self.time_ext, self.limit_configs,
                                                              self.ref_time)
                else:
                    evts_k, _ = calculate_achievements_complex(sub, ALL_RULES, self.time_ext, self.limit_configs,
                                                               self.ref_time)

                score_k = sum(e['rule'].get('reward', 0) * e['sets'] for e in evts_k if e['type'] == 'achieved')

                if score_k > max_score:
                    max_score = score_k
                    best_evts_for_suppress = evts_k
                elif max_score == -1 and k == 0:
                    max_score = score_k
                    best_evts_for_suppress = evts_k

            suppress_map = {}
            if mode == "normal":
                for e in best_evts_for_suppress:
                    if e['type'] == 'achieved':
                        p = e['rule']['priority']
                        for o in e['data']['orders']:
                            oid = str(o['oid'])
                            if oid not in suppress_map or p < suppress_map[oid]:
                                suppress_map[oid] = p

            rows = []
            inserts_s, inserts_l = False, False
            has_10 = False

            for i, row in enumerate(orders):
                sep = None
                if not inserts_l and row['time_obj'] > t_loose:
                    sep = ("-" * 20, f"--- 24H + {self.time_ext}H Tolerance ---", "-", "")
                    inserts_l = True
                if not inserts_s and row['time_obj'] > t_strict:
                    if sep: rows.append({'sep': True, 'val': sep})
                    sep = ("-" * 20, "--- 24H Strict Limit (10+1) ---", "-", "")
                    inserts_s = True
                if sep: rows.append({'sep': True, 'val': sep})

                sub = orders[i:]
                if mode == "normal":
                    evts, _ = calculate_achievements_normal(sub, ALL_RULES, self.time_ext, self.limit_configs,
                                                            self.ref_time)
                else:
                    evts, _ = calculate_achievements_complex(sub, ALL_RULES, self.time_ext, self.limit_configs,
                                                             self.ref_time)

                oid_disp = f"{row['oid']} | {row['amt']}"
                is_marked = str(row.get('is_marked', '否')).strip()
                perf_owner = str(row.get('perf_owner', '')).strip()
                free_evt = str(row.get('free_event', '否')).strip()
                extra_tags = []
                if is_marked not in ['否', 'nan', '', 'None', 'NO']:
                    if perf_owner and perf_owner not in ['nan', 'None', '']:
                        extra_tags.append(perf_owner)
                    else:
                        extra_tags.append("绩效✔")
                if free_evt not in ['否', 'nan', '', 'None', 'NO']: extra_tags.append(free_evt)
                if extra_tags: oid_disp += f" [{'|'.join(extra_tags)}]"

                sc = sum(e['rule'].get('reward', 0) * e['sets'] for e in evts if e['type'] == 'achieved')
                row_sum = generate_summary_grouped(evts, mode, self.ref_time, str(row.get('oid')), suppress_map)

                rows.append({
                    'sep': False,
                    'val': (row['time_str'], oid_disp, row_sum, str(sc) if sc else ""),
                    'evt': evts,
                    'oid': str(row.get('oid', f"idx_{i}"))
                })

                if mode == 'special' and not has_10:
                    if any(e['type'] == 'contact' and e['rule']['name'] == '10+1' for e in evts): has_10 = True

            return rows, max_score, has_10

        n_data, sc_norm, _ = compute_mode("normal")
        s_data, sc_spec, h10 = compute_mode("special")

        return {'normal': n_data, 'special': s_data}, sc_spec - sc_norm, "10+1 Contact" if h10 else ""

    def _precompute_all(self):
        # 缓存最原始数据及生成最初底图计算
        self.original_orders = self.p_data['orders_list']
        self.default_precalc_data, self.default_opt_diff_val, self.default_opt_10_info = self._precompute_data(
            self.original_orders)

        self.precalc_data = self.default_precalc_data
        self.opt_diff_val = self.default_opt_diff_val
        self.opt_10_info = self.default_opt_10_info

        self.btn_mode.config(state=tk.NORMAL)
        self._update_button_style()

        # ================== 【新增功能区域】过滤方法 ==================

    def filter_owner(self, keep_10_1=False):
        filtered_orders = []
        for row in self.original_orders:
            perf_owner = str(row.get('perf_owner', '')).strip()
            has_owner = perf_owner and perf_owner not in ['nan', 'None', '']
            is_marked = is_marked_performance(row)
            free_event_val = str(row.get('free_event', '')).strip()

            # 判断是否是需要被剔除的"已标记归属人"的订单
            if is_marked or has_owner:
                # 如果开启了"保留10+1"模式，且该订单正好是"24小时内买10送1"，则免于剔除
                if keep_10_1 and "24小时内买10送1" in free_event_val:
                    pass  # 放行保留
                else:
                    continue  # 正式剔除
            filtered_orders.append(row)
        self._apply_new_orders(filtered_orders)

    def filter_free_event(self, keep_10_1=False):
        filtered_orders = []
        for row in self.original_orders:
            is_free = is_free_event_pack(row)
            free_event_val = str(row.get('free_event', '')).strip()

            # 判断是否是免费包订单
            if is_free:
                # 如果开启了"保留10+1"模式，且该订单正好是"24小时内买10送1"，则免于剔除
                if keep_10_1 and "24小时内买10送1" in free_event_val:
                    pass  # 放行保留
                else:
                    continue  # 正式剔除
            filtered_orders.append(row)
        self._apply_new_orders(filtered_orders)

    def reset_filter(self):
        self.precalc_data = self.default_precalc_data
        self.opt_diff_val = self.default_opt_diff_val
        self.opt_10_info = self.default_opt_10_info

        # 刷新视图并清理卡片防残影
        self._refresh_view()
        self.f_ach_cards.clear()
        self.f_con_cards.clear()
        for w in self.frm_manual_stats.winfo_children(): w.destroy()
        self._update_button_style()

    def _apply_new_orders(self, modified_orders):
        # 执行实时过滤计算并更新
        self.precalc_data, self.opt_diff_val, self.opt_10_info = self._precompute_data(modified_orders)
        self._refresh_view()
        self.f_ach_cards.clear()
        self.f_con_cards.clear()
        for w in self.frm_manual_stats.winfo_children(): w.destroy()
        self._update_button_style()

        # ==========================================================

    def _refresh_view(self):
        self.tree.delete(*self.tree.get_children())
        self.cache_view = {}
        self.iid_to_oid = {};
        self.oid_to_iid = {}

        data = self.precalc_data[self.calc_mode]
        for item in data:
            if item['sep']:
                self.tree.insert("", tk.END, values=item['val'], tags=('separator',))
            else:
                iid = self.tree.insert("", tk.END, values=item['val'], tags=('normal_row',))
                self.cache_view[iid] = item['evt']
                self.iid_to_oid[iid] = item['oid']
                self.oid_to_iid[item['oid']] = iid

    def toggle_mode(self):
        self.calc_mode = "special" if self.calc_mode == "normal" else "normal"
        self._update_button_style()
        self._refresh_view()

    def _update_button_style(self):
        if self.calc_mode == "normal":
            if self.opt_diff_val > 0:
                self.btn_mode.configure(text=f"🚀 推荐: 特殊模式 (收益 +{self.opt_diff_val})", bootstyle="success")
            elif self.opt_10_info:
                self.btn_mode.configure(text=f"✨ 特殊: {self.opt_10_info} (收益 +{self.opt_diff_val})",
                                        bootstyle="info")
            else:
                self.btn_mode.configure(text="切换: 特殊计算 (优先10+1)", bootstyle="info-outline")
        else:
            self.btn_mode.configure(text="切换: 普通计算 (Standard)", bootstyle="secondary-outline")

    def _update_stats_display(self, selection_only=False):
        for w in self.frm_manual_stats.winfo_children(): w.destroy()

        def extract_price(val_str):
            try:
                price_part = val_str.split('|')[1]
                return float(re.search(r"(\d+\.?\d*)", price_part).group(1))
            except:
                return 0.0

        sel = self.tree.selection()
        sel_total = 0.0
        if sel:
            for iid in sel:
                if 'separator' in self.tree.item(iid, "tags"): continue
                val_str = self.tree.item(iid, 'values')[1]
                sel_total += extract_price(val_str)

        mark_sums = {}
        for child in self.tree.get_children():
            tags = self.tree.item(child, "tags")
            for t in tags:
                if t.startswith('manual_mark_'):
                    parts = t.split('_')
                    if len(parts) > 2:
                        color = parts[2]
                        val_str = self.tree.item(child, 'values')[1]
                        amt = extract_price(val_str)
                        mark_sums[color] = mark_sums.get(color, 0) + amt

        if sel_total > 0:
            self._create_stats_canvas(f"Selection: ${sel_total:.2f}", "white")

        for color, total in mark_sums.items():
            self._create_stats_canvas(f"${total:.2f}", color)

    def _create_stats_canvas(self, text, color):
        w = len(text) * 10
        cvs = tk.Canvas(self.frm_manual_stats, width=w, height=24, highlightthickness=0, bg="#2b2b2b")
        cvs.pack(side=tk.LEFT, padx=5)
        cvs.create_text(w / 2, 12, text=text, fill=color, font=("Helvetica", 10, "bold"))

    def clear_manual_marks(self):
        for item in self.tree.get_children():
            tags = list(self.tree.item(item, "tags"))
            self.tree.item(item, tags=tuple([t for t in tags if not t.startswith('manual_mark')] or ['normal_row']))
        self._update_stats_display(False)

    def mark_selection_logic(self, mode):
        sel = self.tree.selection()
        if not sel: return
        color = colorchooser.askcolor(parent=self)[1]
        if not color: return
        tag = f"manual_mark_{color}_{datetime.now().timestamp()}"
        self.tree.tag_configure(tag, background=color, foreground="white")
        indices = [self.tree.index(i) for i in sel]
        for i in range(min(indices), max(indices) + 1):
            iid = self.tree.get_children()[i]
            if 'separator' in self.tree.item(iid, "tags"): continue
            val = self.tree.item(iid, 'values')[1]
            if mode == "99" and "99.99" not in val: continue
            tags = list(self.tree.item(iid, "tags"))
            if 'normal_row' in tags: tags.remove('normal_row')
            tags = [t for t in tags if not t.startswith('manual_mark')];
            tags.append(tag)
            self.tree.item(iid, tags=tuple(tags))
        self._update_stats_display(False)

    def on_select(self, event):
        sel = self.tree.selection()
        self._update_stats_display(selection_only=True)
        if not sel: return
        iid = sel[0]
        if iid not in self.cache_view: return
        events = sorted(self.cache_view[iid],
                        key=lambda x: (0 if x['type'] == 'achieved' else 1, x['rule']['priority']))

        for child in self.tree.get_children():
            tags = list(self.tree.item(child, "tags"))
            keep = [t for t in tags if t.startswith('separator') or t.startswith('manual_mark')]
            self.tree.item(child, tags=tuple(keep or ['normal_row']))

        self.f_ach_cards.clear();
        self.f_con_cards.clear();
        b_cons = {};
        has_contact = False;
        txt_log = ""

        for e in events:
            if e['type'] == 'achieved':
                txt = format_event_text_full(e)
                oids = [str(o['oid']) for o in e['data']['orders']]
                self.f_ach_cards.add_card(e['rule']['name'], txt, "success", oids)
                txt_log += txt + "\n"

                for o in e['data']['orders']:
                    tgt = self.oid_to_iid.get(str(o.get('oid', '')))
                    if tgt:
                        ct = list(self.tree.item(tgt, "tags"))
                        if any(t.startswith('manual_mark') for t in ct): continue
                        if 'normal_row' in ct: ct.remove('normal_row')
                        if e['rule']['name'] not in ct: ct.append(e['rule']['name'])
                        self.tree.item(tgt, tags=tuple(ct))

            elif e['type'] == 'contact':
                # 修复: 使用 self.ref_time 而不是 self.current_ref_time
                if e['data']['deadline'] <= self.ref_time: continue
                rn = e['rule']['name']
                if rn not in b_cons or e['miss'] < b_cons[rn]['miss']:
                    b_cons[rn] = e
                has_contact = True

        for r_name, e in b_cons.items():
            t_key = "contact_count" if e['rule']['type'] == 'count' else "contact_sum"

            # --- 核心修改开始 ---
            # 重新计算严格截止时间：起始时间 + 规则原生小时数 (不加 time_ext)
            strict_deadline = e['data']['start_t'] + timedelta(hours=e['rule']['hours'])

            msg = tmpl_mgr.render(t_key, bonus_name=r_name, reward=e['rule'].get('reward'), miss=e['miss'],
                                  deadline=str(strict_deadline).split('.')[0])
            # --- 核心修改结束 ---

            oids = [str(o['oid']) for o in e['data']['orders']]
            self.f_con_cards.add_card(r_name, msg, "warning", oids, template_text=msg)
        if has_contact:
            self.nb_cards.select(1)
        else:
            self.nb_cards.select(0)

            # ================= 6. 主程序 (修改版) =================


class PromoAnalyzerApp:
    def __init__(self):
        self.root = ttk.Window(title="黑道英文业绩计算器V7.6", themename="darkly")
        self.root.geometry("1450x980")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.df_raw = None
        self.results_cache = {2: [], 3: [], 4: [], 5: []}

        self.current_ref_time = None
        self.tree_views = []
        self.limit_entries = {}
        self.time_ext_val = 3.0

        self.selected_days = tk.IntVar(value=3)
        self.days_options = [2, 3, 4, 5]

        # [修改点1] 新增：用于全局保存打勾的玩家 PID 的记忆本
        self.checked_pids = set()

        self._customize_styles()
        self._init_ui()

    def on_close(self):
        if messagebox.askyesno("Exit", "Are you sure you want to quit?"): self.root.destroy()

    def _customize_styles(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)
        style.configure("TNotebook.Tab", font=("Helvetica", 10, "bold"))
        style.configure("Day.TButton", font=("Helvetica", 10, "bold"))

    def _init_ui(self):
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.X, padx=15, pady=10)

        # Row 1
        row1 = ttk.Frame(main_container)
        row1.pack(fill=tk.X, pady=5)
        ttk.Button(row1, text="📁 导入 Excel", bootstyle="warning-outline", command=self.load_file).pack(side=tk.LEFT)
        ttk.Button(row1, text="📝 编辑话术", bootstyle="secondary-outline", command=self.open_template_editor).pack(
            side=tk.LEFT, padx=10)

        ttk.Label(row1, text="超时宽限(H):").pack(side=tk.LEFT, padx=(20, 0))
        self.entry_time_ext = ttk.Entry(row1, width=5);
        self.entry_time_ext.insert(0, "3");
        self.entry_time_ext.pack(side=tk.LEFT, padx=5)

        f_rules = ttk.Frame(row1);
        f_rules.pack(side=tk.LEFT, padx=20)
        for rule in ALL_RULES:
            ttk.Label(f_rules, text=f"{rule['name']}:").pack(side=tk.LEFT, padx=(5, 0))
            e = ttk.Entry(f_rules, width=4);
            e.insert(0, str(rule.get('default_min_reached', 0)));
            e.pack(side=tk.LEFT)
            self.limit_entries[rule['name']] = e

            # Row 2
        row2 = ttk.Frame(main_container)
        row2.pack(fill=tk.X, pady=10)

        f_days = ttk.Labelframe(row2, text=" Data Range (Days) ", bootstyle="info")
        f_days.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))

        for d in self.days_options:
            rb = ttk.Radiobutton(
                f_days,
                text=f" {d} Days ",
                variable=self.selected_days,
                value=d,
                bootstyle="info-toolbutton",
                command=self.refresh_current_view
            )
            rb.pack(side=tk.LEFT, padx=2, pady=5, ipady=3)

        f_filter = ttk.Labelframe(row2, text=" ID Filter (One per line) ", bootstyle="secondary")
        f_filter.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.txt_filter = ScrolledText(f_filter, height=3, width=50, font=("Consolas", 9))
        self.txt_filter.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        f_filter_btns = ttk.Frame(f_filter)
        f_filter_btns.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)

        ttk.Button(f_filter_btns, text="Apply Filter", bootstyle="secondary-outline",
                   command=self.refresh_current_view).pack(fill=tk.X, pady=2)
        ttk.Button(f_filter_btns, text="Clear Filter", bootstyle="danger-outline", command=self.clear_filter).pack(
            fill=tk.X, pady=2)

        # Row 3
        row3 = ttk.Frame(main_container)
        row3.pack(fill=tk.X, pady=5)
        ttk.Button(row3, text="⚡ 开始计算 (Calc All Days)", bootstyle="warning", command=self.run_analysis).pack(
            side=tk.LEFT)
        self.lbl_status = ttk.Label(row3, text="READY");
        self.lbl_status.pack(side=tk.LEFT, padx=20)
        self.progress = ttk.Progressbar(row3, length=300, mode="determinate", bootstyle="warning-striped");
        self.progress.pack(side=tk.LEFT)
        ttk.Button(row3, text="🔍 单人详情", command=self.open_single_check).pack(side=tk.RIGHT)

        # List Area
        self.paned = ttk.Panedwindow(self.root, orient=tk.VERTICAL);
        self.paned.pack(fill=tk.BOTH, expand=True, padx=15)
        self.nb_list = ttk.Notebook(self.paned);
        self.paned.add(self.nb_list, weight=1)
        self.tree_all = self._add_tree(self.nb_list, " ALL PLAYERS ")
        self.tree_achieved = self._add_tree(self.nb_list, " ✅ ACHIEVED ")
        self.tree_contact = self._add_tree(self.nb_list, " 📞 CONTACT ")
        self.tree_views = [self.tree_all, self.tree_achieved, self.tree_contact]

        self.nb_cards = ttk.Notebook(self.paned, bootstyle="warning");
        self.paned.add(self.nb_cards, weight=2)
        self.f_ach_cards = SafeScrollableFrame(self.nb_cards);
        self.nb_cards.add(self.f_ach_cards, text=" ACHIEVED ")
        self.f_con_cards = SafeScrollableFrame(self.nb_cards);
        self.nb_cards.add(self.f_con_cards, text=" CONTACT ")

    def _add_tree(self, nb, text):
        f = ttk.Frame(nb);
        nb.add(f, text=text)
        cols = ("check", "pid", "real_id", "server", "total", "summary")
        t = ttk.Treeview(f, columns=cols, show="headings")
        t.column("check", width=40, anchor="center");
        t.heading("check", text="✔")
        t.column("pid", width=150, anchor="center");
        t.heading("pid", text="玩家昵称")
        t.column("real_id", width=110, anchor="center");
        t.heading("real_id", text="Player_id")
        t.column("server", width=80, anchor="center");
        t.heading("server", text="Server")
        t.column("total", width=100, anchor="center");
        t.heading("total", text="Total($)")
        t.column("summary", width=700, anchor="w");
        t.heading("summary", text="SUMMARY")

        sc = ttk.Scrollbar(f, command=t.yview);
        t.configure(yscrollcommand=sc.set)
        t.pack(side=tk.LEFT, fill=tk.BOTH, expand=True);
        sc.pack(side=tk.RIGHT, fill=tk.Y)
        t.bind("<<TreeviewSelect>>", self.on_player_select)
        t.bind("<Button-1>", self.on_tree_click)
        t.bind("<Double-1>", self.open_single_check)

        def _copy_tr(e):
            sel = t.selection()
            if sel: val = t.item(sel[0], 'values'); self.root.clipboard_clear(); self.root.clipboard_append(str(val[2]))
            show_copy_toast("COPIED", "Player ID Copied")

        t.bind("<Control-c>", _copy_tr)
        return t

    def on_tree_click(self, event):
        tree = event.widget
        col_id = tree.identify_column(event.x)
        if col_id == "#1":
            row_id = tree.identify_row(event.y)
            if row_id:
                vals = list(tree.item(row_id, "values"))
                new_state = "☑" if vals[0] == "☐" else "☐"
                vals[0] = new_state
                pid = vals[1]

                # [修改点2] 新增：把状态记入内存。勾选就加入，取消就移除
                if new_state == "☑":
                    self.checked_pids.add(pid)
                else:
                    self.checked_pids.discard(pid)

                for t in self.tree_views:
                    for child in t.get_children():
                        if str(t.item(child, 'values')[1]) == pid:
                            v = list(t.item(child, 'values'))
                            v[0] = new_state
                            tags = ('processed',) if new_state == "☑" else ()
                            t.item(child, values=v, tags=tags)

    def load_internal_player_ids(self):
        bp = get_app_path();
        fp = os.path.join(bp, "内玩ID.xlsx")
        if not os.path.exists(fp): return set()
        try:
            df = pd.read_excel(fp) if fp.endswith('.xlsx') else pd.read_excel(fp)
            return set(str(int(float(x))) if isinstance(x, (int, float)) else str(x).strip() for x in df.iloc[:, 0])
        except:
            return set()

    def load_file(self):
        fp = filedialog.askopenfilename(filetypes=[("Data", "*.csv *.xlsx")])
        if not fp: return
        self.root.config(cursor="watch");
        self.root.update()
        try:
            df = pd.read_excel(fp) if fp.endswith('.xlsx') else pd.read_csv(fp, encoding='gbk')
            df.rename(columns=lambda x: x.strip(), inplace=True)
            rev_map = {v: k for k, v in COLUMN_MAPPING_CONFIG.items()}
            df.rename(columns=rev_map, inplace=True)
            if 'real_id' in df.columns:
                def clean_id_safe(x):
                    s = str(x).strip()
                    if s.endswith('.0'): return s[:-2]
                    return s

                df['real_id'] = df['real_id'].apply(clean_id_safe)
            internals = self.load_internal_player_ids()
            if internals and 'real_id' in df.columns:
                df = df[~df['real_id'].isin(internals)]
            if 'time' in df.columns: df['time'] = pd.to_datetime(df['time'])
            df.sort_values(by=['pid', 'time'], inplace=True)
            self.df_raw = df
            self.lbl_status.config(text=f"LOADED {len(df)} ROWS")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.root.config(cursor="")

    def get_filter_ids(self):
        raw = self.txt_filter.get("1.0", tk.END).strip()
        if not raw: return None
        raw = raw.replace(',', '\n').replace('，', '\n')
        return set(line.strip() for line in raw.split('\n') if line.strip())

    def clear_filter(self):
        self.txt_filter.delete("1.0", tk.END)
        self.refresh_current_view()

    def run_analysis(self):
        if self.df_raw is None: return
        try:
            self.time_ext_val = float(self.entry_time_ext.get())
            self.limit_config_map = {k: int(v.get()) for k, v in self.limit_entries.items()}
        except:
            return

        self.root.config(cursor="watch")
        self.current_ref_time = datetime.utcnow()
        t_str = self.current_ref_time.strftime("%Y-%m-%d %H:%M:%S")

        df = self.df_raw.to_dict('records')
        for r in df: r['time_obj'] = r['time']; r['time_str'] = str(r['time'])
        grouped_raw = {k: list(v) for k, v in groupby(df, key=lambda x: x.get('pid'))}

        self.results_cache = {2: [], 3: [], 4: [], 5: []}
        total_steps = len(grouped_raw) * len(self.days_options)
        current_step = 0

        for days_lookback in self.days_options:
            target_date = self.current_ref_time - timedelta(days=(days_lookback - 1))
            cutoff_time = target_date.replace(hour=0, minute=0, second=0, microsecond=0)

            day_results = []
            for pid, all_orders in grouped_raw.items():
                filtered_orders = [o for o in all_orders if o['time_obj'] >= cutoff_time]

                if not filtered_orders:
                    current_step += 1
                    continue

                if current_step % 100 == 0:
                    self.progress['value'] = (current_step / total_steps) * 100
                    self.root.update()

                total_amt = sum(o['amt'] for o in filtered_orders)

                # 使用 find_best_scenario 寻找最优起点
                evts = find_best_scenario(filtered_orders, ALL_RULES, self.time_ext_val, self.limit_config_map,
                                          self.current_ref_time)

                has_achieved = any(e['type'] == 'achieved' for e in evts)
                active_contacts = [e for e in evts if
                                   e['type'] == 'contact' and e['data']['deadline'] > self.current_ref_time]
                has_active_contact = len(active_contacts) > 0

                if not has_achieved and not has_active_contact:
                    current_step += 1
                    continue

                summary = generate_summary_grouped(evts, calc_mode="normal", ref_time=self.current_ref_time)

                min_ddl = datetime.max
                if has_active_contact:
                    for e in active_contacts:
                        if e['data']['deadline'] < min_ddl: min_ddl = e['data']['deadline']

                res_obj = {
                    'pid': pid, 'real_id': filtered_orders[0].get('real_id'),
                    'server': filtered_orders[0].get('server'),
                    'total_amt': total_amt,
                    'orders_list': filtered_orders, 'events': evts, 'summary': summary,
                    'has_achieved': has_achieved, 'has_contact': has_active_contact, 'min_ddl': min_ddl
                }
                day_results.append(res_obj)
                current_step += 1

            day_results.sort(key=lambda x: (0 if x['has_contact'] else 1, x['min_ddl']))
            self.results_cache[days_lookback] = day_results

        self.lbl_status.config(text=f"DONE @ {t_str} (UTC)")
        self.root.config(cursor="")
        self.refresh_current_view()

    def refresh_current_view(self):
        for t in self.tree_views: t.delete(*t.get_children())
        self.f_ach_cards.clear()
        self.f_con_cards.clear()

        days = self.selected_days.get()
        filter_ids = self.get_filter_ids()

        data_source = self.results_cache.get(days, [])
        if not data_source: return

        count_shown = 0
        for r in data_source:
            if filter_ids:
                rid = str(r['real_id'])
                if rid not in filter_ids:
                    continue

            pid = r['pid']
            amt_str = f"{r['total_amt']:.2f}"

            # [修改点3] 核心：每次查一下“记忆本”，如果在里面就判定为打过勾，并打上对应Tag
            is_checked = pid in self.checked_pids
            check_state = "☑" if is_checked else "☐"
            tags = ('processed',) if is_checked else ()

            # 第一个值变成动态分配的 check_state
            vals = (check_state, pid, r['real_id'], r['server'], amt_str, r['summary'])

            self.tree_all.insert("", "end", values=vals, tags=tags)
            if r['has_achieved']: self.tree_achieved.insert("", "end", values=vals, tags=tags)
            if r['has_contact']: self.tree_contact.insert("", "end", values=vals, tags=tags)
            count_shown += 1

        self.lbl_status.config(text=f"Showing {days} Days | {count_shown} Players")

    def on_player_select(self, event):
        sel = event.widget.selection()
        if not sel: return
        try:
            pid = str(event.widget.item(sel[0], 'values')[1])
        except:
            return

        current_data = self.results_cache.get(self.selected_days.get(), [])
        data = next((item for item in current_data if str(item['pid']) == pid), None)
        if not data: return

        self.f_ach_cards.clear();
        self.f_con_cards.clear()
        best_contacts = {};
        has_contact = False
        for e in data['events']:
            if e['type'] == 'achieved':
                txt = format_event_text_full(e)
                oids = [str(o['oid']) for o in e['data']['orders']]
                self.f_ach_cards.add_card(e['rule']['name'], txt, "success", oids)
            elif e['type'] == 'contact':
                if e['data']['deadline'] <= self.current_ref_time: continue
                rn = e['rule']['name']
                if rn not in best_contacts or e['miss'] < best_contacts[rn]['miss']: best_contacts[rn] = e
                has_contact = True
        for r_name, e in best_contacts.items():
            t_key = "contact_count" if e['rule']['type'] == 'count' else "contact_sum"

            # --- 核心修改开始 ---
            # 同样应用：重新计算严格截止时间
            strict_deadline = e['data']['start_t'] + timedelta(hours=e['rule']['hours'])

            msg = tmpl_mgr.render(t_key, bonus_name=r_name, reward=e['rule'].get('reward'), miss=e['miss'],
                                  deadline=str(strict_deadline).split('.')[0])
            # --- 核心修改结束 ---

            oids = [str(o['oid']) for o in e['data']['orders']]
            self.f_con_cards.add_card(r_name, msg, "warning", oids, template_text=msg)
        if has_contact:
            self.nb_cards.select(1)
        else:
            self.nb_cards.select(0)

    def open_single_check(self, event=None):
        try:
            if event:
                widget = event.widget
            else:
                tab_id = self.nb_list.select()
                tab_frame = self.nb_list.nametowidget(tab_id)
                widget = None
                for child in tab_frame.winfo_children():
                    if isinstance(child, ttk.Treeview):
                        widget = child
                        break

            if not widget: return
            sel = widget.selection()
            if not sel:
                messagebox.showinfo("Hint", "Select a player first")
                return

            pid = str(widget.item(sel[0], 'values')[1])
            current_data = self.results_cache.get(self.selected_days.get(), [])
            player_data = next((item for item in current_data if str(item['pid']) == pid), None)

            if player_data:
                SinglePlayerCheck(self.root, player_data, self.time_ext_val, self.limit_config_map,
                                  self.current_ref_time)
        except Exception as e:
            print(e)
            pass

    def open_template_editor(self):
        TemplateEditorWindow(self.root)

    def main_loop(self):
        self.root.mainloop()


if __name__ == "__main__": app = PromoAnalyzerApp(); app.main_loop()
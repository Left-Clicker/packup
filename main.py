import os
import re
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import timedelta, datetime
import calendar

# -------------------- 配置项（默认值，可在GUI中改） --------------------
DEFAULT_MID_FILTER = '16388'         # MID 预设（可逗号分隔多个）
DEFAULT_REST_THRESHOLD_HOURS = 6.0   # 休息判定阈值（> 此小时数算一次停止）
DEFAULT_REPORT_THRESHOLD_HOURS = 18.0 # 报告阈值（最长不间断 >= 此小时数则加黄底）

YELLOW_HL = '#FFF59D'   # 主结果高亮颜色（柔和黄）
GREEN_DAY = '#C8E6C9'   # 列筛选面板绿色提示底色（已选择）

# -------------------- 工具函数 --------------------
def normalize_col(col: str) -> str:
    if not isinstance(col, str):
        col = str(col)
    return (col.lower()
                .replace('（', '')
                .replace('）', '')
                .replace('(', '')
                .replace(')', '')
                .replace('_', '')
                .replace(' ', ''))

def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for c in df.columns:
        n = normalize_col(c)
        if 'playerid' in n or ('玩家' in n and 'id' in n):
            mapping[c] = 'player_id'
        elif 'playername' in n or '玩家昵称' in n or '玩家名' in n:
            mapping[c] = 'player_name'
        elif n == 'mid':
            mapping[c] = 'mid'
        elif 'createdat' in n or 'createat' in n or '时间' in n:
            mapping[c] = 'created_at'
        elif 'server' in n or '服务器id' in n or '服务器' in n:
            mapping[c] = 'server'
        else:
            mapping[c] = c
    df = df.rename(columns=mapping)
    return df

def parse_datetime(series: pd.Series) -> pd.Series:
    s = pd.to_datetime(series, errors='coerce')
    if s.notna().sum() < max(1, int(0.3 * len(s))):
        s = pd.to_datetime(series, errors='coerce', dayfirst=False)
    return s

def read_one_file(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        for enc in ('utf-8-sig', 'utf-8', 'gb18030'):
            try:
                return pd.read_csv(path, encoding=enc)
            except Exception:
                continue
        return pd.read_csv(path)
    elif ext in ('.xlsx', '.xls'):
        engine = 'openpyxl' if ext == '.xlsx' else None
        try:
            with pd.ExcelFile(path, engine=engine) as xls:
                sheet_names = xls.sheet_names
                target_raw = '队列行动明细'
                target = target_raw if target_raw in sheet_names else None
                if target is None:
                    def norm(s: str) -> str:
                        return str(s).strip().replace(' ', '').replace('\u3000', '')
                    norm_target = norm(target_raw)
                    for s in sheet_names:
                        if norm(s) == norm_target:
                            target = s
                            break
                if target is None and sheet_names:
                    target = sheet_names[0]
                return pd.read_excel(xls, sheet_name=target)
        except Exception:
            try:
                return pd.read_excel(path, sheet_name='队列行动明细', engine=engine)
            except Exception:
                try:
                    return pd.read_excel(path, sheet_name=0, engine=engine)
                except Exception:
                    return pd.DataFrame()
    else:
        return pd.DataFrame()

def parse_mid_list(mid_text: str) -> list:
    if mid_text is None:
        return []
    s = str(mid_text).replace('，', ',')
    return [t.strip() for t in s.split(',') if t.strip() != '']

def filter_by_mids(df: pd.DataFrame, mids: list) -> pd.DataFrame:
    if df.empty or 'mid' not in df.columns or not mids:
        return pd.DataFrame()
    mid_prefix = df['mid'].astype(str).str.split(':', n=1).str[0].str.strip()
    return df.loc[mid_prefix.isin(set(mids))].copy()

def analyze_player_sessions(group: pd.DataFrame, rest_hours: float, report_hours: float = None) -> dict:
    g = group.sort_values('created_at').copy()
    g = g.loc[g['created_at'].notna()]
    if g.empty:
        return {
            'stop_count': 0,
            'max_cont_hours': 0.0,
            'max_session_start': pd.NaT,
            'max_session_end': pd.NaT,
            'rows': len(group),
            'days_exceed': 0
        }
    diffs = g['created_at'].diff()
    rest_threshold = pd.Timedelta(hours=rest_hours)
    stops = diffs > rest_threshold
    session_id = stops.cumsum()
    g['session_id'] = session_id

    agg = g.groupby('session_id')['created_at'].agg(['min', 'max'])
    durations = (agg['max'] - agg['min'])
    if durations.empty:
        max_dur = pd.Timedelta(0)
        max_idx = None
    else:
        max_idx = durations.idxmax()
        max_dur = durations.loc[max_idx]

    max_start = agg.loc[max_idx, 'min'] if max_idx is not None else pd.NaT
    max_end = agg.loc[max_idx, 'max'] if max_idx is not None else pd.NaT
    max_hours = max_dur.total_seconds() / 3600.0 if pd.notna(max_dur) else 0.0
    stop_count = int(stops.fillna(False).sum())

    # 以自然日统计：每天内的“最长不间断时长”是否 >= 报告阈值
    days_exceed = 0
    if report_hours is not None and not agg.empty:
        threshold = pd.Timedelta(hours=report_hours)
        day_max = {}
        for _, row in agg.iterrows():
            s_start = row['min']
            s_end = row['max']
            if pd.isna(s_start) or pd.isna(s_end) or s_end <= s_start:
                continue
            day = pd.Timestamp(s_start.date())
            while day < s_end:
                day_next = day + pd.Timedelta(days=1)
                overlap_start = max(s_start, day)
                overlap_end = min(s_end, day_next)
                if overlap_end > overlap_start:
                    dur = overlap_end - overlap_start
                    dkey = overlap_start.date()
                    if dur > day_max.get(dkey, pd.Timedelta(0)):
                        day_max[dkey] = dur
                day = day_next
        days_exceed = sum(1 for _, v in day_max.items() if v >= threshold)

    return {
        'stop_count': stop_count,
        'max_cont_hours': max_hours,
        'max_session_start': max_start,
        'max_session_end': max_end,
        'rows': len(g),
        'days_exceed': days_exceed
    }

def analyze(df: pd.DataFrame, mids: list, rest_hours: float, report_hours: float) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    focus = filter_by_mids(df, mids)
    if focus.empty:
        return pd.DataFrame()

    if 'player_id' not in focus.columns or 'created_at' not in focus.columns:
        return pd.DataFrame()

    records = []
    def pick_one(s):
        return s.dropna().iloc[0] if s.dropna().size else None

    for pid, g in focus.groupby('player_id'):
        stats = analyze_player_sessions(g, rest_hours, report_hours)
        player_name = pick_one(g.get('player_name', pd.Series(dtype=object)))
        server = pick_one(g.get('server', pd.Series(dtype=object)))
        rec = {
            'player_id': pid,
            'player_name': player_name,
            'server': server,
            'stop_count': stats['stop_count'],
            'max_cont_hours': round(stats['max_cont_hours'], 2),
            'max_session_start': stats['max_session_start'],
            'max_session_end': stats['max_session_end'],
            'rows': stats['rows'],
            'days_exceed': int(stats.get('days_exceed', 0)),
        }
        records.append(rec)

    res = pd.DataFrame(records)
    if res.empty:
        return res

    res['flagged'] = res['max_cont_hours'] >= report_hours
    res = res.sort_values(['flagged', 'max_cont_hours', 'stop_count'],
                          ascending=[False, False, True]).reset_index(drop=True)
    return res

# -------------------- GUI --------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("资源商检索器")
        self.geometry("1200x680")

        self.style = ttk.Style(self)
        try:
            self.style.theme_use('clam')
        except Exception:
            pass
        self.style.map('Treeview',
                       background=[('selected', '#4a90e2')],
                       foreground=[('selected', 'white')])

        self.filepath = tk.StringVar()
        self.mid_filter = tk.StringVar(value=DEFAULT_MID_FILTER)
        self.rest_hours = tk.DoubleVar(value=DEFAULT_REST_THRESHOLD_HOURS)
        self.report_hours = tk.DoubleVar(value=DEFAULT_REPORT_THRESHOLD_HOURS)

        self.data_all = pd.DataFrame()
        self.data_focus = pd.DataFrame()
        self.result_df = pd.DataFrame()

        self.create_widgets()

    def create_widgets(self):
        frm_top = ttk.Frame(self)
        frm_top.pack(fill='x', padx=8, pady=8)

        ttk.Label(frm_top, text="数据文件:").pack(side='left')
        ttk.Entry(frm_top, textvariable=self.filepath, width=60).pack(side='left', padx=4)
        ttk.Button(frm_top, text="浏览...", command=self.browse_file).pack(side='left', padx=4)
        ttk.Button(frm_top, text="读取 + 分析", command=self.load_and_analyze).pack(side='left', padx=8)

        frm_params = ttk.Frame(self)
        frm_params.pack(fill='x', padx=8, pady=4)

        ttk.Label(frm_params, text="MID筛选(逗号分隔):").pack(side='left')
        ttk.Entry(frm_params, textvariable=self.mid_filter, width=24).pack(side='left', padx=4)

        ttk.Label(frm_params, text="休息阈值(小时，>此值算停止):").pack(side='left', padx=(16,0))
        ttk.Entry(frm_params, textvariable=self.rest_hours, width=6).pack(side='left', padx=4)

        ttk.Label(frm_params, text="报告阈值(小时，>=此值黄底):").pack(side='left', padx=(16,0))
        ttk.Entry(frm_params, textvariable=self.report_hours, width=6).pack(side='left', padx=4)

        ttk.Button(frm_params, text="导出结果CSV", command=self.export_results).pack(side='right', padx=4)

        cols = ('player_id', 'player_name', 'server', 'stop_count',
                'max_cont_hours', 'days_exceed',
                'max_session_start', 'max_session_end', 'rows')
        self.tree = ttk.Treeview(self, columns=cols, show='headings', selectmode='extended')
        headings = {
            'player_id': '玩家ID',
            'player_name': '玩家昵称',
            'server': '服务器ID',
            'stop_count': '停止次数(>阈值)',
            'max_cont_hours': '最长不间断(小时)',
            'days_exceed': '超阈天数(天)',
            'max_session_start': '最长会话开始',
            'max_session_end': '最长会话结束',
            'rows': '记录数'
        }
        for c in cols:
            self.tree.heading(c, text=headings.get(c, c))
            width = 120
            if c in ('player_name',):
                width = 180
            if c in ('max_session_start', 'max_session_end'):
                width = 160
            self.tree.column(c, width=width, anchor='center')

        self.tree.tag_configure('flagged', background=YELLOW_HL)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tree.pack(fill='both', expand=True, padx=8, pady=(4,0))
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        self.tree.bind('<Double-1>', self.on_double_click_row)

        frm_bottom = ttk.Frame(self)
        frm_bottom.pack(fill='x', padx=8, pady=8)
        ttk.Button(frm_bottom, text="查看选中玩家原始数据", command=self.view_selected_raw).pack(side='left')
        ttk.Button(frm_bottom, text="复制标黄玩家ID", command=self.copy_flagged_ids).pack(side='right')

    def browse_file(self):
        f = filedialog.askopenfilename(
            title="选择CSV或XLSX文件",
            filetypes=[("数据文件", "*.csv *.xlsx *.xls"), ("CSV 文件", "*.csv"), ("Excel 文件", "*.xlsx *.xls")]
        )
        if f:
            self.filepath.set(f)

    def load_and_analyze(self):
        path = self.filepath.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("提示", "请选择有效的文件。")
            return
        try:
            df = read_one_file(path)
            if df.empty:
                messagebox.showinfo("提示", "未读取到有效数据。请检查文件是否为CSV/XLSX，且包含列：mid、created_at、玩家ID。")
                return
            df = standardize_columns(df)
            if 'created_at' in df.columns:
                df['created_at'] = parse_datetime(df['created_at'])
            self.data_all = df

            mids = parse_mid_list(self.mid_filter.get().strip())
            if not mids:
                mids = parse_mid_list(DEFAULT_MID_FILTER)

            self.data_focus = filter_by_mids(self.data_all, mids)
            if self.data_focus.empty:
                messagebox.showinfo("提示", f"数据中没有所选 MID 的记录：{', '.join(mids)}")
            self.result_df = analyze(self.data_all, mids, self.rest_hours.get(), self.report_hours.get())
            self.refresh_tree()
            total_players = len(self.result_df)
            flagged_count = int(self.result_df['flagged'].sum()) if 'flagged' in self.result_df.columns else 0
            messagebox.showinfo("完成", f"分析完成。玩家总数：{total_players}，达标(黄底)：{flagged_count}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"分析失败：\n{e}")

    def refresh_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if self.result_df.empty:
            return
        for _, row in self.result_df.iterrows():
            vals = [row.get(c, '') for c in self.tree['columns']]
            vals_fmt = []
            for c, v in zip(self.tree['columns'], vals):
                if isinstance(v, pd.Timestamp):
                    v = '' if pd.isna(v) else v.strftime('%Y-%m-%d %H:%M:%S')
                try:
                    is_nan = pd.isna(v) if hasattr(v, '__float__') else False
                except Exception:
                    is_nan = False
                vals_fmt.append('' if is_nan else v)
            tags = ('flagged',) if bool(row.get('flagged', False)) else ()
            self.tree.insert('', 'end', values=vals_fmt, tags=tags)

    def get_selected_player_ids(self):
        sels = self.tree.selection()
        if not sels:
            return []
        pids = []
        for iid in sels:
            vals = self.tree.item(iid, 'values')
            if vals:
                pids.append(vals[0])
        pids = [str(p).strip() for p in pids if str(p).strip() != '']
        return list(dict.fromkeys(pids))

    def on_double_click_row(self, event=None):
        self.view_selected_raw()

    def view_selected_raw(self):
        pids = self.get_selected_player_ids()
        if not pids:
            messagebox.showinfo("提示", "请先在结果列表中选中一位或多位玩家。")
            return
        if self.data_focus.empty:
            messagebox.showinfo("提示", "没有可展示的原始数据。")
            return
        df = self.data_focus
        if 'player_id' not in df.columns:
            messagebox.showinfo("提示", "缺少 player_id 列，无法筛选。")
            return
        sub = df.loc[df['player_id'].astype(str).isin(set(map(str, pids)))].copy()
        if sub.empty:
            messagebox.showinfo("提示", f"未找到所选玩家的原始数据。")
            return
        if 'created_at' in sub.columns:
            sub = sub.sort_values('created_at')
        mids = parse_mid_list(self.mid_filter.get().strip())
        title_ids = ', '.join(pids[:5]) + ('...' if len(pids) > 5 else '')
        RawViewer(self, sub, title=f"玩家({len(pids)}): {title_ids} - 原始数据 (MID筛选: {', '.join(mids) if mids else DEFAULT_MID_FILTER})")

    def copy_flagged_ids(self):
        if self.result_df.empty or 'flagged' not in self.result_df.columns:
            messagebox.showinfo("提示", "没有可复制的ID。")
            return
        ids_series = self.result_df.loc[self.result_df['flagged'], 'player_id'].dropna()

        def to_clean_str(x):
            try:
                if isinstance(x, float):
                    if x.is_integer():
                        return str(int(x))
                    return str(int(round(x)))
                return str(x)
            except Exception:
                return str(x)

        ids = [to_clean_str(v) for v in ids_series.tolist() if str(v).strip() != '']
        if not ids:
            messagebox.showinfo("提示", "当前没有标黄的玩家。")
            return
        text = '\n'.join(ids)
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.update()
            messagebox.showinfo("完成", f"已复制 {len(ids)} 个玩家ID到剪贴板。")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"复制失败：\n{e}")

    def export_results(self):
        if self.result_df.empty:
            messagebox.showinfo("提示", "没有可导出的结果。")
            return
        f = filedialog.asksaveasfilename(
            title="保存结果为CSV",
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv")]
        )
        if not f:
            return
        try:
            out = self.result_df.copy()
            for c in ('max_session_start', 'max_session_end'):
                if c in out.columns and pd.api.types.is_datetime64_any_dtype(out[c]):
                    out[c] = out[c].dt.strftime('%Y-%m-%d %H:%M:%S')
            out.to_csv(f, index=False, encoding='utf-8-sig')
            messagebox.showinfo("完成", f"已导出：{f}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"导出失败：\n{e}")

# -------------------- 原始数据窗口 --------------------
class RawViewer(tk.Toplevel):
    def __init__(self, master, df: pd.DataFrame, title="原始数据"):
        super().__init__(master)
        self.title(title)
        self.geometry("1200x760")

        self.df_original = df.copy()
        self.df_current = self.df_original.copy()
        # 普通列: set[str]
        # created_at: {'type': 'range', 'start': Timestamp, 'end': Timestamp}
        self.col_filters = {}

        cols = list(self.df_original.columns)
        self.tree = ttk.Treeview(self, columns=cols, show='headings', selectmode='extended')
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140, anchor='center')

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tree.pack(fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        # 点击列头打开筛选
        self.tree.bind('<Button-1>', self.on_tree_click_heading)

        self.fill_tree(self.df_current)

        frm = ttk.Frame(self)
        frm.pack(fill='x', padx=8, pady=8)
        ttk.Button(frm, text="导出当前显示数据为CSV", command=self.export_current).pack(side='left')
        ttk.Button(frm, text="清空所有筛选", command=self.clear_all_filters).pack(side='left', padx=8)

        ttk.Label(self, text="提示：点击列头可打开筛选面板；按Ctrl/Shift可多选行。").pack(fill='x', padx=8, pady=(0,8))

    def fill_tree(self, df: pd.DataFrame):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if list(self.tree['columns']) != list(df.columns):
            self.tree['columns'] = list(df.columns)
            for c in df.columns:
                self.tree.heading(c, text=c)
                self.tree.column(c, width=140, anchor='center')
        for _, row in df.iterrows():
            vals = []
            for c in df.columns:
                v = row.get(c, '')
                if isinstance(v, pd.Timestamp):
                    v = '' if pd.isna(v) else v.strftime('%Y-%m-%d %H:%M:%S')
                try:
                    is_nan = pd.isna(v) if hasattr(v, '__float__') else False
                except Exception:
                    is_nan = False
                vals.append('' if is_nan else v)
            self.tree.insert('', 'end', values=vals)

    def apply_filters_and_refresh(self):
        df = self.df_original
        mask_all = pd.Series(True, index=df.index)

        for col, selected in self.col_filters.items():
            if selected is None or col not in df.columns:
                continue
            if col == 'created_at' and isinstance(selected, dict) and selected.get('type') == 'range':
                ser = pd.to_datetime(df[col], errors='coerce')
                mins = ser.dt.floor('min')
                start = selected.get('start')
                end = selected.get('end')
                mask = (mins >= start) & (mins <= end)
            else:
                ser = df[col].astype(str).str.strip()
                mask = ser.isin({str(x).strip() for x in selected})
            mask_all = mask_all & mask

        self.df_current = df.loc[mask_all].copy()
        self.fill_tree(self.df_current)

    def clear_all_filters(self):
        self.col_filters.clear()
        self.df_current = self.df_original.copy()
        self.fill_tree(self.df_current)

    def on_tree_click_heading(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != 'heading':
            return
        col_id = self.tree.identify_column(event.x)
        try:
            col_idx = int(col_id.replace('#', '')) - 1
        except Exception:
            return
        columns = list(self.tree['columns'])
        if col_idx < 0 or col_idx >= len(columns):
            return
        col_name = columns[col_idx]
        df_base = self.df_current.copy()
        x = self.tree.winfo_rootx() + event.x
        y = self.tree.winfo_rooty() + event.y + 24
        FilterPopup(self, column=col_name, df=df_base,
                    pre_selected=self.col_filters.get(col_name),
                    on_ok=self.on_filter_ok,
                    geometry=f"+{x}+{y}")

    def on_filter_ok(self, column, selected_payload):
        # 普通列: set[str] 或 空集合；None 取消该列过滤
        # created_at: {'type':'range','start':ts,'end':ts}；None 取消该列过滤
        if selected_payload is None:
            if column in self.col_filters:
                del self.col_filters[column]
        else:
            self.col_filters[column] = selected_payload
        self.apply_filters_and_refresh()

    def export_current(self):
        df_to_save = self.df_current.copy()
        if 'created_at' in df_to_save.columns and pd.api.types.is_datetime64_any_dtype(df_to_save['created_at']):
            df_to_save['created_at'] = df_to_save['created_at'].dt.strftime('%Y-%m-%d %H:%M:%S')
        f = filedialog.asksaveasfilename(
            title="保存当前显示数据为CSV",
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv")]
        )
        if not f:
            return
        try:
            df_to_save.to_csv(f, index=False, encoding='utf-8-sig')
            messagebox.showinfo("完成", f"已导出：{f}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"导出失败：\n{e}")

# -------------------- 日期时间选择器 --------------------
class DateTimePicker(ttk.Frame):
    def __init__(self, master, years, init_ts=None):
        super().__init__(master)
        self.var_year = tk.IntVar()
        self.var_month = tk.IntVar()
        self.var_day = tk.IntVar()
        self.var_hour = tk.IntVar()
        self.var_min = tk.IntVar()

        years = sorted(set(years)) if years else [datetime.now().year]
        months = list(range(1,13))
        hours = list(range(0,24))
        mins = list(range(0,60))

        ttk.Label(self, text="年").grid(row=0, column=0, padx=(0,2))
        self.cmb_year = ttk.Combobox(self, width=5, state='readonly', values=years, textvariable=self.var_year)
        self.cmb_year.grid(row=0, column=1, padx=(0,8))

        ttk.Label(self, text="月").grid(row=0, column=2, padx=(0,2))
        self.cmb_month = ttk.Combobox(self, width=3, state='readonly', values=months, textvariable=self.var_month)
        self.cmb_month.grid(row=0, column=3, padx=(0,8))

        ttk.Label(self, text="日").grid(row=0, column=4, padx=(0,2))
        self.cmb_day = ttk.Combobox(self, width=3, state='readonly', values=[1], textvariable=self.var_day)
        self.cmb_day.grid(row=0, column=5, padx=(0,8))

        ttk.Label(self, text="时").grid(row=0, column=6, padx=(0,2))
        self.cmb_hour = ttk.Combobox(self, width=3, state='readonly', values=hours, textvariable=self.var_hour)
        self.cmb_hour.grid(row=0, column=7, padx=(0,8))

        ttk.Label(self, text="分").grid(row=0, column=8, padx=(0,2))
        self.cmb_min = ttk.Combobox(self, width=3, state='readonly', values=mins, textvariable=self.var_min)
        self.cmb_min.grid(row=0, column=9)

        self.cmb_year.bind('<<ComboboxSelected>>', self._update_days)
        self.cmb_month.bind('<<ComboboxSelected>>', self._update_days)

        ts = init_ts or datetime.now().replace(second=0, microsecond=0)
        self.var_year.set(ts.year)
        self.var_month.set(ts.month)
        self._update_days()
        self.var_day.set(min(ts.day, int(self.cmb_day['values'][-1])))
        self.var_hour.set(ts.hour)
        self.var_min.set(ts.minute)

    def _update_days(self, event=None):
        y = self.var_year.get() or datetime.now().year
        m = self.var_month.get() or 1
        last = calendar.monthrange(y, m)[1]
        vals = list(range(1, last+1))
        self.cmb_day['values'] = vals
        if self.var_day.get() not in vals:
            self.var_day.set(1)

    def get_timestamp(self) -> datetime:
        return datetime(self.var_year.get(), self.var_month.get(), self.var_day.get(),
                        self.var_hour.get(), self.var_min.get())

# -------------------- 列筛选弹窗 --------------------
class FilterPopup(tk.Toplevel):
    """
    普通列：每次打开显示“原始数据全量取值”，当前已应用的取值以绿色底标记；
           若未选择任何项点击“确定”，则不改变该列筛选；提供“取消本列筛选”按钮。
    created_at：仅“时间区间”参与过滤；分钟勾选仅用于绿色标记。
    """
    def __init__(self, master: RawViewer, column: str, df: pd.DataFrame, pre_selected=None, on_ok=None, geometry=None):
        super().__init__(master)
        self.title(f"筛选 - {column}")
        self.resizable(True, True)
        if geometry:
            self.geometry(geometry)
        self.transient(master)
        self.column = column
        self.df_view = df.copy()          # 当前视图（不用于普通列取值候选）
        self.df_all = master.df_original  # 原始全量，用于普通列取值候选
        self.on_ok_cb = on_ok

        if column == 'created_at':
            self.pre_range = pre_selected if isinstance(pre_selected, dict) and pre_selected.get('type') == 'range' else None
            self._build_datetime_ui()
        else:
            self.pre_selected_values = set(pre_selected) if isinstance(pre_selected, set) else None
            self._build_values_ui()

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    # ---------- 普通列 UI ----------
    def _build_values_ui(self):
        frm_top = ttk.Frame(self)
        frm_top.pack(fill='x', padx=8, pady=6)
        ttk.Label(frm_top, text="搜索：").pack(side='left')
        self.ent_search = ttk.Entry(frm_top)
        self.ent_search.pack(side='left', fill='x', expand=True, padx=4)
        ttk.Button(frm_top, text="查找", command=self._apply_search_values).pack(side='left', padx=4)
        ttk.Button(frm_top, text="重置", command=self._reset_search_values).pack(side='left', padx=4)

        frm_mid = ttk.Frame(self)
        frm_mid.pack(fill='both', expand=True, padx=8, pady=4)
        self.lst = tk.Listbox(frm_mid, selectmode='extended', exportselection=False)
        self.lst.pack(side='left', fill='both', expand=True)
        self._list_default_bg = self.lst.cget('bg')
        sb = ttk.Scrollbar(frm_mid, orient='vertical', command=self.lst.yview)
        sb.pack(side='left', fill='y')
        self.lst.configure(yscrollcommand=sb.set)

        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill='x', padx=8, pady=8)
        ttk.Button(frm_btn, text="全选", command=self._select_all_values).pack(side='left')
        ttk.Button(frm_btn, text="清空选择", command=self._clear_select_values).pack(side='left', padx=6)
        ttk.Button(frm_btn, text="取消本列筛选", command=self._clear_filter_values).pack(side='left', padx=6)
        ttk.Button(frm_btn, text="确定", command=self.on_ok_values).pack(side='right')
        ttk.Button(frm_btn, text="取消", command=self.on_cancel).pack(side='right', padx=6)

        # 取值候选来自“原始全量数据”，而非当前视图
        ser = self.df_all[self.column].astype(str).str.strip().fillna('') if self.column in self.df_all.columns else pd.Series([], dtype=str)
        self._values_all = sorted(set(ser.tolist()))
        self._render_values_list(self._values_all)

    def _render_values_list(self, values):
        self.lst.delete(0, 'end')
        for v in values:
            self.lst.insert('end', v)
        # 绿色底标注已应用的取值
        pre = self.pre_selected_values or set()
        for i, v in enumerate(values):
            if v in pre:
                try:
                    self.lst.itemconfig(i, bg=GREEN_DAY)
                except Exception:
                    pass

    def _apply_search_values(self):
        q = self.ent_search.get().strip().lower()
        if not q:
            self._render_values_list(self._values_all)
            return
        filt = [v for v in self._values_all if q in v.lower()]
        self._render_values_list(filt)

    def _reset_search_values(self):
        self.ent_search.delete(0, 'end')
        self._render_values_list(self._values_all)

    def _select_all_values(self):
        self.lst.selection_set(0, 'end')

    def _clear_select_values(self):
        self.lst.selection_clear(0, 'end')

    def _clear_filter_values(self):
        # 明确取消该列筛选
        if self.on_ok_cb:
            self.on_ok_cb(self.column, None)
        self.destroy()

    def on_ok_values(self):
        sel_idx = self.lst.curselection()
        if sel_idx:
            selected = set(self.lst.get(i) for i in sel_idx)
            if self.on_ok_cb:
                self.on_ok_cb(self.column, selected)
        else:
            # 未选择任何项：不改变外部已有筛选
            if self.on_ok_cb:
                if self.pre_selected_values is None:
                    # 原先无筛选，仍保持无变化
                    self.on_ok_cb(self.column, None)
                else:
                    self.on_ok_cb(self.column, set(self.pre_selected_values))
        self.destroy()

    # ---------- created_at（区间 + 日期/分钟：勾选仅用于标记） ----------
    def _build_datetime_ui(self):
        ser = pd.to_datetime(self.df_view.get('created_at'), errors='coerce')
        mins_all = ser.dt.floor('min').dropna().sort_values()
        self._minutes_all = [pd.Timestamp(t) for t in mins_all.unique()]
        self.selected_minutes = set()  # 仅用于绿色标记

        years = sorted({t.year for t in self._minutes_all}) or [datetime.now().year]
        init_start = self._minutes_all[0] if self._minutes_all else datetime.now().replace(second=0, microsecond=0)
        init_end = self._minutes_all[-1] if self._minutes_all else init_start

        frm_top = ttk.LabelFrame(self, text="时间区间")
        frm_top.pack(fill='x', padx=8, pady=6)

        ttk.Label(frm_top, text="起始").grid(row=0, column=0, sticky='w', padx=(6,4), pady=4)
        self.dtp_start = DateTimePicker(frm_top, years=years, init_ts=(self.pre_range['start'].to_pydatetime() if self.pre_range else (init_start.to_pydatetime() if isinstance(init_start, pd.Timestamp) else init_start)))
        self.dtp_start.grid(row=0, column=1, sticky='w', pady=4)

        ttk.Label(frm_top, text="结束").grid(row=1, column=0, sticky='w', padx=(6,4), pady=4)
        self.dtp_end = DateTimePicker(frm_top, years=years, init_ts=(self.pre_range['end'].to_pydatetime() if self.pre_range else (init_end.to_pydatetime() if isinstance(init_end, pd.Timestamp) else init_end)))
        self.dtp_end.grid(row=1, column=1, sticky='w', pady=4)

        btns = ttk.Frame(frm_top)
        btns.grid(row=0, column=2, rowspan=2, padx=10, pady=4, sticky='nsew')
        ttk.Button(btns, text="应用区间", command=self._apply_range).pack(fill='x', pady=(0,6))
        ttk.Button(btns, text="清空区间", command=self._clear_range).pack(fill='x')

        frm_mid = ttk.Frame(self)
        frm_mid.pack(fill='both', expand=True, padx=8, pady=4)

        lf = ttk.LabelFrame(frm_mid, text='日期')
        lf.pack(side='left', fill='both', expand=False, padx=(0,6))
        self.lst_days = tk.Listbox(lf, selectmode='extended', width=14, exportselection=False)
        self.lst_days.pack(side='left', fill='both', expand=True)
        self._day_default_bg = self.lst_days.cget('bg')
        sb_day = ttk.Scrollbar(lf, orient='vertical', command=self.lst_days.yview)
        sb_day.pack(side='left', fill='y')
        self.lst_days.configure(yscrollcommand=sb_day.set)
        self.lst_days.bind('<<ListboxSelect>>', self._on_day_select)

        rf = ttk.LabelFrame(frm_mid, text='分钟（勾选仅用于标记，不影响过滤）')
        rf.pack(side='left', fill='both', expand=True)
        self.lst_minutes = tk.Listbox(rf, selectmode='extended', exportselection=False)
        self.lst_minutes.pack(side='left', fill='both', expand=True)
        sb_min = ttk.Scrollbar(rf, orient='vertical', command=self.lst_minutes.yview)
        sb_min.pack(side='left', fill='y')
        self.lst_minutes.configure(yscrollcommand=sb_min.set)
        self.lst_minutes.bind('<<ListboxSelect>>', self._on_minutes_select)

        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill='x', padx=8, pady=8)
        ttk.Button(frm_btn, text="全选当前显示分钟", command=self._select_all_minutes_current).pack(side='left')
        ttk.Button(frm_btn, text="清空当前显示分钟", command=self._clear_minutes_current).pack(side='left', padx=6)
        ttk.Button(frm_btn, text="确定", command=self.on_ok_datetime).pack(side='right')
        ttk.Button(frm_btn, text="取消", command=self.on_cancel).pack(side='right', padx=6)

        self._build_index_all()
        self._date_to_minutes_view = dict(self._date_to_minutes_all)
        self._range_applied = False
        self._range_cleared = False

        if self.pre_range:
            self._apply_range_to_view(self.pre_range['start'], self.pre_range['end'])
            self._range_applied = True

        self._update_days_list()
        if self.lst_days.size() > 0:
            self.lst_days.selection_set(0)
            self._on_day_select()

    def _build_index_all(self):
        self._date_to_minutes_all = {}
        for t in self._minutes_all:
            d = t.date().isoformat()
            self._date_to_minutes_all.setdefault(d, []).append(t)
        for d in self._date_to_minutes_all:
            self._date_to_minutes_all[d] = sorted(self._date_to_minutes_all[d])

    def _apply_range_to_view(self, start, end):
        minutes_in = [t for t in self._minutes_all if (t >= start and t <= end)]
        self._date_to_minutes_view = {}
        for t in minutes_in:
            d = t.date().isoformat()
            self._date_to_minutes_view.setdefault(d, []).append(t)
        for d in list(self._date_to_minutes_view.keys()):
            self._date_to_minutes_view[d] = sorted(self._date_to_minutes_view[d])

    def _apply_range(self):
        start = pd.Timestamp(self.dtp_start.get_timestamp())
        end = pd.Timestamp(self.dtp_end.get_timestamp())
        if end < start:
            messagebox.showwarning("提示", "结束时间不能早于起始时间。")
            return
        self._apply_range_to_view(start, end)
        self._range_applied = True
        self._range_cleared = False
        self._update_days_list()
        if self.lst_days.size() > 0:
            self.lst_days.selection_set(0)
            self._on_day_select()
        else:
            self._render_minutes_list([])

    def _clear_range(self):
        # 1) 还原为全量可见
        self._date_to_minutes_view = dict(self._date_to_minutes_all)


        # 3) 刷新日期列表
        self._update_days_list()

        # 清空分钟区（不展示任何分钟，等待用户重新选择日期）
        self._render_minutes_list([])

        # 5) 刷新日期高亮（由于已无选择，全部恢复为默认底色）
        self._refresh_day_highlight()

    def _update_days_list(self):
        days_all = sorted(self._date_to_minutes_view.keys())
        self.lst_days.delete(0, 'end')
        for d in days_all:
            self.lst_days.insert('end', d)
        self._refresh_day_highlight()

    def _minutes_for_days(self, days):
        minutes = []
        for d in days:
            minutes.extend(self._date_to_minutes_view.get(d, []))
        return sorted(set(minutes))

    def _on_day_select(self, event=None):
        sel_idx = self.lst_days.curselection()
        sel_days = [self.lst_days.get(i) for i in sel_idx]
        mins_sorted = self._minutes_for_days(sel_days)
        self._render_minutes_list(mins_sorted)

    def _render_minutes_list(self, mins_list):
        self.lst_minutes.delete(0, 'end')
        self._mins_shown = mins_list
        for t in mins_list:
            self.lst_minutes.insert('end', t.strftime('%Y-%m-%d %H:%M'))
        if self.selected_minutes:
            want = self.selected_minutes
            for i, t in enumerate(self._mins_shown):
                if t in want:
                    self.lst_minutes.selection_set(i)

    def _on_minutes_select(self, event=None):
        current_sel_idx = set(self.lst_minutes.curselection())
        current_sel_minutes = {self._mins_shown[i] for i in current_sel_idx} if self._mins_shown else set()
        for t in self._mins_shown:
            self.selected_minutes.discard(t)
        self.selected_minutes |= current_sel_minutes
        self._refresh_day_highlight()

    def _select_all_minutes_current(self):
        self.lst_minutes.selection_set(0, 'end')
        self._on_minutes_select()

    def _clear_minutes_current(self):
        self.lst_minutes.selection_clear(0, 'end')
        self._on_minutes_select()

    def _refresh_day_highlight(self):
        for i in range(self.lst_days.size()):
            d = self.lst_days.get(i)
            arr = self._date_to_minutes_view.get(d, [])
            has = any((t in self.selected_minutes) for t in arr)
            try:
                self.lst_days.itemconfig(i, bg=(GREEN_DAY if has else self._day_default_bg))
            except Exception:
                pass

    def on_ok_datetime(self):
        if self.on_ok_cb:
            if self._range_applied:
                start = pd.Timestamp(self.dtp_start.get_timestamp()).floor('min')
                end = pd.Timestamp(self.dtp_end.get_timestamp()).floor('min')
                payload = {'type': 'range', 'start': start, 'end': end}
                self.on_ok_cb('created_at', payload)
            elif self._range_cleared:
                self.on_ok_cb('created_at', None)
            else:
                # 未应用也未清空：不改变外部过滤
                pass
        self.destroy()

    def on_cancel(self):
        self.destroy()

# -------------------- 入口 --------------------
if __name__ == '__main__':
    app = App()
    app.mainloop()

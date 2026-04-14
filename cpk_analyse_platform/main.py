"""
main.py
-------
产线数据分析AI平台 – Zillnk Efficiency Improvement Group
Entry point: launches the tkinter GUI.
"""

import os
import sys
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

# Ensure the project root is on sys.path so 'core' package is importable
_ROOT = os.path.dirname(os.path.abspath(__file__))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.data_extractor import read_barcodes, run_extraction, generate_missing_report
from core.cpk_calculator import analyze_xlsx_folder
from core.html_report import generate_report


# ============================================================================
# Helpers
# ============================================================================

def _ts() -> str:
    return datetime.now().strftime('%H:%M:%S')


_HELP_TEXT = """\
产线数据分析AI平台  –  使用说明
=====================================

【功能一：本地数据分析】

1. 发货 Excel 文件（仅"最后一次pass数据"模式必填）
   选择包含序列号的发货清单 Excel 文件（第一列=序列号，第二列=产品编码）。
   其他分析模式无需 Excel，程序将自动扫描工站目录发现条码。

2. 输出目录
   选择分析结果的存放目录。程序将自动创建以下内容：
     <输出目录>/
       <工站类型>/xlsx/          ← 提取出的测试 xlsx 文件（保留原文件名）
       <工站类型>/json/          ← 对应的测试 json 文件（保留原文件名）
       missing_barcodes.xlsx    ← 缺失条码汇总报表
       cpk_report.html          ← 测试数据分析 HTML 报告
       fault_database.db        ← 故障分析定位关系库（SQLite）
       analysis_log_<时间戳>.txt ← 本次运行完整过程日志

3. 测试工站配置
   为每类测试工站填写：
     · 工站类型标签（如 FT1、FT2、Aging …）
     · 测试数据文件夹路径（该工站的数据根目录）
   点击 [+ 添加工站] 可增加配置行。
   在工站类型或路径输入框内，按 ↑ / ↓ 可快速在行间切换焦点。

   工站合并配置（默认隐藏，点击展开）：
     将指定工站类型的数据合并到目标工站类型中分析。
     例如：FT2 → FT1，则 FT2 数据归入 FT1 统一处理。

4. 分析模式说明
   · 最后一次pass数据     — 需要发货Excel，取每条码最新一次成功记录
   · 所选文件夹分析       — 无需Excel，直接对所配置文件夹下的文件做CPK
   · 全部成功数据         — 无需Excel，收集全部成功记录（含同一模块多次）做CPK
   · 所有数据（含失败）   — 无需Excel，含失败全量CPK，自动建立故障分析库
   · 仅失败数据           — 无需Excel，仅失败记录，自动更新故障分析库

5. 分析文件类型
   · xlsx — 读取提取目录下的 .xlsx 文件做 CPK 分析
   · json — 读取提取目录下的 .json 文件做 CPK 分析（默认 xlsx）

6. 开始分析 / 停止分析
   点击"开始分析"后按键变为"停止分析"，可随时中止当前运行。

7. HTML 报告说明
   · 在搜索框输入 "cpk"（不区分大小写）可展示 Cp/Cpl/Cpu/Cpk 列及产品数据检索Tab
   · 含失败数据时，直方图蓝色=通过，红色=失败，统计表显示通过率

【功能二 / 三】深科技 / 立讯 MES 数据分析
   功能待实现，敬请期待。

如有问题，请联系 Zillnk Efficiency Improvement Group。
"""


# ============================================================================
# Tooltip helper
# ============================================================================

_MODE_HINTS = {
    '最后一次pass数据':   '⚠ 需要发货Excel\n仅取每个条码最近一次成功测试记录做CPK',
    '所选文件夹分析':     '✓ 无需Excel\n直接对所配置文件夹下的xlsx/json文件做CPK\n适用于已手动整理好的数据文件夹',
    '全部成功数据':       '✓ 无需Excel，自动扫描\n收集全部工站测试成功记录（含同一模块多次）做CPK',
    '所有数据（含失败）': '✓ 无需Excel，自动扫描\n成功+失败全量CPK，自动建立故障分析库\n重点：跨站比对（同一模块在不同设备的测试数据关联分析）',
    '仅失败数据':         '✓ 无需Excel，自动扫描\n仅失败记录CPK，自动更新故障分析库\n重点：失败规律总结，高频失败项/设备错误模式挖掘',
}


class _ToolTip:
    """Lightweight balloon tooltip for any tkinter widget."""

    def __init__(self, widget, text_fn):
        self._widget = widget
        self._text_fn = text_fn   # callable() → str
        self._tip = None
        widget.bind('<Enter>', self._show)
        widget.bind('<Leave>', self._hide)
        widget.bind('<ButtonPress>', self._hide)

    def _show(self, _e=None):
        text = self._text_fn()
        if not text:
            return
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._tip = tk.Toplevel(self._widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f'+{x}+{y}')
        tk.Label(self._tip, text=text, justify='left',
                 bg='#fffde7', fg='#333', relief='solid', bd=1,
                 font=('Segoe UI', 8), padx=6, pady=4).pack()

    def _hide(self, _e=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None


# ============================================================================
# Product name inference for report title
# ============================================================================

def _infer_product_name(station_configs: list) -> str:
    from pathlib import Path
    names = []
    for cfg in station_configs:
        folder = cfg.get('folder', '').strip()
        if not folder:
            continue
        parts = Path(folder).parts
        for i, part in enumerate(parts):
            if part.lower() in ('testresult', 'test_result', 'testresults',
                                'testdata', 'test_data'):
                if i + 1 < len(parts):
                    names.append(parts[i + 1])
                break
    if names:
        return names[0]
    folders = [cfg.get('folder', '') for cfg in station_configs if cfg.get('folder')]
    if folders:
        try:
            from os.path import commonpath
            common = commonpath(folders)
            return Path(common).name
        except Exception:
            pass
    return ''


# ============================================================================
# Station row widget
# ============================================================================

class StationRow:
    """One row in the station config table: [type entry] [folder entry] [Browse] [Delete]"""

    def __init__(self, parent_frame, delete_callback):
        self._del_cb = delete_callback
        self.frame = tk.Frame(parent_frame, bg='#f5f6fa')
        self.var_type = tk.StringVar()
        self.var_folder = tk.StringVar()

        self._type_entry = tk.Entry(self.frame, textvariable=self.var_type, width=10,
                                    font=('Segoe UI', 9))
        self._type_entry.pack(side='left', padx=(0, 3))

        self._folder_entry = tk.Entry(self.frame, textvariable=self.var_folder, width=28,
                                      font=('Segoe UI', 9))
        self._folder_entry.pack(side='left', padx=(0, 3), fill='x', expand=True)

        tk.Button(self.frame, text='浏览', font=('Segoe UI', 8),
                  command=self._browse, relief='flat', bg='#e0e4f0',
                  padx=5).pack(side='left', padx=(0, 3))

        tk.Button(self.frame, text='✕', font=('Segoe UI', 8, 'bold'), fg='#c62828',
                  command=self._on_delete,
                  relief='flat', bg='#fce4e4', padx=5).pack(side='left')

    def _browse(self):
        d = filedialog.askdirectory(title='选择测试数据文件夹')
        if d:
            self.var_folder.set(d)

    def _on_delete(self):
        self._del_cb(self)

    def pack(self, **kw):
        self.frame.pack(**kw)

    def destroy(self):
        self.frame.destroy()

    def get(self) -> dict:
        return {'type': self.var_type.get().strip(),
                'folder': self.var_folder.get().strip()}


# ============================================================================
# Merge rule row widget
# ============================================================================

class MergeRuleRow:
    """One station merge rule: [source entry] → 合并到 → [target entry] [Delete]"""

    def __init__(self, parent_frame, delete_callback):
        self._del_cb = delete_callback
        self.frame = tk.Frame(parent_frame, bg='#f9f9f9')
        self.var_src = tk.StringVar()
        self.var_dst = tk.StringVar()

        self._src_entry = tk.Entry(self.frame, textvariable=self.var_src, width=10,
                                   font=('Segoe UI', 9))
        self._src_entry.pack(side='left', padx=(0, 4))

        tk.Label(self.frame, text='→ 合并到', bg='#f9f9f9',
                 font=('Segoe UI', 8), fg='#555').pack(side='left', padx=(0, 4))

        self._dst_entry = tk.Entry(self.frame, textvariable=self.var_dst, width=10,
                                   font=('Segoe UI', 9))
        self._dst_entry.pack(side='left', padx=(0, 4))

        tk.Button(self.frame, text='✕', font=('Segoe UI', 8, 'bold'), fg='#c62828',
                  command=lambda: self._del_cb(self),
                  relief='flat', bg='#fce4e4', padx=5).pack(side='left')

    def pack(self, **kw):
        self.frame.pack(**kw)

    def destroy(self):
        self.frame.destroy()

    def get(self) -> dict:
        return {'src': self.var_src.get().strip(),
                'dst': self.var_dst.get().strip()}


# ============================================================================
# Local Analysis Tab
# ============================================================================

class LocalAnalysisTab:
    """Tab 1: 测试站本地测试数据分析"""

    def __init__(self, notebook: ttk.Notebook):
        self.frame = ttk.Frame(notebook)
        notebook.add(self.frame, text='  本地数据分析  ')

        self._report_path = None
        self._station_rows = []
        self._merge_rows = []
        self._merge_expanded = False
        self._stop_event = threading.Event()
        self._config_path = os.path.join(_ROOT, 'app_config.json')

        self._build_ui()
        self._load_config()

    # ── UI construction ──────────────────────────────────────────────────

    def _build_ui(self):
        # Fill the whole tab with the same background color first,
        # so there is no visible empty strip below the content sections.
        bg_fill = tk.Frame(self.frame, bg='#f0f2f5')
        bg_fill.pack(fill='both', expand=True)

        outer = tk.Frame(bg_fill, bg='#f0f2f5')
        outer.pack(fill='x', expand=False, padx=10, pady=6)
        self._outer_frame = outer

        # ── Section 1: Input / Output config ──────────────────────────
        sec1 = self._make_section(outer, '输入 / 输出配置')

        inp_row = tk.Frame(sec1, bg='white')
        inp_row.pack(fill='x', pady=(2, 3))

        tk.Label(inp_row, text='发货Excel\n(可选):', width=9, anchor='w',
                 bg='white', font=('Segoe UI', 9)).pack(side='left')
        self._var_excel = tk.StringVar()
        tk.Entry(inp_row, textvariable=self._var_excel,
                 font=('Segoe UI', 9), width=28).pack(side='left', padx=(0, 3))
        tk.Button(inp_row, text='浏览…', font=('Segoe UI', 8),
                  command=lambda: self._browse_file(
                      self._var_excel, '选择发货 Excel',
                      [('Excel', '*.xlsx *.xls')]),
                  bg='#e0e4f0', relief='flat', padx=5).pack(side='left', padx=(0, 14))

        tk.Label(inp_row, text='输出目录:', width=8, anchor='w',
                 bg='white', font=('Segoe UI', 9)).pack(side='left')
        self._var_outdir = tk.StringVar()
        tk.Entry(inp_row, textvariable=self._var_outdir,
                 font=('Segoe UI', 9), width=28).pack(side='left', padx=(0, 3))
        tk.Button(inp_row, text='浏览…', font=('Segoe UI', 8),
                  command=lambda: self._browse_dir(self._var_outdir, '选择输出目录'),
                  bg='#e0e4f0', relief='flat', padx=5).pack(side='left')

        # ── Sections 2 + 3: Side by side ──────────────────────────────
        h_frame = tk.Frame(outer, bg='#f0f2f5')
        h_frame.pack(fill='x', pady=(0, 6))
        h_frame.columnconfigure(0, weight=1)
        h_frame.columnconfigure(1, weight=0, minsize=270)
        h_frame.rowconfigure(0, weight=1)

        # Left: Station config
        sec2 = tk.LabelFrame(h_frame, text='  测试工站配置  ',
                             font=('Segoe UI', 9, 'bold'),
                             bg='white', fg='#1a237e',
                             relief='groove', bd=1, padx=8, pady=6)
        sec2.grid(row=0, column=0, sticky='nsew', padx=(0, 4))
        self._build_station_list(sec2)
        self._build_merge_config(sec2)

        # Right: Analysis mode
        sec3 = tk.LabelFrame(h_frame, text='  分析模式  ',
                             font=('Segoe UI', 9, 'bold'),
                             bg='white', fg='#1a237e',
                             relief='groove', bd=1, padx=8, pady=6)
        sec3.grid(row=0, column=1, sticky='nsew')
        self._build_section3_content(sec3)

        # ── Section 4: Actions + progress ─────────────────────────────
        sec4 = self._make_section(outer, '操作')

        btn_row = tk.Frame(sec4, bg='white')
        btn_row.pack(fill='x', pady=(0, 4))

        self._btn_run = tk.Button(btn_row, text='开始分析',
                                  font=('Segoe UI', 9, 'bold'),
                                  bg='#3949ab', fg='white',
                                  relief='flat', padx=12, pady=4,
                                  command=self._on_run)
        self._btn_run.pack(side='left')

        tk.Button(btn_row, text='查看运行日志',
                  font=('Segoe UI', 9), bg='#455a64', fg='white',
                  relief='flat', padx=10, pady=4,
                  command=self._show_log).pack(side='left', padx=(8, 0))

        self._progress_label = tk.Label(sec4, text='就绪', anchor='w',
                                        font=('Segoe UI', 8), bg='white', fg='#555')
        self._progress_label.pack(fill='x')

        self._progress_var = tk.DoubleVar(value=0)
        self._progress_bar = ttk.Progressbar(sec4, variable=self._progress_var,
                                             maximum=100, mode='determinate')
        self._progress_bar.pack(fill='x', pady=(2, 0))

        # Build the detached log window (hidden until user opens it)
        self._build_log_window()

    # ── Station list (Section 2 upper part) ──────────────────────────────

    def _build_station_list(self, parent: tk.Frame):
        # Column header
        hdr = tk.Frame(parent, bg='#dde3f0')
        hdr.pack(fill='x', pady=(0, 2))
        tk.Label(hdr, text='工站类型', width=10, anchor='w',
                 bg='#dde3f0', font=('Segoe UI', 8, 'bold')).pack(side='left', padx=4, pady=1)
        tk.Label(hdr, text='测试数据文件夹路径', anchor='w',
                 bg='#dde3f0', font=('Segoe UI', 8, 'bold')).pack(side='left', padx=4)

        scroll_wrap = tk.Frame(parent, bg='#f5f6fa')
        scroll_wrap.pack(fill='x')

        self._station_canvas = tk.Canvas(
            scroll_wrap, bg='#f5f6fa', height=88, highlightthickness=0
        )
        vsb = ttk.Scrollbar(scroll_wrap, orient='vertical',
                             command=self._station_canvas.yview)
        self._station_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self._station_canvas.pack(side='left', fill='both', expand=True)

        self._rows_frame = tk.Frame(self._station_canvas, bg='#f5f6fa')
        self._canvas_win = self._station_canvas.create_window(
            (0, 0), window=self._rows_frame, anchor='nw'
        )

        def _on_rows_configure(_e):
            self._station_canvas.configure(
                scrollregion=self._station_canvas.bbox('all')
            )
            needed = self._rows_frame.winfo_reqheight() + 4
            new_h = max(28, min(needed, 130))
            self._station_canvas.configure(height=new_h)

        def _on_canvas_resize(e):
            self._station_canvas.itemconfig(self._canvas_win, width=e.width)

        self._rows_frame.bind('<Configure>', _on_rows_configure)
        self._station_canvas.bind('<Configure>', _on_canvas_resize)

        def _on_wheel(e):
            self._station_canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')

        self._station_canvas.bind('<MouseWheel>', _on_wheel)
        self._rows_frame.bind('<MouseWheel>', _on_wheel)

        for stype in ('FT1', 'FT2', 'Aging'):
            self._add_station_row(preset_type=stype)

        tk.Button(parent, text='＋ 添加工站', command=self._add_station_row,
                  font=('Segoe UI', 8), bg='#e8f5e9', relief='flat',
                  padx=6, pady=2).pack(anchor='w', pady=(4, 0))

    # ── Station merge config (Section 2 lower part, collapsible) ─────────

    def _build_merge_config(self, parent: tk.Frame):
        # Separator
        tk.Frame(parent, bg='#dde3f0', height=1).pack(fill='x', pady=(8, 0))

        # Toggle header
        merge_hdr = tk.Frame(parent, bg='white')
        merge_hdr.pack(fill='x', pady=(0, 0))
        self._merge_toggle_btn = tk.Button(
            merge_hdr, text='▶  工站合并配置',
            font=('Segoe UI', 8, 'bold'), bg='white', fg='#3949ab',
            relief='flat', anchor='w', cursor='hand2', padx=4, pady=3,
            command=self._toggle_merge,
        )
        self._merge_toggle_btn.pack(fill='x')

        # Merge body (hidden initially)
        self._merge_body = tk.Frame(parent, bg='#f9f9f9', relief='groove', bd=1)

        tk.Label(self._merge_body,
                 text='将指定工站类型的数据合并到目标工站类型中分析',
                 bg='#f9f9f9', font=('Segoe UI', 8), fg='#666').pack(
            anchor='w', padx=6, pady=(4, 2))

        self._merge_rows_frame = tk.Frame(self._merge_body, bg='#f9f9f9')
        self._merge_rows_frame.pack(fill='x', padx=6)

        tk.Button(self._merge_body, text='＋ 添加合并规则',
                  command=self._add_merge_row,
                  font=('Segoe UI', 8), bg='#e3f2fd', relief='flat',
                  padx=6, pady=2).pack(anchor='w', padx=6, pady=(4, 6))

    def _toggle_merge(self):
        self._merge_expanded = not self._merge_expanded
        if self._merge_expanded:
            self._merge_body.pack(fill='x', pady=(0, 4))
            self._merge_toggle_btn.configure(text='▼  工站合并配置')
        else:
            self._merge_body.pack_forget()
            self._merge_toggle_btn.configure(text='▶  工站合并配置')

    def _add_merge_row(self, src: str = '', dst: str = ''):
        row = MergeRuleRow(self._merge_rows_frame, self._delete_merge_row)
        if src:
            row.var_src.set(src)
        if dst:
            row.var_dst.set(dst)
        row.pack(fill='x', pady=1)
        self._merge_rows.append(row)

    def _delete_merge_row(self, row):
        self._merge_rows.remove(row)
        row.destroy()

    # ── Log window (popup) ────────────────────────────────────────────────

    def _build_log_window(self):
        """Create the detached log Toplevel (hidden until user opens it)."""
        self._log_win = tk.Toplevel(self.frame)
        self._log_win.title('运行日志 — 产线数据分析AI平台')
        self._log_win.geometry('760x420')
        self._log_win.configure(bg='#1e1e2e')
        self._log_win.protocol('WM_DELETE_WINDOW', self._log_win.withdraw)
        self._log_win.withdraw()   # hidden until user clicks 查看运行日志

        hdr = tk.Frame(self._log_win, bg='#1a237e')
        hdr.pack(fill='x')
        tk.Label(hdr, text='运行日志', font=('Segoe UI', 9, 'bold'),
                 bg='#1a237e', fg='white', padx=8, pady=4).pack(side='left')
        tk.Button(hdr, text='清空', font=('Segoe UI', 8),
                  bg='#3949ab', fg='white', relief='flat', padx=8,
                  command=self._clear_log).pack(side='right', padx=4, pady=2)

        self._log = scrolledtext.ScrolledText(
            self._log_win, font=('Consolas', 8),
            bg='#1e1e2e', fg='#a8d8a8', insertbackground='white',
            state='disabled', wrap='word',
        )
        self._log.pack(fill='both', expand=True, padx=2, pady=(0, 2))

    def _show_log(self):
        """Show (or bring to front) the log popup window."""
        self._log_win.deiconify()
        self._log_win.lift()
        self._log.see('end')

    # ── Section 3 content (analysis mode) ────────────────────────────────

    _CPK_MODE_LABELS = [
        '最后一次pass数据',
        '所选文件夹分析',
        '全部成功数据',
        '所有数据（含失败）',
        '仅失败数据',
    ]
    _CPK_MODE_VALUES = [
        'latest_pass',
        'folder_direct',
        'all_pass',
        'all_with_fail',
        'fail_only',
    ]

    def _build_section3_content(self, container: tk.Frame):
        """Fill the analysis mode panel (container = the LabelFrame sec3)."""
        # ── File type selection ────────────────────────────────────────
        type_row = tk.Frame(container, bg='white')
        type_row.pack(fill='x', pady=(0, 6))

        tk.Label(type_row, text='分析文件类型：', bg='white',
                 font=('Segoe UI', 9)).pack(side='left')

        self._var_file_type = tk.StringVar(value='xlsx')
        for val, lbl in (('xlsx', 'xlsx'), ('json', 'json')):
            tk.Radiobutton(type_row, text=lbl, variable=self._var_file_type, value=val,
                           bg='white', font=('Segoe UI', 9),
                           activebackground='white').pack(side='left', padx=(0, 10))

        # ── Mode combobox ──────────────────────────────────────────────
        tk.Label(container, text='分析模式：', bg='white',
                 font=('Segoe UI', 9), anchor='w').pack(fill='x', pady=(0, 2))

        self._mode_display_var = tk.StringVar(value=self._CPK_MODE_LABELS[0])
        mode_combo = ttk.Combobox(container, textvariable=self._mode_display_var,
                                  values=self._CPK_MODE_LABELS, state='readonly',
                                  width=22, font=('Segoe UI', 9))
        mode_combo.pack(fill='x', pady=(0, 4))
        _ToolTip(mode_combo, lambda: _MODE_HINTS.get(self._mode_display_var.get(), ''))

        # ── Fault analysis method (shown only for all_with_fail / fail_only) ──
        self._fault_frame = tk.Frame(container, bg='white')

        tk.Frame(self._fault_frame, bg='#cccccc', height=1).pack(fill='x', pady=(2, 6))

        fault_lbl = tk.Label(self._fault_frame, text='故障分析方式：', bg='white',
                             font=('Segoe UI', 9), anchor='w')
        fault_lbl.pack(fill='x')
        _ToolTip(fault_lbl,
                 lambda: '基础版：仅使用规则库匹配\n增强版：规则库 + Ollama LLM 辅助分类（需本地Ollama）')

        self._fault_level_var = tk.StringVar(value='基础版（规则库）')
        ttk.Combobox(self._fault_frame, textvariable=self._fault_level_var,
                     values=['基础版（规则库）', '增强版（规则库+Ollama）'],
                     state='readonly', width=22,
                     font=('Segoe UI', 9)).pack(fill='x', pady=(2, 0))

        # Show/hide fault panel when mode changes
        self._mode_display_var.trace_add('write', lambda *_: self._update_fault_panel())
        self._update_fault_panel()

    def _update_fault_panel(self):
        """Show fault analysis options only for modes that update the fault DB."""
        mode = self._mode_display_var.get()
        if mode in ('所有数据（含失败）', '仅失败数据'):
            self._fault_frame.pack(fill='x')
        else:
            self._fault_frame.pack_forget()

    # ── Section helper ────────────────────────────────────────────────────

    def _make_section(self, parent, title: str, expand: bool = False) -> tk.Frame:
        lf = tk.LabelFrame(parent, text=f'  {title}  ',
                           font=('Segoe UI', 9, 'bold'),
                           bg='white', fg='#1a237e',
                           relief='groove', bd=1, padx=8, pady=6)
        if expand:
            lf.pack(fill='both', expand=True, pady=(0, 6))
        else:
            lf.pack(fill='x', pady=(0, 6))
        return lf

    # ── Station rows ──────────────────────────────────────────────────────

    def _add_station_row(self, preset_type: str = ''):
        row = StationRow(self._rows_frame, self._delete_station_row)
        if preset_type:
            row.var_type.set(preset_type)
        row.pack(fill='x', pady=1, padx=2)
        self._station_rows.append(row)

        def _nav(event, r=row):
            idx = self._station_rows.index(r)
            focused = event.widget
            delta = -1 if event.keysym == 'Up' else 1
            target_idx = idx + delta
            if 0 <= target_idx < len(self._station_rows):
                target = self._station_rows[target_idx]
                if focused is r._folder_entry:
                    target._folder_entry.focus_set()
                else:
                    target._type_entry.focus_set()
            return 'break'

        for widget in (row._type_entry, row._folder_entry):
            widget.bind('<Up>',   _nav)
            widget.bind('<Down>', _nav)

    def _delete_station_row(self, row):
        self._station_rows.remove(row)
        row.destroy()

    # ── Browse helpers ────────────────────────────────────────────────────

    def _browse_file(self, var, title, filetypes):
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if path:
            var.set(path)

    def _browse_dir(self, var, title):
        path = filedialog.askdirectory(title=title)
        if path:
            var.set(path)

    # ── Log helpers ───────────────────────────────────────────────────────

    def _log_msg(self, msg: str):
        def _do():
            self._log.configure(state='normal')
            self._log.insert('end', f'[{_ts()}] {msg}\n')
            self._log.see('end')
            self._log.configure(state='disabled')
        self.frame.after(0, _do)

    def _clear_log(self):
        self._log.configure(state='normal')
        self._log.delete('1.0', 'end')
        self._log.configure(state='disabled')

    # ── Config persistence ────────────────────────────────────────────────

    def _load_config(self):
        import json
        try:
            with open(self._config_path, encoding='utf-8') as f:
                cfg = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return

        self._var_excel.set(cfg.get('excel_path', ''))
        self._var_outdir.set(cfg.get('out_dir', ''))

        stations = cfg.get('stations')
        if stations:
            for row in list(self._station_rows):
                row.destroy()
            self._station_rows.clear()
            for s in stations:
                self._add_station_row(preset_type=s.get('type', ''))
                self._station_rows[-1].var_folder.set(s.get('folder', ''))

        # Restore analysis mode
        saved_mode = cfg.get('cpk_mode', self._CPK_MODE_LABELS[0])
        if saved_mode in self._CPK_MODE_LABELS:
            self._mode_display_var.set(saved_mode)
        saved_level = cfg.get('fault_level', '基础版（规则库）')
        self._fault_level_var.set(saved_level)

        # Restore file type
        saved_type = cfg.get('file_type', 'xlsx')
        if saved_type in ('xlsx', 'json'):
            self._var_file_type.set(saved_type)

        # Restore merge rules
        merge_rules = cfg.get('merge_rules', [])
        for rule in merge_rules:
            self._add_merge_row(src=rule.get('src', ''), dst=rule.get('dst', ''))
        if merge_rules:
            # Auto-expand merge config if there are rules
            self._toggle_merge()

    def save_config(self):
        import json
        cfg = {
            'excel_path':   self._var_excel.get().strip(),
            'out_dir':      self._var_outdir.get().strip(),
            'stations':     [r.get() for r in self._station_rows],
            'cpk_mode':     self._mode_display_var.get(),
            'fault_level':  self._fault_level_var.get(),
            'file_type':    self._var_file_type.get(),
            'merge_rules':  [r.get() for r in self._merge_rows
                             if r.get()['src'] and r.get()['dst']],
        }
        try:
            with open(self._config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except OSError:
            pass

    # ── Progress helpers ──────────────────────────────────────────────────

    def _set_progress(self, pct: float, label: str = ''):
        def _do():
            self._progress_var.set(pct)
            display = f'[{pct:.0f}%]  {label}' if label else f'[{pct:.0f}%]'
            self._progress_label.configure(text=display)
        self.frame.after(0, _do)

    def _on_stop(self):
        self._stop_event.set()
        self._btn_run.configure(state='disabled')
        self._log_msg('[INFO] 正在中止分析，请稍候...')

    def _set_buttons(self, running: bool):
        def _do():
            if running:
                self._btn_run.configure(
                    text='停止分析', bg='#c62828', command=self._on_stop,
                    state='normal'
                )
            else:
                self._btn_run.configure(
                    text='开始分析', bg='#3949ab', command=self._on_run,
                    state='normal'
                )
        self.frame.after(0, _do)

    # ── Main run logic ────────────────────────────────────────────────────

    def _on_run(self):
        excel_path = self._var_excel.get().strip()
        out_dir = self._var_outdir.get().strip()

        # Determine CPK mode
        label = self._mode_display_var.get()
        try:
            cpk_mode = self._CPK_MODE_VALUES[self._CPK_MODE_LABELS.index(label)]
        except (ValueError, IndexError):
            cpk_mode = 'latest_pass'

        fault_enabled = cpk_mode in ('all_with_fail', 'fail_only')
        fault_level = self._fault_level_var.get()
        file_type = self._var_file_type.get()

        # Collect merge rules
        merge_rules = [r.get() for r in self._merge_rows
                       if r.get()['src'] and r.get()['dst']]

        # Validation
        if cpk_mode == 'latest_pass':
            if not excel_path or not os.path.isfile(excel_path):
                messagebox.showerror(
                    '错误',
                    '【最后一次pass数据】模式需要发货 Excel 文件\n'
                    '请选择文件，或切换到其他不需要 Excel 的分析模式'
                )
                return
        elif excel_path and not os.path.isfile(excel_path):
            messagebox.showerror('错误', '发货 Excel 文件路径无效，请重新选择')
            return

        if not out_dir:
            messagebox.showerror('错误', '请选择输出目录')
            return

        station_configs = [r.get() for r in self._station_rows]
        station_configs = [c for c in station_configs if c['type'] and c['folder']]
        if not station_configs:
            messagebox.showerror('错误', '请至少配置一个测试工站（类型 + 文件夹）')
            return

        self.save_config()
        self._report_path = None
        self._stop_event.clear()
        self._set_buttons(running=True)
        self._set_progress(0, '正在准备...')
        self._clear_log()

        threading.Thread(
            target=self._run_analysis,
            args=(excel_path, out_dir, station_configs, cpk_mode,
                  fault_enabled, fault_level, file_type, merge_rules),
            daemon=True,
        ).start()

    def _run_analysis(self, excel_path: str, out_dir: str, station_configs: list,
                      cpk_mode: str = 'latest_pass',
                      fault_enabled: bool = False,
                      fault_level: str = '基础版（规则库）',
                      file_type: str = 'xlsx',
                      merge_rules: list = None):
        import time
        t_start = time.time()
        merge_rules = merge_rules or []

        def elapsed():
            return f'{time.time() - t_start:.1f}s'

        os.makedirs(out_dir, exist_ok=True)
        log_filename = f'analysis_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt'
        log_path = os.path.join(out_dir, log_filename)
        try:
            _log_file = open(log_path, 'w', encoding='utf-8')
        except OSError:
            _log_file = None

        def _log(msg: str):
            self._log_msg(msg)
            if _log_file:
                try:
                    _log_file.write(f'[{_ts()}] {msg}\n')
                    _log_file.flush()
                except OSError:
                    pass

        _MODE_DESC = {
            'latest_pass':   '最后一次pass数据',
            'folder_direct': '所选文件夹分析',
            'all_pass':      '全部成功数据',
            'all_with_fail': '所有数据（含失败）',
            'fail_only':     '仅失败数据',
        }

        try:
            _log(f'日志文件: {log_path}')
            _log(f'分析模式: {_MODE_DESC.get(cpk_mode, cpk_mode)}'
                 + (f'  |  故障分析: {fault_level}' if fault_enabled else '')
                 + f'  |  文件类型: {file_type}')
            if merge_rules:
                merge_desc = ', '.join(f'{r["src"]}→{r["dst"]}' for r in merge_rules)
                _log(f'工站合并规则: {merge_desc}')

            # ──────────────────────────────────────────────────────────
            # MODE: folder_direct — skip discovery/extraction, use folders directly
            # ──────────────────────────────────────────────────────────
            if cpk_mode == 'folder_direct':
                _log('=' * 56)
                _log('【直接分析模式】跳过目录遍历，直接对工站文件夹做CPK')
                all_analysis = {}
                for cfg in station_configs:
                    if self._stop_event.is_set():
                        _log('[INFO] 分析已中止')
                        self._set_progress(0, '已中止')
                        return
                    stype = cfg['type']
                    folder = cfg['folder']
                    if not folder or not os.path.isdir(folder):
                        _log(f'  [WARN] 工站 [{stype}] 文件夹不存在: {folder}')
                        continue
                    # Apply merge rules: if this stype is a source, merge into target
                    effective_type = stype
                    for rule in merge_rules:
                        if rule['src'] == stype:
                            effective_type = rule['dst']
                            _log(f'  [合并] {stype} → {effective_type}')
                            break
                    _log(f'\n  工站 [{effective_type}]  ({folder})')
                    self._set_progress(30, f'CPK 分析: {effective_type}')
                    station_result = analyze_xlsx_folder(folder, log_cb=_log)
                    if station_result:
                        if effective_type in all_analysis:
                            # Merge results from multiple folders into same station type
                            for item, pts in station_result.items():
                                if item in all_analysis[effective_type]:
                                    all_analysis[effective_type][item].extend(pts)
                                else:
                                    all_analysis[effective_type][item] = pts
                        else:
                            all_analysis[effective_type] = station_result
                    else:
                        _log(f'  [WARN] 工站 [{effective_type}] 无可分析数据')

                # Jump directly to HTML report
                self._set_progress(80, '生成 HTML 报告...')
                _log('\n' + '=' * 56)
                _log('【生成HTML报告】')
                report_path = os.path.join(out_dir, 'cpk_report.html')
                from collections import Counter as _Counter
                station_info = dict(_Counter(
                    c['type'] for c in station_configs if c['type'] and c['folder']
                ))
                product_name = _infer_product_name(station_configs)
                report_title = (f'{product_name}测试数据分析报告 - Zillnk'
                                if product_name else '测试数据分析报告 - Zillnk')
                generate_report(
                    analysis_data=all_analysis,
                    output_path=report_path,
                    title=report_title,
                    station_info=station_info,
                )
                self._report_path = report_path
                report_kb = os.path.getsize(report_path) // 1024
                _log(f'  HTML 报告: {report_path}  ({report_kb} KB)')
                _log(f'\n【完成】总耗时: {elapsed()}')
                self._set_progress(100, f'完成！耗时 {elapsed()}')
                self.frame.after(800, lambda: webbrowser.open(
                    'file:///' + report_path.replace(os.sep, '/')
                ))
                return

            # ──────────────────────────────────────────────────────────
            # MODES: latest_pass / all_pass / all_with_fail / fail_only
            # ──────────────────────────────────────────────────────────

            # ── Step 1: read / discover barcodes ──────────────────────
            _log('=' * 56)
            if excel_path:
                _log('【第1步】读取发货 Excel 条码列表')
                _log(f'  文件: {excel_path}')
                try:
                    barcodes = read_barcodes(excel_path)
                except Exception as exc:
                    _log(f'  [ERROR] 读取条码失败: {exc}')
                    self._set_buttons(running=False)
                    self._set_progress(0, '失败 - 请检查 Excel 文件')
                    self.frame.after(0, lambda e=exc: messagebox.showerror(
                        '读取 Excel 失败',
                        f'无法读取发货 Excel 文件：\n\n{e}\n\n请确认文件未被其他程序占用，且格式正确。'
                    ))
                    return
            else:
                _log('【第1步】自动扫描工站目录发现条码')
                from core.data_extractor import discover_barcodes
                valid_folders = [c['folder'] for c in station_configs
                                 if os.path.isdir(c.get('folder', ''))]
                barcodes = discover_barcodes(valid_folders)
                _log(f'  自动发现条码: {len(barcodes)} 个')

            unique_bc = list(dict.fromkeys(barcodes))
            if len(unique_bc) < len(barcodes):
                _log(
                    f'  [WARN] 发现重复条码: 原始 {len(barcodes)} 条 → '
                    f'去重后 {len(unique_bc)} 条'
                )
                barcodes = unique_bc
            _log(f'  条码总数: {len(barcodes)} 个')
            if barcodes:
                _log(f'  样例: {barcodes[:3]} ...')

            from collections import Counter
            type_counts = Counter(c['type'] for c in station_configs if c['type'])
            _log(f'\n  工站配置: {len(station_configs)} 条记录，'
                 f'涉及类型: {dict(type_counts)}')
            for cfg in station_configs:
                exists = '✓' if os.path.isdir(cfg['folder']) else '✗ 不存在'
                _log(f'    [{cfg["type"]}] {cfg["folder"]}  [{exists}]')

            _log(f'  输出目录: {out_dir}')
            self._set_progress(5, f'读取到 {len(barcodes)} 个条码')

            # ── Step 2: extract files ──────────────────────────────────
            _log('\n' + '=' * 56)
            _log(f'【第2步】遍历工站目录，提取测试记录（模式: {_MODE_DESC.get(cpk_mode)}）')

            def progress_cb(done, total, bc):
                pct = 5 + 55 * done / max(total, 1)
                self._set_progress(pct, f'提取中 ({done}/{total}): {bc}')

            # Map all_with_fail → all for data_extractor (extractor uses 'all')
            extractor_mode = 'all' if cpk_mode == 'all_with_fail' else cpk_mode

            extraction_summary = run_extraction(
                barcodes=barcodes,
                station_configs=station_configs,
                output_base_dir=out_dir,
                log_cb=_log,
                progress_cb=progress_cb,
                stop_event=self._stop_event,
                mode=extractor_mode,
            )

            if self._stop_event.is_set():
                _log('[INFO] 分析已中止')
                self._set_progress(0, '已中止')
                return

            # ── Step 3: missing barcodes report ───────────────────────
            _log('\n' + '=' * 56)
            _log('【第3步】生成缺失条码汇总报表')
            missing_path = os.path.join(out_dir, 'missing_barcodes.xlsx')
            try:
                generate_missing_report(
                    summary=extraction_summary,
                    output_path=missing_path,
                    log_cb=_log,
                )
                total_missing = sum(
                    sum(1 for r in info['results'] if r['status'] != 'success')
                    for info in extraction_summary.values()
                )
                if total_missing == 0:
                    _log('  所有条码均已成功提取，缺失报表为空')
                else:
                    _log(
                        f'  [注意] 共 {total_missing} 个条码缺失/异常，'
                        f'详见: {missing_path}'
                    )
            except Exception as exc:
                _log(f'  [ERROR] 缺失报表生成失败: {exc}')

            self._set_progress(62, '缺失报表已生成')

            # ── Step 4: CPK analysis ───────────────────────────────────
            _log('\n' + '=' * 56)
            _log('【第4步】CPK 分析')
            all_analysis = {}
            station_list = list(extraction_summary.keys())

            for idx, stype in enumerate(station_list):
                if self._stop_event.is_set():
                    _log('[INFO] CPK 分析已中止')
                    self._set_progress(0, '已中止')
                    return
                xlsx_dir = extraction_summary[stype]['xlsx_dir']
                try:
                    xlsx_count = len([
                        f for f in os.listdir(xlsx_dir)
                        if f.lower().endswith('.xlsx')
                    ])
                except OSError:
                    xlsx_count = 0

                _log(f'\n  工站 [{stype}]  —  共 {xlsx_count} 个xlsx文件')
                self._set_progress(
                    62 + 28 * idx / max(len(station_list), 1),
                    f'CPK 分析: {stype} ({idx+1}/{len(station_list)})'
                )
                station_result = analyze_xlsx_folder(xlsx_dir, log_cb=_log)
                if station_result:
                    all_analysis[stype] = station_result
                else:
                    _log(f'  [WARN] 工站 [{stype}] 无可分析数据')

            if not all_analysis:
                _log('[WARN] 所有工站均无可分析的 xlsx 数据，HTML 报告将为空')

            # ── Step 5: generate HTML report ───────────────────────────
            _log('\n' + '=' * 56)
            _log('【第5步】生成 HTML 报告')
            self._set_progress(92, '生成 HTML 报告...')

            report_path = os.path.join(out_dir, 'cpk_report.html')
            from collections import Counter as _Counter
            station_info = dict(_Counter(
                c['type'] for c in station_configs if c['type'] and c['folder']
            ))
            product_name = _infer_product_name(station_configs)
            report_title = (f'{product_name}测试数据分析报告 - Zillnk'
                            if product_name else '测试数据分析报告 - Zillnk')
            generate_report(
                analysis_data=all_analysis,
                output_path=report_path,
                title=report_title,
                station_info=station_info,
            )

            report_kb = os.path.getsize(report_path) // 1024
            self._report_path = report_path
            _log(f'  HTML 报告: {report_path}  ({report_kb} KB)')

            # ── Step 6: fault analysis (optional) ─────────────────────
            if fault_enabled and not self._stop_event.is_set():
                _log('\n' + '=' * 56)
                _log(f'【第6步】故障分析及定位关系库建立（{fault_level}）')
                self._set_progress(93, '故障分析中...')
                try:
                    from core.fault_analyzer import run_fault_analysis
                    # Map all_with_fail → all for fault_analyzer (uses 'all' internally)
                    fa_mode = 'all' if cpk_mode == 'all_with_fail' else cpk_mode
                    fault_summary = run_fault_analysis(
                        station_configs=station_configs,
                        out_dir=out_dir,
                        level=fault_level,
                        mode=fa_mode,
                        log_cb=_log,
                        stop_event=self._stop_event,
                    )
                    for stype, stats in fault_summary.items():
                        if stype.startswith('__'):
                            continue
                        _log(
                            f'  [{stype}] 总 {stats["total"]} 条记录，'
                            f'已分类 {stats["classified"]}，'
                            f'未分类 {stats["unclassified"]}'
                        )
                except Exception as exc:
                    import traceback as _tb
                    _log(f'  [ERROR] 故障分析失败: {exc}')
                    _log(_tb.format_exc())

            # ── Final summary ──────────────────────────────────────────
            _log('\n' + '=' * 56)
            _log(f'【完成】总耗时: {elapsed()}')
            for stype, info in extraction_summary.items():
                res = info['results']
                ok  = sum(1 for r in res if r['status'] == 'success')
                bad = len(res) - ok
                cpk_pts = sum(
                    len(pts)
                    for pts in all_analysis.get(stype, {}).values()
                )
                _log(
                    f'  [{stype}] 提取: {ok}/{len(res)} 成功，{bad} 缺失  |  '
                    f'CPK 子项: {cpk_pts}'
                )
            _log('=' * 56)
            _log(f'日志已保存: {log_path}')

            self._set_progress(100, f'完成！耗时 {elapsed()}')

            self.frame.after(800, lambda: webbrowser.open(
                'file:///' + report_path.replace(os.sep, '/')
            ))

        except Exception as exc:
            import traceback
            tb = traceback.format_exc()
            _log(f'[ERROR] 未预期的错误: {exc}')
            _log(tb)
            self._set_progress(0, '发生错误，请查看日志')
            self.frame.after(0, lambda e=exc, t=tb: messagebox.showerror(
                '分析出错',
                f'分析过程中发生错误：\n\n{e}\n\n详细信息已写入运行日志，点击"查看运行日志"可查看完整追踪。'
            ))
        finally:
            if _log_file:
                _log_file.close()
            self._set_buttons(running=False)


# ============================================================================
# Placeholder tabs for future modules
# ============================================================================

class PlaceholderTab:
    def __init__(self, notebook: ttk.Notebook, title: str, description: str):
        self.frame = ttk.Frame(notebook)
        notebook.add(self.frame, text=f'  {title}  ')

        outer = tk.Frame(self.frame, bg='#f0f2f5')
        outer.pack(fill='both', expand=True)

        tk.Label(outer, text=title,
                 font=('Segoe UI', 18, 'bold'),
                 bg='#f0f2f5', fg='#1a237e').pack(pady=(80, 12))

        tk.Label(outer, text=description,
                 font=('Segoe UI', 11), bg='#f0f2f5', fg='#555',
                 justify='center').pack()

        tk.Label(outer, text='（功能待实现）',
                 font=('Segoe UI', 10),
                 bg='#f0f2f5', fg='#aaa').pack(pady=(8, 0))


# ============================================================================
# Help window
# ============================================================================

def _show_help(root: tk.Tk):
    win = tk.Toplevel(root)
    win.title('使用帮助 — 产线数据分析AI平台')
    win.geometry('620x480')
    win.resizable(True, True)
    win.configure(bg='#f0f2f5')
    win.transient(root)
    win.grab_set()

    tk.Label(win, text='使用帮助', font=('Segoe UI', 12, 'bold'),
             bg='#1a237e', fg='white').pack(fill='x', ipady=6)

    txt = scrolledtext.ScrolledText(
        win, font=('Segoe UI', 9), wrap='word',
        bg='white', fg='#212121', relief='flat',
        padx=12, pady=8
    )
    txt.pack(fill='both', expand=True, padx=8, pady=8)
    txt.insert('1.0', _HELP_TEXT)
    txt.configure(state='disabled')

    tk.Button(win, text='关闭', command=win.destroy,
              font=('Segoe UI', 9), bg='#3949ab', fg='white',
              relief='flat', padx=16, pady=4).pack(pady=(0, 10))


# ============================================================================
# Fault rule import dialog
# ============================================================================

def _import_fault_rules(root: tk.Tk, config_path: str):
    """Load a YAML fault rules file and import into the fault database."""
    yaml_path = filedialog.askopenfilename(
        title='选择故障关系描述文件',
        filetypes=[('YAML文件', '*.yml *.yaml'), ('所有文件', '*.*')]
    )
    if not yaml_path:
        return

    # Find the output dir from app_config.json to locate the DB
    import json
    out_dir = ''
    try:
        with open(config_path, encoding='utf-8') as f:
            cfg = json.load(f)
        out_dir = cfg.get('out_dir', '')
    except Exception:
        pass

    if not out_dir:
        out_dir = filedialog.askdirectory(
            title='选择包含 fault_database.db 的输出目录'
        )
        if not out_dir:
            return

    db_path = os.path.join(out_dir, 'fault_database.db')
    if not os.path.isfile(db_path):
        messagebox.showerror(
            '错误',
            f'未找到故障数据库文件:\n{db_path}\n\n请先运行一次"所有数据"或"仅失败数据"分析模式。'
        )
        return

    try:
        # Simple YAML parser for our format (no external dependency)
        rules = _parse_fault_rules_yaml(yaml_path)
        if not rules:
            messagebox.showwarning('警告', '文件中未找到有效规则，请检查YAML格式')
            return

        from core.fault_db import init_db, get_rules, add_rule, update_rule
        init_db(db_path)
        existing = {r['keywords']: r for r in get_rules(db_path)}

        added, updated = 0, 0
        for rule in rules:
            kw = rule.get('keywords', '').strip()
            ft = rule.get('fault_type', '').strip()
            sg = rule.get('suggestion', '').strip()
            if not kw or not ft:
                continue
            if kw in existing:
                update_rule(db_path, existing[kw]['id'],
                            keywords=kw, fault_type=ft, suggestion=sg)
                updated += 1
            else:
                add_rule(db_path, kw, ft, sg)
                added += 1

        messagebox.showinfo(
            '导入完成',
            f'故障规则导入成功！\n\n新增: {added} 条\n更新: {updated} 条\n\n数据库: {db_path}'
        )
    except Exception as exc:
        messagebox.showerror('导入失败', f'导入故障规则时出错:\n{exc}')


def _parse_fault_rules_yaml(path: str) -> list:
    """
    Minimal YAML parser for fault rules files.
    Supports the format:
        rules:
          - keywords: "..."
            fault_type: "..."
            suggestion: "..."
    """
    rules = []
    current = {}
    in_rules = False

    with open(path, encoding='utf-8') as f:
        for line in f:
            stripped = line.rstrip()
            if not stripped or stripped.lstrip().startswith('#'):
                continue

            if stripped.strip() == 'rules:':
                in_rules = True
                continue

            if not in_rules:
                continue

            indent = len(line) - len(line.lstrip())

            if stripped.strip().startswith('- '):
                if current:
                    rules.append(current)
                current = {}
                # Handle inline first key: "- keywords: ..."
                rest = stripped.strip()[2:]
                if ':' in rest:
                    k, v = rest.split(':', 1)
                    current[k.strip()] = v.strip().strip('"').strip("'")
            elif ':' in stripped and indent > 0:
                k, v = stripped.strip().split(':', 1)
                current[k.strip()] = v.strip().strip('"').strip("'")

    if current:
        rules.append(current)

    return [r for r in rules if r.get('keywords') and r.get('fault_type')]


# ============================================================================
# Main application window
# ============================================================================

class CPKAnalysisPlatform:

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title('产线数据分析AI平台 — Zillnk Efficiency Improvement Group')
        root.configure(bg='#1a237e')

        self._apply_style()
        self._build_menu()
        self._build_ui()

        # Auto-fit window height to content; keep width fixed at 1000
        root.update_idletasks()
        root.geometry(f'1000x{root.winfo_reqheight()}')
        root.minsize(820, root.winfo_reqheight())

        root.bind('<F11>', self._toggle_fullscreen)
        root.bind('<Escape>', lambda _e: root.attributes('-fullscreen', False))
        root.protocol('WM_DELETE_WINDOW', self._on_close)

    def _apply_style(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure('TNotebook', background='#1a237e', borderwidth=0)
        style.configure('TNotebook.Tab',
                        background='#3949ab', foreground='white',
                        padding=[12, 5], font=('Segoe UI', 9))
        style.map('TNotebook.Tab',
                  background=[('selected', '#f0f2f5')],
                  foreground=[('selected', '#1a237e')])
        style.configure('TFrame', background='#f0f2f5')
        style.configure('TProgressbar',
                        troughcolor='#e0e0e0', background='#3949ab',
                        thickness=10)

    def _build_menu(self):
        menubar = tk.Menu(self.root, bg='#1a237e', fg='white',
                          activebackground='#3949ab', activeforeground='white',
                          relief='flat')

        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0,
                             bg='white', fg='#212121',
                             activebackground='#3949ab', activeforeground='white')
        tools_menu.add_command(
            label='加载故障关系描述文件…',
            command=lambda: _import_fault_rules(
                self.root,
                os.path.join(_ROOT, 'app_config.json')
            )
        )
        menubar.add_cascade(label=' 工具 ', menu=tools_menu)

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0,
                            bg='white', fg='#212121',
                            activebackground='#3949ab', activeforeground='white')
        help_menu.add_command(
            label='使用帮助',
            command=lambda: _show_help(self.root)
        )
        help_menu.add_separator()
        help_menu.add_command(
            label='关于',
            command=lambda: messagebox.showinfo(
                '关于',
                '产线数据分析AI平台  v2.0\n\nZillnk Efficiency Improvement Group\n\n'
                '本地测试站数据分析、CPK计算、故障定位关系库建立。'
            )
        )
        menubar.add_cascade(label=' 帮助 ', menu=help_menu)
        self.root.configure(menu=menubar)

    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill='both', expand=True)

        self._local_tab = LocalAnalysisTab(nb)
        PlaceholderTab(nb, '深科技 MES 数据分析',
                       '从深科技 MES 导出的测试数据 CPK 分析\n支持批次、工站、产品型号多维度分析')
        PlaceholderTab(nb, '立讯 MES 数据分析',
                       '从立讯 MES 导出的测试数据 CPK 分析\n支持批次、工站、产品型号多维度分析')

    def _on_close(self):
        self._local_tab.save_config()
        self.root.destroy()

    def _toggle_fullscreen(self, _event=None):
        current = self.root.attributes('-fullscreen')
        self.root.attributes('-fullscreen', not current)


# ============================================================================
# Entry point
# ============================================================================

def main():
    root = tk.Tk()
    CPKAnalysisPlatform(root)
    root.mainloop()


if __name__ == '__main__':
    main()

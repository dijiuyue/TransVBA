"""Tkinter view layer for TransVBA.

Corresponds to VBA UserForm1.frm:
  - UserForm_Initialize
  - CreateContentPage / CreateTitlePage / CreateTablePage / CreateFigurePage
  - btnApply_Click / btnOK_Click / btnCancel_Click
  - LoadSettingsToForm / SetEditingEnabled
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path


class TvbaMainWindow(tk.Tk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.title("TransVBA-Pro — Word 格式自动刷新")
        self.geometry("900x650")
        self.minsize(700, 500)

        self._build_layout()
        self._populate_from_settings()

    def _build_layout(self):
        # Top bar: file open
        top_frame = ttk.Frame(self, padding=5)
        top_frame.pack(fill=tk.X)

        self.btn_open = ttk.Button(top_frame, text="打开文件...", command=self._on_open)
        self.btn_open.pack(side=tk.LEFT, padx=5)

        # Template selector
        from tvba_templates import TemplateManager
        templates = TemplateManager.list_templates()
        template_names = [t["name"] for t in templates]
        self._template_ids = [t["id"] for t in templates]
        self._template_names = {t["id"]: t["name"] for t in templates}

        ttk.Label(top_frame, text="模板标准：").pack(side=tk.LEFT, padx=(15, 2))
        self.cmb_template = ttk.Combobox(top_frame, values=template_names, state="readonly", width=22)
        if template_names:
            self.cmb_template.current(0)
        self.cmb_template.pack(side=tk.LEFT, padx=2)
        self.cmb_template.bind("<<ComboboxSelected>>", self._on_template_change)

        self.lbl_file = ttk.Label(top_frame, text="(未选择文件)")
        self.lbl_file.pack(side=tk.LEFT, padx=15)

        # Main paned window
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left: category tree
        left_frame = ttk.Frame(paned, width=180)
        paned.add(left_frame, weight=0)

        self.tree = ttk.Treeview(left_frame, show="tree", selectmode="browse")
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Tree items
        self.tree.insert("", "end", "body", text="正文")
        self.tree.insert("", "end", "titles", text="标题", open=True)
        for i in range(1, 6):
            self.tree.insert("titles", "end", f"title_{i}", text=f"  {i}级标题")
        self.tree.insert("", "end", "table", text="表格")
        self.tree.insert("", "end", "figure", text="图片标题")
        self.tree.insert("", "end", "advanced", text="高级")

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # Right: detail panel with scrollbar
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)

        self.detail_canvas = tk.Canvas(right_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=self.detail_canvas.yview)
        self.detail_frame = ttk.Frame(self.detail_canvas)

        # Use frame's requested size for scrollregion instead of canvas bbox("all").
        # bbox("all") is unreliable because Combobox dropdowns (Toplevel windows)
        # can transiently expand the reported bounds, causing a big empty area.
        self.detail_frame.bind(
            "<Configure>",
            lambda e: self.detail_canvas.configure(
                scrollregion=(0, 0, self.detail_frame.winfo_reqwidth(), self.detail_frame.winfo_reqheight())
            )
        )

        self._detail_window = self.detail_canvas.create_window((0, 0), window=self.detail_frame, anchor="nw")
        self.detail_canvas.configure(yscrollcommand=scrollbar.set)

        # Dynamically resize the inner frame to match canvas width minus scrollbar
        def _on_canvas_resize(event):
            canvas_width = event.width
            self.detail_canvas.itemconfig(self._detail_window, width=canvas_width)
        self.detail_canvas.bind("<Configure>", _on_canvas_resize)

        self.detail_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom bar
        bottom_frame = ttk.Frame(self, padding=5)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.chk_edit = tk.BooleanVar(value=False)
        ttk.Checkbutton(bottom_frame, text="修改模式", variable=self.chk_edit,
                       command=self._on_edit_toggle).pack(side=tk.LEFT, padx=5)

        self.chk_remember = tk.BooleanVar(value=True)
        self._remember_cb = ttk.Checkbutton(bottom_frame, text="记忆本次设置", variable=self.chk_remember)
        self._remember_cb.pack(side=tk.LEFT, padx=5)

        self._reset_btn = ttk.Button(bottom_frame, text="重置为默认", command=self._on_reset)
        self._reset_btn.pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="取消", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="应用", command=self._on_apply).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="应用并关闭", command=self._on_apply_close).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="格式检查", command=self._on_validate).pack(side=tk.RIGHT, padx=5)

        # Progress bar
        self.progress = ttk.Progressbar(self, mode="determinate", maximum=100)
        self.progress.pack(fill=tk.X, padx=5, pady=2)

        self.status = ttk.Label(self, text="就绪", anchor=tk.W)
        self.status.pack(fill=tk.X, padx=5, pady=2)

        # Detail panels (lazy-built)
        self._panels = {}
        self._current_panel = None

        # Initialize: hide "记忆本次设置" and "重置为默认" unless "修改模板"
        if self.controller.current_template_id != "__custom__":
            self._remember_cb.pack_forget()
            self._reset_btn.pack_forget()

    def _get_or_build_panel(self, key: str):
        if key not in self._panels:
            builder = getattr(self, f"_build_{key}_panel", self._build_placeholder_panel)
            panel = builder()
            # Apply current edit state to newly built panel
            state = "normal" if self.chk_edit.get() else "disabled"
            self._set_edit_state_recursive(panel, state)
            self._panels[key] = panel
        return self._panels[key]

    def _build_placeholder_panel(self):
        frame = ttk.Frame(self.detail_frame)
        ttk.Label(frame, text="请在左侧选择你想要修改的标题级别").pack(pady=20)
        return frame

    def _build_body_panel(self):
        frame = ttk.Frame(self.detail_frame)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="正文格式", font=("Microsoft YaHei", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=10)

        row = 1
        ttk.Label(frame, text="中文字体:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_body_font = ttk.Combobox(frame, values=["宋体", "黑体", "楷体", "仿宋"], state="readonly")
        self.cmb_body_font.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="字号:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_body_size = ttk.Combobox(frame, values=["初号", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五"], state="readonly")
        self.cmb_body_size.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="行距(倍):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_body_spacing = ttk.Spinbox(frame, from_=1.0, to=3.0, increment=0.5)
        self.spn_body_spacing.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="对齐方式:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_body_align = ttk.Combobox(frame, values=["左对齐", "居中", "右对齐", "两端对齐"], state="readonly")
        self.cmb_body_align.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="段前行数:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_body_before = ttk.Spinbox(frame, from_=0, to=3, increment=0.5)
        self.spn_body_before.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="段后行数:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_body_after = ttk.Spinbox(frame, from_=0, to=3, increment=0.5)
        self.spn_body_after.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="左缩进(cm):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_body_left = ttk.Spinbox(frame, from_=0.0, to=5.0, increment=0.1)
        self.spn_body_left.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="右缩进(cm):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_body_right = ttk.Spinbox(frame, from_=0.0, to=5.0, increment=0.1)
        self.spn_body_right.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="特殊缩进:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_body_special = ttk.Combobox(frame, values=["无", "首行缩进", "悬挂缩进"], state="readonly")
        self.cmb_body_special.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="缩进值(字符):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_body_special = ttk.Spinbox(frame, from_=0.0, to=10.0, increment=0.5)
        self.spn_body_special.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        self.var_body_modify = tk.BooleanVar()
        ttk.Checkbutton(frame, text="同时修正内容（括号、句号）", variable=self.var_body_modify).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        return frame

    def _build_title_panel(self, level: int = 1):
        frame = ttk.Frame(self.detail_frame)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text=f"{level}级标题格式", font=("Microsoft YaHei", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=10)

        row = 1
        ttk.Label(frame, text="中文字体:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        cmb_font = ttk.Combobox(frame, values=["宋体", "黑体", "楷体", "仿宋", "方正小标宋简体"], state="readonly")
        cmb_font.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)
        setattr(self, f"cmb_title_{level}_font", cmb_font)

        row += 1
        ttk.Label(frame, text="字号:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        cmb_size = ttk.Combobox(frame, values=["三号", "小三", "四号", "小四", "五号"], state="readonly")
        cmb_size.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)
        setattr(self, f"cmb_title_{level}_size", cmb_size)

        row += 1
        ttk.Label(frame, text="加粗:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        var_bold = tk.BooleanVar()
        chk_bold = ttk.Checkbutton(frame, variable=var_bold)
        chk_bold.grid(row=row, column=1, sticky=tk.W, padx=5, pady=3)
        setattr(self, f"var_title_{level}_bold", var_bold)

        row += 1
        ttk.Label(frame, text="段前行数:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        spn_before = ttk.Spinbox(frame, from_=0, to=3, increment=0.5)
        spn_before.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)
        setattr(self, f"spn_title_{level}_before", spn_before)

        row += 1
        ttk.Label(frame, text="段后行数:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        spn_after = ttk.Spinbox(frame, from_=0, to=3, increment=0.5)
        spn_after.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)
        setattr(self, f"spn_title_{level}_after", spn_after)

        row += 1
        ttk.Label(frame, text="行距(倍):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        spn_spacing = ttk.Spinbox(frame, from_=1.0, to=3.0, increment=0.5)
        spn_spacing.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)
        setattr(self, f"spn_title_{level}_spacing", spn_spacing)

        row += 1
        ttk.Label(frame, text="对齐方式:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        cmb_align = ttk.Combobox(frame, values=["左对齐", "居中", "右对齐", "两端对齐"], state="readonly")
        cmb_align.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)
        setattr(self, f"cmb_title_{level}_align", cmb_align)

        return frame

    def _build_table_panel(self):
        frame = ttk.Frame(self.detail_frame)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="表格格式", font=("Microsoft YaHei", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=10)

        row = 1
        ttk.Label(frame, text="标题字体:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_table_title_font = ttk.Combobox(frame, values=["宋体", "黑体", "楷体", "仿宋"], state="readonly")
        self.cmb_table_title_font.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题字号:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_table_title_size = ttk.Combobox(frame, values=["三号", "小三", "四号", "小四", "五号"], state="readonly")
        self.cmb_table_title_size.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题加粗:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.var_table_title_bold = tk.BooleanVar()
        ttk.Checkbutton(frame, variable=self.var_table_title_bold).grid(row=row, column=1, sticky=tk.W, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题行距(倍):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_table_title_spacing = ttk.Spinbox(frame, from_=1.0, to=3.0, increment=0.5)
        self.spn_table_title_spacing.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题对齐:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_table_title_align = ttk.Combobox(frame, values=["左对齐", "居中", "右对齐", "两端对齐"], state="readonly")
        self.cmb_table_title_align.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题段前(行):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_table_title_before = ttk.Spinbox(frame, from_=0.0, to=3.0, increment=0.5)
        self.spn_table_title_before.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题段后(行):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_table_title_after = ttk.Spinbox(frame, from_=0.0, to=3.0, increment=0.5)
        self.spn_table_title_after.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="正文字体:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_table_body_font = ttk.Combobox(frame, values=["宋体", "黑体", "楷体", "仿宋"], state="readonly")
        self.cmb_table_body_font.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="正文字号:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_table_body_size = ttk.Combobox(frame, values=["三号", "小三", "四号", "小四", "五号"], state="readonly")
        self.cmb_table_body_size.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="线宽(pt):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_table_line_width = ttk.Spinbox(frame, from_=0.25, to=3.0, increment=0.25)
        self.spn_table_line_width.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="行高(cm):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_table_row_height = ttk.Spinbox(frame, from_=0.3, to=2.0, increment=0.1)
        self.spn_table_row_height.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="行距(倍):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_table_spacing = ttk.Spinbox(frame, from_=1.0, to=3.0, increment=0.5)
        self.spn_table_spacing.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="列宽模式:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_table_layout = ttk.Combobox(frame, values=["适应窗口", "适应内容", "固定列宽"], state="readonly")
        self.cmb_table_layout.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        return frame

    def _build_figure_panel(self):
        frame = ttk.Frame(self.detail_frame)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="图片标题格式", font=("Microsoft YaHei", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=10)

        row = 1
        ttk.Label(frame, text="标题字体:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_figure_title_font = ttk.Combobox(frame, values=["宋体", "黑体", "楷体", "仿宋"], state="readonly")
        self.cmb_figure_title_font.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题字号:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_figure_title_size = ttk.Combobox(frame, values=["三号", "小三", "四号", "小四", "五号"], state="readonly")
        self.cmb_figure_title_size.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题加粗:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.var_figure_title_bold = tk.BooleanVar()
        ttk.Checkbutton(frame, variable=self.var_figure_title_bold).grid(row=row, column=1, sticky=tk.W, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题行距(倍):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_figure_title_spacing = ttk.Spinbox(frame, from_=1.0, to=3.0, increment=0.5)
        self.spn_figure_title_spacing.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题对齐:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.cmb_figure_title_align = ttk.Combobox(frame, values=["左对齐", "居中", "右对齐", "两端对齐"], state="readonly")
        self.cmb_figure_title_align.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题段前(行):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_figure_title_before = ttk.Spinbox(frame, from_=0.0, to=3.0, increment=0.5)
        self.spn_figure_title_before.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        row += 1
        ttk.Label(frame, text="标题段后(行):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_figure_title_after = ttk.Spinbox(frame, from_=0.0, to=3.0, increment=0.5)
        self.spn_figure_title_after.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

        return frame

    def _build_advanced_panel(self):
        frame = ttk.Frame(self.detail_frame)
        ttk.Label(frame, text="高级设置", font=("Microsoft YaHei", 12, "bold")).pack(anchor=tk.W, pady=10)

        self.var_auto_detect = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="自动识别数字标题", variable=self.var_auto_detect).pack(anchor=tk.W, pady=5)

        self.var_include_list = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="包含列表段落（自动编号标题识别，无 Word COM 时功能受限）", variable=self.var_include_list).pack(anchor=tk.W, pady=5)

        self.var_com_resolver = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="使用 Word COM 读取列表级别（需安装 Microsoft Word）", variable=self.var_com_resolver).pack(anchor=tk.W, pady=5)

        # COM availability indicator
        com_frame = ttk.Frame(frame)
        com_frame.pack(anchor=tk.W, pady=5)
        ttk.Label(com_frame, text="Word COM 状态:").pack(side=tk.LEFT)
        self.lbl_com_status = ttk.Label(com_frame, text="检测中...")
        self.lbl_com_status.pack(side=tk.LEFT, padx=5)
        self.after(100, self._check_com_availability)

        # Separator
        ttk.Separator(frame, orient="horizontal").pack(fill=tk.X, pady=10)

        # Preset management
        ttk.Label(frame, text="预设管理", font=("Microsoft YaHei", 10, "bold")).pack(anchor=tk.W, pady=5)

        preset_row = ttk.Frame(frame)
        preset_row.pack(anchor=tk.W, pady=5, fill=tk.X)
        ttk.Label(preset_row, text="预设名称:").pack(side=tk.LEFT)
        self.ent_preset_name = ttk.Entry(preset_row, width=18)
        self.ent_preset_name.pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_row, text="保存", command=self._on_save_preset).pack(side=tk.LEFT, padx=2)

        load_row = ttk.Frame(frame)
        load_row.pack(anchor=tk.W, pady=5, fill=tk.X)
        ttk.Label(load_row, text="加载预设:").pack(side=tk.LEFT)
        self.cmb_presets = ttk.Combobox(load_row, state="readonly", width=16)
        self.cmb_presets.pack(side=tk.LEFT, padx=5)
        ttk.Button(load_row, text="加载", command=self._on_load_preset).pack(side=tk.LEFT, padx=2)
        ttk.Button(load_row, text="刷新列表", command=self._refresh_preset_list).pack(side=tk.LEFT, padx=2)

        # Refresh preset list on panel build
        self._refresh_preset_list()

        return frame

    def _check_com_availability(self):
        """Check if Word COM is available and update the indicator."""
        try:
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Quit()
            self.lbl_com_status.config(text="可用", foreground="green")
        except Exception:
            self.lbl_com_status.config(text="不可用 (自动编号标题可能无法识别)", foreground="red")

    def _on_save_preset(self):
        name = self.ent_preset_name.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入预设名称")
            return
        try:
            self._sync_settings_to_controller()
        except Exception:
            pass
        if self.controller.save_preset(name):
            self._refresh_preset_list()
            self.status.config(text=f"已保存预设: {name}")
            messagebox.showinfo("完成", f"预设 '{name}' 已保存")
        else:
            messagebox.showerror("错误", "保存预设失败")

    def _on_load_preset(self):
        name = self.cmb_presets.get()
        if not name:
            messagebox.showwarning("提示", "请选择要加载的预设")
            return
        if self.controller.load_preset(name):
            self._populate_from_settings()
            self.status.config(text=f"已加载预设: {name}")
        else:
            messagebox.showerror("错误", f"加载预设 '{name}' 失败")

    def _refresh_preset_list(self):
        presets = self.controller.list_presets()
        if hasattr(self, 'cmb_presets'):
            self.cmb_presets["values"] = presets
            if presets:
                self.cmb_presets.current(0)

    def _on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]

        if self._current_panel:
            self._current_panel.pack_forget()

        is_new = False
        if item == "body":
            is_new = "body" not in self._panels
            self._current_panel = self._get_or_build_panel("body")
        elif item.startswith("title_"):
            level = int(item.split("_")[1])
            key = f"title_{level}"
            is_new = key not in self._panels
            if is_new:
                self._panels[key] = self._build_title_panel(level)
                state = "normal" if self.chk_edit.get() else "disabled"
                self._set_edit_state_recursive(self._panels[key], state)
            self._current_panel = self._panels[key]
        elif item == "table":
            is_new = "table" not in self._panels
            self._current_panel = self._get_or_build_panel("table")
        elif item == "figure":
            is_new = "figure" not in self._panels
            self._current_panel = self._get_or_build_panel("figure")
        elif item == "advanced":
            is_new = "advanced" not in self._panels
            self._current_panel = self._get_or_build_panel("advanced")
        else:
            self._current_panel = self._build_placeholder_panel()

        if self._current_panel:
            self._current_panel.pack(fill=tk.BOTH, expand=True)

        # Populate newly created panels from controller settings so they
        # don't have empty values that crash _sync_settings_to_controller.
        if is_new:
            self._populate_from_settings()

    def _populate_from_settings(self):
        s = self.controller.settings
        # Body
        if hasattr(self, "cmb_body_font"):
            self.cmb_body_font.set(s.body.font)
            self.cmb_body_size.set(s.body.size)
            self.spn_body_spacing.set(str(s.body.spacing))
            self.cmb_body_align.set(s.body.alignment)
            self.spn_body_before.set(str(s.body.before_lines))
            self.spn_body_after.set(str(s.body.after_lines))
            self.spn_body_left.set(str(s.body.left_indent_cm))
            self.spn_body_right.set(str(s.body.right_indent_cm))
            self.cmb_body_special.set(s.body.special_indent)
            self.spn_body_special.set(str(s.body.special_indent_chars))
            if hasattr(self, "var_body_modify"):
                self.var_body_modify.set(s.body.modify_content)

        # Titles
        for i in range(1, 6):
            title = s.titles[i - 1]
            if hasattr(self, f"cmb_title_{i}_font"):
                getattr(self, f"cmb_title_{i}_font").set(title.font)
                getattr(self, f"cmb_title_{i}_size").set(title.size)
                getattr(self, f"var_title_{i}_bold").set(title.bold)
                getattr(self, f"spn_title_{i}_before").set(str(title.before_lines))
                getattr(self, f"spn_title_{i}_after").set(str(title.after_lines))
                getattr(self, f"spn_title_{i}_spacing").set(str(title.line_spacing))
                getattr(self, f"cmb_title_{i}_align").set(title.alignment)

        # Table
        if hasattr(self, "cmb_table_title_font"):
            self.cmb_table_title_font.set(s.table.title_font)
            self.cmb_table_title_size.set(s.table.title_size)
            self.var_table_title_bold.set(s.table.title_bold)
            self.spn_table_title_spacing.set(str(s.table.title_spacing))
            self.cmb_table_title_align.set(s.table.title_alignment)
            self.spn_table_title_before.set(str(s.table.title_before_lines))
            self.spn_table_title_after.set(str(s.table.title_after_lines))
            self.cmb_table_body_font.set(s.table.body_font)
            self.cmb_table_body_size.set(s.table.body_size)
            self.spn_table_line_width.set(str(s.table.line_width_pt))
            self.spn_table_row_height.set(str(s.table.row_height_cm))
            self.spn_table_spacing.set(str(s.table.spacing))
            layout_map = {"window": "适应窗口", "content": "适应内容", "fixed": "固定列宽"}
            self.cmb_table_layout.set(layout_map.get(s.table.auto_fit_mode, "适应窗口"))

        # Figure
        if hasattr(self, "cmb_figure_title_font"):
            self.cmb_figure_title_font.set(s.figure.title_font)
            self.cmb_figure_title_size.set(s.figure.title_size)
            self.var_figure_title_bold.set(s.figure.title_bold)
            self.spn_figure_title_spacing.set(str(s.figure.title_spacing))
            self.cmb_figure_title_align.set(s.figure.title_alignment)
            self.spn_figure_title_before.set(str(s.figure.title_before_lines))
            self.spn_figure_title_after.set(str(s.figure.title_after_lines))

        # Advanced
        if hasattr(self, "var_auto_detect"):
            self.var_auto_detect.set(s.auto_detect_numeric_titles)
            self.var_include_list.set(s.auto_detect_include_list_paragraphs)
            self.chk_remember.set(s.remember_settings)
            self.var_com_resolver.set(s.prefer_com_resolver)

    def _on_template_change(self, event):
        idx = self.cmb_template.current()
        if idx < 0 or idx >= len(self._template_ids):
            return
        template_id = self._template_ids[idx]
        self.controller.switch_template(template_id)
        is_custom = self.controller.is_custom_template()
        # Show "记忆本次设置" and "重置为默认" only for "修改模板"
        if is_custom:
            if not self._remember_cb.winfo_ismapped():
                self._remember_cb.pack(side=tk.LEFT, padx=5)
            if not self._reset_btn.winfo_ismapped():
                self._reset_btn.pack(side=tk.RIGHT, padx=5)
        else:
            self._remember_cb.pack_forget()
            self._reset_btn.pack_forget()
        # Enable editing automatically for 修改模板
        if is_custom and not self.chk_edit.get():
            self.chk_edit.set(True)
            self._on_edit_toggle()
        elif not is_custom and self.chk_edit.get():
            self.chk_edit.set(False)
            self._on_edit_toggle()
        self._populate_from_settings()
        self.status.config(text=f"已切换模板: {self.cmb_template.get()}")

    def _on_validate(self):
        if self.controller.opened_file is None:
            messagebox.showwarning("提示", "请先打开一个 Word 文档")
            return

        if self.chk_edit.get():
            try:
                self._sync_settings_to_controller()
            except Exception as e:
                self.status.config(text="错误: 同步设置失败")
                return

        self.status.config(text="正在检查格式（3个模板）...")
        self.progress["value"] = 0
        self.update()

        from tvba_core_validate import validate_document
        template_ids = self.controller.get_all_template_ids()
        all_issues = []
        template_names = []
        for i, tid in enumerate(template_ids):
            settings = self.controller.load_template_for_validation(tid)
            name = self._template_names.get(tid, tid)
            template_names.append(name)
            self.status.config(text=f"检查 {name} 模板...")
            self.progress["value"] = int((i / len(template_ids)) * 100)
            self.update()
            issues = validate_document(
                self.controller.opened_file,
                settings,
                progress_cb=None
            )
            if issues:
                for issue in issues:
                    all_issues.append((name, issue))

        if not all_issues:
            self.status.config(text="格式检查通过（所有模板）")
            messagebox.showinfo("格式检查", "未发现问题，所有模板格式检查通过！")
        else:
            self.status.config(text=f"发现 {len(all_issues)} 个问题")
            lines = [f"共发现 {len(all_issues)} 个格式问题（跨 {len(template_ids)} 个模板）：", ""]
            current_tpl = None
            for tpl_name, issue in all_issues[:80]:
                if tpl_name != current_tpl:
                    current_tpl = tpl_name
                    lines.append(f"【{tpl_name}】")
                lines.append(f"  [{issue.severity}] {issue.description}")
                if issue.location:
                    lines.append(f"    位置: {issue.location}")
            if len(all_issues) > 80:
                lines.append(f"  ... 还有 {len(all_issues) - 80} 个问题")
            messagebox.showwarning("格式检查结果", "\n".join(lines))

    def _on_open(self):
        path = filedialog.askopenfilename(
            filetypes=[("Word documents", "*.doc;*.docx"), ("All files", "*.*")]
        )
        if path:
            p = Path(path)
            self.controller.open_file(p)
            self.lbl_file.config(text=str(p))
            self.status.config(text=f"已打开: {p.name}")

    def _on_apply(self):
        if self.chk_edit.get():
            try:
                self._sync_settings_to_controller()
            except Exception as e:
                self.status.config(text=f"错误: 同步设置失败")
                messagebox.showerror("错误", f"同步设置失败: {e}")
                return

        self.progress["value"] = 0
        self.status.config(text="正在应用...")
        self.update()

        def progress_cb(msg, pct):
            self.status.config(text=msg)
            self.progress["value"] = pct * 100
            self.update()

        result = self.controller.apply(
            save_settings=(self.chk_edit.get() and self.chk_remember.get()),
            progress_cb=progress_cb,
        )
        if result.success:
            status_text = f"完成: {result.output_path}"
            if result.warnings:
                status_text += f"  ({len(result.warnings)} 条警告)"
            self.status.config(text=status_text)
            if result.elapsed_ms < 1000:
                time_str = f"{result.elapsed_ms} 毫秒"
            else:
                time_str = f"{result.elapsed_ms / 1000:.2f} 秒"
            # Build completion message with warnings if any
            msg = f"格式刷新成功！\n耗时: {time_str}"
            if result.warnings:
                msg += f"\n\n⚠ 警告 ({len(result.warnings)}):"
                for w in result.warnings[:5]:
                    msg += f"\n  • {w}"
                if len(result.warnings) > 5:
                    msg += f"\n  ... 还有 {len(result.warnings) - 5} 条警告"
            os.startfile(str(result.output_path))
            messagebox.showinfo("完成", msg)
        else:
            self.status.config(text=f"错误: {result.message}")
            messagebox.showerror("错误", result.message)

    def _on_apply_close(self):
        self._on_apply()
        self.destroy()

    def _on_cancel(self):
        self.destroy()

    def _on_reset(self):
        self.controller.reset_to_template_defaults()
        self.controller.clear_saved_settings()
        self._populate_from_settings()
        self.status.config(text="已重置为默认值")

    def _set_edit_state_recursive(self, widget, state):
        """Recursively set state on all input widgets."""
        input_types = (ttk.Combobox, ttk.Spinbox, ttk.Checkbutton, tk.Checkbutton)
        if isinstance(widget, input_types):
            try:
                widget.config(state=state)
            except Exception:
                pass
        for child in widget.winfo_children():
            self._set_edit_state_recursive(child, state)

    def _on_edit_toggle(self):
        enabled = self.chk_edit.get()
        state = "normal" if enabled else "disabled"
        if enabled and self.controller.is_custom_template():
            self.controller.load_saved_settings()
        elif not enabled and not self.controller.is_custom_template():
            self.controller.reset_to_template_defaults()
        self._populate_from_settings()
        # Enable/disable all input widgets in all panels
        for panel in self._panels.values():
            self._set_edit_state_recursive(panel, state)
        # Also apply to current panel if not in _panels yet
        if self._current_panel:
            self._set_edit_state_recursive(self._current_panel, state)
        self.status.config(text="修改模式已" + ("开启" if enabled else "关闭"))

    def _sync_settings_to_controller(self):
        def _safe_float(widget):
            try:
                val = widget.get()
                return float(val) if val != "" else None
            except (ValueError, TypeError):
                return None

        def _safe_get(widget):
            try:
                return widget.get()
            except (AttributeError, tk.TclError):
                return None

        # Sync body settings
        if hasattr(self, "cmb_body_font"):
            val = _safe_get(self.cmb_body_font)
            if val:
                self.controller.update_setting("body.font", val)
            val = _safe_get(self.cmb_body_size)
            if val:
                self.controller.update_setting("body.size", val)
            fval = _safe_float(self.spn_body_spacing)
            if fval is not None:
                self.controller.update_setting("body.spacing", fval)
            val = _safe_get(self.cmb_body_align)
            if val:
                self.controller.update_setting("body.alignment", val)
            fval = _safe_float(self.spn_body_before)
            if fval is not None:
                self.controller.update_setting("body.before_lines", fval)
            fval = _safe_float(self.spn_body_after)
            if fval is not None:
                self.controller.update_setting("body.after_lines", fval)
            fval = _safe_float(self.spn_body_left)
            if fval is not None:
                self.controller.update_setting("body.left_indent_cm", fval)
            fval = _safe_float(self.spn_body_right)
            if fval is not None:
                self.controller.update_setting("body.right_indent_cm", fval)
            val = _safe_get(self.cmb_body_special)
            if val:
                self.controller.update_setting("body.special_indent", val)
            fval = _safe_float(self.spn_body_special)
            if fval is not None:
                self.controller.update_setting("body.special_indent_chars", fval)
            if hasattr(self, "var_body_modify"):
                self.controller.update_setting("body.modify_content", self.var_body_modify.get())

        # Sync title settings (all 5 levels)
        for i in range(1, 6):
            if hasattr(self, f"cmb_title_{i}_font"):
                val = _safe_get(getattr(self, f"cmb_title_{i}_font"))
                if val:
                    self.controller.update_setting(f"titles.{i - 1}.font", val)
                val = _safe_get(getattr(self, f"cmb_title_{i}_size"))
                if val:
                    self.controller.update_setting(f"titles.{i - 1}.size", val)
                self.controller.update_setting(f"titles.{i - 1}.bold", getattr(self, f"var_title_{i}_bold").get())
                fval = _safe_float(getattr(self, f"spn_title_{i}_before"))
                if fval is not None:
                    self.controller.update_setting(f"titles.{i - 1}.before_lines", fval)
                fval = _safe_float(getattr(self, f"spn_title_{i}_after"))
                if fval is not None:
                    self.controller.update_setting(f"titles.{i - 1}.after_lines", fval)
                fval = _safe_float(getattr(self, f"spn_title_{i}_spacing"))
                if fval is not None:
                    self.controller.update_setting(f"titles.{i - 1}.line_spacing", fval)
                val = _safe_get(getattr(self, f"cmb_title_{i}_align"))
                if val:
                    self.controller.update_setting(f"titles.{i - 1}.alignment", val)

        # Sync table settings
        if hasattr(self, "cmb_table_title_font"):
            val = _safe_get(self.cmb_table_title_font)
            if val:
                self.controller.update_setting("table.title_font", val)
            val = _safe_get(self.cmb_table_title_size)
            if val:
                self.controller.update_setting("table.title_size", val)
            self.controller.update_setting("table.title_bold", self.var_table_title_bold.get())
            fval = _safe_float(self.spn_table_title_spacing)
            if fval is not None:
                self.controller.update_setting("table.title_spacing", fval)
            val = _safe_get(self.cmb_table_title_align)
            if val:
                self.controller.update_setting("table.title_alignment", val)
            fval = _safe_float(self.spn_table_title_before)
            if fval is not None:
                self.controller.update_setting("table.title_before_lines", fval)
            fval = _safe_float(self.spn_table_title_after)
            if fval is not None:
                self.controller.update_setting("table.title_after_lines", fval)
            val = _safe_get(self.cmb_table_body_font)
            if val:
                self.controller.update_setting("table.body_font", val)
            val = _safe_get(self.cmb_table_body_size)
            if val:
                self.controller.update_setting("table.body_size", val)
            fval = _safe_float(self.spn_table_line_width)
            if fval is not None:
                self.controller.update_setting("table.line_width_pt", fval)
            fval = _safe_float(self.spn_table_row_height)
            if fval is not None:
                self.controller.update_setting("table.row_height_cm", fval)
            fval = _safe_float(self.spn_table_spacing)
            if fval is not None:
                self.controller.update_setting("table.spacing", fval)
            layout_map = {"适应窗口": "window", "适应内容": "content", "固定列宽": "fixed"}
            layout_val = _safe_get(self.cmb_table_layout)
            if layout_val:
                self.controller.update_setting("table.auto_fit_mode", layout_map.get(layout_val, "window"))

        # Sync figure settings
        if hasattr(self, "cmb_figure_title_font"):
            val = _safe_get(self.cmb_figure_title_font)
            if val:
                self.controller.update_setting("figure.title_font", val)
            val = _safe_get(self.cmb_figure_title_size)
            if val:
                self.controller.update_setting("figure.title_size", val)
            self.controller.update_setting("figure.title_bold", self.var_figure_title_bold.get())
            fval = _safe_float(self.spn_figure_title_spacing)
            if fval is not None:
                self.controller.update_setting("figure.title_spacing", fval)
            val = _safe_get(self.cmb_figure_title_align)
            if val:
                self.controller.update_setting("figure.title_alignment", val)
            fval = _safe_float(self.spn_figure_title_before)
            if fval is not None:
                self.controller.update_setting("figure.title_before_lines", fval)
            fval = _safe_float(self.spn_figure_title_after)
            if fval is not None:
                self.controller.update_setting("figure.title_after_lines", fval)

        # Sync advanced settings
        if hasattr(self, "var_auto_detect"):
            self.controller.update_setting("auto_detect_numeric_titles", self.var_auto_detect.get())
            self.controller.update_setting("auto_detect_include_list_paragraphs", self.var_include_list.get())
            self.controller.update_setting("remember_settings", self.chk_remember.get())
            self.controller.update_setting("prefer_com_resolver", self.var_com_resolver.get())

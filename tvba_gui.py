"""Tkinter view layer for TransVBA.

Corresponds to VBA UserForm1.frm:
  - UserForm_Initialize
  - CreateContentPage / CreateTitlePage / CreateTablePage / CreateFigurePage
  - btnApply_Click / btnOK_Click / btnCancel_Click
  - LoadSettingsToForm / SetEditingEnabled
"""
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

        self.lbl_file = ttk.Label(top_frame, text="(未选择文件)")
        self.lbl_file.pack(side=tk.LEFT, padx=5)

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
        self.tree.insert("", "end", "titles", text="标题")
        for i in range(1, 6):
            self.tree.insert("titles", "end", f"title_{i}", text=f"  {i}级标题")
        self.tree.insert("", "end", "table", text="表格")
        self.tree.insert("", "end", "figure", text="图片标题")
        self.tree.insert("", "end", "advanced", text="高级")

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # Right: detail panel with scrollbar
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)

        self.detail_canvas = tk.Canvas(right_frame)
        scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=self.detail_canvas.yview)
        self.detail_frame = ttk.Frame(self.detail_canvas)

        self.detail_frame.bind(
            "<Configure>",
            lambda e: self.detail_canvas.configure(scrollregion=self.detail_canvas.bbox("all"))
        )

        self.detail_canvas.create_window((0, 0), window=self.detail_frame, anchor="nw", width=680)
        self.detail_canvas.configure(yscrollcommand=scrollbar.set)

        self.detail_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom bar
        bottom_frame = ttk.Frame(self, padding=5)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.chk_edit = tk.BooleanVar(value=False)
        ttk.Checkbutton(bottom_frame, text="修改模式", variable=self.chk_edit,
                       command=self._on_edit_toggle).pack(side=tk.LEFT, padx=5)

        self.chk_remember = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom_frame, text="记忆本次设置", variable=self.chk_remember).pack(side=tk.LEFT, padx=5)

        ttk.Button(bottom_frame, text="重置为默认", command=self._on_reset).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="取消", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="应用", command=self._on_apply).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="应用并关闭", command=self._on_apply_close).pack(side=tk.RIGHT, padx=5)

        # Progress bar
        self.progress = ttk.Progressbar(self, mode="determinate", maximum=100)
        self.progress.pack(fill=tk.X, padx=5, pady=2)

        self.status = ttk.Label(self, text="就绪", anchor=tk.W)
        self.status.pack(fill=tk.X, padx=5, pady=2)

        # Detail panels (lazy-built)
        self._panels = {}
        self._current_panel = None

    def _get_or_build_panel(self, key: str):
        if key not in self._panels:
            builder = getattr(self, f"_build_{key}_panel", self._build_placeholder_panel)
            self._panels[key] = builder()
        return self._panels[key]

    def _build_placeholder_panel(self):
        frame = ttk.Frame(self.detail_frame)
        ttk.Label(frame, text="(此面板尚未实现)").pack(pady=20)
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
        ttk.Label(frame, text="缩进值(cm):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.spn_body_special = ttk.Spinbox(frame, from_=0.0, to=3.0, increment=0.1)
        self.spn_body_special.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)

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
        ttk.Label(frame, text="自动适应窗口:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
        self.var_table_auto_fit = tk.BooleanVar()
        ttk.Checkbutton(frame, variable=self.var_table_auto_fit).grid(row=row, column=1, sticky=tk.W, padx=5, pady=3)

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

        return frame

    def _build_advanced_panel(self):
        frame = ttk.Frame(self.detail_frame)
        ttk.Label(frame, text="高级设置", font=("Microsoft YaHei", 12, "bold")).pack(anchor=tk.W, pady=10)

        self.var_auto_detect = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="自动识别数字标题", variable=self.var_auto_detect).pack(anchor=tk.W, pady=5)

        self.var_include_list = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="包含列表段落", variable=self.var_include_list).pack(anchor=tk.W, pady=5)

        return frame

    def _on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]

        if self._current_panel:
            self._current_panel.pack_forget()

        if item == "body":
            self._current_panel = self._get_or_build_panel("body")
        elif item.startswith("title_"):
            level = int(item.split("_")[1])
            key = f"title_{level}"
            if key not in self._panels:
                self._panels[key] = self._build_title_panel(level)
            self._current_panel = self._panels[key]
        elif item == "table":
            self._current_panel = self._get_or_build_panel("table")
        elif item == "figure":
            self._current_panel = self._get_or_build_panel("figure")
        elif item == "advanced":
            self._current_panel = self._get_or_build_panel("advanced")
        else:
            self._current_panel = self._build_placeholder_panel()

        if self._current_panel:
            self._current_panel.pack(fill=tk.BOTH, expand=True)

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
            self.spn_body_special.set(str(s.body.special_indent_cm))

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
            self.cmb_table_body_font.set(s.table.body_font)
            self.cmb_table_body_size.set(s.table.body_size)
            self.spn_table_line_width.set(str(s.table.line_width_pt))
            self.spn_table_row_height.set(str(s.table.row_height_cm))
            self.spn_table_spacing.set(str(s.table.spacing))
            self.var_table_auto_fit.set(s.table.auto_fit_window)

        # Figure
        if hasattr(self, "cmb_figure_title_font"):
            self.cmb_figure_title_font.set(s.figure.title_font)
            self.cmb_figure_title_size.set(s.figure.title_size)
            self.var_figure_title_bold.set(s.figure.title_bold)
            self.spn_figure_title_spacing.set(str(s.figure.title_spacing))

        # Advanced
        if hasattr(self, "var_auto_detect"):
            self.var_auto_detect.set(s.auto_detect_numeric_titles)
            self.var_include_list.set(s.auto_detect_include_list_paragraphs)
            self.chk_remember.set(s.remember_settings)

    def _on_open(self):
        path = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])
        if path:
            p = Path(path)
            self.controller.open_file(p)
            self.lbl_file.config(text=str(p))
            self.status.config(text=f"已打开: {p.name}")

    def _on_apply(self):
        self._sync_settings_to_controller()
        self.progress["value"] = 0
        self.status.config(text="正在应用...")
        self.update()

        def progress_cb(msg, pct):
            self.status.config(text=msg)
            self.progress["value"] = pct * 100
            self.update()

        result = self.controller.apply(
            save_settings=self.chk_remember.get(),
            progress_cb=progress_cb,
        )
        if result.success:
            self.status.config(text=f"完成: {result.output_path}")
            messagebox.showinfo("完成", "格式刷新成功！")
        else:
            self.status.config(text=f"错误: {result.message}")
            messagebox.showerror("错误", result.message)

    def _on_apply_close(self):
        self._on_apply()
        self.destroy()

    def _on_cancel(self):
        self.destroy()

    def _on_reset(self):
        self.controller.reset_to_defaults()
        self._populate_from_settings()
        self.status.config(text="已重置为默认值")

    def _on_edit_toggle(self):
        enabled = self.chk_edit.get()
        state = "normal" if enabled else "disabled"
        # TODO: Enable/disable all input widgets
        self.status.config(text="修改模式已" + ("开启" if enabled else "关闭"))

    def _sync_settings_to_controller(self):
        # Sync body settings
        if hasattr(self, "cmb_body_font"):
            self.controller.update_setting("body.font", self.cmb_body_font.get())
            self.controller.update_setting("body.size", self.cmb_body_size.get())
            self.controller.update_setting("body.spacing", float(self.spn_body_spacing.get()))
            self.controller.update_setting("body.alignment", self.cmb_body_align.get())
            self.controller.update_setting("body.before_lines", float(self.spn_body_before.get()))
            self.controller.update_setting("body.after_lines", float(self.spn_body_after.get()))
            self.controller.update_setting("body.left_indent_cm", float(self.spn_body_left.get()))
            self.controller.update_setting("body.right_indent_cm", float(self.spn_body_right.get()))
            self.controller.update_setting("body.special_indent", self.cmb_body_special.get())
            self.controller.update_setting("body.special_indent_cm", float(self.spn_body_special.get()))

        # Sync title settings (all 5 levels)
        for i in range(1, 6):
            if hasattr(self, f"cmb_title_{i}_font"):
                self.controller.update_setting(f"titles.{i - 1}.font", getattr(self, f"cmb_title_{i}_font").get())
                self.controller.update_setting(f"titles.{i - 1}.size", getattr(self, f"cmb_title_{i}_size").get())
                self.controller.update_setting(f"titles.{i - 1}.bold", getattr(self, f"var_title_{i}_bold").get())
                self.controller.update_setting(f"titles.{i - 1}.before_lines", float(getattr(self, f"spn_title_{i}_before").get()))
                self.controller.update_setting(f"titles.{i - 1}.after_lines", float(getattr(self, f"spn_title_{i}_after").get()))
                self.controller.update_setting(f"titles.{i - 1}.line_spacing", float(getattr(self, f"spn_title_{i}_spacing").get()))
                self.controller.update_setting(f"titles.{i - 1}.alignment", getattr(self, f"cmb_title_{i}_align").get())

        # Sync table settings
        if hasattr(self, "cmb_table_title_font"):
            self.controller.update_setting("table.title_font", self.cmb_table_title_font.get())
            self.controller.update_setting("table.title_size", self.cmb_table_title_size.get())
            self.controller.update_setting("table.title_bold", self.var_table_title_bold.get())
            self.controller.update_setting("table.title_spacing", float(self.spn_table_title_spacing.get()))
            self.controller.update_setting("table.body_font", self.cmb_table_body_font.get())
            self.controller.update_setting("table.body_size", self.cmb_table_body_size.get())
            self.controller.update_setting("table.line_width_pt", float(self.spn_table_line_width.get()))
            self.controller.update_setting("table.row_height_cm", float(self.spn_table_row_height.get()))
            self.controller.update_setting("table.spacing", float(self.spn_table_spacing.get()))
            self.controller.update_setting("table.auto_fit_window", self.var_table_auto_fit.get())

        # Sync figure settings
        if hasattr(self, "cmb_figure_title_font"):
            self.controller.update_setting("figure.title_font", self.cmb_figure_title_font.get())
            self.controller.update_setting("figure.title_size", self.cmb_figure_title_size.get())
            self.controller.update_setting("figure.title_bold", self.var_figure_title_bold.get())
            self.controller.update_setting("figure.title_spacing", float(self.spn_figure_title_spacing.get()))

        # Sync advanced settings
        if hasattr(self, "var_auto_detect"):
            self.controller.update_setting("auto_detect_numeric_titles", self.var_auto_detect.get())
            self.controller.update_setting("auto_detect_include_list_paragraphs", self.var_include_list.get())
            self.controller.update_setting("remember_settings", self.chk_remember.get())

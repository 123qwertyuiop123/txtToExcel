"""桌面图形界面；仅负责交互，业务处理交给各转换模块。"""

import sys
from pathlib import Path

from expense_converter import FILE_DATE_RE, convert_many, validate_many
from game_converter import convert_game_txt_files, validate_game_txt_files
from year_merge import merge_yearly_workbooks, validate_yearly_workbooks


def _file_snapshot(paths: list[Path]) -> tuple[tuple[str, int, int], ...]:
    """记录路径、修改时间和大小，防止检查通过后源文件又被修改。"""
    states: list[tuple[str, int, int]] = []
    for path in paths:
        stat = path.stat()
        states.append((str(path.resolve()), stat.st_mtime_ns, stat.st_size))
    return tuple(states)


def _show_error_list(parent, title: str, intro: str, errors: list[str]) -> None:
    """用可滚动窗口展示多个文件或多行错误。"""
    import tkinter as tk
    from tkinter import scrolledtext, ttk

    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.geometry("720x460")
    dialog.minsize(580, 360)
    dialog.transient(parent)
    dialog.grab_set()
    body = ttk.Frame(dialog, padding=18)
    body.pack(fill="both", expand=True)
    ttk.Label(body, text=intro, font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
    details = scrolledtext.ScrolledText(body, wrap="word", font=("Microsoft YaHei UI", 10))
    details.pack(fill="both", expand=True, pady=10)
    details.insert("1.0", "\n\n".join(f"{index}. {error}" for index, error in enumerate(errors, 1)))
    details.configure(state="disabled")
    ttk.Button(body, text="确定", command=dialog.destroy).pack(anchor="e")


def launch_year_merge_gui(parent) -> None:
    """打开多年度 Excel 合并窗口。"""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    window = tk.Toplevel(parent)
    window.title("合并年度消费统计 Excel")
    window.geometry("720x330")
    window.minsize(620, 300)
    window.transient(parent)
    selected_files: list[Path] = []
    input_var = tk.StringVar()
    output_var = tk.StringVar()
    status_var = tk.StringVar(value="请选择两个或更多年度消费统计 Excel")

    frame = ttk.Frame(window, padding=24)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    ttk.Label(frame, text="合并年度消费统计", font=("Microsoft YaHei UI", 16, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 6)
    )
    ttk.Label(frame, text="只保留每个年份的总情况，并生成跨年份汇总表。", foreground="#5F6B7A").grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(0, 20)
    )

    def choose_sources() -> None:
        """选择多个年度工作簿并自动建议输出路径。"""
        filenames = filedialog.askopenfilenames(
            title="选择年度消费统计 Excel",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if filenames:
            selected_files.clear()
            selected_files.extend(Path(filename) for filename in filenames)
            input_var.set("；".join(path.name for path in selected_files))
            output_var.set(str(selected_files[0].parent / "年份统计.xlsx"))
            status_var.set(f"已选择 {len(selected_files)} 个年度文件")

    def choose_output() -> None:
        """选择合并结果的保存位置。"""
        initial = Path(output_var.get()) if output_var.get().strip() else None
        filename = filedialog.asksaveasfilename(
            title="保存年份统计 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            initialdir=str(initial.parent) if initial else None,
            initialfile=initial.name if initial else "年份统计.xlsx",
        )
        if filename:
            output_var.set(filename)

    def merge_files() -> None:
        """检查输入，确认覆盖后执行年度合并。"""
        if len(selected_files) < 2:
            messagebox.showwarning("文件不足", "请至少选择两个不同年份的消费统计 Excel。", parent=window)
            return
        if not output_var.get().strip():
            messagebox.showwarning("缺少输出位置", "请选择年份统计 Excel 的保存位置。", parent=window)
            return
        errors = validate_yearly_workbooks(selected_files)
        if errors:
            status_var.set(f"检查未通过：发现 {len(errors)} 个问题")
            _show_error_list(window, "年度文件检查未通过", f"发现 {len(errors)} 个问题：", errors)
            return
        output = Path(output_var.get().strip())
        if output.exists() and not messagebox.askyesno("确认覆盖", f"文件已存在，是否覆盖？\n{output}", parent=window):
            return
        merge_button.state(["disabled"])
        status_var.set("正在合并，请稍候…")
        window.update_idletasks()
        try:
            result = merge_yearly_workbooks(selected_files, output)
        except (OSError, ValueError, RuntimeError) as exc:
            status_var.set("合并失败")
            messagebox.showerror("合并失败", str(exc), parent=window)
        else:
            status_var.set(f"合并完成：{result.name}")
            messagebox.showinfo("合并完成", f"年份统计已保存到：\n{result.resolve()}", parent=window)
        finally:
            merge_button.state(["!disabled"])

    ttk.Label(frame, text="年度文件").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=7)
    ttk.Entry(frame, textvariable=input_var, state="readonly").grid(row=2, column=1, sticky="ew", pady=7)
    ttk.Button(frame, text="多选…", command=choose_sources).grid(row=2, column=2, padx=(10, 0), pady=7)
    ttk.Label(frame, text="输出文件").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=7)
    ttk.Entry(frame, textvariable=output_var).grid(row=3, column=1, sticky="ew", pady=7)
    ttk.Button(frame, text="选择…", command=choose_output).grid(row=3, column=2, padx=(10, 0), pady=7)
    ttk.Separator(frame).grid(row=4, column=0, columnspan=3, sticky="ew", pady=(18, 14))
    ttk.Label(frame, textvariable=status_var, foreground="#5F6B7A").grid(row=5, column=0, columnspan=2, sticky="w")
    merge_button = ttk.Button(frame, text="检查并合并", command=merge_files)
    merge_button.grid(row=5, column=2, sticky="e")


def launch_game_gui(parent) -> None:
    """打开年度游戏消费 TXT 转换窗口。"""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    window = tk.Toplevel(parent)
    window.title("年度游戏消费 TXT 转 Excel")
    window.geometry("740x360")
    window.minsize(640, 330)
    window.transient(parent)
    selected_files: list[Path] = []
    validated_snapshot: tuple[tuple[str, int, int], ...] | None = None
    input_var = tk.StringVar()
    output_var = tk.StringVar()
    status_var = tk.StringVar(value="请选择一个或多个年度游戏消费 TXT")

    frame = ttk.Frame(window, padding=24)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    ttk.Label(frame, text="年度游戏消费 TXT 转 Excel", font=("Microsoft YaHei UI", 16, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 4)
    )
    ttk.Label(
        frame,
        text="支持 2025年游戏.txt；可一次选择多个年份，检查通过后生成同一个工作簿。",
        foreground="#5F6B7A",
    ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 20))

    def choose_sources() -> None:
        """选择游戏 TXT；重新选择后立即使旧检查结果失效。"""
        nonlocal validated_snapshot
        filenames = filedialog.askopenfilenames(title="选择年度游戏消费 TXT", filetypes=[("文本文件", "*.txt")])
        if filenames:
            selected_files.clear()
            selected_files.extend(Path(filename) for filename in filenames)
            input_var.set("；".join(path.name for path in selected_files))
            output_var.set(str(selected_files[0].parent / "游戏消费统计.xlsx"))
            validated_snapshot = None
            convert_button.state(["disabled"])
            status_var.set(f"已选择 {len(selected_files)} 个文件，请先检查")

    def choose_output() -> None:
        """选择游戏消费工作簿的保存位置。"""
        initial = Path(output_var.get()) if output_var.get().strip() else None
        filename = filedialog.asksaveasfilename(
            title="保存游戏消费统计 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            initialdir=str(initial.parent) if initial else None,
            initialfile=initial.name if initial else "游戏消费统计.xlsx",
        )
        if filename:
            output_var.set(filename)

    def check_files() -> None:
        """严格检查全部文件，并保存检查通过时的文件状态。"""
        nonlocal validated_snapshot
        if not selected_files:
            messagebox.showwarning("缺少文件", "请先选择年度游戏消费 TXT。", parent=window)
            return
        errors = validate_game_txt_files(selected_files)
        if errors:
            validated_snapshot = None
            convert_button.state(["disabled"])
            status_var.set(f"检查未通过：发现 {len(errors)} 个问题")
            _show_error_list(
                window,
                "游戏消费文件检查未通过",
                f"发现 {len(errors)} 个问题，请修改 TXT 后重新检查：",
                errors,
            )
            return
        validated_snapshot = _file_snapshot(selected_files)
        convert_button.state(["!disabled"])
        status_var.set(f"检查通过：{len(selected_files)} 个文件可以转换")
        messagebox.showinfo("检查通过", "所有游戏消费文件均符合转换标准。", parent=window)

    def run_conversion() -> None:
        """仅在文件状态与检查快照一致时执行转换。"""
        if not selected_files:
            messagebox.showwarning("缺少文件", "请先选择年度游戏消费 TXT。", parent=window)
            return
        try:
            current = _file_snapshot(selected_files)
        except OSError as exc:
            messagebox.showerror("文件不可用", str(exc), parent=window)
            return
        if validated_snapshot is None or current != validated_snapshot:
            convert_button.state(["disabled"])
            messagebox.showwarning("请先检查文件", "文件尚未检查，或检查后内容发生了变化。请重新检查。", parent=window)
            return
        if not output_var.get().strip():
            messagebox.showwarning("缺少输出位置", "请选择 Excel 保存位置。", parent=window)
            return
        output = Path(output_var.get().strip())
        if output.exists() and not messagebox.askyesno("确认覆盖", f"文件已存在，是否覆盖？\n{output}", parent=window):
            return
        convert_button.state(["disabled"])
        status_var.set("正在转换，请稍候…")
        window.update_idletasks()
        try:
            result = convert_game_txt_files(selected_files, output)
        except (OSError, ValueError, RuntimeError) as exc:
            status_var.set("转换失败")
            messagebox.showerror("转换失败", str(exc), parent=window)
        else:
            status_var.set(f"转换完成：{result.name}")
            messagebox.showinfo("转换完成", f"游戏消费统计已保存到：\n{result.resolve()}", parent=window)
        finally:
            if validated_snapshot == current:
                convert_button.state(["!disabled"])

    ttk.Label(frame, text="游戏 TXT").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=7)
    ttk.Entry(frame, textvariable=input_var, state="readonly").grid(row=2, column=1, sticky="ew", pady=7)
    ttk.Button(frame, text="多选…", command=choose_sources).grid(row=2, column=2, padx=(10, 0), pady=7)
    ttk.Label(frame, text="输出文件").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=7)
    ttk.Entry(frame, textvariable=output_var).grid(row=3, column=1, sticky="ew", pady=7)
    ttk.Button(frame, text="选择…", command=choose_output).grid(row=3, column=2, padx=(10, 0), pady=7)
    ttk.Separator(frame).grid(row=4, column=0, columnspan=3, sticky="ew", pady=(18, 14))
    ttk.Label(frame, textvariable=status_var, foreground="#5F6B7A").grid(row=5, column=0, sticky="w")
    ttk.Button(frame, text="检查文件", command=check_files).grid(row=5, column=1, sticky="e", padx=(0, 10))
    convert_button = ttk.Button(frame, text="开始转换", command=run_conversion)
    convert_button.grid(row=5, column=2, sticky="e")
    convert_button.state(["disabled"])


def launch_gui() -> int:
    """启动普通消费转换主窗口。"""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ImportError as exc:
        print(f"无法启动图形窗口：{exc}", file=sys.stderr)
        return 1

    root = tk.Tk()
    root.title("TXT 消费记录转 Excel")
    root.geometry("680x440")
    root.minsize(620, 410)
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    style.configure("Title.TLabel", font=("Microsoft YaHei UI", 16, "bold"))
    style.configure("Hint.TLabel", foreground="#5F6B7A")
    style.configure("Accent.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(16, 8))

    selected_files: list[Path] = []
    validated_snapshot: tuple[object, ...] | None = None
    input_var = tk.StringVar()
    output_var = tk.StringVar()
    year_var = tk.StringVar()
    month_var = tk.StringVar()
    status_var = tk.StringVar(value="请选择消费记录 TXT 文件")
    container = ttk.Frame(root, padding=24)
    container.pack(fill="both", expand=True)
    container.columnconfigure(1, weight=1)
    ttk.Label(container, text="TXT 消费记录转 Excel", style="Title.TLabel").grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 4)
    )
    ttk.Label(
        container,
        text="金额前带“+”表示收入；无符号或负数金额按支出统计。",
        style="Hint.TLabel",
    ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 20))

    def update_suggestions() -> None:
        """从所选文件名自动填写年月和默认输出文件名。"""
        matches = [FILE_DATE_RE.search(path.stem) for path in selected_files]
        years = {match.group("year") for match in matches if match}
        if len(years) == 1 and all(matches):
            selected_year = years.pop()
            year_var.set(selected_year)
            output_var.set(str(selected_files[0].parent / f"{selected_year}年消费统计.xlsx"))
        else:
            year_var.set("")
            output_var.set(str(selected_files[0].parent / "消费统计.xlsx"))
        month_var.set(matches[0].group("month") if len(selected_files) == 1 and matches[0] else "")

    def current_snapshot() -> tuple[object, ...]:
        """普通消费还需把手动年月加入快照，防止检查后被修改。"""
        return _file_snapshot(selected_files) + (year_var.get().strip(), month_var.get().strip())

    def choose_input() -> None:
        """选择一个或多个月份 TXT，并重置转换状态。"""
        nonlocal validated_snapshot
        filenames = filedialog.askopenfilenames(
            title="选择一个或多个月份的消费记录 TXT 文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if filenames:
            selected_files.clear()
            selected_files.extend(Path(filename) for filename in filenames)
            input_var.set("；".join(path.name for path in selected_files))
            update_suggestions()
            validated_snapshot = None
            convert_button.state(["disabled"])
            status_var.set(f"已选择 {len(selected_files)} 个文件，请先检查")

    def choose_output() -> None:
        """选择普通消费工作簿的保存位置。"""
        initial = Path(output_var.get()) if output_var.get().strip() else None
        filename = filedialog.asksaveasfilename(
            title="保存 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            initialdir=str(initial.parent) if initial else None,
            initialfile=initial.name if initial else "消费统计.xlsx",
        )
        if filename:
            output_var.set(filename)

    def read_year_month() -> tuple[int | None, int | None]:
        """把可选年月输入转换成整数，并统一报告非数字错误。"""
        try:
            year = int(year_var.get()) if year_var.get().strip() else None
            month = int(month_var.get()) if month_var.get().strip() else None
        except ValueError as exc:
            raise ValueError("年份和月份必须填写数字") from exc
        return year, month

    def check_files() -> None:
        """检查普通消费文件；全部通过后才启用转换按钮。"""
        nonlocal validated_snapshot
        if not selected_files:
            messagebox.showwarning("缺少文件", "请先选择一个或多个 TXT 文件。", parent=root)
            return
        try:
            year, month = read_year_month()
            errors = validate_many(selected_files, year, month if len(selected_files) == 1 else None)
        except (OSError, ValueError) as exc:
            errors = [str(exc)]
        if errors:
            validated_snapshot = None
            convert_button.state(["disabled"])
            status_var.set(f"检查未通过：发现 {len(errors)} 个问题")
            _show_error_list(root, "文件检查未通过", f"发现 {len(errors)} 个问题，请修改 TXT 后重新检查：", errors)
            return
        validated_snapshot = current_snapshot()
        convert_button.state(["!disabled"])
        status_var.set(f"检查通过：{len(selected_files)} 个文件可以转换")
        messagebox.showinfo("检查通过", "所有文件均符合转换标准，现在可以开始转换。", parent=root)

    def run_conversion() -> None:
        """核对检查快照、覆盖确认和年月后执行转换。"""
        if not selected_files:
            messagebox.showwarning("缺少文件", "请先选择一个或多个 TXT 文件。", parent=root)
            return
        try:
            snapshot = current_snapshot()
        except OSError as exc:
            messagebox.showerror("文件不可用", str(exc), parent=root)
            return
        if validated_snapshot is None or snapshot != validated_snapshot:
            convert_button.state(["disabled"])
            messagebox.showwarning("请先检查文件", "文件尚未检查，或检查后内容/年月发生了变化。请重新检查。", parent=root)
            return
        if not output_var.get().strip():
            messagebox.showwarning("缺少输出位置", "请选择 Excel 文件的保存位置。", parent=root)
            return
        output = Path(output_var.get().strip())
        if output.exists() and not messagebox.askyesno("确认覆盖", f"文件已存在，是否覆盖？\n{output}", parent=root):
            status_var.set("已取消转换")
            return
        try:
            year, month = read_year_month()
        except ValueError as exc:
            messagebox.showerror("年月错误", str(exc), parent=root)
            return
        convert_button.state(["disabled"])
        status_var.set("正在转换，请稍候…")
        root.update_idletasks()
        try:
            result = convert_many(selected_files, output, year, month if len(selected_files) == 1 else None)
        except (OSError, ValueError, RuntimeError) as exc:
            status_var.set("转换失败，请检查输入内容")
            messagebox.showerror("转换失败", str(exc), parent=root)
        else:
            status_var.set(f"转换完成：{result.name}")
            messagebox.showinfo("转换完成", f"Excel 文件已保存到：\n{result.resolve()}", parent=root)
        finally:
            convert_button.state(["!disabled"])

    ttk.Label(container, text="TXT 文件").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=6)
    ttk.Entry(container, textvariable=input_var, state="readonly").grid(row=2, column=1, sticky="ew", pady=6)
    ttk.Button(container, text="多选…", command=choose_input).grid(row=2, column=2, padx=(10, 0), pady=6)
    ttk.Label(container, text="输出文件").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=6)
    ttk.Entry(container, textvariable=output_var).grid(row=3, column=1, sticky="ew", pady=6)
    ttk.Button(container, text="选择…", command=choose_output).grid(row=3, column=2, padx=(10, 0), pady=6)
    ttk.Label(container, text="年月").grid(row=4, column=0, sticky="w", padx=(0, 12), pady=(8, 4))
    date_frame = ttk.Frame(container)
    date_frame.grid(row=4, column=1, sticky="w", pady=(8, 4))
    ttk.Entry(date_frame, textvariable=year_var, width=9).pack(side="left")
    ttk.Label(date_frame, text=" 年 ").pack(side="left")
    ttk.Entry(date_frame, textvariable=month_var, width=5).pack(side="left")
    ttk.Label(date_frame, text=" 月（多文件时按文件名自动识别月份）", style="Hint.TLabel").pack(side="left")
    ttk.Separator(container).grid(row=5, column=0, columnspan=3, sticky="ew", pady=(18, 14))

    action_frame = ttk.Frame(container)
    action_frame.grid(row=6, column=0, columnspan=3, sticky="ew")
    action_frame.columnconfigure(0, weight=1)
    ttk.Label(action_frame, textvariable=status_var, style="Hint.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Button(action_frame, text="检查文件", command=check_files).grid(row=0, column=1, sticky="e", padx=(0, 10))
    convert_button = ttk.Button(action_frame, text="开始转换", command=run_conversion, style="Accent.TButton")
    convert_button.grid(row=0, column=2, sticky="e")
    convert_button.state(["disabled"])

    ttk.Separator(container).grid(row=7, column=0, columnspan=3, sticky="ew", pady=(18, 12))
    ttk.Label(container, text="已有多个年度消费统计 Excel？", style="Hint.TLabel").grid(
        row=8, column=0, columnspan=2, sticky="w"
    )
    ttk.Button(container, text="合并年度 Excel…", command=lambda: launch_year_merge_gui(root)).grid(
        row=8, column=2, sticky="e"
    )
    ttk.Label(container, text="转换每年的游戏消费 TXT？", style="Hint.TLabel").grid(
        row=9, column=0, columnspan=2, sticky="w", pady=(10, 0)
    )
    ttk.Button(container, text="游戏消费 TXT…", command=lambda: launch_game_gui(root)).grid(
        row=9, column=2, sticky="e", pady=(10, 0)
    )
    root.bind("<Return>", lambda _event: run_conversion())
    root.mainloop()
    return 0

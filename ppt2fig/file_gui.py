import json
import os
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .cli import default_output_path
from .core import detect_backends, export_selected_pages, parse_page_range


FILE_MODE_HISTORY_PATH = Path.home() / ".ppt2fig_file_history.json"
MAX_HISTORY_ITEMS = 16


def _load_file_mode_history():
    try:
        if not FILE_MODE_HISTORY_PATH.exists():
            return []
        data = json.loads(FILE_MODE_HISTORY_PATH.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return [item for item in data if isinstance(item, dict)]
    except Exception:
        return []
    return []


def _save_file_mode_history(items):
    try:
        FILE_MODE_HISTORY_PATH.write_text(
            json.dumps(items[:MAX_HISTORY_ITEMS], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def _history_label(item):
    source = item.get("source", "")
    name = Path(source).name if source else "未命名"
    pages = item.get("pages", "")
    backend = item.get("backend", "auto")
    updated_at = item.get("updated_at", "")
    short_time = updated_at[5:16].replace("T", " ") if len(updated_at) >= 16 else ""
    prefix = f"[{short_time}] " if short_time else ""
    return f"{prefix}{name} | 页码 {pages} | {backend}"


def _record_current_settings(
    history_items,
    *,
    source,
    pages,
    output,
    backend,
    office_bin,
    powerpoint_intent,
    bitmap_missing_fonts,
    no_crop,
    percent_retain,
    margin_size,
    use_uniform,
    use_same_size,
    threshold,
):
    record = {
        "source": source,
        "pages": pages,
        "output": output,
        "backend": backend,
        "office_bin": office_bin,
        "powerpoint_intent": powerpoint_intent,
        "bitmap_missing_fonts": bitmap_missing_fonts,
        "no_crop": no_crop,
        "percent_retain": percent_retain,
        "margin_size": margin_size,
        "use_uniform": use_uniform,
        "use_same_size": use_same_size,
        "threshold": threshold,
        "updated_at": datetime.now().isoformat(timespec="minutes"),
    }
    normalized_source = os.path.normcase(os.path.abspath(source)) if source else ""
    filtered = []
    for item in history_items:
        item_source = os.path.normcase(os.path.abspath(item.get("source", ""))) if item.get("source") else ""
        if item_source == normalized_source and item.get("pages") == pages and item.get("backend") == backend:
            continue
        filtered.append(item)
    filtered.insert(0, record)
    return filtered[:MAX_HISTORY_ITEMS]


def _make_output_path(source_text, pages_text):
    source = Path(source_text).resolve()
    pages = parse_page_range(pages_text)
    return str(default_output_path(source, pages))


def main():
    root = tk.Tk()
    root.title(f"PPT2Fig 文件模式 v{__version__}")
    root.geometry("980x690")
    root.minsize(920, 640)

    source_var = tk.StringVar()
    pages_var = tk.StringVar(value="1")
    output_var = tk.StringVar()
    office_bin_var = tk.StringVar()
    backend_var = tk.StringVar(value="auto")
    powerpoint_intent_var = tk.StringVar(value="print")
    bitmap_missing_fonts_var = tk.BooleanVar(value=False)
    no_crop_var = tk.BooleanVar(value=False)
    percent_retain_var = tk.DoubleVar(value=0.0)
    margin_size_var = tk.DoubleVar(value=0.0)
    use_uniform_var = tk.BooleanVar(value=True)
    use_same_size_var = tk.BooleanVar(value=True)
    threshold_var = tk.IntVar(value=191)
    auto_output_var = tk.BooleanVar(value=True)
    status_var = tk.StringVar(value="选择文件后即可导出，常用配置会自动进入左侧历史记录。")
    backend_status_var = tk.StringVar(value="未检测后端")
    exporting = {"active": False}
    history_items = _load_file_mode_history()

    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")

    outer = ttk.Frame(root, padding=14)
    outer.pack(fill=tk.BOTH, expand=True)

    header = ttk.Frame(outer)
    header.pack(fill=tk.X, pady=(0, 10))
    title_row = ttk.Frame(header)
    title_row.pack(fill=tk.X)
    ttk.Label(title_row, text="PPT2Fig 文件模式", font=("Microsoft YaHei UI", 14, "bold")).pack(anchor="w", side=tk.LEFT)
    ttk.Label(title_row, text=f"v{__version__}", foreground="#666666").pack(anchor="e", side=tk.RIGHT)
    ttk.Label(
        header,
        text="面向重复导出场景：左侧保留最近任务，右侧编辑当前配置。",
        foreground="#555555",
    ).pack(anchor="w", pady=(2, 0))

    body = ttk.PanedWindow(outer, orient=tk.HORIZONTAL)
    body.pack(fill=tk.BOTH, expand=True)

    history_panel = ttk.Frame(body, padding=(0, 0, 12, 0))
    body.add(history_panel, weight=28)
    form_panel = ttk.Frame(body)
    body.add(form_panel, weight=72)

    ttk.Label(history_panel, text="最近任务", font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
    ttk.Label(
        history_panel,
        text="双击可直接重导，单击后可载入到右侧编辑。",
        foreground="#666666",
    ).pack(anchor="w", pady=(2, 8))

    history_list = tk.Listbox(history_panel, activestyle="none", height=18)
    history_list.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
    history_scroll = ttk.Scrollbar(history_panel, orient=tk.VERTICAL, command=history_list.yview)
    history_scroll.pack(fill=tk.Y, side=tk.LEFT)
    history_list.config(yscrollcommand=history_scroll.set)

    history_actions = ttk.Frame(history_panel)
    history_actions.pack(fill=tk.X, pady=(8, 0))

    selected_history_index = {"value": None}

    def refresh_history_list(select_index=0):
        history_list.delete(0, tk.END)
        for item in history_items:
            history_list.insert(tk.END, _history_label(item))
        if history_items:
            idx = min(select_index, len(history_items) - 1)
            history_list.selection_clear(0, tk.END)
            history_list.selection_set(idx)
            history_list.activate(idx)
            selected_history_index["value"] = idx
        else:
            selected_history_index["value"] = None

    def get_selected_history_item():
        selection = history_list.curselection()
        if not selection:
            return None
        idx = selection[0]
        selected_history_index["value"] = idx
        if 0 <= idx < len(history_items):
            return history_items[idx]
        return None

    def apply_history_item(item):
        if not item:
            return
        source_var.set(item.get("source", ""))
        pages_var.set(item.get("pages", "1"))
        output_var.set(item.get("output", ""))
        backend_var.set(item.get("backend", "auto"))
        office_bin_var.set(item.get("office_bin", ""))
        powerpoint_intent_var.set(item.get("powerpoint_intent", "print"))
        bitmap_missing_fonts_var.set(bool(item.get("bitmap_missing_fonts", False)))
        no_crop_var.set(bool(item.get("no_crop", False)))
        percent_retain_var.set(float(item.get("percent_retain", 0.0)))
        margin_size_var.set(float(item.get("margin_size", 0.0)))
        use_uniform_var.set(bool(item.get("use_uniform", True)))
        use_same_size_var.set(bool(item.get("use_same_size", True)))
        threshold_var.set(int(item.get("threshold", 191)))
        auto_output_var.set(False)
        update_dynamic_state()
        status_var.set("已载入历史任务，可直接修改后重新导出。")

    def delete_history_item():
        item = get_selected_history_item()
        if not item:
            return
        idx = selected_history_index["value"] or 0
        del history_items[idx]
        _save_file_mode_history(history_items)
        refresh_history_list(max(0, idx - 1))
        status_var.set("已删除一条历史任务。")

    ttk.Button(history_actions, text="载入", command=lambda: apply_history_item(get_selected_history_item())).pack(
        side=tk.LEFT
    )
    ttk.Button(history_actions, text="重导", command=lambda: start_export(from_history=True)).pack(
        side=tk.LEFT, padx=6
    )
    ttk.Button(history_actions, text="删除", command=delete_history_item).pack(side=tk.LEFT)

    history_list.bind("<<ListboxSelect>>", lambda _event: get_selected_history_item())
    history_list.bind(
        "<Double-Button-1>",
        lambda _event: (apply_history_item(get_selected_history_item()), start_export(from_history=True)),
    )

    current_label = ttk.Label(form_panel, text="当前配置", font=("Microsoft YaHei UI", 11, "bold"))
    current_label.pack(anchor="w")

    quick_actions = ttk.Frame(form_panel)
    quick_actions.pack(fill=tk.X, pady=(6, 10))

    def choose_source():
        selected = filedialog.askopenfilename(
            parent=root,
            title="选择输入文件",
            filetypes=[("Presentation", "*.pptx *.ppt *.odp"), ("All files", "*.*")],
        )
        if not selected:
            return
        source_var.set(selected)
        if auto_output_var.get():
            try:
                output_var.set(_make_output_path(selected, pages_var.get()))
            except Exception:
                pass
        status_var.set("已选择输入文件。")

    def choose_output():
        selected = filedialog.asksaveasfilename(
            parent=root,
            title="选择输出 PDF",
            defaultextension=".pdf",
            filetypes=[("PDF file", "*.pdf")],
            initialfile=os.path.basename(output_var.get()) if output_var.get() else "",
            initialdir=os.path.dirname(output_var.get()) if output_var.get() else "",
        )
        if selected:
            output_var.set(selected)
            auto_output_var.set(False)

    def choose_office_bin():
        selected = filedialog.askopenfilename(parent=root, title="选择后端可执行文件")
        if selected:
            office_bin_var.set(selected)
            refresh_backends()

    def set_output_from_input():
        if not source_var.get():
            return
        output_var.set(_make_output_path(source_var.get(), pages_var.get()))

    ttk.Button(quick_actions, text="选择文件", command=choose_source).pack(side=tk.LEFT)
    ttk.Button(quick_actions, text="输出同目录", command=set_output_from_input).pack(side=tk.LEFT, padx=6)
    ttk.Button(quick_actions, text="刷新后端", command=lambda: refresh_backends(verbose=True)).pack(
        side=tk.LEFT, padx=6
    )
    ttk.Checkbutton(
        quick_actions,
        text="自动生成输出文件名",
        variable=auto_output_var,
        command=lambda: set_output_from_input() if auto_output_var.get() and source_var.get() else None,
    ).pack(side=tk.RIGHT)

    basic_frame = ttk.LabelFrame(form_panel, text="基础信息", padding=12)
    basic_frame.pack(fill=tk.X)
    basic_frame.columnconfigure(1, weight=1)

    ttk.Label(basic_frame, text="输入文件").grid(row=0, column=0, sticky="w", pady=5)
    ttk.Entry(basic_frame, textvariable=source_var).grid(row=0, column=1, sticky="ew", padx=8, pady=5)
    ttk.Button(basic_frame, text="浏览", command=choose_source).grid(row=0, column=2, sticky="e", pady=5)

    ttk.Label(basic_frame, text="页码").grid(row=1, column=0, sticky="w", pady=5)
    pages_row = ttk.Frame(basic_frame)
    pages_row.grid(row=1, column=1, sticky="ew", padx=8, pady=5)
    pages_row.columnconfigure(0, weight=1)
    ttk.Entry(pages_row, textvariable=pages_var).grid(row=0, column=0, sticky="ew")
    ttk.Label(pages_row, text="支持 1,3,5-7", foreground="#666666").grid(row=0, column=1, sticky="w", padx=(8, 0))
    ttk.Button(basic_frame, text="第一页", command=lambda: pages_var.set("1")).grid(row=1, column=2, sticky="e", pady=5)

    ttk.Label(basic_frame, text="输出 PDF").grid(row=2, column=0, sticky="w", pady=5)
    ttk.Entry(basic_frame, textvariable=output_var).grid(row=2, column=1, sticky="ew", padx=8, pady=5)
    ttk.Button(basic_frame, text="浏览", command=choose_output).grid(row=2, column=2, sticky="e", pady=5)

    backend_frame = ttk.LabelFrame(form_panel, text="导出后端", padding=12)
    backend_frame.pack(fill=tk.X, pady=(10, 0))
    backend_frame.columnconfigure(1, weight=1)

    ttk.Label(backend_frame, text="后端").grid(row=0, column=0, sticky="w", pady=5)
    backend_select = ttk.Combobox(
        backend_frame,
        textvariable=backend_var,
        values=["auto", "libreoffice", "wps", "powerpoint"],
        state="readonly",
    )
    backend_select.grid(row=0, column=1, sticky="ew", padx=8, pady=5)
    ttk.Button(backend_frame, text="刷新", command=lambda: refresh_backends(verbose=True)).grid(
        row=0, column=2, sticky="e", pady=5
    )

    ttk.Label(backend_frame, text="程序路径").grid(row=1, column=0, sticky="w", pady=5)
    ttk.Entry(backend_frame, textvariable=office_bin_var).grid(row=1, column=1, sticky="ew", padx=8, pady=5)
    ttk.Button(backend_frame, text="浏览", command=choose_office_bin).grid(row=1, column=2, sticky="e", pady=5)

    ttk.Label(backend_frame, text="检测状态").grid(row=2, column=0, sticky="nw", pady=5)
    ttk.Label(backend_frame, textvariable=backend_status_var, foreground="#555555", justify=tk.LEFT).grid(
        row=2, column=1, columnspan=2, sticky="w", padx=8, pady=5
    )

    powerpoint_frame = ttk.LabelFrame(form_panel, text="PowerPoint 特定选项", padding=12)
    powerpoint_frame.pack(fill=tk.X, pady=(10, 0))

    ttk.Label(powerpoint_frame, text="导出意图").grid(row=0, column=0, sticky="w", pady=5)
    powerpoint_intent_select = ttk.Combobox(
        powerpoint_frame,
        textvariable=powerpoint_intent_var,
        values=["print", "screen"],
        state="readonly",
        width=14,
    )
    powerpoint_intent_select.grid(row=0, column=1, sticky="w", padx=8, pady=5)

    bitmap_missing_fonts_check = ttk.Checkbutton(
        powerpoint_frame,
        text="缺字库时将文字位图化",
        variable=bitmap_missing_fonts_var,
    )
    bitmap_missing_fonts_check.grid(row=0, column=2, sticky="w", pady=5)

    crop_frame = ttk.LabelFrame(form_panel, text="裁剪与页面设置", padding=12)
    crop_frame.pack(fill=tk.X, pady=(10, 0))

    preset_row = ttk.Frame(crop_frame)
    preset_row.pack(fill=tk.X, pady=(0, 8))
    ttk.Checkbutton(preset_row, text="不裁剪", variable=no_crop_var).pack(side=tk.LEFT)
    ttk.Button(
        preset_row,
        text="紧密裁剪",
        command=lambda: (percent_retain_var.set(0), margin_size_var.set(0)),
    ).pack(side=tk.LEFT, padx=(12, 4))
    ttk.Button(
        preset_row,
        text="小白边",
        command=lambda: (percent_retain_var.set(0), margin_size_var.set(3)),
    ).pack(side=tk.LEFT, padx=4)
    ttk.Button(
        preset_row,
        text="中白边",
        command=lambda: (percent_retain_var.set(0), margin_size_var.set(6)),
    ).pack(side=tk.LEFT, padx=4)
    ttk.Button(
        preset_row,
        text="保留原边距",
        command=lambda: (percent_retain_var.set(10), margin_size_var.set(0)),
    ).pack(side=tk.LEFT, padx=4)

    crop_grid = ttk.Frame(crop_frame)
    crop_grid.pack(fill=tk.X)
    crop_grid.columnconfigure(1, weight=1)
    crop_grid.columnconfigure(3, weight=1)

    ttk.Label(crop_grid, text="保留原边距(%)").grid(row=0, column=0, sticky="w", pady=4)
    percent_spin = ttk.Spinbox(crop_grid, from_=0, to=100, increment=1, textvariable=percent_retain_var, width=10)
    percent_spin.grid(row=0, column=1, sticky="w", padx=(8, 20), pady=4)

    ttk.Label(crop_grid, text="额外白边(bp)").grid(row=0, column=2, sticky="w", pady=4)
    margin_spin = ttk.Spinbox(crop_grid, from_=0, to=50, increment=0.5, textvariable=margin_size_var, width=10)
    margin_spin.grid(row=0, column=3, sticky="w", padx=8, pady=4)

    ttk.Label(crop_grid, text="检测阈值").grid(row=1, column=0, sticky="w", pady=4)
    threshold_spin = ttk.Spinbox(crop_grid, from_=0, to=255, increment=1, textvariable=threshold_var, width=10)
    threshold_spin.grid(row=1, column=1, sticky="w", padx=(8, 20), pady=4)

    uniform_check = ttk.Checkbutton(crop_grid, text="统一裁剪", variable=use_uniform_var)
    uniform_check.grid(row=1, column=2, sticky="w", pady=4)
    same_size_check = ttk.Checkbutton(crop_grid, text="统一页面大小", variable=use_same_size_var)
    same_size_check.grid(row=1, column=3, sticky="w", padx=8, pady=4)

    footer = ttk.Frame(form_panel)
    footer.pack(fill=tk.X, pady=(12, 0))

    ttk.Label(footer, textvariable=status_var, foreground="#444444", justify=tk.LEFT).pack(
        side=tk.LEFT, fill=tk.X, expand=True
    )

    export_button = ttk.Button(footer, text="导出当前配置")
    export_button.pack(side=tk.RIGHT)

    def update_dynamic_state(*_args):
        crop_enabled = not no_crop_var.get()
        state = tk.NORMAL if crop_enabled else tk.DISABLED
        for widget in (percent_spin, margin_spin, threshold_spin):
            widget.configure(state=state)
        for widget in (uniform_check, same_size_check):
            widget.configure(state=tk.NORMAL if crop_enabled else tk.DISABLED)

        backend = backend_var.get()
        is_powerpoint = backend == "powerpoint"
        ppt_state = "readonly" if is_powerpoint else tk.DISABLED
        powerpoint_intent_select.configure(state=ppt_state)
        bitmap_missing_fonts_check.configure(state=tk.NORMAL if is_powerpoint else tk.DISABLED)

    def refresh_backends(verbose=False):
        labels = []
        for backend in detect_backends(explicit_path=office_bin_var.get() or None):
            mark = "supported" if backend.supported else "detected"
            detail = f" - {backend.detail}" if backend.detail else ""
            labels.append(f"{backend.name} ({mark}){detail}")
        backend_status_var.set("\n".join(labels) if labels else "未检测到可用后端")
        if verbose:
            status_var.set("已刷新后端检测结果。")

    def on_source_or_pages_changed(*_args):
        if auto_output_var.get() and source_var.get():
            try:
                output_var.set(_make_output_path(source_var.get(), pages_var.get()))
            except Exception:
                pass

    def validate_form():
        source = source_var.get().strip()
        if not source:
            raise ValueError("请选择输入文件")
        path = Path(source).resolve()
        if not path.exists():
            raise ValueError("输入文件不存在")
        pages = parse_page_range(pages_var.get().strip())
        output = output_var.get().strip()
        if not output:
            output = _make_output_path(source, pages_var.get().strip())
            output_var.set(output)
        return path, pages, output

    def run_export():
        try:
            source, pages, output = validate_form()
            export_selected_pages(
                source,
                output,
                pages,
                backend=backend_var.get(),
                office_bin=office_bin_var.get().strip() or None,
                powerpoint_intent=powerpoint_intent_var.get(),
                bitmap_missing_fonts=bitmap_missing_fonts_var.get(),
                no_crop=no_crop_var.get(),
                percent_retain=percent_retain_var.get(),
                margin_size=margin_size_var.get(),
                use_uniform=use_uniform_var.get(),
                use_same_size=use_same_size_var.get(),
                threshold=threshold_var.get(),
            )
        except Exception as exc:
            root.after(
                0,
                lambda: (
                    status_var.set(f"导出失败: {exc}"),
                    messagebox.showerror("错误", str(exc), parent=root),
                    exporting.__setitem__("active", False),
                    export_button.configure(state=tk.NORMAL),
                ),
            )
            return

        root.after(
            0,
            lambda: on_export_success(),
        )

    def on_export_success():
        history_items[:] = _record_current_settings(
            history_items,
            source=source_var.get(),
            pages=pages_var.get(),
            output=output_var.get(),
            backend=backend_var.get(),
            office_bin=office_bin_var.get(),
            powerpoint_intent=powerpoint_intent_var.get(),
            bitmap_missing_fonts=bitmap_missing_fonts_var.get(),
            no_crop=no_crop_var.get(),
            percent_retain=percent_retain_var.get(),
            margin_size=margin_size_var.get(),
            use_uniform=use_uniform_var.get(),
            use_same_size=use_same_size_var.get(),
            threshold=threshold_var.get(),
        )
        _save_file_mode_history(history_items)
        refresh_history_list(0)
        status_var.set(f"导出完成: {output_var.get()}")
        exporting["active"] = False
        export_button.configure(state=tk.NORMAL)
        messagebox.showinfo("成功", f"PDF已导出至：\n{output_var.get()}", parent=root)

    def start_export(from_history=False):
        if exporting["active"]:
            return
        if from_history:
            apply_history_item(get_selected_history_item())
        exporting["active"] = True
        export_button.configure(state=tk.DISABLED)
        status_var.set("正在导出，请稍候...")
        threading.Thread(target=run_export, daemon=True).start()

    export_button.configure(command=start_export)

    source_var.trace_add("write", on_source_or_pages_changed)
    pages_var.trace_add("write", on_source_or_pages_changed)
    backend_var.trace_add("write", update_dynamic_state)
    no_crop_var.trace_add("write", update_dynamic_state)

    refresh_backends()
    refresh_history_list()
    update_dynamic_state()
    root.mainloop()

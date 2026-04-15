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
from .i18n import DEFAULT_LANGUAGE, get_translator


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


def _history_label(item, translator):
    source = item.get("source", "")
    name = Path(source).name if source else translator("gui.file.history_untitled")
    pages = item.get("pages", "")
    backend = item.get("backend", "auto")
    updated_at = item.get("updated_at", "")
    short_time = updated_at[5:16].replace("T", " ") if len(updated_at) >= 16 else ""
    prefix = f"[{short_time}] " if short_time else ""
    return translator("gui.file.history_label", prefix=prefix, name=name, pages=pages, backend=backend)


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


def _make_output_path(source_text, pages_text, *, lang=DEFAULT_LANGUAGE):
    source = Path(source_text).resolve()
    pages = parse_page_range(pages_text, lang=lang)
    return str(default_output_path(source, pages))


def main():
    root = tk.Tk()
    root.geometry("1020x730")
    root.minsize(960, 680)

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
    language_var = tk.StringVar(value=DEFAULT_LANGUAGE)
    status_var = tk.StringVar()
    backend_status_var = tk.StringVar()

    exporting = {"active": False}
    history_items = _load_file_mode_history()
    selected_history_index = {"value": None}
    status_state = {"key": "gui.file.status.ready", "kwargs": {}}

    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")

    def translator():
        return get_translator(language_var.get())

    def set_status(key, **kwargs):
        status_state["key"] = key
        status_state["kwargs"] = kwargs
        status_var.set(translator()(key, **kwargs))

    outer = ttk.Frame(root, padding=14)
    outer.pack(fill=tk.BOTH, expand=True)

    header = ttk.Frame(outer)
    header.pack(fill=tk.X, pady=(0, 10))
    title_row = ttk.Frame(header)
    title_row.pack(fill=tk.X)

    title_label = ttk.Label(title_row, font=("Microsoft YaHei UI", 14, "bold"))
    title_label.pack(anchor="w", side=tk.LEFT)

    version_label = ttk.Label(title_row, text=f"v{__version__}", foreground="#666666")
    version_label.pack(anchor="e", side=tk.RIGHT)

    language_row = ttk.Frame(header)
    language_row.pack(fill=tk.X, pady=(6, 0))
    language_label = ttk.Label(language_row)
    language_label.pack(side=tk.LEFT)
    language_select = ttk.Combobox(
        language_row,
        textvariable=language_var,
        values=("zh", "en"),
        state="readonly",
        width=6,
    )
    language_select.pack(side=tk.LEFT, padx=(8, 0))

    subtitle_label = ttk.Label(header, foreground="#555555", wraplength=920, justify=tk.LEFT)
    subtitle_label.pack(anchor="w", pady=(6, 0))

    body = ttk.PanedWindow(outer, orient=tk.HORIZONTAL)
    body.pack(fill=tk.BOTH, expand=True)

    history_panel = ttk.Frame(body, padding=(0, 0, 12, 0))
    body.add(history_panel, weight=28)
    form_panel = ttk.Frame(body)
    body.add(form_panel, weight=72)

    history_title_label = ttk.Label(history_panel, font=("Microsoft YaHei UI", 11, "bold"))
    history_title_label.pack(anchor="w")
    history_subtitle_label = ttk.Label(history_panel, foreground="#666666", wraplength=250, justify=tk.LEFT)
    history_subtitle_label.pack(anchor="w", pady=(2, 8))

    history_list = tk.Listbox(history_panel, activestyle="none", height=18)
    history_list.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
    history_scroll = ttk.Scrollbar(history_panel, orient=tk.VERTICAL, command=history_list.yview)
    history_scroll.pack(fill=tk.Y, side=tk.LEFT)
    history_list.config(yscrollcommand=history_scroll.set)

    history_actions = ttk.Frame(history_panel)
    history_actions.pack(fill=tk.X, pady=(8, 0))
    load_button = ttk.Button(history_actions)
    load_button.pack(side=tk.LEFT)
    rerun_button = ttk.Button(history_actions)
    rerun_button.pack(side=tk.LEFT, padx=6)
    delete_button = ttk.Button(history_actions)
    delete_button.pack(side=tk.LEFT)

    current_label = ttk.Label(form_panel, font=("Microsoft YaHei UI", 11, "bold"))
    current_label.pack(anchor="w")

    quick_actions = ttk.Frame(form_panel)
    quick_actions.pack(fill=tk.X, pady=(6, 10))
    choose_file_button = ttk.Button(quick_actions)
    choose_file_button.pack(side=tk.LEFT)
    same_dir_button = ttk.Button(quick_actions)
    same_dir_button.pack(side=tk.LEFT, padx=6)
    refresh_backends_button = ttk.Button(quick_actions)
    refresh_backends_button.pack(side=tk.LEFT, padx=6)
    auto_output_check = ttk.Checkbutton(quick_actions, variable=auto_output_var)
    auto_output_check.pack(side=tk.RIGHT)

    basic_frame = ttk.LabelFrame(form_panel, padding=12)
    basic_frame.pack(fill=tk.X)
    basic_frame.columnconfigure(1, weight=1)

    source_label = ttk.Label(basic_frame)
    source_label.grid(row=0, column=0, sticky="w", pady=5)
    ttk.Entry(basic_frame, textvariable=source_var).grid(row=0, column=1, sticky="ew", padx=8, pady=5)
    source_browse_button = ttk.Button(basic_frame)
    source_browse_button.grid(row=0, column=2, sticky="e", pady=5)

    pages_label = ttk.Label(basic_frame)
    pages_label.grid(row=1, column=0, sticky="w", pady=5)
    pages_row = ttk.Frame(basic_frame)
    pages_row.grid(row=1, column=1, sticky="ew", padx=8, pady=5)
    pages_row.columnconfigure(0, weight=1)
    ttk.Entry(pages_row, textvariable=pages_var).grid(row=0, column=0, sticky="ew")
    pages_hint_label = ttk.Label(pages_row, foreground="#666666")
    pages_hint_label.grid(row=0, column=1, sticky="w", padx=(8, 0))
    first_page_button = ttk.Button(basic_frame, command=lambda: pages_var.set("1"))
    first_page_button.grid(row=1, column=2, sticky="e", pady=5)

    output_label = ttk.Label(basic_frame)
    output_label.grid(row=2, column=0, sticky="w", pady=5)
    ttk.Entry(basic_frame, textvariable=output_var).grid(row=2, column=1, sticky="ew", padx=8, pady=5)
    output_browse_button = ttk.Button(basic_frame)
    output_browse_button.grid(row=2, column=2, sticky="e", pady=5)

    backend_frame = ttk.LabelFrame(form_panel, padding=12)
    backend_frame.pack(fill=tk.X, pady=(10, 0))
    backend_frame.columnconfigure(1, weight=1)

    backend_label = ttk.Label(backend_frame)
    backend_label.grid(row=0, column=0, sticky="w", pady=5)
    backend_select = ttk.Combobox(
        backend_frame,
        textvariable=backend_var,
        values=["auto", "libreoffice", "wps", "powerpoint"],
        state="readonly",
    )
    backend_select.grid(row=0, column=1, sticky="ew", padx=8, pady=5)
    backend_refresh_button = ttk.Button(backend_frame)
    backend_refresh_button.grid(row=0, column=2, sticky="e", pady=5)

    program_path_label = ttk.Label(backend_frame)
    program_path_label.grid(row=1, column=0, sticky="w", pady=5)
    ttk.Entry(backend_frame, textvariable=office_bin_var).grid(row=1, column=1, sticky="ew", padx=8, pady=5)
    office_bin_browse_button = ttk.Button(backend_frame)
    office_bin_browse_button.grid(row=1, column=2, sticky="e", pady=5)

    detection_status_label = ttk.Label(backend_frame)
    detection_status_label.grid(row=2, column=0, sticky="nw", pady=5)
    ttk.Label(backend_frame, textvariable=backend_status_var, foreground="#555555", justify=tk.LEFT).grid(
        row=2, column=1, columnspan=2, sticky="w", padx=8, pady=5
    )

    powerpoint_frame = ttk.LabelFrame(form_panel, padding=12)
    powerpoint_frame.pack(fill=tk.X, pady=(10, 0))
    export_intent_label = ttk.Label(powerpoint_frame)
    export_intent_label.grid(row=0, column=0, sticky="w", pady=5)
    powerpoint_intent_select = ttk.Combobox(
        powerpoint_frame,
        textvariable=powerpoint_intent_var,
        values=["print", "screen"],
        state="readonly",
        width=14,
    )
    powerpoint_intent_select.grid(row=0, column=1, sticky="w", padx=8, pady=5)
    bitmap_missing_fonts_check = ttk.Checkbutton(powerpoint_frame, variable=bitmap_missing_fonts_var)
    bitmap_missing_fonts_check.grid(row=0, column=2, sticky="w", pady=5)

    crop_frame = ttk.LabelFrame(form_panel, padding=12)
    crop_frame.pack(fill=tk.X, pady=(10, 0))

    preset_row = ttk.Frame(crop_frame)
    preset_row.pack(fill=tk.X, pady=(0, 8))
    no_crop_check = ttk.Checkbutton(preset_row, variable=no_crop_var)
    no_crop_check.pack(side=tk.LEFT)
    tight_crop_button = ttk.Button(
        preset_row,
        command=lambda: (percent_retain_var.set(0), margin_size_var.set(0)),
    )
    tight_crop_button.pack(side=tk.LEFT, padx=(12, 4))
    small_margin_button = ttk.Button(
        preset_row,
        command=lambda: (percent_retain_var.set(0), margin_size_var.set(3)),
    )
    small_margin_button.pack(side=tk.LEFT, padx=4)
    medium_margin_button = ttk.Button(
        preset_row,
        command=lambda: (percent_retain_var.set(0), margin_size_var.set(6)),
    )
    medium_margin_button.pack(side=tk.LEFT, padx=4)
    keep_original_button = ttk.Button(
        preset_row,
        command=lambda: (percent_retain_var.set(10), margin_size_var.set(0)),
    )
    keep_original_button.pack(side=tk.LEFT, padx=4)

    crop_grid = ttk.Frame(crop_frame)
    crop_grid.pack(fill=tk.X)
    crop_grid.columnconfigure(1, weight=1)
    crop_grid.columnconfigure(3, weight=1)

    percent_label = ttk.Label(crop_grid)
    percent_label.grid(row=0, column=0, sticky="w", pady=4)
    percent_spin = ttk.Spinbox(crop_grid, from_=0, to=100, increment=1, textvariable=percent_retain_var, width=10)
    percent_spin.grid(row=0, column=1, sticky="w", padx=(8, 20), pady=4)

    margin_label = ttk.Label(crop_grid)
    margin_label.grid(row=0, column=2, sticky="w", pady=4)
    margin_spin = ttk.Spinbox(crop_grid, from_=0, to=50, increment=0.5, textvariable=margin_size_var, width=10)
    margin_spin.grid(row=0, column=3, sticky="w", padx=8, pady=4)

    threshold_label = ttk.Label(crop_grid)
    threshold_label.grid(row=1, column=0, sticky="w", pady=4)
    threshold_spin = ttk.Spinbox(crop_grid, from_=0, to=255, increment=1, textvariable=threshold_var, width=10)
    threshold_spin.grid(row=1, column=1, sticky="w", padx=(8, 20), pady=4)

    uniform_check = ttk.Checkbutton(crop_grid, variable=use_uniform_var)
    uniform_check.grid(row=1, column=2, sticky="w", pady=4)
    same_size_check = ttk.Checkbutton(crop_grid, variable=use_same_size_var)
    same_size_check.grid(row=1, column=3, sticky="w", padx=8, pady=4)

    footer = ttk.Frame(form_panel)
    footer.pack(fill=tk.X, pady=(12, 0))
    ttk.Label(footer, textvariable=status_var, foreground="#444444", justify=tk.LEFT, wraplength=650).pack(
        side=tk.LEFT, fill=tk.X, expand=True
    )
    export_button = ttk.Button(footer)
    export_button.pack(side=tk.RIGHT)

    def refresh_history_list(select_index=None):
        history_list.delete(0, tk.END)
        current = translator()
        for item in history_items:
            history_list.insert(tk.END, _history_label(item, current))
        if history_items:
            idx = selected_history_index["value"] if select_index is None else select_index
            if idx is None:
                idx = 0
            idx = min(max(idx, 0), len(history_items) - 1)
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

    def update_dynamic_state(*_args):
        crop_enabled = not no_crop_var.get()
        state = tk.NORMAL if crop_enabled else tk.DISABLED
        for widget in (percent_spin, margin_spin, threshold_spin):
            widget.configure(state=state)
        for widget in (uniform_check, same_size_check):
            widget.configure(state=tk.NORMAL if crop_enabled else tk.DISABLED)

        is_powerpoint = backend_var.get() == "powerpoint"
        powerpoint_intent_select.configure(state="readonly" if is_powerpoint else tk.DISABLED)
        bitmap_missing_fonts_check.configure(state=tk.NORMAL if is_powerpoint else tk.DISABLED)

    def refresh_backends(verbose=False):
        current = translator()
        labels = []
        for backend in detect_backends(explicit_path=office_bin_var.get() or None, lang=language_var.get()):
            mark = current("cli.backend.supported") if backend.supported else current("cli.backend.detected")
            detail = f" - {backend.detail}" if backend.detail else ""
            labels.append(f"{backend.name} ({mark}){detail}")
        backend_status_var.set("\n".join(labels) if labels else current("gui.file.backend_none_detected"))
        if verbose:
            set_status("gui.file.status.refreshed_backends")

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
        set_status("gui.file.status.loaded_history")

    def delete_history_item():
        item = get_selected_history_item()
        if not item:
            return
        idx = selected_history_index["value"] or 0
        del history_items[idx]
        _save_file_mode_history(history_items)
        refresh_history_list(max(0, idx - 1))
        set_status("gui.file.status.deleted_history")

    def choose_source():
        current = translator()
        selected = filedialog.askopenfilename(
            parent=root,
            title=current("gui.file.dialog.select_input"),
            filetypes=[
                (current("common.filetype.presentation"), "*.pptx *.ppt *.odp"),
                (current("common.filetype.all_files"), "*.*"),
            ],
        )
        if not selected:
            return
        source_var.set(selected)
        if auto_output_var.get():
            try:
                output_var.set(_make_output_path(selected, pages_var.get(), lang=language_var.get()))
            except Exception:
                pass
        set_status("gui.file.status.selected_file")

    def choose_output():
        current = translator()
        selected = filedialog.asksaveasfilename(
            parent=root,
            title=current("gui.file.dialog.select_output"),
            defaultextension=".pdf",
            filetypes=[(current("common.filetype.pdf"), "*.pdf")],
            initialfile=os.path.basename(output_var.get()) if output_var.get() else "",
            initialdir=os.path.dirname(output_var.get()) if output_var.get() else "",
        )
        if selected:
            output_var.set(selected)
            auto_output_var.set(False)

    def choose_office_bin():
        current = translator()
        selected = filedialog.askopenfilename(parent=root, title=current("gui.file.dialog.select_backend_bin"))
        if selected:
            office_bin_var.set(selected)
            refresh_backends()

    def set_output_from_input():
        if not source_var.get():
            return
        output_var.set(_make_output_path(source_var.get(), pages_var.get(), lang=language_var.get()))

    def on_source_or_pages_changed(*_args):
        if auto_output_var.get() and source_var.get():
            try:
                output_var.set(_make_output_path(source_var.get(), pages_var.get(), lang=language_var.get()))
            except Exception:
                pass

    def validate_form():
        current = translator()
        source = source_var.get().strip()
        if not source:
            raise ValueError(current("gui.file.error.select_input"))
        path = Path(source).resolve()
        if not path.exists():
            raise ValueError(current("gui.file.error.input_not_found"))
        pages = parse_page_range(pages_var.get().strip(), lang=language_var.get())
        output = output_var.get().strip()
        if not output:
            output = _make_output_path(source, pages_var.get().strip(), lang=language_var.get())
            output_var.set(output)
        return path, pages, output

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
        set_status("gui.file.status.export_done", path=output_var.get())
        exporting["active"] = False
        export_button.configure(state=tk.NORMAL)
        current = translator()
        messagebox.showinfo(
            current("common.success"),
            current("gui.quick.success_message", path=output_var.get()),
            parent=root,
        )

    def run_export():
        try:
            source, pages, output = validate_form()
            export_selected_pages(
                source,
                output,
                pages,
                lang=language_var.get(),
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
            error_text = str(exc)
            root.after(0, lambda error_text=error_text: on_export_failure(error_text))
            return

        root.after(0, on_export_success)

    def on_export_failure(error_text):
        exporting["active"] = False
        export_button.configure(state=tk.NORMAL)
        set_status("gui.file.status.export_failed", error=error_text)
        messagebox.showerror(translator()("common.error"), error_text, parent=root)

    def start_export(from_history=False):
        if exporting["active"]:
            return
        if from_history:
            apply_history_item(get_selected_history_item())
        exporting["active"] = True
        export_button.configure(state=tk.DISABLED)
        set_status("gui.file.status.exporting")
        threading.Thread(target=run_export, daemon=True).start()

    def update_texts():
        current = translator()
        root.title(current("gui.file.title", version=__version__))
        language_label.config(text=current("lang.label"))
        title_label.config(text=current("gui.file.header_title"))
        subtitle_label.config(text=current("gui.file.header_subtitle"))
        history_title_label.config(text=current("gui.file.history_title"))
        history_subtitle_label.config(text=current("gui.file.history_subtitle"))
        load_button.config(text=current("common.load"))
        rerun_button.config(text=current("gui.file.history_rerun"))
        delete_button.config(text=current("common.delete"))
        current_label.config(text=current("gui.file.current_title"))
        choose_file_button.config(text=current("gui.file.choose_file"))
        same_dir_button.config(text=current("gui.file.output_same_dir"))
        refresh_backends_button.config(text=current("gui.file.refresh_backends"))
        auto_output_check.config(text=current("gui.file.auto_output_name"))
        basic_frame.config(text=current("gui.file.section_basic"))
        source_label.config(text=current("gui.file.field_source"))
        source_browse_button.config(text=current("common.browse"))
        pages_label.config(text=current("gui.file.field_pages"))
        pages_hint_label.config(text=current("gui.file.pages_hint"))
        first_page_button.config(text=current("gui.file.first_page"))
        output_label.config(text=current("gui.file.field_output"))
        output_browse_button.config(text=current("common.browse"))
        backend_frame.config(text=current("gui.file.section_backend"))
        backend_label.config(text=current("gui.file.field_backend"))
        backend_refresh_button.config(text=current("common.refresh"))
        program_path_label.config(text=current("gui.file.field_program_path"))
        office_bin_browse_button.config(text=current("common.browse"))
        detection_status_label.config(text=current("gui.file.field_detection_status"))
        powerpoint_frame.config(text=current("gui.file.section_powerpoint"))
        export_intent_label.config(text=current("gui.file.field_export_intent"))
        bitmap_missing_fonts_check.config(text=current("gui.file.bitmap_missing_fonts"))
        crop_frame.config(text=current("gui.file.section_crop"))
        no_crop_check.config(text=current("gui.crop.no_crop"))
        tight_crop_button.config(text=current("gui.crop.preset.tight"))
        small_margin_button.config(text=current("gui.crop.preset.small_margin"))
        medium_margin_button.config(text=current("gui.crop.preset.medium_margin"))
        keep_original_button.config(text=current("gui.crop.preset.keep_original"))
        percent_label.config(text=current("gui.crop.percent").rstrip(":"))
        margin_label.config(text=current("gui.crop.margin").rstrip(":"))
        threshold_label.config(text=current("gui.crop.threshold").rstrip(":"))
        uniform_check.config(text=current("gui.crop.uniform"))
        same_size_check.config(text=current("gui.crop.same_size"))
        export_button.config(text=current("gui.file.export_button"))
        refresh_history_list(selected_history_index["value"])
        refresh_backends(verbose=False)
        status_var.set(current(status_state["key"], **status_state["kwargs"]))

    load_button.config(command=lambda: apply_history_item(get_selected_history_item()))
    rerun_button.config(command=lambda: start_export(from_history=True))
    delete_button.config(command=delete_history_item)
    choose_file_button.config(command=choose_source)
    same_dir_button.config(command=set_output_from_input)
    refresh_backends_button.config(command=lambda: refresh_backends(verbose=True))
    source_browse_button.config(command=choose_source)
    output_browse_button.config(command=choose_output)
    backend_refresh_button.config(command=lambda: refresh_backends(verbose=True))
    office_bin_browse_button.config(command=choose_office_bin)
    export_button.config(command=start_export)
    auto_output_check.config(
        command=lambda: set_output_from_input() if auto_output_var.get() and source_var.get() else None
    )

    history_list.bind("<<ListboxSelect>>", lambda _event: get_selected_history_item())
    history_list.bind(
        "<Double-Button-1>",
        lambda _event: (apply_history_item(get_selected_history_item()), start_export(from_history=True)),
    )
    source_var.trace_add("write", on_source_or_pages_changed)
    pages_var.trace_add("write", on_source_or_pages_changed)
    backend_var.trace_add("write", update_dynamic_state)
    no_crop_var.trace_add("write", update_dynamic_state)
    language_select.bind("<<ComboboxSelected>>", lambda _event: update_texts())

    set_status("gui.file.status.ready")
    update_dynamic_state()
    update_texts()
    root.mainloop()


if __name__ == "__main__":
    main()

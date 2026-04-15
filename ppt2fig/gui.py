import os
import platform
import tkinter as tk
import ctypes
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .core import (
    crop_pdf_file,
    current_slide_2_pdf_mac,
    current_slide_2_pdf_windows,
    get_active_presentation_info,
)
from .i18n import DEFAULT_LANGUAGE, get_translator


def _enable_high_dpi_support():
    if os.name != "nt":
        return

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _fit_window(root, *, min_width, min_height):
    root.update_idletasks()
    width = max(min_width, root.winfo_reqwidth() + 24)
    height = max(min_height, root.winfo_reqheight() + 32)
    root.geometry(f"{width}x{height}")


def _quick_gui_min_size(lang, show_advanced):
    if show_advanced:
        if lang == "en":
            return 620, 520
        return 440, 460
    if lang == "en":
        return 520, 180
    return 400, 160


def _apply_crop_preset(percent_var, margin_var, preset_type):
    if preset_type == "tight":
        percent_var.set(0)
        margin_var.set(0)
    elif preset_type == "small_margin":
        percent_var.set(0)
        margin_var.set(3)
    elif preset_type == "medium_margin":
        percent_var.set(0)
        margin_var.set(6)
    elif preset_type == "keep_original":
        percent_var.set(10)
        margin_var.set(0)


def _build_crop_settings(parent, *, no_crop, margin_size, percent_retain, use_uniform, use_same_size, threshold):
    advanced_frame = tk.Frame(parent)

    title_label = tk.Label(advanced_frame, font=("Arial", 10, "bold"))
    title_label.pack(pady=(10, 5))

    preset_frame = tk.LabelFrame(advanced_frame, font=("Arial", 9))
    preset_frame.pack(fill=tk.X, pady=(0, 5))

    preset_buttons_frame = tk.Frame(preset_frame)
    preset_buttons_frame.pack(pady=5)

    tight_button = tk.Button(
        preset_buttons_frame,
        font=("Arial", 8),
        command=lambda: _apply_crop_preset(percent_retain, margin_size, "tight"),
    )
    tight_button.pack(side=tk.LEFT, padx=2)

    small_margin_button = tk.Button(
        preset_buttons_frame,
        font=("Arial", 8),
        command=lambda: _apply_crop_preset(percent_retain, margin_size, "small_margin"),
    )
    small_margin_button.pack(side=tk.LEFT, padx=2)

    medium_margin_button = tk.Button(
        preset_buttons_frame,
        font=("Arial", 8),
        command=lambda: _apply_crop_preset(percent_retain, margin_size, "medium_margin"),
    )
    medium_margin_button.pack(side=tk.LEFT, padx=2)

    keep_original_button = tk.Button(
        preset_buttons_frame,
        font=("Arial", 8),
        command=lambda: _apply_crop_preset(percent_retain, margin_size, "keep_original"),
    )
    keep_original_button.pack(side=tk.LEFT, padx=2)

    params_frame = tk.LabelFrame(advanced_frame, font=("Arial", 9))
    params_frame.pack(fill=tk.X, pady=5)

    percent_frame = tk.Frame(params_frame)
    percent_frame.pack(fill=tk.X, pady=2)
    percent_label = tk.Label(percent_frame, width=18, anchor="w")
    percent_label.pack(side=tk.LEFT)
    tk.Spinbox(
        percent_frame,
        from_=0,
        to=100,
        increment=1,
        width=6,
        textvariable=percent_retain,
        format="%.0f",
    ).pack(side=tk.LEFT, padx=5)

    margin_frame = tk.Frame(params_frame)
    margin_frame.pack(fill=tk.X, pady=2)
    margin_label = tk.Label(margin_frame, width=18, anchor="w")
    margin_label.pack(side=tk.LEFT)
    tk.Spinbox(
        margin_frame,
        from_=0,
        to=50,
        increment=0.5,
        width=6,
        textvariable=margin_size,
        format="%.1f",
    ).pack(side=tk.LEFT, padx=5)

    threshold_frame = tk.Frame(params_frame)
    threshold_frame.pack(fill=tk.X, pady=2)
    threshold_label = tk.Label(threshold_frame, width=18, anchor="w")
    threshold_label.pack(side=tk.LEFT)
    tk.Spinbox(
        threshold_frame,
        from_=0,
        to=255,
        increment=1,
        width=6,
        textvariable=threshold,
    ).pack(side=tk.LEFT, padx=5)

    options_frame = tk.Frame(params_frame)
    options_frame.pack(fill=tk.X, pady=2)

    no_crop_check = tk.Checkbutton(options_frame, variable=no_crop, font=("Arial", 8))
    no_crop_check.pack(side=tk.LEFT, padx=5)
    uniform_check = tk.Checkbutton(options_frame, variable=use_uniform, font=("Arial", 8))
    uniform_check.pack(side=tk.LEFT, padx=5)
    same_size_check = tk.Checkbutton(options_frame, variable=use_same_size, font=("Arial", 8))
    same_size_check.pack(side=tk.LEFT, padx=5)

    text_widgets = {
        "title": (title_label, "gui.crop.title"),
        "preset_frame": (preset_frame, "gui.crop.quick"),
        "tight": (tight_button, "gui.crop.preset.tight"),
        "small_margin": (small_margin_button, "gui.crop.preset.small_margin"),
        "medium_margin": (medium_margin_button, "gui.crop.preset.medium_margin"),
        "keep_original": (keep_original_button, "gui.crop.preset.keep_original"),
        "params_frame": (params_frame, "gui.crop.detail"),
        "percent": (percent_label, "gui.crop.percent"),
        "margin": (margin_label, "gui.crop.margin"),
        "threshold": (threshold_label, "gui.crop.threshold"),
        "no_crop": (no_crop_check, "gui.crop.no_crop"),
        "uniform": (uniform_check, "gui.crop.uniform"),
        "same_size": (same_size_check, "gui.crop.same_size"),
    }
    return advanced_frame, text_widgets


def main():
    _enable_high_dpi_support()
    root = tk.Tk()
    default_path_map = {}

    language_var = tk.StringVar(value=DEFAULT_LANGUAGE)
    no_crop = tk.BooleanVar(value=False)
    margin_size = tk.DoubleVar(value=0.0)
    percent_retain = tk.DoubleVar(value=0.0)
    use_uniform = tk.BooleanVar(value=True)
    use_same_size = tk.BooleanVar(value=True)
    threshold = tk.IntVar(value=191)
    show_advanced = tk.BooleanVar(value=False)

    root.attributes("-topmost", True)
    initial_width, initial_height = _quick_gui_min_size(DEFAULT_LANGUAGE, False)
    root.geometry(f"{initial_width}x{initial_height}")
    root.resizable(False, False)

    main_frame = tk.Frame(root)
    main_frame.pack(padx=15, pady=10, fill=tk.BOTH, expand=True)

    header_row = tk.Frame(main_frame)
    header_row.pack(fill=tk.X)

    version_label = tk.Label(
        header_row,
        text=f"ppt2fig v{__version__}",
        font=("Arial", 8),
        fg="#666666",
    )
    version_label.pack(side=tk.RIGHT)

    language_label = tk.Label(header_row, font=("Arial", 9))
    language_label.pack(side=tk.LEFT)

    language_select = ttk.Combobox(
        header_row,
        textvariable=language_var,
        values=("zh", "en"),
        state="readonly",
        width=6,
    )
    language_select.pack(side=tk.LEFT, padx=(6, 0))

    convert_frame = tk.Frame(main_frame)
    convert_frame.pack(fill=tk.X, pady=(10, 10))

    convert_button = tk.Button(
        convert_frame,
        command=lambda: hello_callback(),
        font=("Arial", 10, "bold"),
        bg="#4CAF50",
        fg="white",
        width=14,
        height=1,
    )
    convert_button.pack()

    toggle_button = tk.Button(
        main_frame,
        command=lambda: [show_advanced.set(not show_advanced.get()), toggle_advanced()],
        font=("Arial", 9),
        relief=tk.FLAT,
        fg="#666666",
    )
    toggle_button.pack()

    advanced_frame, advanced_text_widgets = _build_crop_settings(
        main_frame,
        no_crop=no_crop,
        margin_size=margin_size,
        percent_retain=percent_retain,
        use_uniform=use_uniform,
        use_same_size=use_same_size,
        threshold=threshold,
    )

    def translator():
        return get_translator(language_var.get())

    def update_texts():
        current = translator()
        min_width, min_height = _quick_gui_min_size(language_var.get(), show_advanced.get())
        root.title(current("gui.quick.title", version=__version__))
        language_label.config(text=current("lang.label"))
        convert_button.config(text=current("gui.quick.export_button"))
        toggle_button.config(
            text=current("gui.quick.hide_advanced")
            if show_advanced.get()
            else current("gui.quick.show_advanced")
        )
        for widget, key in advanced_text_widgets.values():
            widget.config(text=current(key))
        _fit_window(root, min_width=min_width, min_height=min_height)

    def toggle_advanced():
        if show_advanced.get():
            advanced_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        else:
            advanced_frame.pack_forget()
        update_texts()

    def hello_callback():
        current = translator()
        try:
            full_name, name = get_active_presentation_info(lang=language_var.get())
            ppt_path = os.path.dirname(full_name)
            ppt_name = os.path.splitext(name)[0]

            if ppt_path not in default_path_map:
                initial_file = os.path.join(ppt_path, ppt_name + ".pdf")
                default_path_map[ppt_path] = initial_file
            else:
                initial_file = default_path_map[ppt_path]

            pdf_file_name = filedialog.asksaveasfilename(
                parent=root,
                title=current("gui.quick.save_title"),
                initialfile=os.path.basename(initial_file),
                initialdir=os.path.dirname(initial_file),
                filetypes=[(current("common.filetype.pdf"), "*.pdf")],
            )

            if not pdf_file_name:
                return

            if not pdf_file_name.endswith(".pdf"):
                pdf_file_name = pdf_file_name + ".pdf"

            if platform.system() == "Windows":
                success = current_slide_2_pdf_windows(pdf_file_name, lang=language_var.get())
            else:
                success = current_slide_2_pdf_mac(pdf_file_name, lang=language_var.get())

            if success:
                crop_pdf_file(
                    pdf_file_name,
                    lang=language_var.get(),
                    no_crop=no_crop.get(),
                    percent_retain=percent_retain.get(),
                    margin_size=margin_size.get(),
                    use_uniform=use_uniform.get(),
                    use_same_size=use_same_size.get(),
                    threshold=threshold.get(),
                )
                messagebox.showinfo(
                    current("common.success"),
                    current("gui.quick.success_message", path=pdf_file_name),
                    parent=root,
                )
            else:
                messagebox.showerror(current("common.error"), current("gui.quick.convert_failed"), parent=root)
        except Exception as exc:
            messagebox.showerror(current("common.error"), str(exc), parent=root)

    language_select.bind("<<ComboboxSelected>>", lambda _event: update_texts())
    update_texts()
    _fit_window(root, min_width=initial_width, min_height=initial_height)
    root.mainloop()


if __name__ == "__main__":
    main()

import os
import platform
import tkinter as tk
from tkinter import filedialog, messagebox

from . import __version__
from .core import (
    crop_pdf_file,
    current_slide_2_pdf_mac,
    current_slide_2_pdf_windows,
    get_active_presentation_info,
)


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

    title_label = tk.Label(advanced_frame, text="PDF裁剪参数设置", font=("Arial", 10, "bold"))
    title_label.pack(pady=(10, 5))

    preset_frame = tk.LabelFrame(advanced_frame, text="快速设置", font=("Arial", 9))
    preset_frame.pack(fill=tk.X, pady=(0, 5))

    preset_buttons_frame = tk.Frame(preset_frame)
    preset_buttons_frame.pack(pady=5)

    tk.Button(
        preset_buttons_frame,
        text="紧密裁剪",
        font=("Arial", 8),
        command=lambda: _apply_crop_preset(percent_retain, margin_size, "tight"),
    ).pack(side=tk.LEFT, padx=2)
    tk.Button(
        preset_buttons_frame,
        text="小白边",
        font=("Arial", 8),
        command=lambda: _apply_crop_preset(percent_retain, margin_size, "small_margin"),
    ).pack(side=tk.LEFT, padx=2)
    tk.Button(
        preset_buttons_frame,
        text="中白边",
        font=("Arial", 8),
        command=lambda: _apply_crop_preset(percent_retain, margin_size, "medium_margin"),
    ).pack(side=tk.LEFT, padx=2)
    tk.Button(
        preset_buttons_frame,
        text="保留原边距",
        font=("Arial", 8),
        command=lambda: _apply_crop_preset(percent_retain, margin_size, "keep_original"),
    ).pack(side=tk.LEFT, padx=2)

    params_frame = tk.LabelFrame(advanced_frame, text="详细参数", font=("Arial", 9))
    params_frame.pack(fill=tk.X, pady=5)

    percent_frame = tk.Frame(params_frame)
    percent_frame.pack(fill=tk.X, pady=2)
    tk.Label(percent_frame, text="保留原始边距(%):", width=15, anchor="w").pack(side=tk.LEFT)
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
    tk.Label(margin_frame, text="额外白边(bp):", width=15, anchor="w").pack(side=tk.LEFT)
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
    tk.Label(threshold_frame, text="检测阈值:", width=15, anchor="w").pack(side=tk.LEFT)
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

    tk.Checkbutton(options_frame, text="不裁剪", variable=no_crop, font=("Arial", 8)).pack(
        side=tk.LEFT, padx=5
    )
    tk.Checkbutton(options_frame, text="统一裁剪", variable=use_uniform, font=("Arial", 8)).pack(
        side=tk.LEFT, padx=5
    )
    tk.Checkbutton(
        options_frame, text="统一页面大小", variable=use_same_size, font=("Arial", 8)
    ).pack(side=tk.LEFT, padx=5)

    return advanced_frame


def main():
    root = tk.Tk()
    default_path_map = {}

    no_crop = tk.BooleanVar(value=False)
    margin_size = tk.DoubleVar(value=0.0)
    percent_retain = tk.DoubleVar(value=0.0)
    use_uniform = tk.BooleanVar(value=True)
    use_same_size = tk.BooleanVar(value=True)
    threshold = tk.IntVar(value=191)
    show_advanced = tk.BooleanVar(value=False)

    def hello_callback():
        try:
            full_name, name = get_active_presentation_info()
            ppt_path = os.path.dirname(full_name)
            ppt_name = os.path.splitext(name)[0]

            if ppt_path not in default_path_map:
                initial_file = os.path.join(ppt_path, ppt_name + ".pdf")
                default_path_map[ppt_path] = initial_file
            else:
                initial_file = default_path_map[ppt_path]

            pdf_file_name = filedialog.asksaveasfilename(
                parent=root,
                initialfile=os.path.basename(initial_file),
                initialdir=os.path.dirname(initial_file),
                filetypes=[("PDF file", "*.pdf")],
            )

            if not pdf_file_name:
                return

            if not pdf_file_name.endswith(".pdf"):
                pdf_file_name = pdf_file_name + ".pdf"

            if platform.system() == "Windows":
                success = current_slide_2_pdf_windows(pdf_file_name)
            else:
                success = current_slide_2_pdf_mac(pdf_file_name)

            if success:
                crop_pdf_file(
                    pdf_file_name,
                    no_crop=no_crop.get(),
                    percent_retain=percent_retain.get(),
                    margin_size=margin_size.get(),
                    use_uniform=use_uniform.get(),
                    use_same_size=use_same_size.get(),
                    threshold=threshold.get(),
                )
                messagebox.showinfo("成功", f"PDF已导出至：\n{pdf_file_name}")
            else:
                messagebox.showerror("错误", "转换失败")
        except Exception as exc:
            messagebox.showerror("错误", str(exc))

    root.attributes("-topmost", True)
    root.title(f"PPT转PDF工具 v{__version__}")
    root.geometry("300x100")
    root.resizable(False, False)

    main_frame = tk.Frame(root)
    main_frame.pack(padx=15, pady=10, fill=tk.BOTH, expand=True)

    tk.Label(
        main_frame,
        text=f"ppt2fig v{__version__}",
        font=("Arial", 8),
        fg="#666666",
    ).pack(anchor="ne")

    convert_frame = tk.Frame(main_frame)
    convert_frame.pack(fill=tk.X, pady=(0, 10))

    tk.Button(
        convert_frame,
        text="转PDF",
        command=hello_callback,
        font=("Arial", 10, "bold"),
        bg="#4CAF50",
        fg="white",
        width=10,
        height=1,
    ).pack()

    def toggle_advanced():
        if show_advanced.get():
            advanced_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
            toggle_button.config(text="▲ 隐藏高级设置")
            root.geometry("320x380")
        else:
            advanced_frame.pack_forget()
            toggle_button.config(text="▼ 显示高级设置")
            root.geometry("300x100")

    toggle_button = tk.Button(
        main_frame,
        text="▼ 显示高级设置",
        command=lambda: [show_advanced.set(not show_advanced.get()), toggle_advanced()],
        font=("Arial", 9),
        relief=tk.FLAT,
        fg="#666",
    )
    toggle_button.pack()

    advanced_frame = _build_crop_settings(
        main_frame,
        no_crop=no_crop,
        margin_size=margin_size,
        percent_retain=percent_retain,
        use_uniform=use_uniform,
        use_same_size=use_same_size,
        threshold=threshold,
    )

    root.mainloop()

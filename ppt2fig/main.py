import os
import shutil
import platform
import tkinter as tk
from tkinter.filedialog import asksaveasfilename
from pdfCropMargins import crop

def current_slide_2_pdf_windows(output_pdf_file):
    """Windows 系统下将 PowerPoint 转换为 PDF"""
    try:
        import comtypes.client
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        ppt_file = powerpoint.ActivePresentation
        output_pdf_file = os.path.abspath(output_pdf_file)
        if os.path.exists(output_pdf_file):
            os.remove(output_pdf_file)
        ppt_file.ExportAsFixedFormat(output_pdf_file, 2, RangeType=3)
        return True
    except Exception as e:
        tk.messagebox.showerror("错误", f"转换过程出错：{str(e)}")
        return False

def current_slide_2_pdf_mac(output_pdf_file):
    """Mac 系统下将 PowerPoint 转换为 PDF"""
    try:
        import subprocess
        script = '''
        tell application "Microsoft PowerPoint"
            if not running then
                return "PowerPoint未启动"
            end if
            
            if (count of presentations) is 0 then
                return "没有打开的PPT文件"
            end if
            
            set pdfPath to "%s"
            set thePresentation to active presentation
            save active presentation in pdfPath as save as PDF
            return "success"
        end tell
        ''' % output_pdf_file
        
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        if "success" in result.stdout:
            return True
        else:
            tk.messagebox.showerror("错误", result.stdout.strip())
            return False
    except Exception as e:
        tk.messagebox.showerror("错误", f"转换过程出错：{str(e)}")
        return False

def get_active_presentation_info():
    """获取当前活动的 PPT 文件信息"""
    if platform.system() == 'Windows':
        import comtypes.client
        try:
            powerpoint = comtypes.client.GetActiveObject("Powerpoint.Application")
            ppt_file = powerpoint.ActivePresentation
            return ppt_file.FullName, ppt_file.Name
        except Exception as e:
            raise Exception("PowerPoint未启动或没有打开的文件")
    else:
        # Mac 系统
        script = '''
        tell application "Microsoft PowerPoint"
            if not running then
                return "PowerPoint未启动"
            end if
            
            if (count of presentations) is 0 then
                return "没有打开的PPT文件"
            end if
            
            set thePresentation to active presentation
            return {full name of thePresentation, name of thePresentation}
        end tell
        '''
        import subprocess
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        if "PowerPoint未启动" in result.stdout or "没有打开的PPT文件" in result.stdout:
            raise Exception(result.stdout.strip())
        
        # 解析返回的文件路径和名称
        output = result.stdout.strip().split(', ')
        if len(output) == 2:
            return output[0], output[1]
        raise Exception("无法获取PPT文件信息")

def main():
    root = tk.Tk()
    defalut_path_map = {} # 对于 每个ppt文件，他的默认导出路径是不同的
    
    # 裁剪参数设置
    margin_size = tk.DoubleVar(value=0.0)  # 白边大小，单位为bp (big points)
    percent_retain = tk.DoubleVar(value=0.0)  # 保留原始边距的百分比
    use_uniform = tk.BooleanVar(value=True)  # 是否统一裁剪
    use_same_size = tk.BooleanVar(value=True)  # 是否统一页面大小
    threshold = tk.IntVar(value=191)  # 阈值设置 (0-255)
    show_advanced = tk.BooleanVar(value=False)  # 是否显示高级设置
    
    def helloCallBack():
        try:
            full_name, name = get_active_presentation_info()
            ppt_path = os.path.dirname(full_name)
            ppt_name = os.path.splitext(name)[0]
            
            if ppt_path not in defalut_path_map:
                initial_file = os.path.join(ppt_path, ppt_name + '.pdf')
                defalut_path_map[ppt_path] = initial_file
            else:
                initial_file = defalut_path_map[ppt_path]

            pdf_file_name = asksaveasfilename(
                parent=root,
                initialfile=os.path.basename(initial_file),
                initialdir=os.path.dirname(initial_file),
                filetypes=[("PDF file", "*.pdf")]
            )
            
            if not pdf_file_name:
                return
                
            if not pdf_file_name.endswith('.pdf'):
                pdf_file_name = pdf_file_name + '.pdf'
                
            success = False
            if platform.system() == 'Windows':
                success = current_slide_2_pdf_windows(pdf_file_name)
            else:
                success = current_slide_2_pdf_mac(pdf_file_name)
                
            if success:
                tmp_pdf_file_name = pdf_file_name + '.crop'
                
                # 构建 pdfCropMargins 参数
                crop_args = []
                
                # 百分比保留参数
                percent = percent_retain.get()
                crop_args.extend(["-p", str(percent)])
                
                # 绝对偏移量（白边）
                margin = margin_size.get()
                if margin > 0:
                    crop_args.extend(["-a", str(-margin)])
                
                # 统一裁剪选项
                if use_uniform.get():
                    crop_args.append("-u")
                
                # 统一页面大小选项
                if use_same_size.get():
                    crop_args.append("-s")
                
                # 阈值设置
                thresh = threshold.get()
                if thresh != 191:  # 只有非默认值时才添加
                    crop_args.extend(["-t", str(thresh)])
                
                # 输入输出文件
                crop_args.extend([pdf_file_name, "-o", tmp_pdf_file_name])
                
                # 执行裁剪
                crop(crop_args)
                shutil.move(tmp_pdf_file_name, pdf_file_name)
            else:
                tk.messagebox.showerror("错误", "转换失败")
        except Exception as e:
            tk.messagebox.showerror("错误", str(e))
    
    root.attributes("-topmost", True)
    root.title("PPT转PDF工具")
    root.geometry("300x100")  # 更小的默认尺寸
    root.resizable(False, False)
    
    # 创建主框架
    main_frame = tk.Frame(root)
    main_frame.pack(padx=15, pady=10, fill=tk.BOTH, expand=True)
    
    # 转换按钮（主要功能）
    convert_frame = tk.Frame(main_frame)
    convert_frame.pack(fill=tk.X, pady=(0, 10))
    
    convert_button = tk.Button(convert_frame, text="转PDF", command=helloCallBack, 
                              font=("Arial", 10, "bold"), bg="#4CAF50", fg="white",
                              width=10, height=1)
    convert_button.pack()
    
    # 高级设置切换按钮
    def toggle_advanced():
        if show_advanced.get():
            # 展开高级设置
            advanced_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
            toggle_button.config(text="▲ 隐藏高级设置")
            root.geometry("320x380")
        else:
            # 隐藏高级设置
            advanced_frame.pack_forget()
            toggle_button.config(text="▼ 显示高级设置")
            root.geometry("300x100")
    
    toggle_button = tk.Button(main_frame, text="▼ 显示高级设置", 
                             command=lambda: [show_advanced.set(not show_advanced.get()), toggle_advanced()],
                             font=("Arial", 9), relief=tk.FLAT, fg="#666")
    toggle_button.pack()
    
    # 高级设置框架（初始隐藏）
    advanced_frame = tk.Frame(main_frame)
    
    # 裁剪参数设置标题
    title_label = tk.Label(advanced_frame, text="PDF裁剪参数设置", font=("Arial", 10, "bold"))
    title_label.pack(pady=(10, 5))
    
    # 快速设置预设
    preset_frame = tk.LabelFrame(advanced_frame, text="快速设置", font=("Arial", 9))
    preset_frame.pack(fill=tk.X, pady=(0, 5))
    
    def apply_preset(preset_type):
        if preset_type == "tight":
            percent_retain.set(0)
            margin_size.set(0)
        elif preset_type == "small_margin":
            percent_retain.set(0)
            margin_size.set(3)
        elif preset_type == "medium_margin":
            percent_retain.set(0)
            margin_size.set(6)
        elif preset_type == "keep_original":
            percent_retain.set(10)
            margin_size.set(0)
    
    preset_buttons_frame = tk.Frame(preset_frame)
    preset_buttons_frame.pack(pady=5)
    
    tk.Button(preset_buttons_frame, text="紧密裁剪", font=("Arial", 8),
              command=lambda: apply_preset("tight")).pack(side=tk.LEFT, padx=2)
    tk.Button(preset_buttons_frame, text="小白边", font=("Arial", 8),
              command=lambda: apply_preset("small_margin")).pack(side=tk.LEFT, padx=2)
    tk.Button(preset_buttons_frame, text="中白边", font=("Arial", 8),
              command=lambda: apply_preset("medium_margin")).pack(side=tk.LEFT, padx=2)
    tk.Button(preset_buttons_frame, text="保留原边距", font=("Arial", 8),
              command=lambda: apply_preset("keep_original")).pack(side=tk.LEFT, padx=2)
    
    # 详细参数设置
    params_frame = tk.LabelFrame(advanced_frame, text="详细参数", font=("Arial", 9))
    params_frame.pack(fill=tk.X, pady=5)
    
    # 百分比保留设置
    percent_frame = tk.Frame(params_frame)
    percent_frame.pack(fill=tk.X, pady=2)
    tk.Label(percent_frame, text="保留原始边距(%):", width=15, anchor='w').pack(side=tk.LEFT)
    percent_spinbox = tk.Spinbox(percent_frame, from_=0, to=100, increment=1, width=6, 
                                textvariable=percent_retain, format="%.0f")
    percent_spinbox.pack(side=tk.LEFT, padx=5)
    
    # 白边设置
    margin_frame = tk.Frame(params_frame)
    margin_frame.pack(fill=tk.X, pady=2)
    tk.Label(margin_frame, text="额外白边(bp):", width=15, anchor='w').pack(side=tk.LEFT)
    margin_spinbox = tk.Spinbox(margin_frame, from_=0, to=50, increment=0.5, width=6, 
                               textvariable=margin_size, format="%.1f")
    margin_spinbox.pack(side=tk.LEFT, padx=5)
    
    # 阈值设置
    threshold_frame = tk.Frame(params_frame)
    threshold_frame.pack(fill=tk.X, pady=2)
    tk.Label(threshold_frame, text="检测阈值:", width=15, anchor='w').pack(side=tk.LEFT)
    threshold_spinbox = tk.Spinbox(threshold_frame, from_=0, to=255, increment=1, width=6,
                                  textvariable=threshold)
    threshold_spinbox.pack(side=tk.LEFT, padx=5)
    
    # 选项设置
    options_frame = tk.Frame(params_frame)
    options_frame.pack(fill=tk.X, pady=2)
    
    uniform_check = tk.Checkbutton(options_frame, text="统一裁剪", variable=use_uniform, font=("Arial", 8))
    uniform_check.pack(side=tk.LEFT, padx=5)
    
    same_size_check = tk.Checkbutton(options_frame, text="统一页面大小", variable=use_same_size, font=("Arial", 8))
    same_size_check.pack(side=tk.LEFT, padx=5)
    
    root.mainloop()

if __name__ == "__main__":
    main()

import os
import shutil
import comtypes.client
from pdfCropMargins import crop
import tkinter as tk
from tkinter.filedialog import asksaveasfilename

def ppt_2_pdf(input_ppt_file, page_index, output_pdf_file):
    """
    Convert a Powerpoint file to a pdf file
    :param input_ppt_file: input Powerpoint file
    :param output_pdf_file: output pdf file
    :param format_type:
    :return: a pdf file written in the directory
    """
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    input_ppt_file=os.path.abspath(input_ppt_file)
    ppt_file = powerpoint.Presentations.Open(input_ppt_file,ReadOnly=True)
    print_range=ppt_file.PrintOptions.Ranges.Add(page_index,page_index)
    output_pdf_file=os.path.abspath(output_pdf_file)
    if os.path.exists(output_pdf_file):
        os.remove(output_pdf_file)
    ppt_file.ExportAsFixedFormat(output_pdf_file,2,PrintRange=print_range,RangeType=4)
    ppt_file.Close()
    
def current_slide_2_pdf(output_pdf_file):
    """
    Convert a Powerpoint file to a pdf file
    :param output_pdf_file: output pdf file
    :return: a pdf file written in the directory
    """
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        # 尝试获取当前活动的PPT文件
        ppt_file = powerpoint.ActivePresentation
        output_pdf_file = os.path.abspath(output_pdf_file)
        if os.path.exists(output_pdf_file):
            os.remove(output_pdf_file)
        ppt_file.ExportAsFixedFormat(output_pdf_file,2,RangeType=3)
        return True
    except Exception as e:
        tk.messagebox.showerror("错误", f"转换过程出错：{str(e)}")
        return False

def main():
    root = tk.Tk()
    defalut_path_map = {} # 对于 每个ppt文件，他的默认导出路径是不同的
    
    def helloCallBack():
        try:
            powerpoint = comtypes.client.GetActiveObject("Powerpoint.Application")
        except Exception as e:
            tk.messagebox.showerror("错误", f"Powerpoint应用程序未启动！")
            return

        try:
            # 先尝试获取当前PPT文件的路径作为默认保存路径
            ppt_file = powerpoint.ActivePresentation
        except Exception as e:
            tk.messagebox.showerror("错误", f"当前没有PPT文件打开！")
            return
                
        # 获取PPT文件的路径和文件名（不带扩展名）
        ppt_path = os.path.dirname(ppt_file.FullName)
        ppt_name = os.path.splitext(ppt_file.Name)[0]

        if ppt_path not in defalut_path_map:
            initial_file = os.path.join(ppt_path, ppt_name + '.pdf')
            defalut_path_map[ppt_path] = initial_file
        else:
            initial_file = defalut_path_map[ppt_path]
        
        # 设置默认保存路径和文件名

        pdf_file_name = asksaveasfilename(
            parent=root,
            initialfile=os.path.basename(initial_file),
            initialdir=os.path.dirname(initial_file),
            filetypes=[("PDF file", "*.pdf")]
        )
        
        if pdf_file_name is None or len(pdf_file_name) == 0:
            return
            
        if not pdf_file_name.endswith('.pdf'):
            pdf_file_name = pdf_file_name + '.pdf'
            
        if current_slide_2_pdf(pdf_file_name):
            tmp_pdf_file_name = pdf_file_name + '.crop'
            crop(["-p", "0", "-u", "-s", pdf_file_name, "-o", tmp_pdf_file_name])
            shutil.move(tmp_pdf_file_name, pdf_file_name)

    root.attributes("-topmost", True)
    B = tk.Button(root, text="转PDF", command=helloCallBack)
    B.pack()
    root.mainloop()

if __name__ == "__main__":
    main()

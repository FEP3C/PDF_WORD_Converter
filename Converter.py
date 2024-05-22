import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
from pdf2docx import Converter

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件格式转换器")
        self.root.geometry("300x200")

        self.file_path_label = tk.Label(root, text="请选择文件：")
        self.file_path_label.pack(pady=10)

        self.select_button = tk.Button(root, text="选择文件", command=self.select_file)
        self.select_button.pack(pady=5)

        self.convert_button = tk.Button(root, text="转换", command=self.convert_file)
        self.convert_button.pack(pady=5)

        self.file_path = ''

    def select_file(self):
        self.file_path = filedialog.askopenfilename()
        if self.file_path:
            self.file_path_label.config(text="选择的文件：" + os.path.basename(self.file_path))

    def convert_file(self):
        if not self.file_path:
            messagebox.showerror("错误", "请先选择文件！")
            return

        file_extension = os.path.splitext(self.file_path)[-1].lower()

        if file_extension == '.pdf':
            target_file = filedialog.asksaveasfilename(defaultextension=".docx",
                                                       filetypes=[("Word Files", "*.docx")])
            if target_file:
                try:
                    with Converter(self.file_path) as converter:
                        converter.convert(target_file, start=0, end=None)
                        converter.close()
                    self.show_popup()
                except Exception as e:
                    messagebox.showerror("错误", str(e))
        elif file_extension == '.docx':
            target_file = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                       filetypes=[("PDF Files", "*.pdf")])
            if target_file:
                try:
                    convert(self.file_path, target_file)
                    self.show_popup()
                except Exception as e:
                    messagebox.showerror("错误", str(e))
        else:
            messagebox.showerror("错误", "只支持Word和PDF格式文件！")

    def show_popup(self):
        popup = tk.Toplevel()
        popup.geometry("300x100")
        popup.title("提示")
        label = tk.Label(popup, text="文件转换好了！\n关注https://github.com/FEP3C吧！\n谢谢喵主人！")
        label.pack()

        open_browser_button = tk.Button(popup, text="打开浏览器", command=self.open_browser)
        open_browser_button.pack(pady=5)

        exit_button = tk.Button(popup, text="退出", command=self.root.destroy)
        exit_button.pack(pady=5)

    def open_browser(self):
        import webbrowser
        webbrowser.open("https://github.com/FEP3C")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()

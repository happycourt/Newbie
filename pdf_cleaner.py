import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import re
import nltk
from docx import Document
from threading import Thread

# 设置 NLTK 数据包路径
nltk.data.path.append('D:\\projects\\nltk_data')

class PDFCleanerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("PDF Cleaner")
        self.master.geometry("400x250")

        self.label = tk.Label(master, text="选择PDF文件进行清理")
        self.label.pack(pady=20)

        self.open_button = tk.Button(master, text="选择文件", command=self.open_file)
        self.open_button.pack(pady=10)

        self.progress = ttk.Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        self.progress_label = tk.Label(master, text="")
        self.progress_label.pack(pady=10)

        self.exit_button = tk.Button(master, text="退出任务", command=self.on_closing)
        self.exit_button.pack(pady=10)

        self.file_path = ""  # 初始化 file_path

        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def open_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.file_path:
            self.progress_label.config(text="开始清理PDF文件...")
            self.progress["value"] = 0
            self.master.update_idletasks()
            self.start_cleaning_thread()

    def start_cleaning_thread(self):
        cleaning_thread = Thread(target=self.clean_pdf, args=(self.file_path,))
        cleaning_thread.start()

    def clean_pdf(self, input_pdf_path):
        try:
            doc = fitz.open(input_pdf_path)
            articles = []

            total_pages = doc.page_count
            self.progress["maximum"] = total_pages

            for page_num in range(total_pages):
                self.update_progress(page_num, total_pages)

                page = doc.load_page(page_num)
                text = page.get_text()
                cleaned_text = self.clean_text(text)
                articles.extend(self.split_into_articles(cleaned_text))

            long_articles = [article for article in articles if self.word_count(article) > 300]

            output_docx_name = self.get_output_filename(input_pdf_path, "docx")
            self.save_as_word(long_articles, output_docx_name)

            self.progress_label.config(text="PDF清理完成")
            messagebox.showinfo("完成", f"PDF清理完成，文件已保存为 {output_docx_name}")

        except Exception as e:
            messagebox.showerror("错误", str(e))

    def update_progress(self, current_page, total_pages):
        self.progress["value"] = current_page + 1
        self.progress_label.config(text=f"处理第 {current_page + 1} 页，共 {total_pages} 页")
        self.master.update_idletasks()

    @staticmethod
    def clean_text(text):
        # 去除段落内的换行符
        cleaned_text = re.sub(r'(?<!\n)\n(?!\n)', ' ', text)
        # 去除多余的空白字符
        cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
        return cleaned_text

    @staticmethod
    def split_into_articles(text):
        # 保留真正的段落分隔符
        return text.split('\n\n')

    @staticmethod
    def word_count(text):
        words = nltk.word_tokenize(text)
        return len(words)

    @staticmethod
    def save_as_word(articles, output_path):
        doc = Document()
        for article in articles:
            doc.add_paragraph(article)
        doc.save(output_path)

    @staticmethod
    def get_output_filename(input_path, ext):
        base_name = os.path.basename(input_path)
        safe_base_name = re.sub(r'\W+', '_', base_name)
        return f"Output_{'_'.join(safe_base_name.split()[:5])}.{ext}"

    def on_closing(self):
        if messagebox.askokcancel("退出", "你确定要退出吗？"):
            self.master.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCleanerApp(root)
    root.mainloop()

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import pdfplumber
from PIL import Image, ImageTk
import os
from pdf2image import convert_from_path

def remove_illegal_characters(text):
    return re.sub(r'[^\x20-\x7E]', '', text)

def create_dataframe(num_list, str_list):
    rows = []
    for num, doc in zip(num_list, str_list):
        rows.append([num, doc])
    return pd.DataFrame(rows, columns=['No', 'Document'])

def save_to_excel(df, excel_path):
    df.to_excel(excel_path, index=False)

def classify_number(s):
    return re.match(r'^\d+(\.\d+)*\.?$', s) is not None

def process_text(text):
    lines = text.split('\n')
    num_list = []
    str_list = []
    
    current_num = None
    current_str = ""

    for line in lines:
        match = re.match(r'^(\d+(\.\d+)*)(\s+|\s*[-:)]\s*)(.*)', line.strip())
        
        if match:
            if current_num is not None:
                num_list.append(current_num)
                str_list.append(current_str.strip())
            
            current_num = match.group(1).strip()
            current_str = match.group(4).strip()
        else:
            if current_num is not None:
                current_str += " " + line.strip()

    if current_num is not None:
        num_list.append(current_num)
        str_list.append(current_str.strip())

    return num_list, str_list

class PDFBoundingBoxSelector(tk.Tk): 
    def __init__(self):
        super().__init__()
        self.title("PDF to Excel Converter")
        icon_path = os.path.abspath("induk.ico")  # 절대 경로 설정
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)  # 아이콘 파일 경로를 설정합니다.
        else:
            print(f"Icon file not found: {icon_path}")
        self.pdf_path = None
        self.save_dir = None
        self.save_filename = None
        self.bbox = None
        self.rect = None
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="Select PDF File:").grid(row=0, column=0, padx=10, pady=5)
        self.pdf_entry = tk.Entry(self, width=50)
        self.pdf_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(self, text="Browse", command=self.select_pdf).grid(row=0, column=2, padx=10, pady=5)

        tk.Label(self, text="Save Location:").grid(row=1, column=0, padx=10, pady=5)
        self.save_entry = tk.Entry(self, width=50)
        self.save_entry.grid(row=1, column=1, padx=10, pady=5)
        tk.Button(self, text="Browse", command=self.select_save_location).grid(row=1, column=2, padx=10, pady=5)

        tk.Label(self, text="File Name:").grid(row=2, column=0, padx=10, pady=5)
        self.filename_entry = tk.Entry(self, width=50)
        self.filename_entry.grid(row=2, column=1, padx=10, pady=5)
        self.filename_entry.insert(0, "extracted_file.xlsx")

        tk.Label(self, text="Start Page:").grid(row=3, column=0, padx=10, pady=5)
        self.start_entry = tk.Entry(self, width=5)
        self.start_entry.grid(row=3, column=1, padx=17, pady=5, sticky='w')
        self.start_entry.insert(0, "1")

        tk.Label(self, text="End Page:").grid(row=4, column=0, padx=10, pady=5)
        self.end_entry = tk.Entry(self, width=5)
        self.end_entry.grid(row=4, column=1, padx=17, pady=5, sticky='w')
        self.end_entry.insert(0, "1")

        tk.Button(self, text="Apply", command=self.display_pdf_page).grid(row=3, column=2, rowspan=2, padx=8, pady=6, sticky='nsew')

        tk.Label(self, text="Warnings:").grid(row=5, column=0, padx=10, pady=5, sticky='w')
        warnings = (
            "1. 특수문자/기호는 복사되지 않습니다.\n"
            "2. 문단 내 소번호는 각 소번호 항목이 분리되지 않고 하나의 문장으로 출력됩니다."
        )
        tk.Label(self, text=warnings, justify=tk.LEFT, fg="red").grid(row=5, column=1, columnspan=2, padx=10, pady=5)

        self.canvas = tk.Canvas(self, cursor="cross")
        self.canvas.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

        tk.Button(self, text="Convert", command=self.convert_pdf_to_excel).grid(row=7, column=0, columnspan=3, pady=10)

        self.canvas.bind("<Button-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

    def select_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not self.pdf_path:
            return
        self.pdf_entry.delete(0, tk.END)
        self.pdf_entry.insert(0, self.pdf_path)
        self.display_pdf_page()

    def display_pdf_page(self):
        start_page = int(self.start_entry.get())
        if not self.pdf_path:
            return
        try:
            images = convert_from_path(self.pdf_path, first_page=start_page, last_page=start_page)
            if images:
                self.image = images[0]
                screen_width = self.winfo_screenwidth()
                screen_height = self.winfo_screenheight()
                scale_factor = min(screen_width / self.image.width, screen_height / self.image.height) * 0.5
                new_width = int(self.image.width * scale_factor)
                new_height = int(self.image.height * scale_factor)
                self.image = self.image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                self.tk_image = ImageTk.PhotoImage(self.image)
                self.canvas.config(width=new_width, height=new_height)
                self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)

                if self.rect:
                    self.canvas.delete(self.rect)
                    self.rect = None
        except Exception as e:
            messagebox.showerror("Error", f"Could not display page: {e}")

    def select_save_location(self):
        self.save_dir = filedialog.askdirectory()
        if self.save_dir:
            self.save_entry.delete(0, tk.END)
            self.save_entry.insert(0, self.save_dir)

    def on_button_press(self, event):
        if self.rect:
            self.canvas.delete(self.rect)
            self.rect = None
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red")

    def on_mouse_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_button_release(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        self.bbox = (self.start_x, self.start_y, end_x, end_y)

    def convert_pdf_to_excel(self):
        if not self.pdf_path or not self.save_dir or not self.bbox:
            messagebox.showerror("Error", "Please select a PDF, save location, and draw a bounding box.")
            return

        start_page = int(self.start_entry.get())
        end_page = int(self.end_entry.get())
        save_filename = self.filename_entry.get()
        save_path = os.path.join(self.save_dir, save_filename)

        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                num_list = []
                str_list = []
                current_num = None
                current_str = ""

                for i in range(start_page - 1, end_page):
                    page = pdf.pages[i]
                    page_width = page.width
                    page_height = page.height

                    x0, y0, x1, y1 = self.bbox
                    x0 = (x0 / self.tk_image.width()) * page_width
                    y0 = (y0 / self.tk_image.height()) * page_height
                    x1 = (x1 / self.tk_image.width()) * page_width
                    y1 = (y1 / self.tk_image.height()) * page_height
                    bbox = (x0, y0, x1, y1)

                    page_text = page.within_bbox(bbox).extract_text()
                    if page_text:
                        lines = page_text.split('\n')

                        for line in lines:
                            match = re.match(r'^(\d+(\.\d+)*)(\s+|\s*[-:)]\s*)(.*)', line.strip())
                            if match:
                                if current_num is not None:
                                    num_list.append(current_num)
                                    str_list.append(current_str.strip())
                                
                                current_num = match.group(1).strip()
                                current_str = match.group(4).strip()
                            else:
                                if current_num is not None:
                                    current_str += " " + line.strip()

                # 마지막 문단 번호 처리
                if current_num is not None:
                    num_list.append(current_num)
                    str_list.append(current_str.strip())

                df = create_dataframe(num_list, str_list)
                save_to_excel(df, save_path)
                messagebox.showinfo("Success", f"PDF converted to Excel successfully!\nFile saved at: {save_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = PDFBoundingBoxSelector()
    app.mainloop()

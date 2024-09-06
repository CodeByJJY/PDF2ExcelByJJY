import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import pdfplumber
from PIL import Image, ImageTk
import os

# extract_text_from_pdf: PDF 파일의 특정 페이지 범위에서 지정된 bounding box(bbox) 내의 텍스트 추출
def extract_text_from_pdf(pdf_path, start_page, end_page, bbox):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for i in range(start_page - 1, end_page):
            page = pdf.pages[i]
            page_text = page.within_bbox(bbox).extract_text()
            if page_text:
                text += page_text + '\n'
    return text

def classify_number(s):
    if re.match(r'^\d+\.$', s):
        return True  # "1."과 같은 형식
    else:
        return False

# list_to_string: 리스트의 요소를 문자열로 변환하고, 구분자로 합침
def list_to_string(input_list, delimiter=' '):
    return delimiter.join(map(str, input_list))

# process_text: 추출된 텍스트를 번호 - 문서 부분으로 분리하여 각각의 리스트에 저장.
def process_text(text):
    lines = text.split('\n')
    num_list = []
    str_list = []
    current_num = ""
    current_str = ""
    
    for line in lines:
        parts = line.split(maxsplit=1)
        # print(parts)
        if len(parts) == 2 and parts[0].replace(".", "").isdigit() and not classify_number(list_to_string(parts[0],'')):
            if current_num and current_str:
                num_list.append(current_num)
                str_list.append(current_str.strip())
                current_num = ""
                current_str = ""
            current_num = parts[0]
            current_str = parts[1]
        else:
            if current_str:
                current_str += " " + line
            else:
                current_str = line
        print(current_str)

    if current_num and current_str:
        num_list.append(current_num)
        str_list.append(current_str.strip())

    return num_list, str_list

# remove_illegal_characters: 문자열에서 허용되지 않는 문자 제거.
def remove_illegal_characters(text):
    return re.sub(r'[^\x20-\x7E]', '', text)

# create_dataframe: 번호 리스트와 텍스트 리스트를 데이터프레임으로 변환.
def create_dataframe(num_list, str_list):
    str_list = [remove_illegal_characters(s) for s in str_list]
    rows = []
    for num, doc in zip(num_list, str_list):
        sub_lines = doc.split('\n')
        rows.append([num, sub_lines[0]])
        for sub_line in sub_lines[1:]:
            rows.append(["", sub_line])
    return pd.DataFrame(rows, columns=['No', 'Document'])

# save_to_excel: 데이터프레임을 엑셀 파일로 저장.
def save_to_excel(df, excel_path):
    df.to_excel(excel_path, index=False)


# PDFBoundingBoxSelector 클래스: Tkinter를 사용하여 GUI를 생성
#                               사용자가 PDF 파일을 선택하고 bounding box를 설정하여 텍스트 추출
#                               엑셀로 변환. 
class PDFBoundingBoxSelector(tk.Tk):
    # init: GUI의 초기 설정 수행.
    def __init__(self):
        super().__init__()
        self.title("PDF to Excel Converter")

        self.pdf_path = None
        self.save_dir = None
        self.save_filename = None
        self.bbox = None

        self.create_widgets()

    # create_widgets: GUI 위젯 생성
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
        self.filename_entry.insert(0, "Excelfilename.xlsx")

        tk.Label(self, text="Start Page:").grid(row=3, column=0, padx=10, pady=5)
        self.start_entry = tk.Entry(self, width=10)
        self.start_entry.grid(row=3, column=1, padx=10, pady=5, sticky='w')
        self.start_entry.insert(0, "1")
        self.start_entry.bind("<Return>", lambda event: self.display_pdf_page())

        tk.Label(self, text="End Page:").grid(row=4, column=0, padx=10, pady=5)
        self.end_entry = tk.Entry(self, width=10)
        self.end_entry.grid(row=4, column=1, padx=10, pady=5, sticky='w')
        self.end_entry.insert(0, "1")

        self.canvas = tk.Canvas(self, cursor="cross")
        self.canvas.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

        tk.Button(self, text="Convert", command=self.convert_pdf_to_excel).grid(row=6, column=0, columnspan=3, pady=10)

        self.canvas.bind("<Button-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

        self.start_x = None
        self.start_y = None
        self.rect = None

    # select_pdf: PDF 파일 및 저장 경로 선택
    def select_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not self.pdf_path:
            return

        self.pdf_entry.delete(0, tk.END)
        self.pdf_entry.insert(0, self.pdf_path)

        self.display_pdf_page()

    # display_pdf_page: 선택한 PDF 페이지를 화면에 표시
    def display_pdf_page(self):
        start_page = int(self.start_entry.get())
        if not self.pdf_path:
            return
        
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                page = pdf.pages[start_page - 1]
                page_image = page.to_image()
                page_image_path = 'page_preview.png'
                page_image.save(page_image_path)

                self.image = Image.open(page_image_path)
                screen_width = self.winfo_screenwidth()
                screen_height = self.winfo_screenheight()

                scale_factor = min(screen_width / self.image.width, screen_height / self.image.height) * 0.6
                new_width = int(self.image.width * scale_factor)
                new_height = int(self.image.height * scale_factor)
                self.image = self.image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                self.tk_image = ImageTk.PhotoImage(self.image)

                self.canvas.config(width=new_width, height=new_height)
                self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)
        except Exception as e:
            messagebox.showerror("Error", f"Could not display page: {e}")

    # select_save_location: 사용자가 PDF 파일의 텍스트를 변환한 후 저장할 디렉토리를 선택할 수 있도록 함.
    def select_save_location(self):
        self.save_dir = filedialog.askdirectory()
        if self.save_dir:
            self.save_entry.delete(0, tk.END)
            self.save_entry.insert(0, self.save_dir)

    # PDF를 Excel로 변환: 사용자가 마우스를 사용하여 선택한 영역을 bounding box으로 설정.
    def on_button_press(self, event):
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
        print(f"Selected bounding box: {self.bbox}")

    # convert_pdf_to_excel: Bounding box를 기반으로 PDF에서 텍스트를 추출하고 엑셀 파일로 저장.
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
                text = ""
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
                        text += page_text + '\n'

            num_list, str_list = process_text(text)
            df = create_dataframe(num_list, str_list)
            save_to_excel(df, save_path)
            messagebox.showinfo("Success", f"PDF converted to Excel successfully!\nFile saved at: {save_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = PDFBoundingBoxSelector()
    app.mainloop()
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from PyPDF2 import PdfReader  # Используем PyPDF2 для работы с PDF
import openpyxl


class PDFProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Processor")

        # Переменные для хранения путей
        self.input_pdf_path = ""
        self.output_excel_path = ""

        # Создание GUI элементов
        self.create_widgets()

    def create_widgets(self):
        # Кнопка для выбора PDF
        tk.Button(self.root, text="Выбрать PDF", command=self.load_pdf).pack(pady=10)

        # Поле для отображения пути к PDF
        self.pdf_path_label = tk.Label(self.root, text="Файл не выбран")
        self.pdf_path_label.pack()

        # Кнопка обработки
        tk.Button(self.root, text="Распознать документ", command=self.process_pdf).pack(pady=10)

        # Кнопка для выбора Excel
        tk.Button(self.root, text="Выбрать Excel для сохранения", command=self.select_excel).pack(pady=10)

        # Поле для отображения пути к Excel
        self.excel_path_label = tk.Label(self.root, text="Файл не выбран")
        self.excel_path_label.pack()

    def load_pdf(self):
        self.input_pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        self.pdf_path_label.config(text=self.input_pdf_path)

    def select_excel(self):
        self.output_excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                              filetypes=[("Excel Files", "*.xlsx")])
        self.excel_path_label.config(text=self.output_excel_path)

    def process_pdf(self):
        if not self.input_pdf_path:
            messagebox.showerror("Ошибка", "Сначала выберите файл PDF!")
            return

        data = []
        try:
            with open(self.input_pdf_path, 'rb') as file:
                reader = PdfReader(file)
                for page in reader.pages:
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        for line in lines:
                            if line.strip():
                                parts = line.split()
                                if len(parts) >= 6:
                                    data.append({
                                        'Код расценки': parts[1],
                                        'Комплексная расценка, ед.изм': parts[2],
                                        'Ед. изм.': parts[3],
                                        'Физический объем конструктива': parts[4],
                                        'Стоимость за единицу объема с НДС': parts[5]
                                    })

            # Проверка, выбран ли файл для сохранения
            if not self.output_excel_path:
                messagebox.showerror("Ошибка", "Сначала выберите файл Excel для сохранения!")
                return

            # Сохранение в Excel
            df = pd.DataFrame(data)
            df.to_excel(self.output_excel_path, index=False, sheet_name='Данные')

            messagebox.showinfo("Успех", "Данные успешно сохранены!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при обработке PDF: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessor(root)
    root.mainloop()

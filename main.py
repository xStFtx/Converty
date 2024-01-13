import tkinter as tk
from tkinter import filedialog, messagebox
import ebooklib
from ebooklib import epub
import pdfkit
from pdf2docx import Converter
import os

wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
import subprocess

def convert_epub_to_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("EPUB files", "*.epub")])
    if file_path:
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        save_path = os.path.join(os.path.dirname(file_path), f"{base_name}.pdf")
        
        try:
            # Using calibre's ebook-convert command line tool to convert EPUB to PDF
            subprocess.run(['ebook-convert', file_path, save_path], check=True)
            messagebox.showinfo("Success", f"File saved as {save_path}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Conversion Error", f"An error occurred: {e}")
        except Exception as e:
            messagebox.showerror("Conversion Error", str(e))


def convert_pdf_to_epub():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        save_path = os.path.join(os.path.dirname(file_path), f"{base_name}.epub")
        try:
            converter = Converter(file_path)
            converter.convert(save_path, start=0, end=None)
            converter.close()
            messagebox.showinfo("Success", f"File saved as {save_path}")
        except Exception as e:
            messagebox.showerror("Conversion Error", str(e))

root = tk.Tk()
root.title("File Format Converter")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

epub_to_pdf_button = tk.Button(frame, text="Convert EPUB to PDF", command=convert_epub_to_pdf)
epub_to_pdf_button.pack(pady=(0, 5))

pdf_to_epub_button = tk.Button(frame, text="Convert PDF to EPUB", command=convert_pdf_to_epub)
pdf_to_epub_button.pack()

root.mainloop()

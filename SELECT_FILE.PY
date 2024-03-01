# import tkinter as tk
# from tkinter import filedialog

# def browse_file():
#     root = tk.Tk()
#     root.withdraw()
#     file_path = filedialog.askopenfilename()
#     return file_path

# def main():
#     file_path = browse_file()
#     if file_path:
#         print("Selected file:", file_path)
#     else:
#         print("No file selected!")

# if __name__ == "__main__":
    # main()

import os
import win32com.client
from docx import Document
from fpdf import FPDF
import tkinter as tk
from tkinter import filedialog

# Define conversion functions for DOC files
def convert_doc_to_docx(doc_file):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_file)
    doc.SaveAs(doc_file.replace(".doc", ".docx"), 12)  # 12 for .docx format
    doc.Close()
    word.Quit()

def convert_doc_to_pdf(doc_file):
    docx_file = doc_file.replace(".doc", ".docx")
    convert_doc_to_docx(doc_file)
    convert_docx_to_pdf(docx_file)

# Define conversion functions for DOCX files
def convert_docx_to_pdf(docx_file):
    pdf = FPDF()
    doc = Document(docx_file)
    for para in doc.paragraphs:
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt=para.text, ln=True)
    pdf.output(docx_file.replace(".docx", ".pdf"))

# GUI to browse files
def browse_file(file_type):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[(f"{file_type.upper()} files", f"*.{file_type}")])
    return file_path

def display_menu():
    print("Select a file type to convert:")
    print("1. DOC")
    print("2. DOCX")
    choice = input("Enter your choice (1-2): ")
    return choice

def main():
    choice = display_menu()
    file_types = {1: "doc", 2: "docx"}

    file_type = file_types.get(int(choice))
    if file_type is None:
        print("Invalid choice!")
        return

    file_path = browse_file(file_type)
    if not file_path:
        print("No file selected!")
        return

    if file_type == "doc":
        convert_doc_to_pdf(file_path)
        print("Conversion complete!")
    else:
        print("Unsupported file type!")

if __name__ == "__main__":
    main()

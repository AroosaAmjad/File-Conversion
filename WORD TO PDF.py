import tkinter as tk
from tkinter import filedialog
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

def convert_to_pdf(file_path):
    # Load the Word document
    doc = Document(file_path)

    # Extract text from the Word document
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"

    # Create PDF
    pdf_path = file_path.replace(".docx", ".pdf")
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    styles = getSampleStyleSheet()
    content = [Paragraph(text, styles["Normal"])]
    doc.build(content)

    return pdf_path

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        pdf_path = convert_to_pdf(file_path)
        status_label.config(text=f"File converted to PDF: {pdf_path}", fg="green")

def create_gui():
    root = tk.Tk()
    root.title("Word to PDF Converter")
    root.geometry("400x200")
    root.configure(bg="#f0f0f0")

    header_label = tk.Label(root, text="Convert Word to PDF", font=("Arial", 14, "bold"), bg="#f0f0f0")
    header_label.pack(pady=10)

    instruction_label = tk.Label(root, text="Select a Word file (.docx) to convert:", font=("Arial", 10), bg="#f0f0f0")
    instruction_label.pack()

    select_button = tk.Button(root, text="Select File", command=select_file, bg="#4CAF50", fg="white", font=("Arial", 10))
    select_button.pack(pady=10)

    global status_label
    status_label = tk.Label(root, text="", font=("Arial", 10), bg="#f0f0f0")
    status_label.pack()

    root.mainloop()

if __name__ == "__main__":
    create_gui()

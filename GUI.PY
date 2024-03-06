import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas

def convert_to_pdf(text, output_file):
    c = canvas.Canvas(output_file, pagesize=letter)
    c.drawString(100, 750, text)
    c.save()

def convert_word_to_pdf(file_path):
    try:
        # Load the Word document
        doc = Document(file_path)

        # Extract text from the Word document
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"

        # Create PDF
        pdf_path = file_path.replace(".docx", ".pdf")
        convert_to_pdf(text, pdf_path)

        return pdf_path
    except Exception as e:
        return str(e)

def display_option(option):
    if option == "Word to PDF":
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if file_path:
            pdf_path = convert_word_to_pdf(file_path)
            if pdf_path:
                status_label.config(text=f"File converted to PDF: {pdf_path}", foreground="green")
            else:
                status_label.config(text="Conversion failed.", foreground="red")
    elif option == "Text to PDF":
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, 'r') as file:
                text = file.read()
                output_file = file_path.replace('.txt', '.pdf')
                convert_to_pdf(text, output_file)
                status_label.config(text=f"File converted to PDF: {output_file}", foreground="green")
    elif option == "CSV to Excel":
        csv_to_excel()
    elif option == "Excel to CSV":  # Integration for Excel to CSV
        excel_to_csv()
    else:
        status_label.config(text=f"Selected Option: {option}")

def csv_to_excel():
    csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if csv_file_path:
        try:
            df = pd.read_csv(csv_file_path)
            excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if excel_file_path:
                df.to_excel(excel_file_path, index=False)
                messagebox.showinfo("Success", f"Excel file saved successfully: {excel_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

def excel_to_csv():  # New function for Excel to CSV
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            csv_file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if csv_file_path:
                df.to_csv(csv_file_path, index=False)
                messagebox.showinfo("Success", f"CSV file saved successfully: {csv_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

def create_gui():
    root = tk.Tk()
    root.title("Convert Your Files")
    root.geometry("600x600")

    options = [
        "Word to PDF",
        "Text to PDF",
        "CSV to Excel",
        "Excel to CSV",
        "PDF to Text",
        "PDF to Image",
        "Image to PDF",
        "HTML to PDF",
        "HTML to DOCS",
        "HTML to Text",
        "PPT to PDF",
        "Audio/Video Conversion"
    ]

    heading_label = ttk.Label(root, text="CONVERT YOUR FILES", font=("Arial", 20, "bold"), foreground="#333")
    heading_label.pack(pady=20)

    for option in options:
        frame = ttk.Frame(root)
        frame.pack(pady=5, padx=10, anchor="w")

        label = ttk.Label(frame, text=option, font=("Arial", 14), foreground="#333")
        label.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)

        button = ttk.Button(frame, text="Select", command=lambda opt=option: display_option(opt))
        button.pack(side=tk.LEFT, fill=tk.X, expand=True)

    global status_label
    status_label = ttk.Label(root, text="", font=("Arial", 12))
    status_label.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_gui()

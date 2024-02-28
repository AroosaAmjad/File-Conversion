import tkinter as tk
from tkinter import filedialog, messagebox
import pdfkit

def html_to_pdf(html_file, output_file):
    try:
        pdfkit.from_file(html_file, output_file)
        messagebox.showinfo("Success", f"PDF file saved successfully: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def select_html():
    html_file = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
    if html_file:
        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_file:
            html_to_pdf(html_file, output_file)

def create_gui():
    root = tk.Tk()
    root.title("HTML to PDF Converter")
    root.geometry("400x150")
    root.configure(bg="#f0f0f0")

    header_label = tk.Label(root, text="Convert HTML to PDF", font=("Arial", 14, "bold"), bg="#f0f0f0")
    header_label.pack(pady=10)

    select_button = tk.Button(root, text="Select HTML File", command=select_html, bg="#4CAF50", fg="white", font=("Arial", 10))
    select_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()

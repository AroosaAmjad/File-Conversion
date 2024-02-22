import tkinter as tk
from tkinter import filedialog
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def convert_to_pdf(text, output_file):
    c = canvas.Canvas(output_file, pagesize=letter)
    c.drawString(100, 750, text)
    c.save()

def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        with open(file_path, 'r') as file:
            text = file.read()
            output_file = file_path.replace('.txt', '.pdf')
            convert_to_pdf(text, output_file)
            status_label.config(text=f"File converted to PDF: {output_file}", fg="green")

def create_gui():
    root = tk.Tk()
    root.title("Text to PDF Converter")
    root.geometry("800x300")
    root.configure(bg="#f0f0f0")

    header_label = tk.Label(root, text="Convert Text to PDF", font=("Arial", 14, "bold"), bg="#f0f0f0")
    header_label.pack(pady=10)

    instruction_label = tk.Label(root, text="Select a text file (.txt) to convert:", font=("Arial", 10), bg="#f0f0f0")
    instruction_label.pack()

    select_button = tk.Button(root, text="Select File", command=select_file, bg="#4CAF50", fg="white", font=("Arial", 10))
    select_button.pack(pady=10)

    global status_label
    status_label = tk.Label(root, text="", font=("Arial", 10), bg="#f0f0f0")
    status_label.pack()

    root.mainloop()

if __name__ == "__main__":
    create_gui()

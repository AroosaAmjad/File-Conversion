# import tkinter as tk
# from tkinter import ttk

# def create_gui():
#     root = tk.Tk()
#     root.title("Convert Your Files")
#     root.geometry("400x600")

#     def display_option(option):
#         label.config(text=f"Selected Option: {option}")

#     options = [
#         "Text to PDF", 
#         "CSV to Excel", 
#         "Excel to CSV", 
#         "Word to PDF",
#         "PDF to Text",
#         "PDF to Image",
#         "Image to PDF",
#         "HTML to PDF",
#         "HTML to DOCS",
#         "HTML to Text",
#         "PPT to PDF",
#         "Audio/Video Conversion"
#     ]

#     heading_label = ttk.Label(root, text="CONVERT YOUR FILES", font=("Arial", 16, "bold"))
#     heading_label.pack(pady=10)

#     for option in options:
#         frame = ttk.Frame(root)
#         frame.pack(pady=5, padx=10, anchor="w")

#         label = ttk.Label(frame, text=option, font=("Arial", 12))
#         label.pack(side=tk.LEFT)

#         button = ttk.Button(frame, text="Select", command=lambda opt=option: option_selected(opt))
#         button.pack(side=tk.LEFT, padx=10)

#     def option_selected(option):
#         display_option(option)

#     label = ttk.Label(root, text="", font=("Arial", 12))
#     label.pack(pady=10)

#     root.mainloop()

# if __name__ == "__main__":
#     create_gui()

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def convert_to_pdf(text, output_file):
    c = canvas.Canvas(output_file, pagesize=letter)
    c.drawString(100, 750, text)
    c.save()

def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()
        except UnicodeDecodeError:
            try:
                with open(file_path, 'r', encoding='latin-1') as file:
                    text = file.read()
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while reading the file: {e}")
                return

        output_file = file_path.replace('.txt', '.pdf')
        convert_to_pdf(text, output_file)
        status_label.config(text=f"File converted to PDF: {output_file}", fg="green")

def create_gui():
    root = tk.Tk()
    root.title("Convert Your Files")
    root.geometry("400x600")

    def display_option(option):
        label.config(text=f"Selected Option: {option}")
        if option == "Text to PDF":
            select_file()

    options = [
        "Text to PDF", 
        "CSV to Excel", 
        "Excel to CSV", 
        "Word to PDF",
        "PDF to Text",
        "PDF to Image",
        "Image to PDF",
        "HTML to PDF",
        "HTML to DOCS",
        "HTML to Text",
        "PPT to PDF",
        "Audio/Video Conversion"
    ]

    heading_label = ttk.Label(root, text="CONVERT YOUR FILES", font=("Arial", 16, "bold"))
    heading_label.pack(pady=10)

    for option in options:
        frame = ttk.Frame(root)
        frame.pack(pady=5, padx=10, anchor="w")

        label = ttk.Label(frame, text=option, font=("Arial", 12))
        label.pack(side=tk.LEFT)

        button = ttk.Button(frame, text="Select", command=lambda opt=option: option_selected(opt))
        button.pack(side=tk.LEFT, padx=10)

    def option_selected(option):
        display_option(option)

    label = ttk.Label(root, text="", font=("Arial", 12))
    label.pack(pady=10)

    global status_label
    status_label = ttk.Label(root, text="", font=("Arial", 10), background="#f0f0f0")
    status_label.pack()

    root.mainloop()

if __name__ == "__main__":
    create_gui()

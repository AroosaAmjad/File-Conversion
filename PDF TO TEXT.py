import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF

def pdf_to_text():
    file_path = filedialog.askopenfilename()
    if file_path:
        try:
            text = ''
            with fitz.open(file_path) as pdf_file:
                for page_num in range(len(pdf_file)):
                    page = pdf_file[page_num]
                    text += page.get_text()
            text_output.delete('1.0', tk.END)
            text_output.insert(tk.END, text)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create main window
root = tk.Tk()
root.title("PDF to Text Converter")

# Create button to open PDF file
open_button = tk.Button(root, text="Open PDF", command=pdf_to_text)
open_button.pack(pady=10)

# Create text widget to display the extracted text with scrollbar
text_frame = tk.Frame(root)
text_frame.pack(padx=10, pady=10)

text_output = tk.Text(text_frame, height=20, width=50)
text_output.pack(side=tk.LEFT, fill=tk.Y)

scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_output.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
text_output.config(yscrollcommand=scrollbar.set)

# Run the main event loop
root.mainloop()

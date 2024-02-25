import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from bs4 import BeautifulSoup
import requests

def html_to_docx(url):
    try:
        # Fetch the HTML content from the URL
        response = requests.get(url)
        response.raise_for_status()  # Raise an error for bad status codes
        
        # Parse the HTML content
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Extract text from the parsed HTML
        text = soup.get_text(separator='\n', strip=True)
        
        # Create a new Document
        doc = Document()
        
        # Add the extracted text to the document
        doc.add_paragraph(text)
        
        return doc
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        return None

def convert_html_to_docx():
    url = url_entry.get()
    if url:
        doc = html_to_docx(url)
        if doc:
            output_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
            if output_file:
                doc.save(output_file)
                messagebox.showinfo("Success", f"DOCX file saved successfully: {output_file}")

# Create main window
root = tk.Tk()
root.title("HTML to DOCX Converter")

# URL input
url_label = tk.Label(root, text="Enter URL:", font=("Arial", 10))
url_label.pack(pady=5)
url_entry = tk.Entry(root, width=50)
url_entry.pack(pady=5)

# Convert button
convert_button = tk.Button(root, text="Convert", command=convert_html_to_docx)
convert_button.pack(pady=5)

root.mainloop()

import tkinter as tk
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup

def html_to_text(url):
    try:
        # Fetch the HTML content from the URL
        response = requests.get(url)
        response.raise_for_status()  # Raise an error for bad status codes
        
        # Parse the HTML content
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Extract text from the parsed HTML
        text = soup.get_text(separator='\n', strip=True)
        
        return text
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        return None

def convert_html_to_text():
    url = url_entry.get()
    if url:
        text = html_to_text(url)
        if text:
            text_output.delete('1.0', tk.END)
            text_output.insert(tk.END, text)

# Create main window
root = tk.Tk()
root.title("HTML to Text Converter")

# URL input
url_label = tk.Label(root, text="Enter URL:", font=("Arial", 10))
url_label.pack(pady=5)
url_entry = tk.Entry(root, width=50)
url_entry.pack(pady=5)

# Convert button
convert_button = tk.Button(root, text="Convert", command=convert_html_to_text)
convert_button.pack(pady=5)

# Text output
text_output = tk.Text(root, height=20, width=80)
text_output.pack(pady=10)

root.mainloop()

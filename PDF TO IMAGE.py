import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF

def pdf_to_images():
    file_path = filedialog.askopenfilename()
    if file_path:
        try:
            images = []
            with fitz.open(file_path) as pdf_file:
                for page_num in range(len(pdf_file)):
                    page = pdf_file.load_page(page_num)
                    image = page.get_pixmap()
                    images.append(image)
            
            save_images(images)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

def save_images(images):
    output_folder = filedialog.askdirectory()
    if output_folder:
        try:
            for i, image in enumerate(images):
                image.save(f"{output_folder}/page_{i+1}.png")
            messagebox.showinfo("Success", "Images saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving images: {str(e)}")

# Create main window
root = tk.Tk()
root.title("PDF to Image Converter")

# Styling
root.configure(bg="#f0f0f0")
font_style = ("Arial", 12)

# Adjust window size
window_width = 300
window_height = 150
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = (screen_width/2) - (window_width/2)
y_coordinate = (screen_height/2) - (window_height/2)
root.geometry(f"{window_width}x{window_height}+{int(x_coordinate)}+{int(y_coordinate)}")

# Create button to open PDF file and convert it to images
convert_button = tk.Button(root, text="Convert PDF to Images", command=pdf_to_images, font=font_style, bg="#4CAF50", fg="white")
convert_button.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

# Run the main event loop
root.mainloop()

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def excel_to_csv():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            csv_file_path = file_path.rsplit(".", 1)[0] + ".csv"
            df.to_csv(csv_file_path, index=False)
            messagebox.showinfo("Success", f"CSV file saved successfully: {csv_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

def create_gui():
    root = tk.Tk()
    root.title("Excel to CSV Converter")
    root.geometry("400x150")
    root.configure(bg="#f0f0f0")

    header_label = tk.Label(root, text="Convert Excel to CSV", font=("Arial", 14, "bold"), bg="#f0f0f0")
    header_label.pack(pady=10)

    convert_button = tk.Button(root, text="Convert Excel to CSV", command=excel_to_csv, bg="#4CAF50", fg="white", font=("Arial", 10))
    convert_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()

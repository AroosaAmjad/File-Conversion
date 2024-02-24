import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

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

def create_gui():
    root = tk.Tk()
    root.title("CSV to Excel Converter")
    root.geometry("400x150")
    root.configure(bg="#f0f0f0")

    header_label = tk.Label(root, text="Convert CSV to Excel", font=("Arial", 14, "bold"), bg="#f0f0f0")
    header_label.pack(pady=10)

    convert_button = tk.Button(root, text="Convert CSV to Excel", command=csv_to_excel, bg="#4CAF50", fg="white", font=("Arial", 10))
    convert_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()

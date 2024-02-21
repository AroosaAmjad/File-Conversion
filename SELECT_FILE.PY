import tkinter as tk
from tkinter import filedialog

def browse_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path

def main():
    file_path = browse_file()
    if file_path:
        print("Selected file:", file_path)
    else:
        print("No file selected!")

if __name__ == "__main__":
    main()

import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter

class PDFtoWordConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Word Converter")

        self.label = tk.Label(root, text="Select a PDF file to convert:")
        self.label.pack(pady=10)

        self.select_button = tk.Button(root, text="Browse", command=self.browse_file)
        self.select_button.pack(pady=5)

        self.convert_button = tk.Button(root, text="Convert to Word", command=self.convert_to_word)
        self.convert_button.pack(pady=20)

        self.filepath = ''

    def browse_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.filepath:
            self.label.config(text=f"Selected file: {self.filepath.split('/')[-1]}")

    def convert_to_word(self):
        if not self.filepath:
            messagebox.showwarning("No file selected", "Please select a PDF file first.")
            return

        docx_filepath = self.filepath.replace('.pdf', '.docx')

        try:
            cv = Converter(self.filepath)
            cv.convert(docx_filepath, start=0, end=None)
            cv.close()
            messagebox.showinfo("Success", f"File converted successfully: {docx_filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoWordConverter(root)
    root.mainloop()
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap import ttk
import fitz  # PyMuPDF
from docx import Document
import os


def load_pdf(file_path):
    """Load text from a PDF file."""
    text = ""
    try:
        pdf_document = fitz.open(file_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text += page.get_text()
        pdf_document.close()
    except Exception as e:
        return f"Error loading PDF: {e}"
    return text


def load_word(file_path):
    """Load text from a Word file."""
    text = ""
    try:
        doc = Document(file_path)
        for para in doc.paragraphs:
            text += para.text + "\n"
    except Exception as e:
        return f"Error loading Word document: {e}"
    return text


def load_cv(file_path):
    """Load text from a CV file (PDF or Word)."""
    _, file_extension = os.path.splitext(file_path)
    file_extension = file_extension.lower()

    if file_extension == '.pdf':
        return load_pdf(file_path)
    elif file_extension == '.docx':
        return load_word(file_path)
    else:
        return "Unsupported file format. Please upload a PDF or Word document."


def open_file():
    """Open file dialog and display CV content."""
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf"), ("Word files", "*.docx")]
    )
    if file_path:
        cv_text = load_cv(file_path)
        if cv_text.startswith("Error") or cv_text.startswith("Unsupported"):
            messagebox.showerror("Error", cv_text)
        else:
            text_box.delete(1.0, tk.END)  # Clear the text box
            text_box.insert(tk.END, cv_text)
            status_bar.config(
                text=f"Loaded file: {os.path.basename(file_path)}")


def quit_app():
    """Quit the application."""
    root.quit()


# Create the main window
root = tk.Tk()
root.title("CV Loader")
root.geometry("800x600")

# Apply the modern theme
# Change 'flatly' to any other theme provided by ttkbootstrap
style = Style(theme='flatly')

# Create and pack widgets
header_frame = ttk.Frame(root)
header_frame.pack(pady=10, fill=tk.X)

open_button = ttk.Button(header_frame, text="Open CV", command=open_file)
open_button.pack(side=tk.LEFT, padx=10)

quit_button = ttk.Button(header_frame, text="Exit", command=quit_app)
quit_button.pack(side=tk.LEFT, padx=10)

# Create a frame for text box and scrollbar
frame = ttk.Frame(root)
frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

text_box = ScrolledText(frame, wrap=tk.WORD, width=80,
                        height=20, font=('Segoe UI', 12))
text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Status bar
status_bar = ttk.Label(root, text="Welcome to CV Loader",
                       anchor=tk.W, relief=tk.SUNKEN, padding=5)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# Start the main event loop
root.mainloop()

import os
import PyPDF2
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox

def search_pdf_for_keyword(pdf_path, keyword):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page_num in range(len(reader.pages)):
                text = reader.pages[page_num].extract_text()
                if keyword.lower() in text.lower():
                    return True
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    return False

def search_docx_for_keyword(docx_path, keyword):
    try:
        doc = Document(docx_path)
        for para in doc.paragraphs:
            if keyword.lower() in para.text.lower():
                return True
    except Exception as e:
        print(f"Error reading {docx_path}: {e}")
    return False

def search_files_in_folder(folder_path, keyword):
    results = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            if file.lower().startswith('~$'):
                continue
            if file.lower().endswith('.pdf'):
                if search_pdf_for_keyword(file_path, keyword):
                    results.append(file_path)
            elif file.lower().endswith('.docx'):
                if search_docx_for_keyword(file_path, keyword):
                    results.append(file_path)
    return results

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_path_var.set(folder_selected)

def search_keyword():
    folder_path = folder_path_var.get()
    keyword = keyword_var.get()
    if not folder_path or not keyword:
        messagebox.showwarning("Input Error", "Please provide both folder path and keyword.")
        return

    results = search_files_in_folder(folder_path, keyword)
    
    if results:
        result_text.set(f"Found the keyword '{keyword}' in the following files:\n" + "\n".join(results))
    else:
        result_text.set(f"No files found containing the keyword '{keyword}'.")

# Creating GUI
root = tk.Tk()
root.title("PDF and DOCX Keyword Search")

tk.Label(root, text="Folder Path:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
folder_path_var = tk.StringVar()
tk.Entry(root, textvariable=folder_path_var, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_folder).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Keyword:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
keyword_var = tk.StringVar()
tk.Entry(root, textvariable=keyword_var, width=50).grid(row=1, column=1, padx=10, pady=10)

tk.Button(root, text="Search", command=search_keyword).grid(row=2, column=0, columnspan=3, pady=10)

result_text = tk.StringVar()
tk.Label(root, textvariable=result_text, justify='left').grid(row=3, column=0, columnspan=3, padx=10, pady=10)

root.mainloop()
    
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert

def browse_word_file():
    global word_file
    word_file = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    print("Selected Word file:", word_file)

def browse_output_folder():
    global output_folder
    output_folder = filedialog.askdirectory()
    print("Selected output folder:", output_folder)

def convert_to_pdf():
    if word_file and output_folder:
        output_file = os.path.join(output_folder, os.path.splitext(os.path.basename(word_file))[0] + ".pdf")
        convert(word_file, output_file)
        print("PDF file saved at:", output_file)
        messagebox.showinfo("Conversion Complete", f"The PDF file has been saved at {output_file}")
    else:
        print("Please select a Word file and output folder first.")

# Create the main window
window = tk.Tk()
window.title("Word to PDF Converter")
window.geometry("600x400")
window.configure(bg="#FFF8DC")

# Add buttons and labels
select_file_btn = tk.Button(window, text="Select Word Document", command=browse_word_file, bd=2)
select_file_btn.place(relx=0.5, rely=0.3, anchor="center")

select_output_btn = tk.Button(window, text="Output File Location", command=browse_output_folder, bd=2)
select_output_btn.place(relx=0.5, rely=0.5, anchor="center")

convert_btn = tk.Button(window, text="Convert", command=convert_to_pdf, bd=2)
convert_btn.place(relx=0.5, rely=0.7, anchor="center")

# Initialize variables
word_file = None
output_folder = None

# Start the main loop
window.mainloop()

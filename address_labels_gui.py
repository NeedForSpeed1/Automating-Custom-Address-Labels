import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def create_address_label_pdf(image_path, excel_path, output_pdf_path):
    # Check if the provided files exist
    if not os.path.isfile(image_path):
        messagebox.showerror("Error", f"Image file '{image_path}' not found.")
        return
    if not os.path.isfile(excel_path):
        messagebox.showerror("Error", f"Excel file '{excel_path}' not found.")
        return

    # Read data from the Excel file
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        messagebox.showerror("Error", f"Error reading the Excel file: {e}")
        return

    # Create a PDF canvas
    c = canvas.Canvas(output_pdf_path, pagesize=A4)
    width, height = A4

    # Iterate through each row in the DataFrame to create individual address labels
    for index, row in df.iterrows():
        # Extract the data from the current row
        name = row.get("Name", "Unknown")
        address = row.get("Address", "Unknown Address")
        return_address = row.get("Return Address", "Unknown Return Address")

        # Clear the page for a new label (or save and create a new page if multiple labels per page)
        c.showPage()

        # Add the image to the top of the label
        image_width = 3 * inch
        image_height = 1 * inch
        x_position = (width - image_width) / 2
        y_position = height - image_height
        c.drawImage(image_path, x=x_position, y=y_position, width=image_width, height=image_height)

        # Set the starting y-position for the text
        text_y_start = y_position - 0.5 * inch
        line_height = 0.5 * inch

        # Add the label text (Name, Address, Return Address)
        c.setFont("Helvetica", 12)
        c.drawString(1 * inch, text_y_start, f"Name: {name}")
        c.drawString(1 * inch, text_y_start - line_height, f"Address: {address}")
        c.drawString(1 * inch, text_y_start - 2 * line_height, f"Return Address: {return_address}")

    # Save the PDF
    c.save()

    messagebox.showinfo("Success", f"PDF '{output_pdf_path}' created successfully.")

def select_image():
    file_path = filedialog.askopenfilename(title="Select Image", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")])
    image_entry.delete(0, tk.END)
    image_entry.insert(0, file_path)

def select_excel():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
    excel_entry.delete(0, tk.END)
    excel_entry.insert(0, file_path)

def select_output():
    file_path = filedialog.asksaveasfilename(title="Save Output PDF", defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, file_path)

def run_create_pdf():
    image_path = image_entry.get()
    excel_path = excel_entry.get()
    output_pdf_path = output_entry.get()

    if not image_path or not excel_path or not output_pdf_path:
        messagebox.showerror("Error", "Please select all the required files.")
        return

    create_address_label_pdf(image_path, excel_path, output_pdf_path)

# Create the main application window
app = tk.Tk()
app.title("Address Label Creator")

# Image file selection
tk.Label(app, text="Select Image:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
image_entry = tk.Entry(app, width=50)
image_entry.grid(row=0, column=1, padx=10, pady=5)
tk.Button(app, text="Browse", command=select_image).grid(row=0, column=2, padx=10, pady=5)

# Excel file selection
tk.Label(app, text="Select Excel File:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
excel_entry = tk.Entry(app, width=50)
excel_entry.grid(row=1, column=1, padx=10, pady=5)
tk.Button(app, text="Browse", command=select_excel).grid(row=1, column=2, padx=10, pady=5)

# Output file selection
tk.Label(app, text="Save Output PDF As:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
output_entry = tk.Entry(app, width=50)
output_entry.grid(row=2, column=1, padx=10, pady=5)
tk.Button(app, text="Browse", command=select_output).grid(row=2, column=2, padx=10, pady=5)

# Create PDF button
tk.Button(app, text="Create PDF", command=run_create_pdf, width=20).grid(row=3, column=1, padx=10, pady=20)

# Run the application
app.mainloop()

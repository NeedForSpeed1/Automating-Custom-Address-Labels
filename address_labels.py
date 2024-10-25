import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch

def create_address_label_pdf(image_path, excel_path, output_pdf_path):
    # Read data from the Excel file
    df = pd.read_excel(excel_path)

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
        c.drawImage(image_path, x=width/2 - 1.5*inch, y=height - 2*inch, width=3*inch, height=1*inch)

        # Add the label text (Name, Address, Return Address)
        c.setFont("Helvetica", 12)
        c.drawString(1*inch, height - 2.5*inch, f"Name: {name}")
        c.drawString(1*inch, height - 3*inch, f"Address: {address}")
        c.drawString(1*inch, height - 3.5*inch, f"Return Address: {return_address}")

    # Save the PDF
    c.save()

    print(f"PDF '{output_pdf_path}' created successfully.")

# Example usage
image_path = "path/to/your/image.png"        # Replace with your image file path
excel_path = "path/to/your/excel_file.xlsx"  # Replace with your Excel file path
output_pdf_path = "address_labels.pdf"       # Desired output PDF path

create_address_label_pdf(image_path, excel_path, output_pdf_path)

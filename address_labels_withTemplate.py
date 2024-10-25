import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO

def create_address_label_with_template(template_path, excel_path, output_pdf_path):
    # Read data from the Excel file
    df = pd.read_excel(excel_path)

    # Read the existing PDF template
    template_reader = PdfReader(template_path)
    template_page = template_reader.pages[0]
    width, height = A4

    # Create a PDF writer for the final output
    output_writer = PdfWriter()

    # Iterate through each row in the DataFrame to create individual address labels
    for index, row in df.iterrows():
        # Extract the data from the current row
        name = row.get("Name", "Unknown")
        address = row.get("Address", "Unknown Address")
        return_address = row.get("Return Address", "Unknown Return Address")

        # Create a new PDF with the address information overlay
        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=A4)

        # Set the position to account for the space the image on the template takes up
        # Adjust y-position if the image occupies a different amount of space
        y_start = height - 2.5 * inch

        # Add the label text (Name, Address, Return Address)
        c.setFont("Helvetica", 12)
        c.drawString(1*inch, y_start, f"Name: {name}")
        c.drawString(1*inch, y_start - 0.5*inch, f"Address: {address}")
        c.drawString(1*inch, y_start - 1*inch, f"Return Address: {return_address}")

        # Save the overlay PDF
        c.save()
        packet.seek(0)

        # Merge the template and the overlay
        overlay_reader = PdfReader(packet)
        overlay_page = overlay_reader.pages[0]
        template_page_copy = template_page

        # Add the overlay content to the template page
        template_page_copy.merge_page(overlay_page)

        # Add the modified page to the output writer
        output_writer.add_page(template_page_copy)

    # Save the final PDF to the output path
    with open(output_pdf_path, "wb") as output_file:
        output_writer.write(output_file)

    print(f"PDF '{output_pdf_path}' created successfully.")

# Example usage
template_path = "path/to/your/template.pdf"  # Replace with your template PDF path
excel_path = "path/to/your/excel_file.xlsx"  # Replace with your Excel file path
output_pdf_path = "address_labels_with_template.pdf"  # Desired output PDF path

create_address_label_with_template(template_path, excel_path, output_pdf_path)

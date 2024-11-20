import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image

# Paths to input files and output folder
excel_file = "your_excel_file.xlsx"  # Replace with your Excel file name
image_a_path = "image_a.jpg"  # Replace with your image a file path
image_b_path = "image_b.jpg"  # Replace with your image b file path
output_folder = "team_pdfs"  # Folder to save PDFs

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Load the Excel spreadsheet
df = pd.read_excel(excel_file)

# Group rows by team
grouped = df.groupby('Team')  # Assuming the column name for teams is 'Team'

# Function to generate PDF for a team
def create_team_pdf(team_name, team_data, output_folder):
    pdf_path = os.path.join(output_folder, f"{team_name}.pdf")
    pdf_canvas = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter

    # Load images
    image_a = Image.open(image_a_path)
    image_b = Image.open(image_b_path)
    image_a_width, image_a_height = image_a.size
    image_b_width, image_b_height = image_b.size

    # Iterate through rows and create a page for each record
    for _, row in team_data.iterrows():
        # Add image A to the top center
        pdf_canvas.drawImage(
            image_a_path, x=(width - image_a_width) / 2, y=height - 100 - image_a_height,
            width=image_a_width, height=image_a_height
        )

        # Write details to the PDF
        pdf_canvas.setFont("Helvetica", 12)
        pdf_canvas.drawString(100, height / 2, f"Name: {row['Name']}")
        pdf_canvas.drawString(100, height / 2 - 20, f"Age: {row['Age']}")
        pdf_canvas.drawString(100, height / 2 - 40, f"Gender: {row['Gender']}")

        # Add image B to the bottom center
        pdf_canvas.drawImage(
            image_b_path, x=(width - image_b_width) / 2, y=50,
            width=image_b_width, height=image_b_height
        )

        # Finish the page and start a new one for the next record
        pdf_canvas.showPage()

    # Save the PDF
    pdf_canvas.save()
    print(f"PDF created for team: {team_name} at {pdf_path}")

# Generate a PDF for each team
for team_name, team_data in grouped:
    create_team_pdf(team_name, team_data, output_folder)

print(f"All PDFs have been saved in the folder: {output_folder}")

import pandas as pd
from openpyxl import load_workbook
from fpdf import FPDF

# Step 1: Read data from the source Excel file
def read_source_data(source_file):
    # Load the data into a pandas DataFrame
    df = pd.read_excel(source_file)
    return df

# Step 2: Fill the Excel form template
def fill_excel_form(template_file, output_file, data):
    # Load the template workbook
    wb = load_workbook(template_file)
    sheet = wb.active

    # Example: Assuming the form has placeholders like "{{name}}" in cell A1
    # Replace placeholders with actual data
    sheet["A1"] = data["name"]  # Replace "name" with the actual key from your data
    sheet["B1"] = data["age"]   # Replace "age" with the actual key from your data
    sheet["C1"] = data["email"] # Replace "email" with the actual key from your data

    # Save the filled form
    wb.save(output_file)

# Step 3: Generate PDF from the filled data
def generate_pdf(output_file, data):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Add content to the PDF
    pdf.cell(200, 10, txt="Form Data", ln=True, align="C")
    pdf.ln(10)  # Line break
    for key, value in data.items():
        pdf.cell(0, 10, txt=f"{key}: {value}", ln=True)

    # Save the PDF
    pdf.output(output_file)

# Main function
def main():
    source_file = "source.xlsx"       # Source Excel file with data
    template_file = "template.xlsx"  # Excel form template
    filled_excel = "filled_form.xlsx" # Filled form output
    pdf_output = "output_form.pdf"   # PDF output

    # Read source data
    source_data = read_source_data(source_file)

    # Iterate over each row in the source data
    for index, row in source_data.iterrows():
        # Convert the row to a dictionary
        data = row.to_dict()

        # Fill the Excel form
        fill_excel_form(template_file, filled_excel, data)

        # Generate PDF from the data
        generate_pdf(pdf_output, data)

        print(f"Processed record {index + 1}: {data}")

if __name__ == "__main__":
    main()

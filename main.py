import os
import pandas as pd
from openpyxl import load_workbook
from fpdf import FPDF
from utils import spell_date
from fpdf import FPDF

# Step 1: Read data from a single source file
def read_source_data(source_file):
    # Load the data into a pandas DataFrame
    return pd.read_excel(source_file)

# Step 2: Fill the Excel form template
def fill_excel_form(template_file, output_file, data):
    # Load the template workbook
    wb = load_workbook(template_file)
    sheet = wb["MASTER"]

    # Get BAPWP and BAUT opener
    day = spell_date(data["Tanggal PKS"], return_format="day")
    date = spell_date(data["Tanggal PKS"], return_format="date")
    month = spell_date(data["Tanggal PKS"], return_format="month")
    year = spell_date(data["Tanggal PKS"], return_format="year")

    BAPWP_opener = f"'Pada hari ini {day} Tanggal {date} Bulan {month} Tahun {year}, telah "
    BAUT_opener = f"''Pada hari ini {day} Tanggal {date} Bulan {month} Tahun {year}, kami yang bertanda tangan"

    # Replace placeholders with actual data (update the cell references as needed)
    sheet["F5"] = data.get("Project ID", "")  
    sheet["F6"] = data.get("Site Name PO", "")  
    sheet["F7"] = data.get("Site Name Tenant", "")  
    sheet["F8"] = data.get("Site ID", "")  
    sheet["F9"] = data.get("PKS No", "")  
    sheet["H9"] = data.get("Tanggal PKS", "")  
    sheet["F10"] = data.get("PO No", "")  
    sheet["H10"] = data.get("Tanggal PO", "")  
    sheet["F11"] = data.get("Tipe Tower", "")  
    sheet["G11"] = data.get("Height", "")  
    sheet["F12"] = data.get("Alamat", "")  
    sheet["G13"] = data.get("Long", "")  
    sheet["G14"] = data.get("Lat", "")  
    sheet["F15"] = data.get("SOW", "")  
    sheet["F16"] = data.get("Area", "")  
    sheet["F17"] = data.get("Mitra", "")  
    sheet["F18"] = data.get("Nilai Kontrak", "")  
    sheet["C25"] = data.get("Jabatan GM", "")  
    sheet["F25"] = data.get("Nama GM", "")  
    sheet["C26"] = data.get("Jabatan Manager Asset", "")  
    sheet["F26"] = data.get("Nama Manager Asset", "")  
    sheet["C27"] = data.get("Jabatan Asman Asset", "")  
    sheet["F27"] = data.get("Nama Asman Asset", "")
    sheet["J10"] = data.get("Jangka Waktu Kerja", "")
    sheet["F31"] = BAPWP_opener
    sheet["F32"] = BAUT_opener

    # Save the filled form with a custom filename
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
    source_folder = "sources_file"  # Folder containing source.xlsx and template.xlsx
    source_file = os.path.join(source_folder, "source.xlsx")
    template_file = os.path.join(source_folder, "template.xlsx")

    # Ensure source and template files exist
    if not os.path.exists(source_file):
        print(f"Source file not found: {source_file}")
        return
    if not os.path.exists(template_file):
        print(f"Template file not found: {template_file}")
        return

    # Ensure output directory exists
    output_folder = "output_file"
    os.makedirs(output_folder, exist_ok=True)

    # Read source data
    source_data = read_source_data(source_file)

    # Process each row in the source data
    for index, row in source_data.iterrows():
        data = row.to_dict()  # Convert row to dictionary
        record_number = index + 1

        # Generate output file paths
        filled_excel = os.path.join(output_folder, f"BAST_{data['SOW']}_{data['Project ID']}_{data['Site Name Tenant']}.xlsx")
        pdf_output = os.path.join(output_folder, f"form_data{record_number}.pdf")

        # Fill the Excel form
        fill_excel_form(template_file, filled_excel, data)

        # Generate PDF
        # generate_pdf(pdf_output, data)

        print(f"Processed Record {record_number}: {data['SOW']} - {data['Project ID']} - {data['Site Name Tenant']}")

if __name__ == "__main__":
    main()

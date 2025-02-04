import os
import pandas as pd
import win32com.client
import fitz
from openpyxl import load_workbook
from fpdf import FPDF
from utils import spell_date
from fpdf import FPDF
from PyPDF2 import PdfMerger

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
    day = spell_date(data["Tanggal PO"], return_format="day")
    date = spell_date(data["Tanggal PO"], return_format="date")
    month = spell_date(data["Tanggal PO"], return_format="month")
    year = spell_date(data["Tanggal PO"], return_format="year")

    BAPWP_opener = f"Pada hari ini {day} Tanggal {date} Bulan {month} Tahun {year}, telah "
    BAUT_opener = f"Pada hari ini {day} Tanggal {date} Bulan {month} Tahun {year}, kami yang bertanda tangan"

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
def generate_pdf(input_excel, output_pdf):
    # Create an Excel application instance
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Run in the background

    try:
        # Open the workbook
        workbook = excel.Workbooks.Open(os.path.abspath(input_excel))
        
        # Export as PDF
        workbook.ExportAsFixedFormat(0, os.path.abspath(output_pdf))  # 0 = PDF format
        
        print(f"Successfully converted {input_excel} to {output_pdf}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Close the workbook and quit Excel
        workbook.Close(False)
        excel.Quit()

def highlight_text_in_pdf(input_pdf, output_pdf, search_text):
    doc = fitz.open(input_pdf)
    for page in doc:
        text_instances = page.search_for(search_text)
        for inst in text_instances:
            rect = inst
            # Adjust the rectangle to only cover the table width
            table_margin_left = rect.x0 - 30  # Adjust as needed
            table_margin_right = rect.x1 + 250  # Adjust based on table structure
            rect.x0 = table_margin_left
            rect.x1 = table_margin_right
            highlight = page.add_highlight_annot(rect)
            highlight.update()

            # Try to highlight the row below by shifting the rectangle downward
            line_height = rect.y1 - rect.y0  # Estimate row height
            for _ in range(2):  # Highlight two rows below
                rect.y0 += line_height
                rect.y1 += line_height
                highlight_below = page.add_highlight_annot(rect)
                highlight_below.update()

    doc.save(output_pdf)
    doc.close()

def merge_pdfs(pdf_files, output_pdf):
    """Merges multiple PDFs into a single PDF"""
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)
    merger.write(output_pdf)
    merger.close()
    # print(f"Merged PDFs into {output_pdf}")

# Main function
def main():
    source_folder = "sources_file"  # Folder containing source.xlsx and template.xlsx
    source_file = os.path.join(source_folder, "source.xlsx")
    template_file = os.path.join(source_folder, "template.xlsx")
    po_file = os.path.join(source_folder, "PO.pdf")

    # Ensure source and template files exist
    if not os.path.exists(source_file):
        print(f"Source file not found: {source_file}")
        return
    if not os.path.exists(template_file):
        print(f"Template file not found: {template_file}")
        return
    if not os.path.exists(po_file):
        print(f"PO file not found: {po_file}")
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
        pdf_output = os.path.join(output_folder, f"BAST_{data['SOW']}_{data['Project ID']}_{data['Site Name Tenant']}_excel.pdf")
        po_output = os.path.join(output_folder, "PO_Highlight.pdf")
        final_pdf = os.path.join(output_folder, f"BAST_{data['SOW']}_{data['Project ID']}_{data['Site Name Tenant']}.pdf")

        # Fill the Excel form
        fill_excel_form(template_file, filled_excel, data)

        # Generate PDF
        generate_pdf(filled_excel, pdf_output)

        # Highlight PO
        highlight_text_in_pdf(po_file, po_output, data['Project ID'])

        # Merge Excel PDF and Additional PDF
        merge_pdfs([pdf_output, po_output], final_pdf)

        print(f"Processed BAST {record_number}: {data['SOW']} - {data['Project ID']} - {data['Site Name Tenant']}")

if __name__ == "__main__":
    main()

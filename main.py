import os
import pandas as pd
import win32com.client
import fitz
from openpyxl import load_workbook
from openpyxl.styles import Font
from PyPDF2 import PdfMerger
from utils import spell_date

def read_source_data(source_file):
    """ Reads the source Excel file containing project data. """
    print(f"Reading source data from {source_file}...")
    return pd.read_excel(source_file)

def fill_excel_template(template_path, output_path, data):
    """ Fills the Excel template with project data and applies formatting. """
    print(f"Filling Excel template for {data['Project ID']}...")
    wb = load_workbook(template_path)
    sheet = wb["MASTER"]
    
    # Extract and format date
    dt = data["Tanggal PO"]
    day = spell_date(dt, "day")
    date = spell_date(dt, "date")
    month = spell_date(dt, "month")
    year = spell_date(dt, "year")

    # Mapping Excel cells to data keys
    mappings = {
        "F5": "Project ID", "F6": "Site Name PO", "F7": "Site Name Tenant",
        "F8": "Site ID", "F9": "PKS No", "H9": "Tanggal PKS", "F10": "PO No",
        "H10": "Tanggal PO", "F11": "Tipe Tower", "G11": "Height", "F12": "Alamat",
        "G13": "Long", "G14": "Lat", "F15": "SOW", "F16": "Area", "F17": "Mitra",
        "F18": "Nilai Kontrak", "C25": "Jabatan GM", "F25": "Nama GM",
        "C26": "Jabatan Manager Asset", "F26": "Nama Manager Asset",
        "C27": "Jabatan Asman Asset", "F27": "Nama Asman Asset", "J10": "Jangka Waktu Kerja"
    }
    
    # Apply values to Excel cells
    for cell, key in mappings.items():
        sheet[cell] = data.get(key, "")
    
    # Apply bold and strikethrough formatting (crossed-out text)
    # Uncomment the following lines if you want to apply formatting
    # for cell in ["F5", "F6", "F7"]:  # Modify the cell list as needed
    #     sheet[cell].font = Font(bold=True, strike=True)  # Bold & crossed-out text

    # Insert date-related sentences
    sheet["F31"] = f"Pada hari ini {day} Tanggal {date} Bulan {month} Tahun {year}, telah "
    sheet["F32"] = f"Pada hari ini {day} Tanggal {date} Bulan {month} Tahun {year}, kami yang bertanda tangan"
    
    wb.save(output_path)
    print(f"Saved filled template to {output_path}")

def excel_to_pdf_batch(conversion_list):
    """ Converts Excel files to PDFs using Excel automation. """
    print("Converting Excel files to PDF...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        for xl_path, pdf_path in conversion_list:
            print(f"Processing {xl_path} -> {pdf_path}")
            wb = excel.Workbooks.Open(os.path.abspath(xl_path))
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            wb.Close(False)
    finally:
        excel.Quit()
    print("PDF conversion completed.")

def process_po_template(po_doc, project_id, output_path):
    """ Processes PO template and highlights project ID occurrences. """
    print(f"Processing PO template for {project_id}...")
    temp_doc = fitz.open()
    temp_doc.insert_pdf(po_doc)
    for page in temp_doc:
        text_instances = page.search_for(project_id)
        for inst in text_instances:
            rect = fitz.Rect(inst.x0 - 30, inst.y0, inst.x1 + 250, inst.y1)
            highlight = page.add_highlight_annot(rect)
            highlight.update()
            line_height = inst.y1 - inst.y0
            for _ in range(2):
                rect.y0 += line_height
                rect.y1 += line_height
                highlight_below = page.add_highlight_annot(rect)
                highlight_below.update()
    temp_doc.save(output_path)
    temp_doc.close()
    print(f"Saved highlighted PO to {output_path}")

def process_documents(use_po=True):
    """ Main process to generate Excel & PDF files, optionally using PO. """
    BASE_DIR = "sources_file"
    OUTPUT_DIR = "output_file"
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    required_files = {
        "source": os.path.join(BASE_DIR, "source.xlsx"),
        "template": os.path.join(BASE_DIR, "template.xlsx"),
        "po": os.path.join(BASE_DIR, "PO.pdf")
    }
    
    # Ensure necessary files exist
    for name, path in required_files.items():
        if name != "po" or use_po:  # Skip checking PO if not used
            if not os.path.exists(path):
                raise FileNotFoundError(f"Missing {name} file: {path}")

    df = read_source_data(required_files["source"])
    po_doc = fitz.open(required_files["po"]) if use_po else None
    processing_queue = []
    
    for idx, row in df.iterrows():
        data = row.to_dict()
        base_name = f"BAST_{data['SOW']}_{data['Project ID']}_{data['Site Name Tenant']}"
        record = {
            "xl": os.path.join(OUTPUT_DIR, f"{base_name}.xlsx"),
            "pdf": os.path.join(OUTPUT_DIR, f"{base_name}.pdf"),
            "temp_pdf": os.path.join(OUTPUT_DIR, f"{base_name}_TEMP.pdf"),
            "project_id": data["Project ID"]
        }
        fill_excel_template(required_files["template"], record["xl"], data)
        processing_queue.append(record)
    
    # Convert Excel to PDF
    excel_to_pdf_batch([(r["xl"], r["temp_pdf"]) for r in processing_queue])
    
    for record in processing_queue:
        if use_po:
            po_output = os.path.join(OUTPUT_DIR, f"PO_Highlight_{record['project_id']}.pdf")
            process_po_template(po_doc, record["project_id"], po_output)
            
            # Merge PO document with generated PDF
            merger = PdfMerger()
            merger.append(record["temp_pdf"])
            merger.append(po_output)
            merger.write(record["pdf"])
            merger.close()

            # Cleanup temporary PO file
            if os.path.exists(po_output):
                os.remove(po_output)
        else:
            os.rename(record["temp_pdf"], record["pdf"])

        # Remove temporary PDFs
        if os.path.exists(record["temp_pdf"]):
            os.remove(record["temp_pdf"])
        
        print(f"Created: {record['pdf']}")
    
    if use_po and po_doc:
        po_doc.close()

    print("\nProcessing complete.")

def main():
    """ Entry point to start document processing. """
    print("\n===== BAST Document Generator =====")
    choice = input("Do you want to use the PO file? (yes/no): ").strip().lower()
    use_po = choice == "yes"
    
    process_documents(use_po=use_po)

    # Prevent terminal from closing immediately
    input("\nProcess completed. Press Enter to exit.")

if __name__ == "__main__":
    main()

import os
import pandas as pd
import win32com.client
import fitz
from openpyxl import load_workbook
from PyPDF2 import PdfMerger
from utils import spell_date

def read_source_data(source_file):
    print(f"Reading source data from {source_file}...")
    return pd.read_excel(source_file)

def fill_excel_template(template_path, output_path, data):
    print(f"Filling Excel template for Project ID: {data['Project ID']}...")
    wb = load_workbook(template_path)
    sheet = wb["MASTER"]
    
    dt = data["Tanggal PO"]
    day = spell_date(dt, "day")
    date = spell_date(dt, "date")
    month = spell_date(dt, "month")
    year = spell_date(dt, "year")
    
    mappings = {
        "F5": "Project ID", "F6": "Site Name PO", "F7": "Site Name Tenant",
        "F8": "Site ID", "F9": "PKS No", "H9": "Tanggal PKS", "F10": "PO No",
        "H10": "Tanggal PO", "F11": "Tipe Tower", "G11": "Height", "F12": "Alamat",
        "G13": "Long", "G14": "Lat", "F15": "SOW", "F16": "Regional", "F17": "Mitra",
        "F18": "Nilai Kontrak", "C25": "Jabatan GM", "F25": "Nama GM",
        "C26": "Jabatan Manager Asset", "F26": "Nama Manager Asset",
        "C27": "Jabatan Asman Asset", "F27": "Nama Asman Asset", "J10": "Jangka Waktu Kerja"
    }
    
    for cell, key in mappings.items():
        sheet[cell] = data.get(key, "")
    
    sheet["F31"] = f"Pada hari ini {day} Tanggal {date} Bulan {month} Tahun {year}, telah "
    sheet["F32"] = f"Pada hari ini {day} Tanggal {date} Bulan {month} Tahun {year}, kami yang bertanda tangan"
    
    wb.save(output_path)
    print(f"Excel file saved: {output_path}")

def excel_to_pdf_batch(conversion_list):
    print("Converting Excel files to PDFs...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        for xl_path, pdf_path in conversion_list:
            print(f"Converting {xl_path} to {pdf_path}...")
            wb = excel.Workbooks.Open(os.path.abspath(xl_path))
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            wb.Close(False)
    finally:
        excel.Quit()
    print("Excel to PDF conversion completed.")

def process_po_template(po_doc, project_id, output_path):
    print(f"Processing PO template for Project ID: {project_id}...")
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
    print(f"PO template processed and saved: {output_path}")

def process_documents(use_po=True):
    print("\nStarting document processing...\n")
    
    BASE_DIR = "sources_file"
    OUTPUT_DIR = "output_file"
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    required_files = {
        "source": os.path.join(BASE_DIR, "source.xlsx"),
        "template": os.path.join(BASE_DIR, "template.xlsx"),
        "po": os.path.join(BASE_DIR, "PO.pdf")
    }

    for name, path in required_files.items():
        if name != "po" or use_po:
            if not os.path.exists(path):
                raise FileNotFoundError(f"Error: Missing {name} file: {path}")
            print(f"Found {name} file: {path}")

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
        print(f"Preparing document for: {base_name}")
        fill_excel_template(required_files["template"], record["xl"], data)
        processing_queue.append(record)
    
    excel_to_pdf_batch([(r["xl"], r["temp_pdf"]) for r in processing_queue])
    
    for record in processing_queue:
        if use_po:
            po_output = os.path.join(OUTPUT_DIR, f"PO_Highlight_{record['project_id']}.pdf")
            process_po_template(po_doc, record["project_id"], po_output)
            merger = PdfMerger()
            merger.append(record["temp_pdf"])
            merger.append(po_output)
            merger.write(record["pdf"])
            merger.close()
            if os.path.exists(po_output):
                os.remove(po_output)
        else:
            os.rename(record["temp_pdf"], record["pdf"])

        if os.path.exists(record["temp_pdf"]):
            os.remove(record["temp_pdf"])

        print(f"Successfully created: {record['pdf']}")
    
    if use_po and po_doc:
        po_doc.close()

    print("\nDocument processing completed successfully!")

def main():
    print("Welcome to the BAST Document Generator!\n")
    choice = input("Do you want to use the PO file? (yes/no): ").strip().lower()
    use_po = choice == "yes"
    process_documents(use_po=use_po)
    
    input("\nPress Enter to exit...")  # Keeps the terminal open after execution

if __name__ == "__main__":
    main()

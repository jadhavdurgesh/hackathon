
import fitz 
import pandas as pd
import os
import re

def clean_text(text):
    """Remove illegal characters that cause issues in Excel."""
    if not isinstance(text, str):
        return text
    return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F-\x9F]', '', text)

def detect_tables_strict(page):
    """Detect tables with a header (5+ columns) and aligned rows."""
    blocks = page.get_text("dict")["blocks"]
    lines = []
    for block in blocks:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    x0, y0, x1, y1 = span["bbox"]
                    text = clean_text(span["text"].strip())
                    if text:
                        lines.append({"text": text, "x0": x0, "y0": y0, "x1": x1})

    if not lines:
        print("No text extracted from page.")
        return []

    rows = {}
    tolerance = 15
    for line in lines:
        y_key = round(line["y0"] / tolerance) * tolerance
        if y_key not in rows:
            rows[y_key] = []
        rows[y_key].append(line)

    sorted_rows = sorted(rows.items(), key=lambda x: x[0])
    
    table = []
    header = None
    for y_key, row_lines in sorted_rows:
        row_lines.sort(key=lambda x: x["x0"])
        row_text = [line["text"] for line in row_lines]
        
        if not header and len(row_text) >= 5: 
            header = row_text
            table.append(header)
        elif header and len(row_text) >= 1:
            if len(row_text) < len(header):
                row_text.extend([''] * (len(header) - len(row_text)))
            elif len(row_text) > len(header):
                row_text = row_text[:len(header)]
            table.append(row_text)
    
    if len(table) > 1:
        return [table]
    return []

def detect_tables_flexible(page):
    """Detect tables based on consistent column counts."""
    blocks = page.get_text("dict")["blocks"]
    lines = []
    for block in blocks:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    x0, y0, x1, y1 = span["bbox"]
                    text = clean_text(span["text"].strip())
                    if text:
                        lines.append({"text": text, "x0": x0, "y0": y0})

    if not lines:
        print("No text extracted from page.")
        return []

    rows = {}
    for line in lines:
        y_key = round(line["y0"], -1) 
        if y_key not in rows:
            rows[y_key] = []
        rows[y_key].append(line)

    sorted_rows = sorted(rows.items(), key=lambda x: x[0])
    
    tables = []
    current_table = []
    for y_key, row_lines in sorted_rows:
        row_lines.sort(key=lambda x: x["x0"])
        row_text = [line["text"] for line in row_lines]
        
        if not current_table or len(row_text) == len(current_table[-1]): 
            current_table.append(row_text)
        else:
            if len(current_table) > 1: 
                tables.append(current_table)
            current_table = [row_text]
    
    if len(current_table) > 1:
        tables.append(current_table)
    
    return tables

def detect_tables(page):
    """Try strict detection first, fall back to flexible if no tables found."""
    tables = detect_tables_strict(page)
    if tables:
        print("Found tables using strict method.")
        return tables
    tables = detect_tables_flexible(page)
    if tables:
        print("Found tables using flexible method.")
    return tables

def process_pdf(pdf_path):
    """Process the PDF and extract all tables."""
    doc = fitz.open(pdf_path)
    all_tables = []
    
    for page_num in range(len(doc)):
        print(f"Processing page {page_num + 1} of {len(doc)}...")
        page = doc[page_num]
        tables = detect_tables(page)
        if tables:
            print(f"Found {len(tables)} table(s) on page {page_num + 1}")
            all_tables.extend(tables)
        else:
            print(f"No tables detected on page {page_num + 1}")
    
    doc.close()
    return all_tables

def save_to_excel(tables, output_path):
    """Save extracted tables to an Excel file."""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            cleaned_table = [[clean_text(cell) for cell in row] for row in table]
            df = pd.DataFrame(cleaned_table[1:], columns=cleaned_table[0]) if len(cleaned_table) > 1 else pd.DataFrame(cleaned_table)
            df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False, header=len(cleaned_table) > 1)

def main(pdf_path):
    """Main function to run the extraction."""
    if not os.path.exists(pdf_path):
        print(f"Error: {pdf_path} not found.")
        return
    
    print(f"Processing {pdf_path}...")
    tables = process_pdf(pdf_path)
    if not tables:
        print("No tables detected in the entire PDF.")
        return
    
    output_path = pdf_path.replace(".pdf", "_tables.xlsx")
    save_to_excel(tables, output_path)
    print(f"Tables saved to {output_path}. Total tables extracted: {len(tables)}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python pdf_table_extractor.py <pdf_file>")
    else:
        main(sys.argv[1])
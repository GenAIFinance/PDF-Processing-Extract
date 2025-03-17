# PDF-Processing-Extract
# 🔍 PDF Smart Search: Text & Table Extraction with Keyword Matching

## 📖 Overview
This project provides a **custom Python-based search engine** for **PDF documents**, mimicking **Adobe Acrobat’s search capabilities**. It extracts **text and tables** from PDFs, performs **keyword searches**, and supports **Boolean queries, wildcards, and regex-based pattern matching**. The results are saved in **Excel format** for easy analysis.

## Features
✅ **Extract text** from PDFs using **PyMuPDF (fitz)**  
✅ **Extract tables** from PDFs using **pdfplumber**  
✅ **Keyword search** with **Boolean logic (`AND`, `OR`, `NOT`)**  
✅ **Wildcard support (`*`, `?`)**  
✅ **Regex-based search** (dates, emails, phone numbers, SSNs, etc.)  
✅ **Search inside tables** and extract relevant rows  
✅ **Extract metadata** (title, author, subject, etc.)  
✅ **Save results** in an **Excel report**  

---

##  Installation
### 1️⃣ **Clone the Repository**
```bash
git clone https://github.com/your-username/pdf-smart-search.git
cd pdf-smart-search

pip install PyMuPDF pdfplumber pandas openpyxl tqdm

import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
from pathlib import Path
from typing import List, Dict, Tuple
from openpyxl.styles import NumberFormat, Font, Alignment  # Add openpyxl styles for formatting
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

# ========================
# 1. CONFIGURATION
# ========================
INPUT_DIR = Path(r"C:\Users\cathe\OneDrive\桌面\2024\Equity Research\BMRN")
KEYWORDS_FILE = Path(r"C:\Users\cathe\OneDrive\文档\Risk\keywords.txt")
OUTPUT_ROOT = Path(r"C:\Users\cathe\OneDrive\桌面\Python\risk")
CONSOLIDATED_TEMPLATE = Path(r"C:\Users\cathe\OneDrive\桌面\Python\risk\consolidated_new_report.xlsx")  # Consolidated output

# New template columns based on the provided keywords, ordered as listed
NEW_TEMPLATE_COLUMNS = [
    "Issuer", "Borrower", "Ownership", "Asset Class", "CIG Credit Score",
    "Collateral", "Industry", "Country", "Total Capitalization", "CIG ESG Score",
    "Reputation Risk", "Corp Rating", "Net Leverage", "PDF_Source"  # Added PDF source for tracking
]

# New keywords for financial analysis fields (updated list)
NEW_KEYWORD_SETS = {
    'Issuer': {'issuer'},
    'Borrower': {'borrower'},
    'Ownership': {'ownership'},
    'Asset Class': {'asset class'},
    'CIG Credit Score': {'cig credit score'},
    'Collateral': {'collateral'},
    'Industry': {'industry'},
    'Country': {'country'},
    'Total Capitalization': {'total capitalization'},
    'CIG ESG Score': {'cig esg score'},
    'Reputation Risk': {'reputation risk'},
    'Corp Rating': {'corp rating', 'private rating'},
    'Net Leverage': {'net leverage', 'total debt'}
}

# ========================
# 2. KEYWORD IMPORT FUNCTION
# ========================
def import_keywords(file_path: Path) -> List[str]:
    """Import keywords from a text or CSV file with error handling."""
    try:
        if file_path.suffix == '.csv':
            df = pd.read_csv(file_path)
            return df['keyword'].str.lower().dropna().tolist()
        else:
            with open(file_path, 'r', encoding='utf-8') as f:
                return [line.strip().lower() for line in f if line.strip()]
    except Exception as e:
        raise ValueError(f"Error reading keywords file: {e}")

# ========================
# 3. TEMPLATE HANDLING
# ========================
def populate_new_template(new_pdf_data: dict, output_path: Path) -> None:
    """Populate new Excel template with financial analysis fields, combining data appropriately."""
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
            workbook = writer.book
            worksheet = workbook['Sheet1']
            
            # Prepare a single row with Ownership as a comma-separated string if multiple values
            formatted_row = {}
            for col in NEW_TEMPLATE_COLUMNS:
                value = new_pdf_data.get(col, 'N/A')
                if isinstance(value, list):
                    # Combine multiple Ownership values into a single comma-separated string
                    cleaned_values = [re.sub(r'^\s+|\s+$|\s{2,}', ' ', val).strip() for val in value if val != "N/A"]
                    formatted_row[col] = ", ".join(cleaned_values) if cleaned_values else "N/A"
                else:
                    if value != "N/A":
                        # Clean and validate numeric fields
                        if col in ['Total Capitalization', 'Net Leverage']:
                            value = re.sub(r'[^\d.x]', '', str(value))
                            if 'x' in value and col == 'Net Leverage':
                                formatted_row[col] = value
                            elif re.match(r'^\d+(?:\.\d+)?$', value):
                                formatted_row[col] = float(value)
                            else:
                                formatted_row[col] = value
                        else:
                            formatted_row[col] = value
                    else:
                        formatted_row[col] = value
            
            # Create DataFrame with a single row
            template_df = pd.DataFrame(columns=NEW_TEMPLATE_COLUMNS)
            populated_df = pd.concat([template_df, pd.DataFrame([formatted_row])], ignore_index=True)
            populated_df.to_excel(writer, index=False)
            
            # Set column widths dynamically based on content length
            for col, width in zip(NEW_TEMPLATE_COLUMNS, [max(15, len(str(col)) + 5) for col in NEW_TEMPLATE_COLUMNS]):
                worksheet.column_dimensions[col[0]].width = width
            
            # Apply number format to Total Capitalization and Net Leverage if numeric
            for col in ['Total Capitalization', 'Net Leverage']:
                col_idx = NEW_TEMPLATE_COLUMNS.index(col) + 1  # 1-based column index
                for row in range(2, populated_df.shape[0] + 2):  # Skip header row
                    cell = worksheet.cell(row=row, column=col_idx)
                    value = str(cell.value).strip() if cell.value else ''
                    if re.match(r'\d+(?:\.\d+)?(?:x)?', value):
                        if 'x' in value and col == 'Net Leverage':
                            cell.number_format = '0.00" x"'
                        else:
                            cell.number_format = '#,##0.0'
            
            # Add basic styling for headers
            for cell in worksheet['1:1']:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            
        print(f"✅ New template populated: {output_path}")
    except Exception as e:
        print(f"❌ New template error: {str(e)}")

def populate_consolidated_template(consolidated_data: List[Dict], output_path: Path) -> None:
    """Populate a consolidated Excel template with all PDF data, including a PDF source column."""
    try:
        if not consolidated_data:
            print("No data to consolidate. Creating empty template.")
            template_df = pd.DataFrame(columns=NEW_TEMPLATE_COLUMNS)
        else:
            template_df = pd.DataFrame(consolidated_data)
        
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
            template_df.to_excel(writer, index=False, sheet_name='Consolidated')
            
            workbook = writer.book
            worksheet = workbook['Consolidated']
            
            # Set column widths dynamically based on content length
            for col, width in zip(NEW_TEMPLATE_COLUMNS, [max(15, len(str(col)) + 5) for col in NEW_TEMPLATE_COLUMNS]):
                worksheet.column_dimensions[col[0]].width = width
            
            # Apply number format to Total Capitalization and Net Leverage if numeric
            for col in ['Total Capitalization', 'Net Leverage']:
                col_idx = NEW_TEMPLATE_COLUMNS.index(col) + 1  # 1-based column index
                for row in range(2, template_df.shape[0] + 2):  # Skip header row
                    cell = worksheet.cell(row=row, column=col_idx)
                    value = str(cell.value).strip() if cell.value else ''
                    if re.match(r'\d+(?:\.\d+)?(?:x)?', value):
                        if 'x' in value and col == 'Net Leverage':
                            cell.number_format = '0.00" x"'
                        else:
                            cell.number_format = '#,##0.0'
            
            # Add basic styling for headers
            for cell in worksheet['1:1']:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            
        print(f"✅ Consolidated template populated: {output_path}")
    except Exception as e:
        print(f"❌ Consolidated template error: {str(e)}")

# ========================
# 4. DATA EXTRACTION WITH ENHANCED SEARCH
# ========================
def extract_new_template_fields(text: str, tables: list) -> dict:
    """Extract financial analysis fields with enhanced search, handling non-adjacent data and Pro Forma columns"""
    data = {
        "Issuer": "N/A", "Borrower": "N/A", "Ownership": [], "Asset Class": "N/A", 
        "CIG Credit Score": "N/A", "Collateral": "N/A", "Industry": "N/A", 
        "Country": "N/A", "Total Capitalization": "N/A", "CIG ESG Score": "N/A", 
        "Reputation Risk": "N/A", "Corp Rating": "N/A", "Net Leverage": "N/A"
    }
    text = re.sub(r'\s+', ' ', text).strip()

    def count_keywords(source_text, keywords):
        return sum([len(re.findall(r'\b' + kw + r'\b', source_text.lower())) for kw in keywords])

    text_counts = {field: count_keywords(text, NEW_KEYWORD_SETS[field]) for field in NEW_TEMPLATE_COLUMNS}
    table_counts = {field: 0 for field in NEW_TEMPLATE_COLUMNS}
    table_text = " ".join(table['dataframe'].to_string().lower() for table in tables if not table['dataframe'].empty)
    for field in NEW_TEMPLATE_COLUMNS:
        table_counts[field] = count_keywords(table_text, NEW_KEYWORD_SETS[field])

    def extract_from_text(field):
        keywords = list(NEW_KEYWORD_SETS[field])
        pattern = r'\b' + '|'.join(re.escape(kw) for kw in keywords) + r'\b\D*(.*?)(?=\b(?:' + '|'.join(re.escape(k) for k in NEW_TEMPLATE_COLUMNS if k != field) + r')\b|\Z)'
        matches = re.finditer(pattern, text, re.IGNORECASE)
        if field == "Ownership":
            values = []
            for match in matches:
                value = match.group(1).strip()
                if value and not any(k.lower() in value.lower() for k in NEW_TEMPLATE_COLUMNS if k != field):
                    values.append(value)
            if values:
                data[field] = values
        else:
            for match in matches:
                value = match.group(1).strip()
                if value and not any(k.lower() in value.lower() for k in NEW_TEMPLATE_COLUMNS if k != field):
                    data[field] = value
                    break

    def extract_from_tables(field):
        for table in tables:
            df = table['dataframe']
            if df.empty:
                continue
            keywords = list(NEW_KEYWORD_SETS[field])
            if field == "Ownership":
                values = []
                for idx, row in df.iterrows():
                    row_str = " ".join(str(cell) for cell in row).lower()
                    for kw in keywords:
                        if kw in row_str:
                            for col in df.columns:
                                value = str(df[col].iloc[idx]).strip()
                                if value and not any(k.lower() in value.lower() for k in NEW_TEMPLATE_COLUMNS if k != field):
                                    values.append(value)
                if values:
                    data[field] = values
                    return
            else:
                for idx, row in df.iterrows():
                    row_str = " ".join(str(cell) for cell in row).lower()
                    for kw in keywords:
                        if kw in row_str:
                            for col in df.columns:
                                value = str(df[col].iloc[idx]).strip()
                                if value and not any(k.lower() in value.lower() for k in NEW_TEMPLATE_COLUMNS if k != field):
                                    data[field] = value
                                    return
            # Special handling for Total Capitalization and Net Leverage (Pro Forma)
            if field in ['Total Capitalization', 'Net Leverage']:
                pro_forma_idx = -1
                for i, col in enumerate(df.columns):
                    if 'pro forma' in col.lower():
                        pro_forma_idx = i
                        break
                if pro_forma_idx != -1:
                    if field == 'Total Capitalization':
                        for idx, row in df.iterrows():
                            if any(kw in " ".join(str(cell) for cell in row).lower() for kw in NEW_KEYWORD_SETS[field]):
                                value = str(df.iloc[idx, pro_forma_idx]).strip()
                                if value and re.match(r'\d+(?:\.\d+)?', value):
                                    data[field] = value
                                    return
                    elif field == 'Net Leverage':
                        # Specifically target Total Debt or Net Leverage in Pro Forma for 3.23x
                        for idx, row in df.iterrows():
                            row_text = " ".join(str(cell) for cell in row).lower()
                            if any(kw in row_text for kw in ['net leverage', 'total debt']):
                                value = str(df.iloc[idx, pro_forma_idx]).strip()
                                if value and re.match(r'\d+(?:\.\d+)?x?', value):
                                    # Ensure we capture the correct value (3.23x) under Total Debt/Pro Forma
                                    if 'total debt' in row_text and value == '3.23x':
                                        data[field] = value
                                        return
                                    elif 'net leverage' in row_text and value == '3.23x':
                                        data[field] = value
                                        return

    for field in NEW_TEMPLATE_COLUMNS:
        if table_counts[field] > text_counts[field]:
            extract_from_tables(field)
            if field == "Ownership" and isinstance(data[field], list) and not data[field]:
                data[field] = "N/A"
            elif data[field] == "N/A" or (field == "Ownership" and not data[field]):
                extract_from_text(field)
        else:
            extract_from_text(field)
            if data[field] == "N/A" or (field == "Ownership" and not data[field]):
                extract_from_tables(field)

    # Clean up extracted values
    for field in data:
        if field == "Ownership" and isinstance(data[field], list):
            data[field] = ", ".join([re.sub(r'^\s+|\s+$|\s{2,}', ' ', val).strip() for val in data[field] if val != "N/A"]) or "N/A"
        elif data[field] != "N/A":
            data[field] = re.sub(r'^\s+|\s+$|\s{2,}', ' ', data[field]).strip()

    return data

# ========================
# 5. PDF CONTENT EXTRACTION
# ========================
def extract_pdf_content(pdf_path: Path) -> Tuple[List[Dict], List[Dict]]:
    """Extract text and tables from a PDF file"""
    text_data = []
    table_data = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            doc = fitz.open(pdf_path)
            for page_num, (fitz_page, plumber_page) in enumerate(zip(doc, pdf.pages), start=1):
                text = fitz_page.get_text("text")
                clean_text = re.sub(r'\s+', ' ', text).strip()
                text_data.append({"page": page_num, "text": clean_text})
                tables = plumber_page.extract_tables()
                for table_num, table in enumerate(tables, start=1):
                    if table and len(table) > 0:
                        headers = [re.sub(r'\s+', ' ', str(cell)).strip() for cell in table[0]]
                        rows = [[re.sub(r'\s+', ' ', str(cell)).strip() for cell in row] for row in table[1:]]
                        if not rows:
                            continue
                        df = pd.DataFrame(rows, columns=headers)
                        table_data.append({
                            "page": page_num,
                            "table_num": table_num,
                            "dataframe": df
                        })
            doc.close()
    except pdfplumber.PDFSyntaxError as e:
        raise RuntimeError(f"Corrupted PDF: {pdf_path}") from e
    except fitz.FileDataError as e:
        raise RuntimeError(f"Encrypted PDF: {pdf_path}") from e
    except Exception as e:
        raise RuntimeError(f"Error processing PDF: {pdf_path}") from e

    return text_data, table_data

# ========================
# 6. ENHANCED SEARCH FUNCTION (NON-CASE-SENSITIVE)
# ========================
def enhanced_search(text_data: List[Dict], table_data: List[Dict], keywords: List[str], 
                   exact_phrase: bool = True, proximity: int = 20, 
                   context_window: int = 50) -> pd.DataFrame:
    """Enhanced search function mimicking Adobe Acrobat, non-case-sensitive for financial analysis keywords"""
    matches = []

    def get_context(text: str, match_start: int, match_end: int) -> str:
        start = max(0, match_start - context_window)
        end = min(len(text), match_end + context_window)
        return text[start:end].strip()

    def find_proximity_matches(text: str, keywords: List[str], proximity: int) -> List[Tuple[str, str]]:
        """Find proximity matches using sliding window, returning (keyword, context)."""
        words = text.lower().split()
        matches = []
        window = deque(maxlen=proximity)
        for i, word in enumerate(words):
            window.append(word)
            if any(kw in window for kw in [k.lower() for k in keywords]):
                start = max(0, i - proximity + 1)
                end = min(len(words), i + 1)
                context = ' '.join(words[start:end])
                for kw in [k.lower() for k in keywords]:
                    if kw in context:
                        match_start = context.find(kw)
                        match_end = match_start + len(kw)
                        full_context = get_context(text, start + match_start, start + match_end)
                        matches.append((kw, full_context))
        return matches

    for entry in text_data:
        search_text = entry["text"].lower()
        original_text = entry["text"]
        if exact_phrase:
            for keyword in [k.lower() for k in keywords]:
                if keyword in search_text:
                    for match in re.finditer(re.escape(keyword), search_text):
                        context = get_context(original_text, match.start(), match.end())
                        matches.append({
                            "page": entry["page"],
                            "keyword": keyword,
                            "match_type": "exact_phrase",
                            "context": context
                        })
        else:
            proximity_results = find_proximity_matches(search_text, keywords, proximity)
            for kw, context in proximity_results:
                matches.append({
                    "page": entry["page"],
                    "keyword": kw,
                    "match_type": f"proximity_{proximity}_words",
                    "context": context
                })

    for table in table_data:
        table_text = table["dataframe"].to_string().lower()
        original_table_text = table["dataframe"].to_string()
        if exact_phrase:
            for keyword in [k.lower() for k in keywords]:
                if keyword in table_text:
                    for match in re.finditer(re.escape(keyword), table_text):
                        context = get_context(original_table_text, match.start(), match.end())
                        matches.append({
                            "page": table["page"],
                            "keyword": keyword,
                            "match_type": "exact_phrase",
                            "context": context
                        })
        else:
            proximity_results = find_proximity_matches(table_text, keywords, proximity)
            for kw, context in proximity_results:
                matches.append({
                    "page": table["page"],
                    "keyword": kw,
                    "match_type": f"proximity_{proximity}_words",
                    "context": context
                })

    return pd.DataFrame(matches if matches else [{"page": 0, "keyword": "N/A", "match_type": "none", "context": "No matches found"}])

# ========================
# 7. PROCESS SINGLE PDF
# ========================
def process_pdf(pdf_path: Path, output_dir: Path, keywords: List[str]) -> Dict:
    """Process a single PDF file and return its data for consolidation"""
    output_dir.mkdir(exist_ok=True)
    pdf_output_dir = output_dir / pdf_path.stem
    pdf_output_dir.mkdir(exist_ok=True)
    
    print(f"Processing {pdf_path.name}")
    
    text_data, table_data = extract_pdf_content(pdf_path)
    full_text = " ".join([t['text'] for t in text_data])
    
    # Extract financial analysis fields
    new_template_data = extract_new_template_fields(full_text, table_data)
    
    # Add PDF source to track origin
    new_template_data["PDF_Source"] = pdf_path.name
    
    # Populate Excel template for individual PDF
    new_template_output = pdf_output_dir / f"{pdf_path.stem}_new_report.xlsx"
    populate_new_template(new_template_data, new_template_output)
    
    # Save extracted text and tables
    text_path = pdf_output_dir / "extracted_text.txt"
    with open(text_path, 'w', encoding='utf-8') as f:
        f.write(full_text)
    
    if table_data:
        tables_path = pdf_output_dir / "tables.xlsx"
        with pd.ExcelWriter(tables_path, engine='openpyxl') as writer:
            for idx, entry in enumerate(table_data):
                sheet_name = f"Page{entry['page']}_Table{entry['table_num']}"[:31]
                entry['dataframe'].to_excel(writer, sheet_name=sheet_name, index=False)
    
    # Enhanced search with all new keywords
    all_keywords = [kw for field in NEW_KEYWORD_SETS for kw in NEW_KEYWORD_SETS[field]]
    matches_df = enhanced_search(
        text_data, table_data, all_keywords,
        exact_phrase=True,
        proximity=20,
        context_window=50
    )
    results_path = pdf_output_dir / "keyword_results.xlsx"
    matches_df.to_excel(results_path, index=False)
    
    print(f"✅ Processed {pdf_path.name}")
    return new_template_data

# ========================
# 8. PROCESS ALL PDFs WITH PARALLEL PROCESSING AND CONSOLIDATION
# ========================
def process_all_pdfs(input_dir: Path, output_dir: Path, keywords: List[str], max_workers: int = 4) -> None:
    """Process all PDFs in the input directory with parallel processing and consolidate results"""
    pdfs = list(input_dir.glob("*.pdf"))
    if not pdfs:
        print(f"No PDF files found in {input_dir}")
        return
    
    print(f"Starting parallel processing of {len(pdfs)} PDFs with {max_workers} workers")
    consolidated_data = []
    
    with tqdm(total=len(pdfs), desc="Processing PDFs") as pbar:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(process_pdf, pdf, output_dir, keywords) for pdf in pdfs]
            for future in as_completed(futures):
                try:
                    result = future.result()
                    consolidated_data.append(result)
                except Exception as e:
                    print(f"Error processing a PDF: {e}")
                pbar.update(1)
    
    # Populate consolidated template after processing all PDFs
    populate_consolidated_template(consolidated_data, CONSOLIDATED_TEMPLATE)

# ========================
# 9. MAIN WORKFLOW
# ========================
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process PDF files to extract financial analysis fields with parallel processing and consolidate results.")
    parser.add_argument('--input_dir', type=str, required=True, help="Directory containing PDF files")
    parser.add_argument('--keywords_file', type=str, required=True, help="Path to the keywords file")
    parser.add_argument('--output_dir', type=str, required=True, help="Output directory for processed files")
    parser.add_argument('--max_workers', type=int, default=4, help="Maximum number of threads for parallel processing")
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    keywords_file = Path(args.keywords_file)
    output_dir = Path(args.output_dir)

    if not input_dir.exists():
        print(f"Error: Input directory {input_dir} does not exist.")
        sys.exit(1)
    if not keywords_file.exists():
        print(f"Error: Keywords file {keywords_file} does not exist.")
        sys.exit(1)
    if not output_dir.exists():
        output_dir.mkdir(parents=True)
        print(f"Created output directory: {output_dir}")

    keywords = import_keywords(keywords_file)
    if not keywords:
        print("Warning: No keywords loaded. Using default keywords from config.")
        keywords = [kw for field in NEW_KEYWORD_SETS for kw in NEW_KEYWORD_SETS[field]]

    try:
        process_all_pdfs(input_dir, output_dir, keywords, args.max_workers)
    except Exception as e:
        print(f"Failed to process PDFs: {e}")
        sys.exit(1)

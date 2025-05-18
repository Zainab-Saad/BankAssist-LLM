import pandas as pd
import re

def is_question(text):
    text = str(text).strip()
    if not text or text == 'nan':
        return False
    patterns = [
        r"\?$", 
        r"^(what|how|is|are|do|does|can|who|when|where|why|shall|will|has|have)",
        r"\b(explain|describe|tell me about)\b"
    ]
    return any(re.search(pattern, text.lower()) for pattern in patterns)

def process_row(row):
    for idx, cell in enumerate(row):
        cell_str = str(cell).strip()
        if cell_str.lower() == 'main':
            continue
        if is_question(cell_str):
            return cell_str, list(row[idx+1:])
    return None, []

def format_answer(answer_data):
    formatted = []
    i = 0
    while i < len(answer_data):
        # Clean current row
        row = [
            f"{float(cell)*100:.2f}%" if (isinstance(cell, float) and cell <= 1) else str(cell).strip()
            for cell in answer_data[i]
            if not pd.isnull(cell) and str(cell).strip() not in ['', 'nan']
        ]
        if not row:
            i += 1
            continue

        # Detect table headers with at least 2 columns
        if i + 1 < len(answer_data) and len(row) >= 2:
            # Clean next row
            next_row = [
                f"{float(cell)*100:.2f}%" if (isinstance(cell, float) and cell <= 1) else str(cell).strip()
                for cell in answer_data[i+1]
                if not pd.isnull(cell) and str(cell).strip() not in ['', 'nan']
            ]
            
            if len(next_row) == len(row):
                # Build markdown table
                headers = row
                rows = [next_row]
                i += 2
                
                # Add subsequent matching rows
                while i < len(answer_data):
                    current_row = [
                        f"{float(cell)*100:.2f}%" if (isinstance(cell, float) and cell <= 1) else str(cell).strip()
                        for cell in answer_data[i]
                        if not pd.isnull(cell) and str(cell).strip() not in ['', 'nan']
                    ]
                    if len(current_row) == len(headers):
                        rows.append(current_row)
                        i += 1
                    else:
                        break
                
                # Format table
                table = [
                    f"| {' | '.join(headers)} |",
                    f"| {' | '.join(['---']*len(headers))} |"
                ]
                table.extend([f"| {' | '.join(row)} |" for row in rows])
                formatted.append('\n'.join(table))
                continue

        # Format as list items if not a table
        formatted.extend(f"- {item}" for item in row)
        i += 1
    
    return '\n'.join(formatted)

def process_sheet(sheet_name, df, source):
    markdown = []
    current_q = None
    current_a = []

    for _, row in df.iterrows():
        q, a = process_row(row)
        if q:
            if current_q:
                answer = format_answer(current_a)
                markdown.append(create_block(sheet_name, current_q, answer, source))
            current_q = q
            current_a = [a]
        else:
            cleaned = [cell for cell in row if not pd.isnull(cell)]
            if cleaned:
                current_a.append(cleaned)
    
    if current_q:
        answer = format_answer(current_a)
        markdown.append(create_block(sheet_name, current_q, answer, source))
    
    return '\n'.join(markdown)

def create_block(sheet, q, a, src):
    return f"""---
sheet_name: "{sheet}"
question: "{q}"
source: "{src}"
---

**Answer:**  
{a}

---
"""

# Main execution
excel_file = "NUST Bank-Product-Knowledge.xlsx"
output_file = "output.md"

# Read all sheets
all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)

with open(output_file, "w", encoding="utf-8") as f:
    for sheet_name, df in all_sheets.items():
        if sheet_name in ["Sheet3", "Sheet1", "Main"]:
            continue
        
        content = process_sheet(sheet_name, df, excel_file)
        if content:
            f.write(content + "\n")

print(f"Markdown output saved to {output_file}")

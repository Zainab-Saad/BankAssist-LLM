import pandas as pd
import re
from itertools import zip_longest

def is_question(text):
    """Enhanced question detection that handles empty cells"""
    text = str(text).strip()
    if not text or text == 'nan':
        return False
    question_patterns = [
        r"\?$",  # Ends with question mark
        r"^(what|how|is|are|do|does|can|who|when|where|why|shall|will|has|have)",
        r"\b(explain|describe|tell me about)\b",
        # r":$"  # Pattern like "Main features:"
    ]
    return any(re.search(pattern, text.lower()) for pattern in question_patterns)

def process_row(row):
    """Process a row to find questions in any column"""
    question = None
    answer_columns = []
    
    # Check each cell in the row for questions
    for idx, cell in enumerate(row):
        # print(cell)
        if str(cell).strip().lower() == 'Main':
            continue
        if is_question(str(cell)):
            question = str(cell).strip()
            answer_columns = list(row[idx+1:])  # Get all subsequent columns
            break
    
    return question, answer_columns

def format_answer(answer_data):
    """Handle multi-column answers with improved formatting"""
    cleaned = []
    for item in answer_data:
        if isinstance(item, list):
            cleaned.extend([str(cell).strip() for cell in item if pd.notnull(cell) and str(cell).strip() and str(cell).strip() != 'Main'])
        else:
            if pd.notnull(item) and str(item).strip() and str(item).strip() != 'Main':
                cleaned.append(str(item).strip())
    
    # Handle tabular data
    if any('|' in cell for cell in cleaned):
        return "\n".join(cleaned)
    
    # Handle key-value pairs
    if len(cleaned) > 1 and ':' in cleaned[0]:
        return "\n".join([f"- **{k}**: {v}" for k, v in zip(cleaned[::2], cleaned[1::2])])
    
    return "\n".join([f"- {line}" for line in cleaned if line])

def process_sheet(sheet_name, df, source_filename):
    """Process any sheet with dynamic column handling"""
    markdown_blocks = []
    current_question = None
    current_answer = []
    
    for _, row in df.iterrows():
        # Convert row to list and process
        row_list = [cell for cell in row]
        question, answer_columns = process_row(row_list)
        
        if question:
            if current_question:
                # Save previous Q&A
                answer = format_answer(current_answer)
                if answer:
                    block = create_block(sheet_name, current_question, answer, source_filename)
                    markdown_blocks.append(block)
            current_question = question
            current_answer = answer_columns
        else:
            # Collect answer data from all columns
            row_data = [str(cell).strip() for cell in row_list if pd.notnull(cell) and str(cell).strip()]
            if row_data:
                current_answer.extend(row_data)
    
    # Process remaining content
    if current_question and current_answer:
        answer = format_answer(current_answer)
        if answer:
            block = create_block(sheet_name, current_question, answer, source_filename)
            markdown_blocks.append(block)
    
    return "\n".join(markdown_blocks)

def create_block(sheet_name, question, answer, source):
    """Create markdown block with proper formatting"""
    return f"""---
sheet_name: "{sheet_name}"
question: "{question}"
source: "{source}"
---

**Answer:**  
{answer}

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
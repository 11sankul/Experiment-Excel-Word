import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------------- CONFIG ----------------
EXCEL_PATH = "data/input.xlsx"
SHEET_NAME = "Sheet1"
OUTPUT_PATH = "output/output.docx"

# IMPORTANT:
# - This script now automatically picks ALL columns from Excel
# - No manual field list is required
# - Blank Excel cells remain blank in Word

# ---------------- LOAD EXCEL ----------------
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

# Replace NaN with empty string so blanks stay blanks
df = df.fillna("")

# ---------------- CREATE WORD DOCUMENT ----------------
document = Document()

for idx, row in df.iterrows():

    # CHANGE 1:
    # Instead of "Record 1", we now use:
    # School Name – <value from Excel>
    school_name = row.get("Name of School", "")
    heading_text = f"School Name – {school_name}"

    heading = document.add_heading(heading_text, level=2)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # CHANGE 2:
    # Table rows = number of columns in Excel
    table = document.add_table(rows=len(df.columns), cols=2)
    table.style = "Table Grid"

    # Loop through ALL Excel columns automatically
    for i, column_name in enumerate(df.columns):
        table.cell(i, 0).text = str(column_name)   # Column name
        table.cell(i, 1).text = str(row[column_name])  # Cell value (blank if blank)

    # Page break after each school except last
    if idx < len(df) - 1:
        document.add_page_break()

# ---------------- SAVE OUTPUT ----------------
document.save(OUTPUT_PATH)

print("Word file generated successfully.")

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

EXCEL_PATH = "data/input.xlsx"
SHEET_NAME = "Sheet1"
OUTPUT_PATH = "output/output.docx"

FIELDS = [
    ("Sr. No", "Sr. no"),
    ("Name of School", "Name of School"),
]

df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
df = df.fillna("")

document = Document()

for idx, row in df.iterrows():
    heading = document.add_heading(f"Record {idx + 1}", level=2)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table = document.add_table(rows=len(FIELDS), cols=2)
    table.style = "Table Grid"

    for i, (label, column) in enumerate(FIELDS):
        table.cell(i, 0).text = label
        table.cell(i, 1).text = str(row[column]) if column in df.columns else ""

    if idx < len(df) - 1:
        document.add_page_break()

document.save(OUTPUT_PATH)
print("Word file generated successfully.")

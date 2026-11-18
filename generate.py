from pathlib import Path
import re

import pandas as pd
from docx import Document
from docx2pdf import convert


TEMPLATE = Path("template.docx")
DATA_FILE = Path("data.xlsx")
OUTPUT = Path("output")
OUTPUT.mkdir(exist_ok=True)

def replace_placeholders(paragraph, values):

    for placeholder, replacement in values.items():
        token = '{{' + placeholder + '}}'

        while token in paragraph.text:
            full_text = ''.join(run.text for run in paragraph.runs)
            start = full_text.find(token)

            if start == -1:
                break

            end = start + len(token)
            position = 0

            for run in paragraph.runs:
                run_start = position
                run_end = position + len(run.text)

                if run_end <= start or run_start >= end:
                    continue

                left = max(0, start - run_start)
                right = min(len(run.text), end - run_start)

                if run_start <= start < run_end:
                    run.text = run.text[:left] + str(replacement) + run.text[right:]
                else:
                    run.text = run.text[:left] + run.text[right:]


def fill_document(document, values):
    
    for paragraph in document.paragraphs:
        replace_placeholders(paragraph, values)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_placeholders(paragraph, values)


def fill_equipment_table(document, equipment):
    
    for table in document.tables:
        for row in table.rows:
            cell_text = [cell.text for cell in row.cells]

            if (
                len(cell_text) == 3
                and '{{item}}' in cell_text[0]
                and '{{model}}' in cell_text[1]
                and '{{sn}}' in cell_text[2]
            ):
                for _, equipment_row in equipment.iterrows():
                    new_row = table.add_row()

                    for i, field in enumerate(['item', 'model', 'sn']):
                        new_row.cells[i].text = str(equipment_row[field])

                table._tbl.remove(row._tr)
                return

    raise ValueError('Equipment placeholder row not found')


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*]', '_', str(value))

data = pd.read_excel(DATA_FILE, dtype=str).fillna('')
data['date'] = pd.to_datetime(data["date"], format="%Y-%m-%d %H:%M:%S").dt.strftime('%d-%m-%Y')

required = {
    'id', 'date', 'name', 'department', 'email', 'site', 'item',
    'model', 'sn'
}
missing = required - set(data.columns)
if missing:
    raise ValueError(f'Missing Excel columns: {missing}')

for document_id, group in data.groupby('id', sort=False):
    first = group.iloc[0]

    values = {
        field: first[field]
        for field in ['id', 'date', 'name', 'department',
                      'email', 'site']
    }

    document = Document(TEMPLATE)

    fill_equipment_table(document, group)
    fill_document(document, values)

    filename = safe_filename('PV-IT-' + document_id)
    docx_path = OUTPUT / f'{filename}.docx'
    pdf_path = OUTPUT / f'{filename}.pdf'

    document.save(docx_path)
    convert(str(docx_path), str(pdf_path))

    if pdf_path.exists():
        docx_path.unlink()
        print(f'Created: {pdf_path}')

print("Finished!")
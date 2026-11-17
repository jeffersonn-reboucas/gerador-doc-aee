import os
import re
from docx import Document

templates_dir = "./templates"

def extract_labels_from_docx(file_path):
    doc = Document(file_path)
    labels = []
    for para in doc.paragraphs:
        text = para.text.strip()
        # Find placeholders "( X )" or "(   )"
        for match in re.finditer(r"\( *\)", text):
            placeholder = match.group()
            # label is text before placeholder
            start = match.start()
            label_text = text[:start].strip()
            labels.append({
                'type': 'checkbox',
                'label': label_text,
                'full_text': text,
                'placeholder': placeholder
            })
    # Check tables
    for tbl in doc.tables:
        for row_idx, row in enumerate(tbl.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                for match in re.finditer(r"\( *\)", cell_text):
                    placeholder = match.group()
                    start = cell_text.find(placeholder)
                    if label_text is None:
                        label_text = cell_text[:label_start].strip()
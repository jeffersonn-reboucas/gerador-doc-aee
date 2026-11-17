import os
import re
import json
from docx import Document

# Mapping of original file names to (profile, doc_type)
FILE_MAP = {
    "ESTUDO DE CASO - CODA.docx": ("coda", "estudo_caso"),
    "ESTUDO DE CASO - Perda Auditiva.docx": ("perda_auditiva", "estudo_caso"),
    "ESTUDO DE CASO - Surdez.docx": ("surdez", "estudo_caso"),
    "PAEE - CODA.docx": ("coda", "paee"),
    "PAEE - Perda Auditiva.docx": ("perda_auditiva", "paee"),
    "PAEE - Surdez.docx": ("surdez", "paee"),
    "PEI - CODA.docx": ("coda", "pei"),
    "PEI - Perda Auditiva.docx": ("perda_auditiva", "pei"),
    "PEI - Surdez.docx": ("surdez", "pei"),
}

templates_dir = "./templates"

def extract_labels_from_docx(file_path):
    doc = Document(file_path)
    labels = []  # list of dicts
    # iterate paragraphs
    for para in doc.paragraphs:
        text = para.text.strip()
        # find all "( X )" patterns
        for match in re.finditer(r"\( *\)", text):
            placeholder = match.group()
            start = match.start()
            label_text = text[:start].strip()
            labels.append({
                'type': 'checkbox',
                'label': label_text,
                'full_text': text,
                'placeholder': placeholder
            })
    # tables
    for tbl in doc.tables:
        for row_idx, row in enumerate(tbl.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                # find placeholders in cell text
                for match in re.finditer(r"\( *\)", cell_text):
                    placeholder = match.group()
                    start = cell_text.find(placeholder)
                    label_text = cell_text[:start].strip()
                    labels.append({
                        'type': 'cell',
                        'row': row_idx,
                        'col': col_idx,
                        'text': cell_text,
                        'placeholder': placeholder,
                        'label': label_text
                    })
    return labels

def main():
    mapping = {}
    for filename, (profile, doc_type) in FILE_MAP.items():
        # Build template file name: original filename without .docx, add _template.docx
        base_name = os.path.splitext(filename)[0]
        template_filename = f"{base_name}_template.docx"
        tmpl_path = os.path.join(templates_dir, profile, template_filename)
        if not os.path.exists(tmpl_path):
            print(f"Template not found: {tmpl_path}")
            continue
        labels = extract_labels_from_docx(tmpl_path)
        for idx, label_info in enumerate(labels):
            # Create a unique key for the placeholder
            key = f"{profile}_{doc_type}_{label_info['type']}_{idx}"
            mapping[key] = label_info['label']
    # Save mapping to JSON
    with open('labels_mapping.json', 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    print("Labels mapping saved to labels_mapping.json")

if __name__ == "__main__":
    main()
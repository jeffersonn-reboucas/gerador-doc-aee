import os
from docx import Document

origins_dir = "./originais"
templates_dir = "./templates"

# Map original file names to (profile, doc_type)
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

def extract_placeholders(doc_path):
    doc = Document(doc_path)
    placeholders = []  # list of dicts: {'type': 'checkbox'|'fill', 'label': str, 'line': str}
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        # Checkboxes: pattern "( )" (space optional)
        if "( )" in text:
            # extract label before "( )"
            # find the position of "( )"
            start = text.find("( ")
            end = text.find(")", start)
            if start != -1 and end != -1:
                label = text[:start].strip()
                placeholders.append({
                    'type': 'checkbox',
                    'label': label,
                    'full_line': text
                })
        # Fill placeholders: lines containing only underscores (e.g., "____")
        if "_" in text and text.strip() != "":
            # check if line consists mainly of underscores
            if all(c == "_" or c.isspace() for c in text):
                placeholders.append({
                    'type': 'fill',
                    'label': text.strip(),
                    'full_line': text
                })
    # tables
    for tbl in doc.tables:
        for row_idx, row in enumerate(tbl.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                # check for "( )" inside cell
                if "( )" in cell_text:
                    start = cell_text.find("( ")
                    end = cell_text.find(")", start)
                    if start != -1 and end != -1:
                        label = cell_text[:start].strip()
                        placeholders.append({
                            'type': 'checkbox',
                            'label': label,
                            'full_line': cell_text,
                            'row': row_idx,
                            'col': col_idx
                        })
                # check for underscores in cell
                if "_" in cell_text and all(c == "_" or c.isspace() for c in cell_text):
                    placeholders.append({
                        'type': 'fill',
                        'label': cell_text.strip(),
                        'full_line': cell_text,
                        'row': row_idx,
                        'col': col_idx
                    })
    return placeholders

def main():
    for filename, (profile, doc_type) in FILE_MAP.items():
        src_path = os.path.join(origins_dir, filename)
        if not os.path.exists(src_path):
            print(f"File not found: {src_path}")
            continue
        placeholders = extract_placeholders(src_path)
        print(f"\n=== {filename} ({profile}/{doc_type}) ===")
        for ph in placeholders:
            print(f"[{ph['type']}] label: {ph['label']}  line: {ph['full_line']}")
            if 'row' in ph:
                print(f"   -> cell (row {ph['row']}, col {ph['col']})")
        print("-" * 60)

if __name__ == "__main__":
    main()
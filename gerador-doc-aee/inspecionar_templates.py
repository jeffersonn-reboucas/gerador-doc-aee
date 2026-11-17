import os
import re
from docx import Document

ORIGINALS_DIR = "./originais"
TEMPLATES_DIR = "./templates"

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

def extract_placeholders_and_labels(doc_path):
    doc = Document(doc_path)
    placeholders = []
    current_label = None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        # Check for checkboxes ( )
        if "( )" in text:
            # Extract label before ( )
            label = text.replace("( )", "").strip()
            # Find all occurrences
            matches = list(re.finditer(r"\( \)", text))
            for m in matches:
                placeholders.append({
                    "type": "checkbox",
                    "label": label,
                    "full_text": text
                })
        # Check for underscores
        if "_" * 3 in text:
            placeholders.append({
                "type": "fill",
                "text": text
            })

    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text and ("( )" in text or "_" * 3 in text or text == ""):
                    placeholders.append({
                        "type": "cell",
                        "row": row_idx,
                        "col": col_idx,
                        "text": text
                    })

    return placeholders

def main():
    for filename, (profile, doc_type) in FILE_MAP.items():
        src_path = os.path.join(ORIGINALS_DIR, filename)
        if not os.path.exists(src_path):
            print(f"Arquivo não encontrado: {src_path}")
            continue
        print(f"\n{'='*80}")
        print(f"ARQUIVO: {filename} -> {profile}/{doc_type}")
        print(f"{'='*80}")
        placeholders = extract_placeholders_and_labels(src_path)
        for p in placeholders:
            print(p)

if __name__ == "__main__":
    main()
import json
import os
from pathlib import Path
from docxtpl import DocxTemplate
import zipfile

# Paths
JSON_PATH = Path("./rascunhos/teste_lucas_mendes.json")
TEMPLATES_DIR = Path("./templates/surdez")
OUTPUT_DIR = Path("./rascunhos/teste_saida")
TEMPLATES_FILES = [
    "ESTUDO DE CASO - Surdez_template.docx",
    "PAEE - Surdez_template.docx",
    "PEI - Surdez_template.docx",
]

def load_data():
    with open(JSON_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def ensure_output_dir():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def render_and_save(data, template_name, output_filename):
    template_path = TEMPLATES_DIR / template_name
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")
    doc = DocxTemplate(str(template_path))
    doc.render(data)
    out_path = OUTPUT_DIR / output_filename
    doc.save(str(out_path))
    return out_path

def has_placeholders(docx_path):
    # Open the docx as a zip and read the main document XML
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')
            return '{{' in content or '}}' in content

def main():
    data = load_data()
    ensure_output_dir()

    for tmpl in TEMPLATES_FILES:
        out_filename = f"{Path(tmpl).stem}_rendered.docx"
        out_path = render_and_save(data, tmpl, out_filename)
        if has_placeholders(out_path):
            raise AssertionError(f"Placeholders still present in {out_path}")
        print(f"Rendered {out_filename} – no placeholders found.")

    print("All documents rendered and verified. Output saved to:", OUTPUT_DIR)

if __name__ == "__main__":
    main()
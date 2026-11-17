import zipfile
import re
from pathlib import Path

templates_dir = Path("./templates/surdez")
template_files = [
    "ESTUDO DE CASO - Surdez_template.docx",
    "PAEE - Surdez_template.docx",
    "PEI - Surdez_template.docx",
]

placeholder_counts = {}

def extract_placeholders_from_docx(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        # Find the main document part
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')
            # Find all {{ ... }} patterns
            matches = re.findall(r'\{\{[\s\w\/\.\-\']+\}\}', content)
            # Clean up the matches (remove surrounding spaces)
            cleaned = [m.strip() for m in matches]
            return cleaned

for tmpl in template_files:
    path = templates_dir / tmpl
    placeholders = extract_placeholders_from_docx(path)
    placeholder_counts[tmpl] = len(placeholders)
    print(f"{tmpl}: {len(placeholders)} placeholders")
    # Optionally print first few
    # print(placeholders[:5])
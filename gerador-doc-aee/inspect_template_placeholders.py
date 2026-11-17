import zipfile
import re
from pathlib import Path

def extract_placeholders(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')
    # Find all {{ ... }} patterns
    matches = re.findall(r'\{\{[\s\w\/\.\-\']+\}\}', content)
    return matches

template_files = [
    "ESTUDO DE CASO - Surdez_template.docx",
    "PAEE - Surdez_template.docx",
    "PEI - Surdez_template.docx",
]

for fname in template_files:
    path = Path("./templates/surdez") / fname
    placeholders = extract_placeholders(path)
    print(f"{fname}: {len(placeholders)} placeholders")
    # Show first few
    for i, ph in enumerate(placeholders[:10]):
        print(f"  {i+1}. {ph}")
    if len(placeholders) > 10:
        print("  ... (more)")
    print()
import zipfile
from pathlib import Path
import re

def extract_text_from_docx(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')
    # Extract text from XML (simplified)
    # We'll look for text between <w:t> tags
    text_chunks = re.findall(r'<w:t>(.*?)</w:t>', content)
    # Clean up the text (remove extra spaces, newlines)
    text = ' '.join(chunk.strip() for chunk in text_chunks)
    return text

rendered_files = [
    "ESTUDO DE CASO - Surdez_template_rendered.docx",
    "PAEE - Surdez_template_rendered.docx",
    "PEI - Surdez_template_rendered.docx"
]

for fname in rendered_files:
    path = Path("./rascunhos/teste_saida") / fname
    text = extract_text_from_docx(path)
    # Look for patterns that indicate checkboxes
    x_pattern = r'\( *X *\)'
    paren_pattern = r'\( *\)'
    x_match = re.search(x_pattern, text)
    paren_match = re.search(paren_pattern, text)
    print(f"{fname}: X pattern {'found' if x_match else 'not found'}, generic parens {'found' if paren_match else 'not found'}")
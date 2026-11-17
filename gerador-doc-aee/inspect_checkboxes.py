import zipfile
import re
from pathlib import Path

def check_checkbox_patterns(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')
            # Look for patterns: "( X )", "(X)", "( X)", "(   )", "(   )", etc.
            # We'll search for a space, then X, then optional spaces and closing parenthesis
            pattern_x = re.search(r'\( *X *\)', content)  # allows spaces around X
            pattern_x2 = re.search(r'\(X\)', content)    # no spaces
            pattern_paren = re.search(r'\( *\)', content)  # generic placeholder
            return pattern_x is not None, pattern_x2 is not None, pattern_paren is not None

rendered_files = [
    "ESTUDO DE CASO - Surdez_template_rendered.docx",
    "PAEE - Surdez_template_rendered.docx",
    "PEI - Surdez_template_rendered.docx"
]

for fname in rendered_files:
    path = Path("./rascunhos/teste_saida") / fname
    has_x, has_x2, has_paren = check_checkbox_patterns(path)
    print(f"{fname}: X pattern (with spaces) {'found' if has_x else 'not found'}, X without spaces {'found' if has_x2 else 'not found'}, generic parens {'found' if has_paren else 'not found'}")
import os
import re
from docx import Document

# --------------------------------------------------------------
# Configurações
# --------------------------------------------------------------
ORIGINALS_DIR = "./originais"
TEMPLATES_DIR = "./templates"

# Mapeamento de nomes de arquivos para (perfil, tipo_de_documento)
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

# --------------------------------------------------------------
# Funções auxiliares
# --------------------------------------------------------------
def unique_tag(profile: str, doc_type: str, tag_type: str, counters: dict, row: int = None, col: int = None) -> str:
    """
    Cria uma tag Jinja2 única.
    tag_type = 'chk'  → {{ chk_<profile>_<doc_type>_<idx> }}
    tag_type = 'cell' → {{ cell_<profile>_<doc_type>_<row>_<col> }}
    """
    idx = counters[tag_type]
    counters[tag_type] += 1
    if tag_type == "chk":
        return f"{{{{ chk_{profile}_{doc_type}_{idx} }}}}"
    else:  # cell
        return f"{{{{ cell_{profile}_{doc_type}_{row}_{col} }}}}"

def replace_placeholders_in_paragraph(paragraph, profile: str, doc_type: str, counters: dict):
    """
    Substitui todas as ocorrências de "( )" dentro do texto do parágrafo.
    """
    def repl(match):
        counters["chk"] += 1
        return unique_tag(profile, doc_type, "chk", counters)

    # regex que captura "( )" (space opcional dentro dos parênteses)
    new_text = re.sub(r"\( *\)", repl, paragraph.text)
    paragraph.text = new_text

def process_document(doc_path: str, profile: str, doc_type: str):
    """
    Carrega o .docx, substitui placeholders e salva o template.
    """
    doc = Document(doc_path)

    # contadores para gerar índices únicos
    counters = {"chk": 0, "cell": 0}

    # ---------- Parágrafos ----------
    for para in doc.paragraphs:
        replace_placeholders_in_paragraph(para, profile, doc_type, counters)

    # ---------- Tabelas ----------
    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                # Substitui "( )" dentro do texto da célula
                def repl_cell(match):
                    counters["chk"] += 1
                    return unique_tag(profile, doc_type, "chk", counters)

                cell.text = re.sub(r"\( *\)", repl_cell, cell.text)

                # Substitui underscores ("____") ou células vazias por tags de célula
                if cell.text.strip() == "" or re.fullmatch(r"_{2,}", cell.text):
                    counters["cell"] += 1
                    cell.text = unique_tag(profile, doc_type, "cell", counters, row_idx, col_idx)

    # --------------------------------------------------------------
    # Salva o template na pasta correspondente ao perfil
    # --------------------------------------------------------------
    profile_dir = os.path.join(TEMPLATES_DIR, profile)
    os.makedirs(profile_dir, exist_ok=True)

    # Mantemos o mesmo nome de arquivo, mas colocamos um sufixo "_template"
    base_name = os.path.splitext(os.path.basename(doc_path))[0]
    new_name = f"{base_name}_template.docx"
    out_path = os.path.join(profile_dir, new_name)

    doc.save(out_path)
    print(f"Template gerado: {out_path}")

# --------------------------------------------------------------
# Main
# --------------------------------------------------------------
def main():
    for filename, (profile, doc_type) in FILE_MAP.items():
        src_path = os.path.join(ORIGINALS_DIR, filename)
        if not os.path.exists(src_path):
            print(f"Arquivo não encontrado: {src_path}")
            continue
        print(f"Processando {src_path} → {profile}/{doc_type}")
        process_document(src_path, profile, doc_type)

if __name__ == "__main__":
    main()
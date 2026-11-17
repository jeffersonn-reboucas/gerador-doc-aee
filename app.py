import base64
import json
import os
import re
import tempfile
import urllib.error
import urllib.request
import zipfile
from pathlib import Path
import streamlit as st
from docx import Document
from docxtpl import DocxTemplate

st.set_page_config(page_title="Gerador de Documentos AEE", layout="wide")
st.title("Gerador de Documentos AEE (CODA / Perda Auditiva / Surdez)")

RASCUNHOS_DIR = Path("./rascunhos")
TEMPLATES_DIR = Path("./templates")
DOSSIES_DIR = Path("./dossies")
RASCUNHOS_DIR.mkdir(exist_ok=True)
DOSSIES_DIR.mkdir(exist_ok=True)

TEMPLATES_MAP = {
    "coda": {
        "estudo_caso": "ESTUDO DE CASO - CODA_template.docx",
        "paee": "PAEE - CODA_template.docx",
        "pei": "PEI - CODA_template.docx",
    },
    "perda_auditiva": {
        "estudo_caso": "ESTUDO DE CASO - Perda Auditiva_template.docx",
        "paee": "PAEE - Perda Auditiva_template.docx",
        "pei": "PEI - Perda Auditiva_template.docx",
    },
    "surdez": {
        "estudo_caso": "ESTUDO DE CASO - Surdez_template.docx",
        "paee": "PAEE - Surdez_template.docx",
        "pei": "PEI - Surdez_template.docx",
    },
}

CAMPOS_IDENTIFICACAO = {
    "nome_completo",
    "data_nascimento",
    "idade",
    "turma",
    "escola_cdi",
    "cdi",
    "municipio_uf",
    "professor_regente",
    "professor_aee_libras",
    "tilsp",
    "data_elaboracao",
    "diagnostico_dispositivos",
    "medicacoes",
    "profissionais_acompanhamento",
}

if "dados" not in st.session_state:
    st.session_state["dados"] = {}
if "arquivo_carregado" not in st.session_state:
    st.session_state["arquivo_carregado"] = ""
if "anexos_aluno" not in st.session_state:
    st.session_state["anexos_aluno"] = []

TAG_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")


def limpar_rotulo_chk(texto_apos: str, tag: str) -> str:
    linhas = [ln.strip() for ln in texto_apos.splitlines() if ln.strip()]
    if not linhas:
        return tag
    primeira = linhas[0]
    for marcador in [
        "Descrição:",
        "Quais?",
        "Faz acompanhamento",
        "Apresenta alguma",
        "Necessita de suporte",
        "Coordenação motora",
        "Leitura e Escrita Visuais:",
        "Noção de Operações Básicas:",
        "Resolução de Problemas Matemáticos:",
    ]:
        if marcador in primeira and not primeira.startswith(marcador):
            primeira = primeira.split(marcador)[0].strip()
    primeira = re.sub(r"_+", "", primeira).strip(" ;:()[]{}-–—.")
    return primeira if primeira else tag


def limpar_rotulo_texto(texto_antes: str, ultimo_subtitulo: str, tag: str) -> str:
    linhas = [ln.strip() for ln in texto_antes.splitlines() if ln.strip()]
    if not linhas:
        return ultimo_subtitulo or tag
    ultima = re.sub(r"_+", "", linhas[-1]).strip(" ;:()[]{}-–—")
    for marcador in [
        "Descrição",
        "Outros",
        "Outras",
        "Quais?",
        "Nome completo",
        "Data de nascimento",
        "Idade",
        "Turma",
        "CDI",
        "Município/UF",
        "Professor(a) regente",
        "Professor(a) do AEE - Libras",
        "Data de elaboração",
        "Professor Regente",
        "Profissional AEE",
        "Responsável Familiar",
        "Coordenação Pedagógica",
    ]:
        if marcador.lower() in ultima.lower():
            if marcador in ("Descrição", "Outros", "Outras", "Quais?") and ultimo_subtitulo:
                return f"{marcador} ({ultimo_subtitulo})"
            return marcador
    if len(ultima) < 75:
        return ultima
    return f"Descrição / Observações ({ultimo_subtitulo})" if ultimo_subtitulo else tag


def detectar_subtitulo(trecho: str, atual: str) -> str:
    for ln in trecho.splitlines():
        s = re.sub(r"_+", "", ln).strip()
        if not s:
            continue
        if s.endswith(":") and len(s) < 85 and not s.lower().startswith(("outros:", "outras:", "descrição:")):
            atual = s.rstrip(":")
        elif s.isupper() and 4 < len(s) < 80:
            atual = s.title()
    return atual


@st.cache_data
def extrair_blocos_documento(caminho_str: str):
    caminho = Path(caminho_str)
    if not caminho.exists():
        return []

    doc = Document(str(caminho))
    blocos = []
    secao_atual = "Geral"

    def extrair_de_texto_corrido(texto: str, secao_base: str):
        matches = list(TAG_RE.finditer(texto))
        if not matches:
            return [], detectar_subtitulo(texto, secao_base)

        itens = []
        sub = secao_base
        pos_anterior = 0
        for i, m in enumerate(matches):
            tag = m.group(1)
            trecho_antes = texto[pos_anterior : m.start()]
            sub = detectar_subtitulo(trecho_antes, sub)
            fim_atual = matches[i + 1].start() if i + 1 < len(matches) else len(texto)
            trecho_depois = texto[m.end() : fim_atual]

            if tag.startswith("chk_"):
                rotulo = limpar_rotulo_chk(trecho_depois, tag)
            else:
                rotulo = limpar_rotulo_texto(trecho_antes, sub, tag)

            itens.append({"tag": tag, "label": rotulo, "subsecao": sub})
            pos_anterior = m.end()

        return itens, sub

    itens_soltos = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        if not TAG_RE.search(txt):
            secao_atual = detectar_subtitulo(txt, secao_atual)
        else:
            novos, secao_atual = extrair_de_texto_corrido(txt, secao_atual)
            itens_soltos.extend(novos)
    if itens_soltos:
        blocos.append({"tipo": "lista", "titulo": secao_atual, "itens": itens_soltos})

    for table in doc.tables:
        num_cols = len(table.columns)
        headers = [c.text.strip().replace("\n", " ") for c in table.rows[0].cells] if table.rows else []

        if num_cols >= 4 and any("Objetivo" in h for h in headers):
            linhas_matriz = []
            for r_idx, row in enumerate(table.rows[1:], start=1):
                celulas = [c.text.strip() for c in row.cells]
                nome_area = TAG_RE.sub("", celulas[0]).replace("\n", " ").strip(" :|-") or f"Linha {r_idx}"
                cols_tags = []
                for c_idx in range(1, len(celulas)):
                    m = TAG_RE.search(celulas[c_idx])
                    if m:
                        col_titulo = headers[c_idx] if c_idx < len(headers) and headers[c_idx] else f"Coluna {c_idx}"
                        cols_tags.append({"tag": m.group(1), "coluna": col_titulo})
                if cols_tags:
                    linhas_matriz.append({"area": nome_area, "colunas": cols_tags})
            if linhas_matriz:
                titulo_matriz = headers[0] if headers and headers[0] else "Planejamento"
                blocos.append({"tipo": "matriz_pei", "titulo": f"Quadro — {titulo_matriz}", "linhas": linhas_matriz})
            continue

        if num_cols == 2:
            itens_kv = []
            for row in table.rows:
                c0 = row.cells[0].text.strip()
                c1 = row.cells[1].text.strip()
                if c0 != c1 and TAG_RE.search(c1):
                    titulo_linha = TAG_RE.sub("", c0).replace("\n", " ").strip(" :|-")
                    novos, _ = extrair_de_texto_corrido(c1, titulo_linha)
                    for it in novos:
                        if not it["tag"].startswith("chk_") and titulo_linha:
                            it["label"] = titulo_linha
                    itens_kv.extend(novos)
                else:
                    novos, secao_atual = extrair_de_texto_corrido(c0, secao_atual)
                    itens_kv.extend(novos)
            if itens_kv:
                blocos.append({"tipo": "lista", "titulo": "Itens do Quadro", "itens": itens_kv})
            continue

        itens_quadro = []
        sub_quadro = secao_atual
        celulas_vistas = set()
        for row in table.rows:
            for cell in row.cells:
                if id(cell._tc) in celulas_vistas:
                    continue
                celulas_vistas.add(id(cell._tc))
                txt_cell = cell.text.strip()
                if not txt_cell:
                    continue
                novos, sub_quadro = extrair_de_texto_corrido(txt_cell, sub_quadro)
                itens_quadro.extend(novos)
        if itens_quadro:
            blocos.append({"tipo": "lista", "titulo": sub_quadro, "itens": itens_quadro})

    return blocos


def mapear_todas_tags_perfil(perfil_key: str) -> dict:
    mapa = {
        "nome_completo": "Nome completo do estudante",
        "data_nascimento": "Data de nascimento (DD/MM/AAAA)",
        "idade": "Idade",
        "turma": "Turma / Série / Turno",
        "escola_cdi": "Escola ou CDI",
        "municipio_uf": "Município/UF",
        "professor_regente": "Professor regente",
        "professor_aee_libras": "Professor AEE Libras",
        "tilsp": "Intérprete / Auxiliar Libras (TILSP)",
        "data_elaboracao": "Data de elaboração",
        "diagnostico_dispositivos": "Diagnóstico auditivo/médico e uso de AASI/IC/FM",
        "medicacoes": "Medicações contínuas",
        "profissionais_acompanhamento": "Profissionais que acompanham (Fono, Otorrino, etc.)",
    }
    for doc_type, filename in TEMPLATES_MAP[perfil_key].items():
        tpath = TEMPLATES_DIR / perfil_key / filename
        for b in extrair_blocos_documento(str(tpath)):
            if b["tipo"] == "matriz_pei":
                for ln in b["linhas"]:
                    for c in ln["colunas"]:
                        mapa[c["tag"]] = f"PEI {ln['area']} - {c['coluna']}"
            else:
                for it in b["itens"]:
                    mapa[it["tag"]] = f"{doc_type.upper()} [{it['subsecao']}]: {it['label']}"
    return mapa


def extrair_texto_local(uploaded_files) -> str:
    """Extrai texto de arquivos .docx localmente."""
    textos = []
    for uf in uploaded_files:
        raw_bytes = uf.getvalue()
        nome_lower = uf.name.lower()
        if nome_lower.endswith(".docx"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                tmp.write(raw_bytes)
                tmp_path = tmp.name
            d = Document(tmp_path)
            txt = "\n".join(p.text for p in d.paragraphs if p.text.strip())
            for t in d.tables:
                for row in t.rows:
                    txt += "\n" + " | ".join(c.text.strip() for c in row.cells)
            textos.append(f"--- ARQUIVO {uf.name} ---\n{txt}")
            Path(tmp_path).unlink(missing_ok=True)
    return "\n\n".join(textos)


def montar_prompt_extracao(perfil_key: str, uploaded_files=None) -> str:
    mapa_tags = mapear_todas_tags_perfil(perfil_key)
    prompt = (
        "Atue como Professor de Educação Especial - AEE Libras. "
        "Analise os documentos anexados do(a) estudante (Ficha de Matrícula, Entrevista Familiar manuscrita, "
        "Laudo Audiológico/Multidisciplinar e relatórios anteriores) e retorne APENAS um código JSON válido "
        "(sem comentários fora do JSON) preenchendo as chaves abaixo:\n\n"
        "REGRAS:\n"
        "1. Para chaves iniciadas por 'chk_', preencha com '( X )' se for verdadeiro para o aluno, ou omita a chave.\n"
        "2. Para campos de texto, Descrições, Objetivos, PAEE e Tabelas do PEI, redija respostas pedagógicas completas "
        "e coerentes com o quadro clínico/linguístico do aluno.\n\n"
        f"MAPA DE CHAVES DO SISTEMA (Perfil: {perfil_key.upper()}):\n"
        f"{json.dumps(mapa_tags, ensure_ascii=False, indent=2)}\n"
    )
    if uploaded_files:
        txt_local = extrair_texto_local(uploaded_files)
        if txt_local:
            prompt += f"\nTEXTO EXTRAÍDO DOS ARQUIVOS WORD:\n{txt_local}\n"
    return prompt


def limpar_e_carregar_json_texto(texto_bruto: str) -> dict:
    limpo = re.sub(r"^```(?:json)?|```$", "", texto_bruto.strip(), flags=re.MULTILINE).strip()
    inicio = limpo.find("{")
    fim = limpo.rfind("}")
    if inicio != -1 and fim != -1:
        limpo = limpo[inicio : fim + 1]
    return json.loads(limpo)


def chamar_gemini_rest_nativo(uploaded_files, perfil_key: str, api_key: str) -> dict:
    """Chama a API do Gemini via HTTPS puro (suporta PDFs escaneados, imagens e Word)."""
    prompt = montar_prompt_extracao(perfil_key, uploaded_files)
    parts = []
    for uf in uploaded_files:
        raw_bytes = uf.getvalue()
        nome_lower = uf.name.lower()
        if nome_lower.endswith(".pdf"):
            parts.append(
                {
                    "inline_data": {
                        "mime_type": "application/pdf",
                        "data": base64.b64encode(raw_bytes).decode("utf-8"),
                    }
                }
            )
        elif nome_lower.endswith((".png", ".jpg", ".jpeg")):
            mime = "image/png" if nome_lower.endswith(".png") else "image/jpeg"
            parts.append(
                {
                    "inline_data": {
                        "mime_type": mime,
                        "data": base64.b64encode(raw_bytes).decode("utf-8"),
                    }
                }
            )
    parts.append({"text": prompt})

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key.strip()}"
    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {"responseMimeType": "application/json"},
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        res_data = json.loads(resp.read().decode("utf-8"))
    texto_json = res_data["candidates"][0]["content"]["parts"][0]["text"]
    return limpar_e_carregar_json_texto(texto_json)


def chamar_openai_compativel_rest(uploaded_files, perfil_key: str, base_url: str, api_key: str, model_name: str) -> dict:
    """Conecta com OmniRoute local, NVIDIA NIM, DeepSeek ou OpenRouter via HTTPS puro."""
    prompt = montar_prompt_extracao(perfil_key, uploaded_files)
    url = base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": model_name.strip(),
        "messages": [
            {"role": "system", "content": "Retorne apenas um objeto JSON válido, sem comentários fora do JSON."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.2,
        "max_tokens": 4096,
    }
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0",
    }
    if api_key.strip():
        headers["Authorization"] = f"Bearer {api_key.strip()}"

    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            res_data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as err:
        detalhe = err.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"HTTP {err.code} ({err.reason}): {detalhe}") from err

    conteudo = res_data["choices"][0]["message"]["content"].strip()
    return limpar_e_carregar_json_texto(conteudo)


def aplicar_json(novos_dados: dict, nome_origem: str):
    st.session_state["dados"] = novos_dados
    st.session_state["arquivo_carregado"] = nome_origem
    for k, v in novos_dados.items():
        wkey = f"w_{k}"
        if v in (True, "( X )", "( x )", "X", "x"):
            st.session_state[wkey] = True
        elif v in (False, "(   )", "( )", ""):
            st.session_state[wkey] = False if k.startswith("chk_") else ""
        else:
            st.session_state[wkey] = "" if v is None else str(v)


# --- BARRA LATERAL ---
st.sidebar.header("Configurações")
perfil_label = st.sidebar.selectbox("Perfil do Estudante", ["CODA", "Perda Auditiva", "Surdez"], index=2)
perfil_key = {"CODA": "coda", "Perda Auditiva": "perda_auditiva", "Surdez": "surdez"}[perfil_label]

st.sidebar.divider()
st.sidebar.subheader("📂 1. Importar Documentos / IA")

arquivos_brutos = st.sidebar.file_uploader(
    "Anexar PDFs / DOCXs do aluno",
    type=["pdf", "docx", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
)
if arquivos_brutos:
    st.session_state["anexos_aluno"] = arquivos_brutos

modo_ia = st.sidebar.radio(
    "Como deseja extrair os dados?",
    [
        "💬 Modo Chat (Copiar Prompt p/ Gemini/GPT/DeepSeek)",
        "🔑 Gemini API Direta (Sem instalar pacotes)",
        "🌐 OmniRoute / NVIDIA / DeepSeek / OpenAI",
    ],
)

if modo_ia.startswith("💬"):
    st.sidebar.caption("1. Copie o prompt abaixo e envie junto com os PDFs do aluno no chat do Gemini, ChatGPT, Claude ou DeepSeek.")
    prompt_pronto = montar_prompt_extracao(perfil_key, arquivos_brutos)
    with st.sidebar.expander("📋 Ver / Copiar Prompt para o Chat"):
        st.code(prompt_pronto, language="markdown")

    json_colado = st.sidebar.text_area(
        "2. Cole aqui o JSON que o Chat respondeu:",
        height=110,
        placeholder='{"nome_completo": "...", "chk_...": "( X )"}',
    )
    if st.sidebar.button("✅ Aplicar JSON do Chat", type="primary", use_container_width=True):
        if json_colado.strip():
            try:
                dados_chat = limpar_e_carregar_json_texto(json_colado)
                aplicar_json({**st.session_state["dados"], **dados_chat}, "JSON colado do Chat")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"JSON inválido: {e}")

elif modo_ia.startswith("🔑"):
    api_key_input = st.sidebar.text_input(
        "Chave Gemini API",
        value=os.environ.get("GEMINI_API_KEY", ""),
        type="password",
    )
    if st.sidebar.button("✨ Extrair com Gemini REST", type="primary", use_container_width=True):
        if not api_key_input or not arquivos_brutos:
            st.sidebar.error("Anexe os arquivos e informe a chave Gemini.")
        else:
            with st.sidebar.spinner("Lendo PDFs e manuscritos via HTTPS..."):
                try:
                    extraido = chamar_gemini_rest_nativo(arquivos_brutos, perfil_key, api_key_input)
                    aplicar_json({**st.session_state["dados"], **extraido}, "Extraído via Gemini")
                    st.rerun()
                except Exception as e:
                    st.sidebar.error(f"Erro na chamada: {e}")

else:
    provedor = st.sidebar.selectbox(
        "Provedor",
        ["OmniRoute Local", "NVIDIA NIM", "DeepSeek API", "Outro (OpenAI Compatível)"],
    )
    defaults_url = {
        "OmniRoute Local": ("http://localhost:8000/v1", "gemini-2.5-flash"),
        "NVIDIA NIM": ("https://integrate.api.nvidia.com/v1", "qwen/qwen2.5-72b-instruct"),
        "DeepSeek API": ("https://api.deepseek.com/v1", "deepseek-chat"),
        "Outro (OpenAI Compatível)": ("https://api.openai.com/v1", "gpt-4o-mini"),
    }
    url_padrao, mod_padrao = defaults_url[provedor]
    base_url = st.sidebar.text_input("Base URL", value=url_padrao)
    model_name = st.sidebar.text_input("Modelo", value=mod_padrao)
    api_key_oa = st.sidebar.text_input("API Key (se necessário)", type="password")

    if st.sidebar.button("🚀 Extrair via Endpoint", type="primary", use_container_width=True):
        if not arquivos_brutos:
            st.sidebar.error("Anexe os arquivos do aluno primeiro.")
        else:
            with st.sidebar.spinner("Consultando endpoint e preenchendo formulário..."):
                try:
                    extraido = chamar_openai_compativel_rest(
                        arquivos_brutos, perfil_key, base_url, api_key_oa, model_name
                    )
                    aplicar_json({**st.session_state["dados"], **extraido}, f"Extraído via {provedor}")
                    st.rerun()
                except Exception as e:
                    st.sidebar.error(f"Erro no endpoint: {e}")

st.sidebar.divider()
st.sidebar.subheader("💾 2. Rascunhos Salvos (JSON)")

arquivos_json = sorted([f.name for f in RASCUNHOS_DIR.glob("*.json")])
escolhido = st.sidebar.selectbox("Selecionar rascunho", ["<Nenhum>"] + arquivos_json)

if st.sidebar.button("Carregar Selecionado", use_container_width=True) and escolhido != "<Nenhum>":
    conteudo = json.loads((RASCUNHOS_DIR / escolhido).read_text(encoding="utf-8"))
    aplicar_json(conteudo, escolhido)
    st.rerun()

upload = st.sidebar.file_uploader("Ou envie um arquivo .json", type=["json"])
if upload is not None and upload.name != st.session_state.get("ultimo_upload"):
    conteudo = json.load(upload)
    st.session_state["ultimo_upload"] = upload.name
    aplicar_json(conteudo, upload.name)
    st.rerun()

if st.session_state["arquivo_carregado"]:
    st.sidebar.success(f"Ativo: {st.session_state['arquivo_carregado']}")

# --- ABAS ---
tab1, tab2, tab3, tab4 = st.tabs(["1. Identificação e Saúde", "2. Estudo de Caso", "3. PAEE", "4. PEI"])
dados_form = {}


def render_widget(tag: str, label: str, usar_textarea: bool = False):
    if tag in dados_form:
        return
    wkey = f"w_{tag}"
    valor_json = st.session_state["dados"].get(tag, "")
    if tag.startswith("chk_"):
        if wkey not in st.session_state or not isinstance(st.session_state[wkey], bool):
            st.session_state[wkey] = valor_json in (True, "( X )", "( x )", "X", "x")
        res = st.checkbox(label, key=wkey)
        dados_form[tag] = "( X )" if res else "(   )"
    else:
        if wkey not in st.session_state or not isinstance(st.session_state[wkey], str):
            st.session_state[wkey] = "" if valor_json is None else str(valor_json)
        if usar_textarea or "Descri" in label or len(st.session_state[wkey]) > 70:
            res = st.text_area(label, height=80, key=wkey)
        else:
            res = st.text_input(label, key=wkey)
        dados_form[tag] = res


def render_aba_documento(doc_type: str):
    tpath = TEMPLATES_DIR / perfil_key / TEMPLATES_MAP[perfil_key][doc_type]
    blocos = extrair_blocos_documento(str(tpath))

    sub_atual = None
    for b in blocos:
        if b["tipo"] == "matriz_pei":
            st.markdown(f"### {b['titulo']}")
            for linha in b["linhas"]:
                with st.container(border=True):
                    st.markdown(f"**{linha['area']}**")
                    cols = st.columns(len(linha["colunas"]))
                    for idx_c, col_info in enumerate(linha["colunas"]):
                        with cols[idx_c]:
                            render_widget(col_info["tag"], col_info["coluna"], usar_textarea=True)
        else:
            for item in b["itens"]:
                if item["tag"] in CAMPOS_IDENTIFICACAO:
                    continue
                if item["subsecao"] and item["subsecao"] != sub_atual:
                    sub_atual = item["subsecao"]
                    st.markdown("---")
                    st.markdown(f"#### {sub_atual}")
                render_widget(item["tag"], item["label"])


with tab1:
    st.subheader("Dados de Identificação (Comuns aos 3 documentos)")
    campos_comuns = [
        ("nome_completo", "Nome completo do(a) estudante"),
        ("data_nascimento", "Data de nascimento (DD/MM/AAAA)"),
        ("idade", "Idade"),
        ("turma", "Turma / Ano"),
        ("escola_cdi", "Escola / CDI"),
        ("municipio_uf", "Município / UF"),
        ("professor_regente", "Professor(a) Regente"),
        ("professor_aee_libras", "Professor(a) do AEE - Libras"),
        ("tilsp", "Tradutor e Intérprete de Libras (TILSP)"),
        ("data_elaboracao", "Data de elaboração (DD/MM/AAAA)"),
        ("diagnostico_dispositivos", "Diagnóstico / Dispositivos auditivos (AASI, IC, Sistema FM)"),
        ("medicacoes", "Medicações de uso contínuo"),
        ("profissionais_acompanhamento", "Profissionais de acompanhamento"),
    ]
    cols = st.columns(2)
    for idx, (chave, rotulo) in enumerate(campos_comuns):
        with cols[idx % 2]:
            render_widget(chave, rotulo)
    if "escola_cdi" in dados_form:
        dados_form["cdi"] = dados_form["escola_cdi"]

with tab2:
    st.subheader(f"Estudo de Caso — {perfil_label}")
    render_aba_documento("estudo_caso")

with tab3:
    st.subheader(f"Plano de AEE (PAEE) — {perfil_label}")
    render_aba_documento("paee")

with tab4:
    st.subheader(f"Plano Educacional Individualizado (PEI) — {perfil_label}")
    render_aba_documento("pei")

for k, v in st.session_state["dados"].items():
    if k not in dados_form:
        dados_form[k] = v

# --- AÇÕES FINAIS ---
st.divider()
col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if st.button("Salvar Rascunho (JSON)", use_container_width=True):
        nome_base = (dados_form.get("nome_completo") or "aluno_sem_nome").strip().replace(" ", "_")
        destino = RASCUNHOS_DIR / f"{nome_base}.json"
        destino.write_text(json.dumps(dados_form, ensure_ascii=False, indent=2), encoding="utf-8")
        st.session_state["dados"] = dados_form
        st.success(f"Rascunho salvo em: {destino}")

with col_btn2:
    if st.button("Gerar Pasta Completa (.ZIP)", type="primary", use_container_width=True):
        nome_base = (dados_form.get("nome_completo") or "aluno").strip().replace(" ", "_")
        pasta_dossie = DOSSIES_DIR / nome_base
        pasta_dossie.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            gerados = []
            for doc_type, filename in TEMPLATES_MAP[perfil_key].items():
                tpath = TEMPLATES_DIR / perfil_key / filename
                if not tpath.exists():
                    st.error(f"Template ausente: {tpath}")
                    continue
                doc = DocxTemplate(str(tpath))
                doc.render(dados_form)
                out_doc = tmp / f"{nome_base}_{doc_type}.docx"
                doc.save(str(out_doc))
                gerados.append(out_doc)

            if gerados:
                zip_bytes_path = tmp / f"{nome_base}_Dossie_Completo_AEE.zip"
                with zipfile.ZipFile(zip_bytes_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for arq in gerados:
                        zf.write(arq, arcname=f"1_Documentos_AEE_Gerados/{arq.name}")
                    for uf in st.session_state.get("anexos_aluno", []):
                        caminho_salvo = pasta_dossie / uf.name
                        caminho_salvo.write_bytes(uf.getvalue())
                        zf.write(caminho_salvo, arcname=f"2_Anexos_e_Laudos_Originais/{uf.name}")

                st.download_button(
                    label=f"Baixar {nome_base}_Dossie_Completo_AEE.zip",
                    data=zip_bytes_path.read_bytes(),
                    file_name=f"{nome_base}_Dossie_Completo_AEE.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
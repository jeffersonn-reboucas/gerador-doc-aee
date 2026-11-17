import json

data = {}

# Helper to add placeholder keys
def add_placeholders(profile, doc_type, max_idx):
    for i in range(1, max_idx + 1):
        key = f"chk_{profile}_{doc_type}_{i}"
        # For simplicity, set all to True (checked)
        data[key] = True

# Add placeholders for each template
add_placeholders("surdez", "estudo_caso", 99)   # 99 placeholders
add_placeholders("surdez", "paee", 11)        # 11 placeholders
add_placeholders("surdez", "pei", 55)         # 55 placeholders

# Add other required fields (minimal set)
data.update({
    "nome_completo": "Lucas Mendes",
    "idade": 9,
    "turma": "4º ano",
    "escola_cdi": "Escola Exemplo",
    "municipio_uf": "São Paulo, SP",
    "professor_regente": "Ana Silva",
    "professor_aee_libras": "Carlos Silva",
    "tilsp": "Maria Oliveira",
    "data_elaboracao": "2026-01-01",
    "diagnostico_dispositivos": "L1 (Libras) e L2 (Português escrito)",
    "medicacoes": "Nenhuma",
    "profissionais_acompanhamento": "Equipe multiprofissional",
    "utiliza_aasi": True,
    "utiliza_ic": False,
    "utiliza_fm": False,
    "usa_medicacao": False,
    "acompanhamento_multi": True,
    "metodologia": "Teste",
    "socioemocional": "Teste",
    "motor_percepcao_visual": "Teste",
    "linguagem_comunicacao": "Teste",
    "sondagem_portugues": "Teste",
    "interpretacao_expressao": "Teste",
    "cognitivo": "Teste",
    "sondagem_matematica": "Teste",
    "texto_modelo": "Teste",
    "eixo": "Comunicação",
    "periodo": "Semestral",
    "frequencia": "Semanal",
    "duracao": "45 min",
    "modalidade": "Presencial",
    "local": "Sala de Recursos",
    "objetivos_especificos": "Teste",
    "estrategias_recursos": "Teste",
    "articulacao_sala_comum": "Teste",
    "periodicidade_revisao": "Semestral",
    "objetivo_geral": "Teste",
    "objetivo_socializacao": "Teste",
    "objetivo_linguagem": "Teste",
    "objetivo_autonomia": "Teste",
    "objetivo_motoras": "Teste",
    "objetivo_portugues": "Teste",
    "objetivo_matematica": "Teste",
    "objetivo_ciencias": "Teste",
    "objetivo_hist_geog": "Teste",
    "objetivo_ed_fisica": "Teste",
    "objetivo_artes": "Teste",
    "objetivo_lingua_estrangeira": "Teste",
    "estrategia": "Teste",
    "adaptacao_curricular": "Teste",
    "recurso": "Teste",
    "articulacao_intersetorial": "Teste",
    "avaliacao": "Teste"
})

with open('teste_lucas_mendes.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
print("JSON file generated")
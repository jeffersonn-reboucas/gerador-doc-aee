import json

data = {
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

    # Estudo de Caso fields
    "metodologia": "Abordagem qualitativa baseada em observação e entrevista",
    "socioemocional": "Observações de interação social e emoções",
    "motor_percepcao_visual": "Descrição das habilidades motoras e perceptivas",
    "linguagem_comunicacao": "Formas de comunicação verbal e não verbal",
    "sondagem_portugues": "Avaliação da produção oral em Português",
    "interpretacao_expressao": "Uso de Libras para expressão oral",
    "cognitivo": "Avaliação das habilidades cognitivas",
    "sondagem_matematica": "Avaliação de raciocínio lógico e resolução de problemas",
    "texto_modelo": "Exemplo de texto modelo a ser inserido nos documentos",

    # Checkboxes for Estudo de Caso (profile Surdez)
    "chk_surdez_estudo_caso_1": True,
    "chk_surdez_estudo_caso_2": True,
    "chk_surdez_estudo_caso_5": False,
    "chk_surdez_estudo_caso_6": True,
    "chk_surdez_estudo_caso_8": True,

    # PAEE fields
    "eixo": "Comunicação",
    "periodo": "Semestral",
    "frequencia": "Semanal",
    "duracao": "45 min",
    "modalidade": "Presencial",
    "local": "Sala de Recursos",
    "objetivos_especificos": "Desenvolver habilidades de comunicação e autonomia",
    "estrategias_recursos": "Uso de recursos visuais e auditivos",
    "articulacao_sala_comum": "Integração com aulas regulares",
    "periodicidade_revisao": "Semestral",

    # PEI fields
    "objetivo_geral": "Promover inclusão e aprendizagem significativa",
    "objetivo_socializacao": "Participação en atividades grupais",
    "objetivo_linguagem": "Uso de Libras para comunicação",
    "objetivo_autonomia": "Independência nas atividades de autocuidado",
    "objetivo_motoras": "Movimentação segura no ambiente escolar",
    "objetivo_portugues": "Compreensão de textos escritos",
    "objetivo_matematica": "Resolução de problemas de adição e subtração",
    "objetivo_ciencias": "Observação de fenômenos naturais",
    "objetivo_hist_geog": "Identificação de eventos históricos",
    "objetivo_ed_fisica": "Participação em atividades físicas",
    "objetivo_artes": "Expressão artística",
    "objetivo_lingua_estrangeira": "Vocabulário básico em Inglês",

    # Table area keys (for data_editor defaults)
    "objetivo_socializacao": "Participação em atividades grupais",
    "objetivo_linguagem": "Uso de Libras para comunicação",
    "objetivo_autonomia": "Independência nas atividades de autocuidado",
    "objetivo_motoras": "Movimentação segura no ambiente escolar",

    # Table discipline keys
    "objetivo_portugues": "Compreensão de textos escritos",
    "objetivo_matematica": "Resolução de problemas de adição e subtração",
    "objetivo_ciencias": "Observação de fenômenos naturais",
    "objetivo_hist_geog": "Identificação de eventos históricos",
    "objetivo_ed_fisica": "Participação em atividades físicas",
    "objetivo_artes": "Expressão artística",
    "objetivo_lingua_estrangeira": "Vocabulário básico em Inglês",

    # Additional PEI fields
    "estrategia": "Aplicar recursos multimodais",
    "adaptacao_curricular": "Adaptar material didático",
    "recurso": "Uso de technology aids",
    "articulacao_intersetorial": "Coordenação com equipe pedagógica",
    "avaliacao": "Avaliação formativa contínua"
}

with open('teste_lucas_mendes.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
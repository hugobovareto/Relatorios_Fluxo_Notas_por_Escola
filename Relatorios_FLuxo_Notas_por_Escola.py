'''
Objetivo: gerar relatórios por escola com os alunos em tabelas diferentes para identificação dos estudantes com notas abaixo de 6 em cada componente

Fonte de Dados: 
- Relatório de Notas dos Estudantes.

Usar tratamentos e processos semelhantes ao usado no site de rendimento escolar RN:
https://rendimento-escolar-rn.streamlit.app/
https://github.com/hugobovareto/Rendimento_Escolar_RN/blob/main/processamento_local.py

Considerar somente as Séries dos Anos Finais do EF e Ensino Médio:
- 1ª Série;
- 2ª Série;
- 3ª Série;
- 6º Ano;
- 7º Ano;
- 8º Ano;
- 9º Ano.

O identificador único usado é o CPF:
- Relatório Geral de Matrículas: 'CPF PESSOA'.

Considerar somente os componentes da BNCC para o df_notas:
bncc = ['Arte',
        'Biologia',
        'Ciências',
        'Educação Física',
        'Filosofia',
        'Física',
        'Geografia',
        'História',
        'Língua Inglesa',
        'Língua Portuguesa',
        'Matemática',
        'Química',
        'Sociologia']

O Relatório de Notas está segmentado por componentes, então nesse, cria a variável:
- nota média.

Para notas, considerar somente as notas do 1º, 2º e 3º bimestres.
Caso o estudante não tenha nota em algum bimestre, essa nota é trasnformada em 0, logo impacta na média.
    A média é sempre feita considerando as notas dos 3 bimestres, mesmo que a nota seja 0.

Para valores duplicatos de CPF + COMPONENTE + SÉRIE, foi mantida a nota mais baixa. 
    

Gerar apresentações por escola.

Nas tabelas dos componentes, os estudantes de diferentes séries estarão juntos, mas organizar para ter agrupado os estudantes da mesma série em ordem crescente de nota. Ex.:
1º) Estudantes do 6º ano em ordem crescente de notas (piores notas primeiro);
2º) Estudantes do 7º ano em ordem crescente de notas (piores notas primeiro);
3º) Estudantes do 8º ano em ordem crescente de notas (piores notas primeiro);
etc



'''
# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
from datetime import datetime
from dateutil.relativedelta import relativedelta
import openpyxl
import re
import copy
from pathlib import Path
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.util import Pt


# Base inicial geral de Notas
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Notas"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df = pd.concat(dfs, ignore_index=True)


# Excluir colunas que não são de interesse
df = df.drop(columns=['ID DIREC', 'ID MUNICÍPIO', 'ID ESCOLA', 'ID ETAPA ENSINO', 'PERIODICIDADE ETAPA ENSINO', 'ID SÉRIE', 'ID PESSOA (PROFESSOR)', 'MATRICULA (PROFESSOR)', 'VÍNCULO', 'NOME DO PROFESSOR', 'DATA INÍCIO ALOCAÇÃO', 'DATA FIM ALOCAÇÃO', 'ID COMPONENTE CURRICULAR', 'PERIODICIDADE COMPONENTE CURRICULAR', 'ID PESSOA', 'APROVEITAMENTO DE ESTUDO'])

# Substituir vírgula por ponto para reconhecimento das notas como números:
colunas_para_converter = [
    "NOTA 1º BIMESTRE",
    "NOTA 2º BIMESTRE",
    "NOTA 3º BIMESTRE",
    "NOTA 4º BIMESTRE",
    "MÉDIA ANUAL",
    "EXAME FINAL",
    "AVALIAÇÃO ESPECIAL",
    "MÉDIA FINAL"
] 

for col in colunas_para_converter:
    if col in df.columns:  # só executa se a coluna estiver no DataFrame
        # Converte para texto e substitui vírgula por ponto
        df[col] = df[col].astype("string").str.replace(",", ".", regex=False)

        # Converte para numérico (valores inválidos viram NaN)
        df[col] = pd.to_numeric(df[col], errors="coerce")


# Manter só Anos Finais e Ensino Médio:
valores_desejados = ['1ª SÉRIE',
                    '2ª SÉRIE',
                    '3ª SÉRIE',
                    '6º Ano',
                    '7º Ano',
                    '8º Ano',
                    '9º Ano',
                    '6º ANO',
                    '7º ANO',
                    '8º ANO',
                    '9º ANO']

df_EF_EM = df[df['SÉRIE'].isin(valores_desejados)]

# substituição das séries e manter padronização
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df_EF_EM['SÉRIE'] = df_EF_EM['SÉRIE'].replace(mapeamento)


# Manter só componentes da BNCC:
bncc = ['Arte',
        'Biologia',
        'Educação Física',
        'Filosofia',
        'Física',
        'Geografia',
        'História',
        'Língua Inglesa',
        'Língua Portuguesa',
        'Matemática',
        'Química',
        'Sociologia', 
        'Ciências']

df_EF_EM_bncc = df_EF_EM[df_EF_EM['COMPONENTE CURRICULAR'].isin(bncc)]

# Criar coluna de "ETAPA_RESUMIDA" para indicar Anos Finais ou Ensino Médio, de acordo com a série
# (São 46 etapas de ensino na base já filtrada pelas séries dos Anos Finais e Ensino Médio e pelos componentes da BNCC)
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df_EF_EM_bncc['ETAPA_RESUMIDA'] = df_EF_EM_bncc['SÉRIE'].map(mapeamento_etapa)

# Criar coluna com nota final média, considerando as notas do 1º, 2º e 3º bimestres
# Notas NaN transformadas em 0 para as notas do 1º, 2º e 3º bimestres
cols_bimestres = ['NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE', 'NOTA 3º BIMESTRE']

df_EF_EM_bncc[cols_bimestres] = (
    df_EF_EM_bncc[cols_bimestres]
    .apply(pd.to_numeric, errors='coerce')
    .fillna(0)
)



                                    ###### MODIFICAR AQUI QUANDO TIVER MAIS NOTAS LANÇADAS ######
df_EF_EM_bncc['MEDIA_NOTAS'] = df_EF_EM_bncc[['NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE', 'NOTA 3º BIMESTRE']].mean(axis=1, skipna=True)

# Criar uma coluna para Aprovado ou Reprovado por componente (reprovação caso a média seja menor que 6)
# (sem nota caso os dois bimestres sejam NaN)
df_EF_EM_bncc['STATUS'] = np.where(
    df_EF_EM_bncc['MEDIA_NOTAS'].isna(),           # 1️⃣ caso: sem média
    'Sem nota',
    np.where(
        df_EF_EM_bncc['MEDIA_NOTAS'].fillna(0) >= 6,         # 2️⃣ caso: média suficiente
        'Aprovado',
        'Reprovado'                                  # 3️⃣ caso: média < 6
    )
)


# Garantir que cada estudante seja contado apenas uma vez por componente e por série
# Excluir duplicatas e manter somente o CPF + SÉRIE + COMPONENTE CURRICULAR com valor menor para MEDIA_NOTAS
chaves = ['CPF PESSOA', 'COMPONENTE CURRICULAR', 'SÉRIE']

# Garante que MEDIA_NOTAS seja numérica (evita ordenação alfabética onde "10" vem antes de "2")
df_EF_EM_bncc['MEDIA_NOTAS'] = pd.to_numeric(df_EF_EM_bncc['MEDIA_NOTAS'], errors='coerce')

# Coloca os valores ausentes (NaN) no topo e ordena os demais do menor para o maior
df_ordenado = (
    df_EF_EM_bncc
    .sort_values(by='MEDIA_NOTAS', ascending=True, na_position='first')
)

# Mantém a 1ª ocorrência de cada grupo (NaN se existir; caso contrário, a menor nota)
df_sem_duplicatas = (
    df_ordenado
    .drop_duplicates(subset=chaves, keep='first')
    .sort_index()
)

# Excel com os valores que aparecem duplicados na base
# Filtra todas as linhas que pertencem a alguma combinação duplicada
df_duplicados = df_ordenado[df_ordenado.duplicated(subset=chaves, keep=False)].copy()

# Identifica qual linha foi mantida ('Mantida') e quais foram descartadas ('Excluída')
df_duplicados['ACAO_DUPLICATA'] = np.where(
    df_duplicados.duplicated(subset=chaves, keep='first'),
    'Excluída',
    'Mantida'
)

# Ordena por estudante, componente, série e coloca 'Mantida' antes de 'Excluída'
df_duplicados = df_duplicados.sort_values(
    by=chaves + ['ACAO_DUPLICATA'],
    ascending=[True, True, True, False]
)

caminho_duplicatas = r"D:\Scripts_Python\FGV\Relatorios_Fluxo_Notas_por_Escola\registros_duplicados.xlsx"

with pd.ExcelWriter(caminho_duplicatas, engine="xlsxwriter") as writer:
    df_duplicados.to_excel(writer, sheet_name="Duplicados", index=False)

print(f"Total de linhas duplicadas exportadas: {len(df_duplicados)}")


# Salvar em Excel o dataframe
caminho_xlsx = r"D:\Scripts_Python\FGV\Relatorios_Fluxo_Notas_por_Escola\notas_por_componentes.xlsx"

with pd.ExcelWriter(caminho_xlsx, engine="xlsxwriter") as writer:
    for direc, grupo in df_sem_duplicatas.groupby("DIREC", dropna=False, sort=True):
        if pd.isna(direc) or str(direc).strip() == "":
            nome_aba = "Sem_DIREC"
        else:
            # Pega apenas o trecho antes do hífen e remove espaços nas extremidades
            nome_aba = str(direc).split("-")[0].strip()

        grupo.to_excel(writer, sheet_name=nome_aba, index=False)
        print(f"Aba '{nome_aba}' gravada: {len(grupo)} linhas")



# Gerar base de dados somente dos reprovados
df_reprovados = df_sem_duplicatas[df_sem_duplicatas['STATUS'].str.strip() == 'Reprovado'].copy()








































###########################################################
# Criação das planilhas para cada escola com 2 abas: Reprovados e Todos os Estudantes

pasta_raiz = Path(r"D:\Scripts_Python\FGV\Relatorios_Fluxo_Notas_por_Escola\relatorios_gerados")
pasta_raiz.mkdir(parents=True, exist_ok=True)

# Remove sufixo '.0' caso INEP ESCOLA tenha sido lido como float e trata nulos
def limpar_inep(serie):
    return (
        serie.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .replace({"nan": "Sem_INEP", "None": "Sem_INEP", "": "Sem_INEP"})
    )

df_todos = df_sem_duplicatas.copy()
df_rep = df_reprovados.copy()

df_todos["INEP_STR"] = limpar_inep(df_todos["INEP ESCOLA"])
df_rep["INEP_STR"] = limpar_inep(df_rep["INEP ESCOLA"])

# Pré-agrupa por (DIREC, INEP) em dicionários para evitar filtros repetitivos dentro do loop
grupos_todos = {
    chave: grupo.drop(columns=["INEP_STR"])
    for chave, grupo in df_todos.groupby(["DIREC", "INEP_STR"], dropna=False)
}
grupos_reprovados = {
    chave: grupo.drop(columns=["INEP_STR"])
    for chave, grupo in df_rep.groupby(["DIREC", "INEP_STR"], dropna=False)
}

# DataFrame vazio com as colunas originais para escolas sem nenhum reprovado
df_reprovados_vazio = df_reprovados.iloc[0:0].copy()

total_arquivos = 0

for (direc, inep_str), df_escola_todos in grupos_todos.items():
    # Define o nome da pasta da DIREC mantendo o texto original (higienizando apenas caracteres proibidos no Windows)
    if pd.isna(direc) or str(direc).strip() == "":
        nome_pasta_direc = "Sem_DIREC"
    else:
        nome_pasta_direc = re.sub(r'[\\/*?:"<>|]', "_", str(direc).strip())

    pasta_direc = pasta_raiz / nome_pasta_direc
    pasta_direc.mkdir(parents=True, exist_ok=True)

    # Recupera os reprovados da escola (ou aba vazia com cabeçalho se houver 0 reprovados)
    df_escola_rep = grupos_reprovados.get((direc, inep_str), df_reprovados_vazio)

    caminho_arquivo = pasta_direc / f"{inep_str}_notas_estudantes.xlsx"

    with pd.ExcelWriter(caminho_arquivo, engine="xlsxwriter") as writer:
        df_escola_rep.to_excel(writer, sheet_name="Reprovados", index=False)
        df_escola_todos.to_excel(writer, sheet_name="Todos alunos", index=False)

    total_arquivos += 1

print(f"Concluído: {total_arquivos} planilhas geradas em '{pasta_raiz}'.")





































###########################################################
# Apresentações para cada escola

# ==============================================================================
# 1. CONFIGURAÇÕES E MAPEAMENTOS
# ==============================================================================
caminho_template = Path(r"D:\Scripts_Python\FGV\Relatorios_Fluxo_Notas_por_Escola\template_notas_escolas.pptx")
pasta_raiz = Path(r"D:\Scripts_Python\FGV\Relatorios_Fluxo_Notas_por_Escola\relatorios_gerados")
data_dados = "25/11/2026"

mapa_componentes = {
    "Arte": "tabela_arte",
    "Biologia": "tabela_biologia",
    "Ciências": "tabela_ciencias",
    "Educação Física": "tabela_ed_fisica",
    "Filosofia": "tabela_filosofia",
    "Física": "tabela_fisica",
    "Geografia": "tabela_geografia",
    "História": "tabela_historia",
    "Língua Inglesa": "tabela_ingles",
    "Língua Portuguesa": "tabela_portugues",
    "Matemática": "tabela_matematica",
    "Química": "tabela_quimica",
    "Sociologia": "tabela_sociologia",
}

ordem_series = [
    "6º ANO",
    "7º ANO",
    "8º ANO",
    "9º ANO",
    "1ª SÉRIE",
    "2ª SÉRIE",
    "3ª SÉRIE",
]

# ==============================================================================
# 2. FUNÇÕES AUXILIARES DE MANIPULAÇÃO DO POWERPOINT
# ==============================================================================
def substituir_texto_placeholders(prs, substituicoes):
    """Substitui placeholders de texto preservando a formatação original."""
    for slide in prs.slides:
        for shp in slide.shapes:
            if not shp.has_text_frame:
                continue
            for p in shp.text_frame.paragraphs:
                texto_original = p.text
                texto_novo = texto_original
                for chave, valor in substituicoes.items():
                    if chave in texto_novo:
                        texto_novo = texto_novo.replace(chave, str(valor))
                if texto_novo != texto_original:
                    if p.runs:
                        p.runs[0].text = texto_novo
                        for run in p.runs[1:]:
                            run.text = ""
                    else:
                        p.text = texto_novo


def escrever_celula(cell, texto):
    """Escreve o valor na célula forçando a fonte Nunito tamanho 20."""
    tf = cell.text_frame
    p = tf.paragraphs[0]
    p.font.name = "Nunito"
    p.font.size = Pt(20)

    if p.runs:
        run = p.runs[0]
        run.text = texto
        for r_extra in p.runs[1:]:
            r_extra.text = ""
    else:
        run = p.add_run()
        run.text = texto

    run.font.name = "Nunito"
    run.font.size = Pt(20)

    # Remove parágrafos extras caso existam na célula
    for p_extra in list(tf.paragraphs)[1:]:
        p_extra._element.getparent().remove(p_extra._element)


def encontrar_tabela_no_slide(slide, nome_tabela):
    """Localiza o objeto de tabela pelo nome definido no Painel de Seleção."""
    for shp in slide.shapes:
        if shp.has_table and shp.name == nome_tabela:
            return shp.table
    return None


def preencher_tabela(table, df_bloco):
    """Preenche a partir da 1ª linha (índice 0) até o limite de 15 linhas e limpa as restantes."""
    max_linhas = min(15, len(table.rows))

    for idx_linha in range(max_linhas):
        row = table.rows[idx_linha]

        if idx_linha < len(df_bloco):
            registro = df_bloco.iloc[idx_linha]

            # Coluna 0: NOME PESSOA (máximo de 30 caracteres)
            nome = "" if pd.isna(registro["NOME PESSOA"]) else str(registro["NOME PESSOA"]).strip()[:30]

            # Coluna 1: MATRÍCULA ESTUDANTE
            if pd.isna(registro["MATRÍCULA ESTUDANTE"]):
                matricula = ""
            else:
                matricula = re.sub(r"\.0$", "", str(registro["MATRÍCULA ESTUDANTE"]).strip())

            # Coluna 2: SÉRIE
            serie = "" if pd.isna(registro["SÉRIE"]) else str(registro["SÉRIE"]).strip()

            # Coluna 3: MEDIA_NOTAS (2 casas decimais)
            if pd.isna(registro["MEDIA_NOTAS"]):
                media = ""
            else:
                valor_num = round(float(registro["MEDIA_NOTAS"]), 2)
                media = f"{valor_num:.2f}".replace(".", ",")

            valores = [nome, matricula, serie, media]
        else:
            valores = ["", "", "", ""]

        for idx_col, val in enumerate(valores):
            escrever_celula(row.cells[idx_col], val)


def duplicar_slide(prs, slide_origem):
    """Cria uma cópia exata do slide_origem e a insere imediatamente após ele."""
    slide_novo = prs.slides.add_slide(slide_origem.slide_layout)

    for shp in list(slide_novo.shapes):
        sp = shp._element
        sp.getparent().remove(sp)

    mapa_rids = {}
    for rId, rel in slide_origem.part.rels.items():
        if "notesSlide" not in rel.reltype and "slideLayout" not in rel.reltype:
            novo_rId = slide_novo.part.rels.get_or_add(rel.reltype, rel._target)
            mapa_rids[rId] = novo_rId

    for shp in slide_origem.shapes:
        el_copia = copy.deepcopy(shp._element)
        if mapa_rids:
            xml_str = el_copia.xml
            for antigo_id, novo_id in mapa_rids.items():
                if antigo_id != novo_id:
                    xml_str = xml_str.replace(f'"{antigo_id}"', f'"{novo_id}"')
            el_copia = parse_xml(xml_str)
        slide_novo.shapes._spTree.insert_element_before(el_copia, "p:extLst")

    sldIdLst = prs.slides._sldIdLst
    lista_ids = list(sldIdLst)
    idx_origem = prs.slides.index(slide_origem)
    novo_sldId = lista_ids[-1]
    sldIdLst.remove(novo_sldId)
    sldIdLst.insert(idx_origem + 1, novo_sldId)

    return slide_novo


# ==============================================================================
# 3. GERAÇÃO DAS APRESENTAÇÕES POR ESCOLA
# ==============================================================================
total_apresentacoes = 0

for (direc, inep_str), df_escola_todos in grupos_todos.items():
    if pd.isna(direc) or str(direc).strip() == "":
        nome_pasta_direc = "Sem_DIREC"
    else:
        nome_pasta_direc = re.sub(r'[\\/*?:"<>|]', "_", str(direc).strip())

    pasta_direc = pasta_raiz / nome_pasta_direc
    pasta_direc.mkdir(parents=True, exist_ok=True)

    series_escola = df_escola_todos["ESCOLA"].dropna()
    nome_escola = str(series_escola.iloc[0]).strip() if not series_escola.empty else f"Escola INEP {inep_str}"

    prs = Presentation(str(caminho_template))

    substituir_texto_placeholders(
        prs,
        {
            "{{nome_escola}}": nome_escola,
            "nome_escola": nome_escola,
            "{{data_dados}}": data_dados,
            "data_dados": data_dados,
        },
    )

    slides_base_tabelas = {}
    for slide in prs.slides:
        for shp in slide.shapes:
            if shp.has_table and shp.name in mapa_componentes.values():
                slides_base_tabelas[shp.name] = slide

    df_escola_rep = grupos_reprovados.get((direc, inep_str), df_reprovados_vazio).copy()
    if not df_escola_rep.empty:
        df_escola_rep["SERIE_ORD"] = pd.Categorical(
            df_escola_rep["SÉRIE"].astype(str).str.strip(),
            categories=ordem_series,
            ordered=True,
        )
        df_escola_rep["MEDIA_NOTAS"] = pd.to_numeric(df_escola_rep["MEDIA_NOTAS"], errors="coerce")

    for componente, nome_tabela in mapa_componentes.items():
        slide_origem = slides_base_tabelas.get(nome_tabela)
        if slide_origem is None:
            print(f"Aviso: tabela '{nome_tabela}' não encontrada no template.")
            continue

        tabela_origem = encontrar_tabela_no_slide(slide_origem, nome_tabela)
        max_linhas = min(15, len(tabela_origem.rows))  # 15 linhas começando do índice 0

        if df_escola_rep.empty:
            df_comp = df_escola_rep
        else:
            df_comp = (
                df_escola_rep[df_escola_rep["COMPONENTE CURRICULAR"].astype(str).str.strip() == componente]
                .sort_values(by=["SERIE_ORD", "MEDIA_NOTAS"], ascending=[True, True], na_position="last")
            )

        if len(df_comp) == 0:
            blocos = [df_comp]
        else:
            blocos = [df_comp.iloc[i : i + max_linhas] for i in range(0, len(df_comp), max_linhas)]

        slide_atual = slide_origem
        for idx_bloco, df_bloco in enumerate(blocos):
            if idx_bloco > 0:
                slide_atual = duplicar_slide(prs, slide_atual)
            tabela_atual = encontrar_tabela_no_slide(slide_atual, nome_tabela)
            preencher_tabela(tabela_atual, df_bloco)

    caminho_saida_pptx = pasta_direc / f"{inep_str}_notas_escolas.pptx"
    prs.save(str(caminho_saida_pptx))
    total_apresentacoes += 1

print(f"Concluído: {total_apresentacoes} apresentações geradas em '{pasta_raiz}'.")
























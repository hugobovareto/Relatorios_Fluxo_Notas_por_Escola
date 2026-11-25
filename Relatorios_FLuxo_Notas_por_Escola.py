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

O Relatório de Notas está segmentado por componentes, então nesse, cria a variável:
- nota média.

Para notas, considerar somente as notas do 1º, 2º e 3º bimestres.
Caso o estudante não tenha nota em algum bimestre, considerar como NaN, logo não impacta na média.
    Se as duas colunas têm valores → média das duas.
    Se apenas uma tem valor → retorna esse valor.
    Se ambas são NaN → retorna NaN.

Para valores duplicatos 
    

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

# Criar coluna com nota final média, considerando as notas do 1º, 2º e 3º bimestres:
# (ignora os valores NaN e fazem a média somente com os valores presentes. Se só tiver 1 nota disponível, a média será essa nota)
'''
Se as duas colunas têm valores → média das duas.
Se apenas uma tem valor → retorna esse valor.
Se ambas são NaN → retorna NaN.
'''
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
# Excluir duplicatas e manter somente o CPF + SÉRIE + COMPONENTE CURRICULAR com menor MEDIA_NOTAS
chaves = ['CPF PESSOA', 'COMPONENTE CURRICULAR', 'SÉRIE']

# Garante que MEDIA_NOTAS seja numérica (evita ordenação alfabética onde "10" vem antes de "2")
df_EF_EM_bncc['MEDIA_NOTAS'] = pd.to_numeric(df_EF_EM_bncc['MEDIA_NOTAS'], errors='coerce')

df_sem_duplicatas = (
    df_EF_EM_bncc
    .sort_values(by='MEDIA_NOTAS', ascending=True, na_position='last')
    .drop_duplicates(subset=chaves, keep='first')
    .sort_index()  # opcional: restaura a ordem original das linhas
)


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

















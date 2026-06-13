import pandas as pd
from datetime import datetime
import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import sys
from pathlib import Path

if getattr(sys, 'frozen', False):
    pasta_base = Path(sys.executable).parent
else:
    pasta_base = Path(__file__).parent

arquivo_origem = pasta_base / "baseorigem.xlsx"
arquivo_saida = pasta_base / "base_formatada.xlsx"

aba_origem = None
data_hoje = datetime.today().strftime("%d/%m/%Y")

df_origem = pd.read_excel(arquivo_origem, sheet_name=aba_origem, engine="openpyxl")

if isinstance(df_origem, dict):
    primeira_aba = list(df_origem.keys())[0]
    df_origem = df_origem[primeira_aba]

df_origem.columns = df_origem.columns.astype(str).str.strip()

colunas_normalizadas = {str(col).strip().lower(): col for col in df_origem.columns}

def buscar_coluna(*nomes_possiveis):
    for nome in nomes_possiveis:
        chave = str(nome).strip().lower()
        if chave in colunas_normalizadas:
            return colunas_normalizadas[chave]
    return None

def puxar_valor(linha, *nomes_possiveis):
    coluna = buscar_coluna(*nomes_possiveis)
    if coluna is None:
        return ""
    valor = linha[coluna]
    if pd.isna(valor):
        return ""
    return valor

def formatar_data_brasileira(valor):
    if pd.isna(valor) or valor == "":
        return ""
    data = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.isna(data):
        return ""
    return data.strftime("%d/%m/%Y")

def limpar_nome_aba(nome):
    nome = str(nome).strip()
    if nome == "" or nome.lower() == "nan":
        nome = "Sem Jornada"
    nome = re.sub(r'[\[\]\:\*\?\/\\]', '-', nome)
    nome = nome[:31]
    return nome

def tratar_origem(valor):
    if pd.isna(valor) or valor == "":
        return ""
    valor = str(valor).strip()
    regras = {
        "Amil 360": "Captado pela Amil/Call Center",
        "Forms Eric (Desosp Total Care)": "Forms"
    }
    return regras.get(valor, valor)

def posicoes_por_jornada(jornada):
    jornada_tratada = str(jornada).strip().lower()
    if jornada_tratada == "anticoagulante seguro":
        return 38, 39
    elif jornada_tratada == "emagrecimento":
        return 42, 43
    else:
        return 36, 37

col_jornada = buscar_coluna("Jornada")

if col_jornada is None:
    raise Exception("A coluna 'Jornada' não foi encontrada na planilha de origem.")

df_origem["_Jornada_Tratada"] = df_origem[col_jornada].fillna("Sem Jornada").astype(str).str.strip()
df_origem.loc[df_origem["_Jornada_Tratada"] == "", "_Jornada_Tratada"] = "Sem Jornada"

wb = Workbook()
ws_inicial = wb.active
wb.remove(ws_inicial)

abas_criadas = set()

for jornada in df_origem["_Jornada_Tratada"].unique():
    df_filtrado = df_origem[df_origem["_Jornada_Tratada"] == jornada].copy()
    nome_aba = limpar_nome_aba(jornada)
    nome_aba_original = nome_aba
    contador = 1

    while nome_aba in abas_criadas:
        sufixo = f"_{contador}"
        nome_aba = nome_aba_original[:31 - len(sufixo)] + sufixo
        contador += 1

    abas_criadas.add(nome_aba)
    ws = wb.create_sheet(title=nome_aba)

    col_flag, col_encaminhado = posicoes_por_jornada(jornada)
    col_final_inicio = col_encaminhado + 1

    cabecalhos_fixos = {
        1: "Prioridade no Contato?",
        2: "Data Envio Paciente",
        3: "Motivo Envio",
        4: "Origem",
        5: "MO",
        6: "Nome do paciente",
        7: "CPF",
        8: "Data de Nascimento",
        9: "Idade",
        10: "Genero",
        11: "Estado",
        12: "Município",
        13: "Bairro",
        14: "Nível de Rede",
        15: "Telefone atualizado",
        col_flag: "Flag",
        col_encaminhado: "Encaminhado p/ Consolidado?"
    }

    cabecalhos_finais = [
        "Linha",
        "Mês/Ano",
        "Semana",
        "Mês/Ano Inclusão",
        "Semana Inclusão",
        "SLA Captação",
        "SLA Exclusão",
        "SLA Tent. Contatos",
        "SLA Consulta",
        "Ativo/Inativo"
    ]

    for coluna, cabecalho in cabecalhos_fixos.items():
        ws.cell(row=1, column=coluna, value=cabecalho)

    for i, cabecalho in enumerate(cabecalhos_finais):
        ws.cell(row=1, column=col_final_inicio + i, value=cabecalho)

    linha_excel = 2

    for _, linha in df_filtrado.iterrows():
        ws.cell(row=linha_excel, column=1, value="Sim")
        ws.cell(row=linha_excel, column=2, value=data_hoje)
        ws.cell(row=linha_excel, column=3, value="Navegação")
        ws.cell(row=linha_excel, column=4, value=tratar_origem(puxar_valor(linha, "Origem")))
        ws.cell(row=linha_excel, column=5, value=puxar_valor(linha, "MO"))
        ws.cell(row=linha_excel, column=6, value=puxar_valor(linha, "Nome do paciente", "Nome Paciente", "Paciente"))
        ws.cell(row=linha_excel, column=7, value=puxar_valor(linha, "CPF"))
        ws.cell(row=linha_excel, column=8, value=formatar_data_brasileira(puxar_valor(linha, "Data de Nascimento", "Nascimento")))
        ws.cell(row=linha_excel, column=9, value="")
        ws.cell(row=linha_excel, column=10, value=puxar_valor(linha, "Genero", "Gênero"))
        ws.cell(row=linha_excel, column=11, value=puxar_valor(linha, "Estado", "UF"))
        ws.cell(row=linha_excel, column=12, value=puxar_valor(linha, "Município", "Municipio", "Cidade"))
        ws.cell(row=linha_excel, column=13, value=puxar_valor(linha, "Bairro"))
        ws.cell(row=linha_excel, column=14, value=puxar_valor(linha, "Nível de Rede", "Nivel de Rede"))
        ws.cell(row=linha_excel, column=15, value=puxar_valor(linha, "Telefone atualizado", "Telefone", "Celular"))
        ws.cell(row=linha_excel, column=col_flag, value="Sim")
        ws.cell(row=linha_excel, column=col_encaminhado, value="Sim")
        linha_excel += 1

    for coluna in range(1, ws.max_column + 1):
        letra = get_column_letter(coluna)
        ws.column_dimensions[letra].width = 22

    ws.freeze_panes = "A2"

wb.save(arquivo_saida)

print(f"Arquivo gerado com sucesso: {arquivo_saida}")
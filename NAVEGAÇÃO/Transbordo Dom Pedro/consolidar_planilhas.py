import glob
import os
import re
import sys
import unicodedata
import warnings
from datetime import datetime
import pandas as pd

NOME_PASTA_INPUT = "planilhas"
ARQUIVO_OUTPUT = "Planilha_Consolidada.xlsx"

CONFIGURACAO_NOME = "1 - Aceitou e não foi encaminhado"
NOVA_ABA_NOME = "Acompanhamento Atual"

JORNADAS_POR_PALAVRA_CHAVE = {
    "antigoaculante": "Anticoagulante Seguro",
    "anticoagulante": "Anticoagulante Seguro",
    "arritmia": "Ritmo Certo",
    "cardio ic": "Insuficiência Cardíaca Controlada",
    "avc": "Pós AVC",
    "iam": "Cuidados Pós Infarto",
    "valvulopatia": "Cuidado Cardíaco Valvar",
    "coluna": "Saúde da Coluna",
    "emagrecimento": "Emagrecimento",
    "nefropatia": "Saúde Renal",
    "renal": "Saúde Renal",
    "mama": "Cuidado Integral da Mama",
    "colorretal": "Cuidado Oncológico Colorretal",
    "coloretal": "Cuidado Oncológico Colorretal",
    "colo retal": "Cuidado Oncológico Colorretal",
    "prostata": "Cuidado Oncológico Próstata",
    "pulmonar": "Cuidado Oncológico Pulmonar",
    "pulmao": "Cuidado Oncológico Pulmonar",
    "dpoc": "DPOC",
    "asma": "ASMA",
    "diabetes": "DIABETES",
    "autoimunidades": "Autoimunidades",
    "hepatico": "Pós transplante hepático",
    "tmo": "TMO",
    "cefaleia": "Cefaléia",
}

warnings.filterwarnings("ignore")
pd.options.mode.chained_assignment = None


def normalizar_texto(valor):
    texto = str(valor).lower().strip()
    texto = (
        unicodedata.normalize("NFKD", texto)
        .encode("ASCII", "ignore")
        .decode("ASCII")
    )
    texto = texto.replace("-", " ").replace("_", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def formatar_data(serie):
    if serie is None or isinstance(serie, str):
        return ""
    return (
        pd.to_datetime(serie, errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
    )


def limpar_nome_sheet(nome):
    nome = str(nome).strip()
    nome = re.sub(r"[\[\]\:\*\?\/\\]", "-", nome)
    nome = nome[:31]
    return nome if nome else "Sem Jornada"


def identificar_jornada(nome_arquivo):
    nome_normalizado = normalizar_texto(nome_arquivo)
    palavras_arquivo = nome_normalizado.split()

    if "ic" in palavras_arquivo:
        return "Insuficiência Cardíaca Controlada"

    for palavra_chave, jornada in JORNADAS_POR_PALAVRA_CHAVE.items():
        if normalizar_texto(palavra_chave) in nome_normalizado:
            return jornada

    return "Jornada não identificada"


def mapear_origem(motivo):
    if pd.isna(motivo):
        return ""
    val = str(motivo).lower().strip()
    if "capta" in val:
        return "Captação Dom Pedro"
    if "navega" in val:
        return "Transbordo Dom Pedro"
    return ""


def processar_flag(val):
    if pd.isna(val):
        return ""
    texto = str(val).strip()
    return "Ativo" if texto.lower() == "sim" else texto


def processar_mo(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    texto = str(val).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto.zfill(9)


# Mantém o formato original completo para a primeira aba (com filtro)
def preparar_formato_consolidado(df_filtrado, jornada_encontrada, data_hoje):
    col_motivo_nome = None
    for col in df_filtrado.columns:
        if "motivo" in str(col).lower():
            col_motivo_nome = col
            break

    if col_motivo_nome:
        origem_col = df_filtrado[col_motivo_nome].apply(mapear_origem)
    else:
        origem_col = pd.Series([""] * len(df_filtrado), index=df_filtrado.index)

    novo_df = pd.DataFrame(index=df_filtrado.index)
    novo_df["Origem"] = origem_col
    novo_df["Jornada"] = jornada_encontrada
    novo_df["Flag"] = (
        df_filtrado["Flag"].apply(processar_flag)
        if "Flag" in df_filtrado.columns
        else ""
    )
    novo_df["Data Input"] = data_hoje
    novo_df["MO"] = (
        df_filtrado["MO"].apply(processar_mo)
        if "MO" in df_filtrado.columns
        else ""
    )
    novo_df["Nome do paciente"] = df_filtrado.get("Nome do paciente", "")
    novo_df["CPF"] = df_filtrado.get("CPF", "")
    novo_df["Data de Nascimento"] = formatar_data(
        df_filtrado.get("Data de Nascimento")
    )
    novo_df["Idade"] = ""
    novo_df["Genero"] = df_filtrado.get("Genero", "")
    novo_df["Estado"] = df_filtrado.get("Estado", "")
    novo_df["Município"] = df_filtrado.get("Município", "")
    novo_df["Bairro"] = df_filtrado.get("Bairro", "")
    novo_df["Nível de Rede"] = ""
    novo_df["Telefone atualizado"] = df_filtrado.get("Telefone atualizado", "")
    novo_df["Hub nvg"] = "CMDP"
    novo_df["Enviado nvg?"] = "Sim"
    novo_df["data envio ncg"] = formatar_data(
        df_filtrado.get("Data Envio Paciente")
    )
    novo_df["Status nvg"] = ""
    novo_df["Data Atualização Status nvg"] = ""
    novo_df["Data de Entrada na Linha"] = formatar_data(
        df_filtrado.get("Data de Inclusão")
    )
    novo_df["Data da Consulta"] = ""
    novo_df["Porta de Entrada"] = ""

    return novo_df


# Nova aba sem filtro (usa o dataframe original completo)
def preparar_aba_resumida_sem_filtro(df_completo, jornada_encontrada):
    novo_df = pd.DataFrame(index=df_completo.index)
    
    novo_df["Jornada"] = jornada_encontrada
    novo_df["MO"] = (
        df_completo["MO"].apply(processar_mo)
        if "MO" in df_completo.columns
        else ""
    )
    novo_df["Nome"] = df_completo.get("Nome do paciente", "")
    novo_df["Local de Acomapanhamento atual"] = df_completo.get("Local de Acomapanhamento atual", "")
    
    return novo_df


def consolidar_planilhas(pasta_origem, caminho_destino):
    arquivos_excel = glob.glob(os.path.join(pasta_origem, "*.xlsx"))
    data_hoje = datetime.now().strftime("%d/%m/%Y")

    resultados_completos = []
    resultados_resumidos = []
    logs_contagem = {}

    total_arquivos = len(arquivos_excel)
    if not arquivos_excel:
        print(f"Nenhum arquivo .xlsx encontrado na pasta: {pasta_origem}")
        return

    print(f"Encontrados {total_arquivos} arquivos para processar.\n")

    for idx, arquivo in enumerate(arquivos_excel, start=1):
        nome_arquivo = os.path.basename(arquivo)
        print(f"[{idx}/{total_arquivos}] Lendo: {nome_arquivo}...", end="", flush=True)

        try:
            df = pd.read_excel(arquivo, sheet_name=0, engine="openpyxl")
        except Exception as e:
            print(f" -> Erro ao ler: {e}")
            continue

        jornada_encontrada = identificar_jornada(nome_arquivo)

        # Filtros apenas para a aba 1
        if "Status BNF Linha" in df.columns:
            filtro_status = (
                df["Status BNF Linha"].fillna("").astype(str).str.strip()
                == "1. Aceitou o programa"
            )
        else:
            filtro_status = pd.Series([False] * len(df), index=df.index)

        if "Encaminhado p/ Consolidado?" in df.columns:
            filtro_encaminhado_vazio = (
                df["Encaminhado p/ Consolidado?"].fillna("").astype(str).str.strip()
                == ""
            )
        else:
            filtro_encaminhado_vazio = pd.Series([True] * len(df), index=df.index)

        filtro_final = filtro_status & filtro_encaminhado_vazio
        df_filtrado = df[filtro_final].copy()

        # 1. Processa dados para a primeira aba (com o filtro aplicado)
        if not df_filtrado.empty:
            df_comp = preparar_formato_consolidado(df_filtrado, jornada_encontrada, data_hoje)
            resultados_completos.append(df_comp)
            logs_contagem[jornada_encontrada] = (
                logs_contagem.get(jornada_encontrada, 0) + len(df_filtrado)
            )

        # 2. Processa dados para a segunda aba (SEM FILTRO - traz todo mundo da planilha)
        if not df.empty:
            df_res = preparar_aba_resumida_sem_filtro(df, jornada_encontrada)
            resultados_resumidos.append(df_res)

        print(f" Concluído! ({len(df)} linhas totais lidas)")

    if not resultados_resumidos:
        print("\nNenhum dado válido foi encontrado para gerar o arquivo.")
        return

    print("\nGravando dados no arquivo final...")
    
    nome_sheet_completa = limpar_nome_sheet(CONFIGURACAO_NOME)
    nome_sheet_resumida = limpar_nome_sheet(NOVA_ABA_NOME)
    
    with pd.ExcelWriter(caminho_destino, engine="openpyxl") as writer:
        if resultados_completos:
            df_final_completo = pd.concat(resultados_completos, ignore_index=True)
            df_final_completo.to_excel(writer, sheet_name=nome_sheet_completa, index=False)
        else:
            # Cria aba vazia caso ninguém se encaixe no filtro 1
            pd.DataFrame().to_excel(writer, sheet_name=nome_sheet_completa, index=False)
            
        df_final_resumido = pd.concat(resultados_resumidos, ignore_index=True)
        df_final_resumido.to_excel(writer, sheet_name=nome_sheet_resumida, index=False)

    print(f"\nSucesso! Arquivo consolidado salvo em: {caminho_destino}")
    print(f"\n--- Resumo de extração da aba com filtros ({CONFIGURACAO_NOME}) ---")
    if logs_contagem:
        for jornada, quantidade in sorted(logs_contagem.items()):
            print(f"  {jornada}: {quantidade} registros")
    else:
        print("  Nenhum registro encontrado nesta regra.")
    print("-" * 35)


if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        diretorio_atual = os.path.dirname(sys.executable)
    else:
        diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    pasta_input = os.path.join(diretorio_atual, NOME_PASTA_INPUT)
    arquivo_output = os.path.join(diretorio_atual, ARQUIVO_OUTPUT)

    os.makedirs(pasta_input, exist_ok=True)
    consolidar_planilhas(pasta_input, arquivo_output)

    input("\nPressione Enter para fechar a janela...")

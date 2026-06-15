import pandas as pd
import glob
import os
import sys

# =========================
# CONFIGURAÇÕES
# =========================

NOME_PASTA_ENTRADA = "planilhas"
NOME_ARQUIVO_SAIDA = "Planilha_Elisa_Status_5_6.xlsx"

COLUNA_STATUS = "Status BNF Linha"
COLUNA_MO = "MO"
COLUNA_NOME = "Nome do paciente"
COLUNA_CPF = "CPF"

STATUS_DESEJADOS = ["5. Alta por múltiplas tentativas", "6. Contato sem sucesso"]

# =========================
# FUNÇÕES
# =========================

def processar_mo(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    texto = str(val).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto.zfill(9)

def consolidar_planilhas(pasta_origem, caminho_destino):
    arquivos_excel = glob.glob(os.path.join(pasta_origem, "*.xlsx"))
    dataframes_processados = []
    log_jornadas = {}

    for arquivo in arquivos_excel:
        nome_arquivo = os.path.basename(arquivo)

        try:
            df = pd.read_excel(arquivo, sheet_name=0)
        except Exception as e:
            print(f"Erro ao ler {nome_arquivo}: {e}")
            continue

        if COLUNA_STATUS not in df.columns:
            print(f"[{nome_arquivo}] Ignorado: coluna '{COLUNA_STATUS}' não encontrada.")
            continue

        filtro_status = df[COLUNA_STATUS].fillna("").astype(str).str.strip().isin(STATUS_DESEJADOS)
        df_filtrado = df[filtro_status].copy()

        print(f"\n[{nome_arquivo}]")
        print(f"Total de linhas: {len(df)}")
        print(f"Status 5 ou 6 encontrados: {len(df_filtrado)}")

        if df_filtrado.empty:
            print(f"[{nome_arquivo}] Nenhum registro passou no filtro.")
            continue

        nome_lower = nome_arquivo.lower()
        jornada_encontrada = ""

        if "antigoaculante" in nome_lower: jornada_encontrada = "Anticoagulante Seguro"
        elif "arritmia" in nome_lower: jornada_encontrada = "Ritmo Certo"
        elif "ic" in nome_lower.replace("-", " ").split() or "cardio ic" in nome_lower: jornada_encontrada = "Insuficiência Cardíaca Controlada"
        elif "avc" in nome_lower: jornada_encontrada = "Pós AVC"
        elif "iam" in nome_lower: jornada_encontrada = "Cuidados Pós Infarto"
        elif "valvulopatia" in nome_lower: jornada_encontrada = "Cuidado Cardíaco Valvar"
        elif "coluna" in nome_lower: jornada_encontrada = "Saúde da Coluna"
        elif "emagrecimento" in nome_lower: jornada_encontrada = "Emagrecimento"
        elif "nefropatia" in nome_lower or "renal" in nome_lower: jornada_encontrada = "Saúde Renal"
        elif "mama" in nome_lower: jornada_encontrada = "Cuidado Integral da Mama"
        elif "colorretal" in nome_lower or "coloretal" in nome_lower or "colo retal" in nome_lower: jornada_encontrada = "Cuidado Oncológico Colorretal"
        elif "prostata" in nome_lower or "próstata" in nome_lower: jornada_encontrada = "Cuidado Oncológico Próstata"
        elif "pulmonar" in nome_lower or "pulmao" in nome_lower or "pulmão" in nome_lower: jornada_encontrada = "Cuidado Oncológico Pulmonar"

        if jornada_encontrada == "":
            print(f"[{nome_arquivo}] Aviso: jornada não identificada pelo nome do arquivo.")

        if COLUNA_MO not in df_filtrado.columns:
            print(f"[{nome_arquivo}] Aviso: coluna '{COLUNA_MO}' não encontrada.")
        if COLUNA_NOME not in df_filtrado.columns:
            print(f"[{nome_arquivo}] Aviso: coluna '{COLUNA_NOME}' não encontrada.")
        if COLUNA_CPF not in df_filtrado.columns:
            print(f"[{nome_arquivo}] Aviso: coluna '{COLUNA_CPF}' não encontrada.")

        novo_df = pd.DataFrame()
        novo_df["Jornada"] = [jornada_encontrada] * len(df_filtrado)
        novo_df["MO"] = df_filtrado[COLUNA_MO].apply(processar_mo).values if COLUNA_MO in df_filtrado.columns else ""
        novo_df["Nome"] = df_filtrado[COLUNA_NOME].values if COLUNA_NOME in df_filtrado.columns else ""
        novo_df["CPF"] = df_filtrado[COLUNA_CPF].values if COLUNA_CPF in df_filtrado.columns else ""

        dataframes_processados.append(novo_df)
        log_jornadas[jornada_encontrada if jornada_encontrada else "Jornada não identificada"] = log_jornadas.get(jornada_encontrada if jornada_encontrada else "Jornada não identificada", 0) + len(novo_df)

    if dataframes_processados:
        df_final = pd.concat(dataframes_processados, ignore_index=True)
        df_final.to_excel(caminho_destino, index=False)

        print(f"\nSucesso! Arquivo salvo em: {caminho_destino}")
        print("\n--- Resumo por Jornada ---")
        for jornada, quantidade in sorted(log_jornadas.items()):
            print(f"{jornada}: {quantidade} registros")
        print("-" * 30)
    else:
        print("\nNenhum dado válido para consolidar.")

# =========================
# EXECUÇÃO
# =========================

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))

    pasta_input = os.path.join(diretorio_script, NOME_PASTA_ENTRADA)
    arquivo_output = os.path.join(diretorio_script, NOME_ARQUIVO_SAIDA)

    os.makedirs(pasta_input, exist_ok=True)

    print(f"Lendo arquivos da pasta '{pasta_input}'...")
    consolidar_planilhas(pasta_input, arquivo_output)

    input("\nPressione Enter para fechar a janela...")
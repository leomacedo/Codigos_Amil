import pandas as pd
import glob
import os
import sys
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================
# FUNÇÃO PARA TRATAR MO
# =========================
def processar_mo(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    texto = str(val).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto.zfill(9)

# =========================
# FUNÇÃO PARA IDENTIFICAR PROGRAMA PELO NOME DO ARQUIVO
# =========================
def identificar_programa(nome_arquivo):
    nome_lower = nome_arquivo.lower()

    if "antigoaculante" in nome_lower or "anticoagulante" in nome_lower:
        return "Anticoagulante Seguro"

    elif "arritmia" in nome_lower:
        return "Ritmo Certo"

    elif "ic" in nome_lower.replace("-", " ").split() or "cardio ic" in nome_lower:
        return "Insuficiência Cardíaca Controlada"

    elif "avc" in nome_lower:
        return "Pós AVC"

    elif "iam" in nome_lower:
        return "Cuidados Pós Infarto"

    elif "valvulopatia" in nome_lower:
        return "Cuidado Cardíaco Valvar"

    elif "coluna" in nome_lower:
        return "Saúde da Coluna"

    elif "emagrecimento" in nome_lower:
        return "Emagrecimento"

    elif "nefropatia" in nome_lower or "renal" in nome_lower:
        return "Saúde Renal"
    elif "mama" in nome_lower:
        return "Cuidado Integral da Mama"

    elif "colorretal" in nome_lower or "coloretal" in nome_lower or "colo retal" in nome_lower:
        return "Cuidado Oncológico Colorretal"

    elif "prostata" in nome_lower or "próstata" in nome_lower:
        return "Cuidado Oncológico Próstata"

    elif "pulmonar" in nome_lower or "pulmao" in nome_lower or "pulmão" in nome_lower:
        return "Cuidado Oncológico Pulmonar"

    return ""

# =========================
# FUNÇÃO PRINCIPAL DE CONSOLIDAÇÃO
# =========================
def consolidar_planilhas(pasta_origem, caminho_destino, tipo_processamento):
    arquivos_excel = glob.glob(os.path.join(pasta_origem, "*.xlsx"))
    dataframes_processados = []
    log_programas = {}

    for arquivo in arquivos_excel:
        nome_arquivo = os.path.basename(arquivo)

        # Evita ler o próprio arquivo de saída, caso esteja na pasta
        if nome_arquivo.lower() == os.path.basename(caminho_destino).lower():
            continue

        # =========================
        # LEITURA DA PLANILHA
        # =========================
        try:
            df = pd.read_excel(arquivo, sheet_name=0, engine="openpyxl")
        except Exception as e:
            print(f"Erro ao ler {nome_arquivo}: {e}")
            continue

        df.columns = df.columns.astype(str).str.strip()
        programa_encontrado = identificar_programa(nome_arquivo)

        print(f"\n[{nome_arquivo}]")
        print(f"Total de linhas: {len(df)}")

        # =========================
        # REGRA 1 - FLAG
        # Status BNF Linha = 1. Aceitou o programa
        # Flag vazia ou Pendente
        # =========================
        if tipo_processamento == "1":
            colunas_obrigatorias = ["Status BNF Linha", "Flag", "MO"]
            colunas_faltantes = [col for col in colunas_obrigatorias if col not in df.columns]

            if colunas_faltantes:
                print(f"Ignorado: colunas não encontradas: {', '.join(colunas_faltantes)}")
                continue

            filtro_status = df["Status BNF Linha"].fillna("").astype(str).str.strip() == "1. Aceitou o programa"
            filtro_flag = df["Flag"].isna() | (df["Flag"].fillna("").astype(str).str.strip() == "") | (df["Flag"].fillna("").astype(str).str.strip().str.lower() == "pendente")
            df_filtrado = df[filtro_status & filtro_flag].copy()

            print(f"Status BNF Linha = 1. Aceitou o programa: {filtro_status.sum()}")
            print(f"Flag vazia ou Pendente: {filtro_flag.sum()}")
            print(f"Passou nos dois filtros: {len(df_filtrado)}")

        # =========================
        # REGRA 2 - DESFLAG
        # Situação Atual Do Paciente = Fora Da Linha
        # Flag = Sim
        # =========================
        elif tipo_processamento == "2":
            colunas_obrigatorias = ["Situação Atual Do Paciente", "Flag", "MO"]
            colunas_faltantes = [col for col in colunas_obrigatorias if col not in df.columns]

            if colunas_faltantes:
                print(f"Ignorado: colunas não encontradas: {', '.join(colunas_faltantes)}")
                continue

            filtro_situacao = df["Situação Atual Do Paciente"].fillna("").astype(str).str.strip().str.lower() == "fora da linha"
            filtro_flag = df["Flag"].fillna("").astype(str).str.strip().str.lower() == "sim"
            df_filtrado = df[filtro_situacao & filtro_flag].copy()

            print(f"Situação Atual Do Paciente = Fora Da Linha: {filtro_situacao.sum()}")
            print(f"Flag = Sim: {filtro_flag.sum()}")
            print(f"Passou nos dois filtros: {len(df_filtrado)}")

        else:
            print("Opção inválida.")
            return

        if df_filtrado.empty:
            print(f"[{nome_arquivo}] Nenhum registro passou nos filtros.")
            continue


        # =========================
        # MONTA PLANILHA FINAL
        # Programa | pula | pula | MO
        # =========================
        novo_df = pd.DataFrame({
            "Programa": [programa_encontrado] * len(df_filtrado),
            "pula_1": [""] * len(df_filtrado),
            "pula_2": [""] * len(df_filtrado),
            "MO": df_filtrado["MO"].apply(processar_mo).values
        })


        dataframes_processados.append(novo_df)

        # =========================
        # LOG POR PROGRAMA
        # =========================
        if programa_encontrado:
            log_programas[programa_encontrado] = log_programas.get(programa_encontrado, 0) + len(novo_df)
        else:
            log_programas["Programa não identificado"] = log_programas.get("Programa não identificado", 0) + len(novo_df)

    # =========================
    # SALVAR ARQUIVO FINAL
    # =========================
    if dataframes_processados:
        df_final = pd.concat(dataframes_processados, ignore_index=True)
        df_final.to_excel(caminho_destino, index=False)

        print(f"\nSucesso! Arquivo salvo em: {caminho_destino}")
        print("\n--- Resumo por Programa ---")

        for programa, quantidade in sorted(log_programas.items()):
            print(f"{programa}: {quantidade} registros")

        print("-" * 35)
    else:
        print("\nNenhum dado válido para consolidar nas planilhas fornecidas.")

# =========================
# EXECUÇÃO DO SCRIPT
# =========================
if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))

    pasta_input = os.path.join(diretorio_script, "planilhas")
    os.makedirs(pasta_input, exist_ok=True)

    print("Escolha o tipo de processamento:")
    print("1 - FLAG")
    print("2 - DESFLAG")

    tipo_processamento = input("Digite 1 ou 2: ").strip()

    if tipo_processamento == "1":
        nome_saida = "Planilha_FLAG.xlsx"
    elif tipo_processamento == "2":
        nome_saida = "Planilha_DESFLAG.xlsx"
    else:
        print("Opção inválida. Digite apenas 1 ou 2.")
        input("\nPressione Enter para fechar a janela...")
        sys.exit()

    arquivo_output = os.path.join(diretorio_script, nome_saida)

    print(f"\nLendo arquivos da pasta '{pasta_input}'...")
    consolidar_planilhas(pasta_input, arquivo_output, tipo_processamento)

    input("\nPressione Enter para fechar a janela...")
import pandas as pd
import glob
import os
import sys
from datetime import datetime

def formatar_data(serie):
    if serie is None or (isinstance(serie, str) and serie == ""):
        return ""
    return pd.to_datetime(serie, errors='coerce').dt.strftime('%d/%m/%Y').fillna("")

def consolidar_planilhas(pasta_origem, caminho_destino):

    arquivos_excel = glob.glob(os.path.join(pasta_origem, "*.xlsx"))
    data_hoje = datetime.now().strftime("%d/%m/%Y")
    dataframes_processados = []
    log_jornadas = {} # Dicionário para armazenar a contagem por jornada

    for arquivo in arquivos_excel:
        nome_arquivo = os.path.basename(arquivo)

        try:
            df = pd.read_excel(arquivo, sheet_name=0)
        except Exception as e:
            print(f"Erro ao ler {nome_arquivo}: {e}")
            continue

        # Valida se as colunas necessárias para o filtro existem na planilha
        if "Status BNF Linha" not in df.columns or "Encaminhado p/ Consolidado?" not in df.columns:
            print(f"[{nome_arquivo}] Ignorado: colunas de filtro não encontradas.")
            continue

        # Aplica filtros com tratamento seguro para vazios ou nulos (NaN)
        filtro_status = df["Status BNF Linha"].fillna("").astype(str).str.strip() == "1. Aceitou o programa"
        filtro_encaminhado = df["Encaminhado p/ Consolidado?"].fillna("").astype(str).str.strip() == ""

        df_filtrado = df[filtro_status & filtro_encaminhado].copy()

        print(f"\n[{nome_arquivo}]")
        print(f"Total de linhas: {len(df)}")
        print(f"Aceitou o programa: {filtro_status.sum()}")
        print(f"Encaminhado vazio: {filtro_encaminhado.sum()}")
        print(f"Passou nos dois filtros: {len(df_filtrado)}")

        if df_filtrado.empty:
            print(f"[{nome_arquivo}] Nenhum registro passou nos filtros.")
            continue

        # Identifica a jornada buscando palavras-chave (trechos) no nome do arquivo
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


        # Tenta achar a coluna de motivo de envio (caça a palavra "motivo" na primeira linha)
        col_motivo_nome = None
        for col in df_filtrado.columns:
            if "motivo" in str(col).lower():
                col_motivo_nome = col
                break

        def mapear_origem(motivo):
            if pd.isna(motivo): return ""
            val = str(motivo).lower().strip()
            if "capta" in val:  # Pega automaticamente 'Captação' e 'Captação + Navegação'
                return "Captação Dom Pedro"
            elif "navega" in val:  # Pega automaticamente 'Navegação'
                return "Transbordo Dom Pedro"
            return ""

        if col_motivo_nome:
            origem_col = df_filtrado[col_motivo_nome].apply(mapear_origem)
        else:
            origem_col = ""

        def processar_flag(val):
            if pd.isna(val): return ""
            texto = str(val).strip()
            return "Ativo" if texto.lower() == "sim" else texto

        def processar_mo(val):
            if pd.isna(val) or str(val).strip() == "": return ""
            texto = str(val).strip()
            if texto.endswith(".0"): texto = texto[:-2]
            return texto.zfill(9)

        # Estrutura o novo DataFrame com base nos de-para solicitados
        novo_df = pd.DataFrame()
        novo_df["Origem"] = origem_col
        novo_df["Jornada"] = jornada_encontrada
        novo_df["Flag"] = df_filtrado["Flag"].apply(processar_flag) if "Flag" in df_filtrado.columns else ""
        novo_df["Data Input"] = data_hoje
        novo_df["MO"] = df_filtrado["MO"].apply(processar_mo) if "MO" in df_filtrado.columns else ""
        novo_df["Nome do paciente"] = df_filtrado.get("Nome do paciente", "")
        novo_df["CPF"] = df_filtrado.get("CPF", "")
        novo_df["Data de Nascimento"] = formatar_data(df_filtrado.get("Data de Nascimento"))
        novo_df["Idade"] = ""
        novo_df["Genero"] = df_filtrado.get("Genero", "")
        novo_df["Estado"] = df_filtrado.get("Estado", "")
        novo_df["Município"] = df_filtrado.get("Município", "")
        novo_df["Bairro"] = df_filtrado.get("Bairro", "")
        novo_df["Nível de Rede"] = ""
        novo_df["Telefone atualizado"] = df_filtrado.get("Telefone atualizado", "")
        novo_df["Hub nvg"] = "CMDP"
        novo_df["Enviado nvg?"] = "Sim"
        novo_df["data envio ncg"] = formatar_data(df_filtrado.get("Data Envio Paciente"))
        novo_df["Status nvg"] = ""
        novo_df["Data Atualização Status nvg"] = ""
        novo_df["Data de Entrada na Linha"] = formatar_data(df_filtrado.get("Data de Inclusão"))
        novo_df["Data da Consulta"] = ""
        novo_df["Porta de Entrada"] = ""

        dataframes_processados.append(novo_df)

        # Atualiza o log de contagem por jornada
        if jornada_encontrada:
            contagem_atual = log_jornadas.get(jornada_encontrada, 0)
            log_jornadas[jornada_encontrada] = contagem_atual + len(novo_df)

    if dataframes_processados:
        df_final = pd.concat(dataframes_processados, ignore_index=True)
        df_final.to_excel(caminho_destino, index=False)
        print(f"\nSucesso! Arquivo consolidado salvo em: {caminho_destino}")

        # Imprime o log de contagem
        print("\n--- Resumo por Linha de Cuidado ---")
        if log_jornadas:
            for jornada, quantidade in sorted(log_jornadas.items()):
                print(f"{jornada}: {quantidade} registros")
        print("-" * 35)
    else:
        print("\nNenhum dado válido para consolidar nas planilhas fornecidas.")

if __name__ == "__main__":
    # Garante que o código encontre a pasta exata onde este script está salvo
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))

    # Diretório onde suas planilhas devem ser colocadas
    pasta_input = os.path.join(diretorio_script, "planilhas")
    arquivo_output = os.path.join(diretorio_script, "Planilha_Consolidada.xlsx")

    # Cria a pasta caso não exista
    os.makedirs(pasta_input, exist_ok=True)

    print(f"Lendo arquivos da pasta '{pasta_input}'...")
    consolidar_planilhas(pasta_input, arquivo_output)

    input("\nPressione Enter para fechar a janela...")
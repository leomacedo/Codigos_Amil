import pandas as pd
import os
import glob
import sys
import re
from datetime import datetime

# Configurações principais
NOME_PASTA_RELATORIOS = "relatorios_entrada"
PREFIXO_ARQUIVO_TOTAL = "relatorioTOTAL"
PREFIXO_ARQUIVO_LIMPO = "relatorioLIMPO"
EXTENSAO_SAIDA = ".xlsx"
FORMATO_DATA_ARQUIVO = "%Y%m%d_%H%M"

# Configurações de leitura dos arquivos
SEPARADOR_CSV = ";"
ENCODING_CSV_PRINCIPAL = "utf-8-sig"
ENCODING_CSV_ALTERNATIVO = "latin1"

# Colunas usadas nas regras
COLUNA_AGENDOU_OUTRO = "Agendou em outro programa/linha de cuidado?"
COLUNA_PROGRAMA = "O paciente entrou em contato para mais informações/agendamento de qual Programa?"
COLUNA_QUAL_PROGRAMA = "Qual programa?"
COLUNA_DATA_CRIACAO = "Data de Criação"
COLUNA_MARCA_OTICA = "Marca Ótica"
COLUNA_CHAVE_PRIMARIA = "Chave Primária"

# Colunas criadas para auditoria
COLUNA_PROGRAMA_ORIGINAL = "PROGRAMA_ORIGINAL_ANTES_DUPLICAR"
COLUNA_COMPARACAO_PROGRAMAS = "COMPARACAO_PROGRAMA_X_QUAL_PROGRAMA"
COLUNA_LINHA_DUPLICADA = "LINHA_DUPLICADA_OUTRO_PROGRAMA"

# Normaliza textos para comparar sem diferença de espaço ou letra maiúscula/minúscula
def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = re.sub(r"\s+", " ", valor)
    return valor

# Duplica linhas quando o paciente agendou em outro programa diferente do principal
def preparar_relatorio_outro_programa(df):
    colunas_necessarias = [COLUNA_AGENDOU_OUTRO, COLUNA_PROGRAMA, COLUNA_QUAL_PROGRAMA]

    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            print(f"ATENÇÃO: A coluna '{coluna}' não foi encontrada.")
            print("A regra de duplicar linhas por outro programa será ignorada.")
            return df

    print("\nAplicando regra de outro programa/linha de cuidado...")
    df = df.copy()

    # Cria colunas para acompanhar a regra aplicada
    df[COLUNA_PROGRAMA_ORIGINAL] = df[COLUNA_PROGRAMA]
    df[COLUNA_COMPARACAO_PROGRAMAS] = ""
    df[COLUNA_LINHA_DUPLICADA] = "NÃO"

    # Filtra somente linhas marcadas como SIM
    mascara_sim = df[COLUNA_AGENDOU_OUTRO].astype(str).str.strip().str.upper().eq("SIM")
    linhas_sim = df[mascara_sim].copy()

    if linhas_sim.empty:
        print("Nenhuma linha com 'Agendou em outro programa/linha de cuidado?' = SIM.")
        return df

    # Compara o programa principal com o outro programa informado
    programa_original_normalizado = linhas_sim[COLUNA_PROGRAMA].apply(normalizar_texto)
    qual_programa_normalizado = linhas_sim[COLUNA_QUAL_PROGRAMA].apply(normalizar_texto)
    linhas_sim[COLUNA_COMPARACAO_PROGRAMAS] = ["IGUAL" if prog == qual else "DIFERENTE" for prog, qual in zip(programa_original_normalizado, qual_programa_normalizado)]
    df.loc[linhas_sim.index, COLUNA_COMPARACAO_PROGRAMAS] = linhas_sim[COLUNA_COMPARACAO_PROGRAMAS]

    # Mantém apenas casos com programas diferentes e preenchidos
    programas_diferentes = linhas_sim[linhas_sim[COLUNA_COMPARACAO_PROGRAMAS] == "DIFERENTE"].copy()
    programas_diferentes = programas_diferentes[programas_diferentes[COLUNA_QUAL_PROGRAMA].notna() & (programas_diferentes[COLUNA_QUAL_PROGRAMA].astype(str).str.strip() != "")].copy()

    # Cria as linhas duplicadas com o novo programa
    duplicadas = programas_diferentes.copy()
    duplicadas[COLUNA_LINHA_DUPLICADA] = "SIM"
    duplicadas[COLUNA_PROGRAMA] = duplicadas[COLUNA_QUAL_PROGRAMA]
    df_expandido = pd.concat([df, duplicadas], ignore_index=True)

    print(f"Linhas com SIM encontradas: {len(linhas_sim)}")
    print(f"Linhas com programas diferentes: {len(programas_diferentes)}")
    print(f"Linhas duplicadas criadas: {len(duplicadas)}")

    return df_expandido

# Consolida os arquivos da pasta e gera os relatórios TOTAL e LIMPO
def consolidar_relatorios(pasta_entrada, caminho_saida_total, caminho_saida_limpo):
    print(f"Buscando arquivos na pasta: {pasta_entrada}")

    # Busca arquivos Excel e CSV na pasta
    arquivos_excel = glob.glob(os.path.join(pasta_entrada, "*.xlsx"))
    arquivos_csv = glob.glob(os.path.join(pasta_entrada, "*.csv"))
    todos_arquivos = arquivos_excel + arquivos_csv

    if not todos_arquivos:
        print("Nenhum arquivo encontrado. Verifique se os relatórios estão na pasta correta.")
        return

    lista_dataframes = []

    # Lê cada arquivo encontrado
    for arquivo in todos_arquivos:
        print(f"Lendo: {os.path.basename(arquivo)}")
        try:
            if arquivo.endswith(".csv"):
                try:
                    df = pd.read_csv(arquivo, sep=SEPARADOR_CSV, encoding=ENCODING_CSV_PRINCIPAL)
                except UnicodeDecodeError:
                    df = pd.read_csv(arquivo, sep=SEPARADOR_CSV, encoding=ENCODING_CSV_ALTERNATIVO)
            else:
                df = pd.read_excel(arquivo)

            df.columns = df.columns.str.strip()
            lista_dataframes.append(df)

        except Exception as e:
            print(f"Erro ao ler o arquivo {arquivo}: {e}")

    if lista_dataframes:
        print("\nEmpilhando todos os relatórios...")

        # Junta todos os relatórios em uma única base
        df_consolidado = pd.concat(lista_dataframes, ignore_index=True)
        print(f"Total de linhas após empilhar os relatórios: {len(df_consolidado)}")

        # Aplica a regra de duplicação por outro programa
        df_consolidado = preparar_relatorio_outro_programa(df_consolidado)
        print(f"Total de linhas após aplicar regra de outro programa: {len(df_consolidado)}")

        # Padroniza a coluna Data de Criação
        if COLUNA_DATA_CRIACAO in df_consolidado.columns:
            datas_texto = df_consolidado[COLUNA_DATA_CRIACAO].astype(str).str.split(",").str[0].str.split(" ").str[0].str.strip()
            df_consolidado[COLUNA_DATA_CRIACAO] = pd.to_datetime(datas_texto, errors="coerce", dayfirst=True).dt.date

        linhas_antes = len(df_consolidado)
        print(f"Total de linhas antes da remoção de duplicatas exatas: {linhas_antes}")

        # Remove apenas linhas completamente iguais
        df_total = df_consolidado.drop_duplicates().copy()
        duplicatas_exatas = linhas_antes - len(df_total)

        print(f"Total de duplicatas exatas removidas: {duplicatas_exatas}")
        print(f"Total de linhas no relatório TOTAL: {len(df_total)}")

        # Formata Marca Ótica com 9 dígitos
        if COLUNA_MARCA_OTICA in df_total.columns:
            df_total[COLUNA_MARCA_OTICA] = df_total[COLUNA_MARCA_OTICA].astype(object)
            df_total[COLUNA_MARCA_OTICA] = df_total[COLUNA_MARCA_OTICA].apply(lambda x: 0 if pd.isna(x) or str(x).strip().lower() in ["nan", "nat", ""] else x)
            df_total[COLUNA_MARCA_OTICA] = df_total[COLUNA_MARCA_OTICA].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(9)

        # Cria a Chave Primária usando Programa e Marca Ótica
        if COLUNA_MARCA_OTICA in df_total.columns and COLUNA_PROGRAMA in df_total.columns:
            df_total[COLUNA_CHAVE_PRIMARIA] = df_total[COLUNA_PROGRAMA].astype(str).str.strip().str.upper() + "_" + df_total[COLUNA_MARCA_OTICA]

        # Salva o relatório TOTAL
        try:
            df_total.to_excel(caminho_saida_total, index=False)
            print("-" * 50)
            print("SUCESSO! Relatório TOTAL gerado em:")
            print(caminho_saida_total)
            print("-" * 50)

        except Exception as e:
            print(f"Erro ao salvar o relatório TOTAL: {e}")

        # Cria uma cópia para gerar o relatório LIMPO
        df_limpo = df_total.copy()

        # Ordena pela data mais antiga
        if COLUNA_DATA_CRIACAO in df_limpo.columns:
            df_limpo = df_limpo.sort_values(by=COLUNA_DATA_CRIACAO, ascending=True)

        # Remove duplicadas pela Chave Primária mantendo a mais antiga
        if COLUNA_CHAVE_PRIMARIA in df_limpo.columns:
            linhas_antes_chave = len(df_limpo)
            df_limpo.drop_duplicates(subset=[COLUNA_CHAVE_PRIMARIA], keep="first", inplace=True)
            duplicatas_chave = linhas_antes_chave - len(df_limpo)
            print(f"Total de duplicatas de Chave Primária removidas no relatório LIMPO, mantendo a mais antiga: {duplicatas_chave}")

        print(f"Total de linhas no relatório LIMPO: {len(df_limpo)}")

        # Salva o relatório LIMPO
        try:
            df_limpo.to_excel(caminho_saida_limpo, index=False)
            print("-" * 50)
            print("SUCESSO! Relatório LIMPO gerado em:")
            print(caminho_saida_limpo)
            print("-" * 50)

        except Exception as e:
            print(f"Erro ao salvar o relatório LIMPO: {e}")

# Executa o script e define as pastas de entrada e saída
if __name__ == "__main__":
    # Define o diretório atual do script ou executável
    if getattr(sys, "frozen", False):
        diretorio_atual = os.path.dirname(sys.executable)
    else:
        diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    # Define a pasta onde os relatórios devem ficar
    pasta_relatorios = os.path.join(diretorio_atual, NOME_PASTA_RELATORIOS)

    if not os.path.exists(pasta_relatorios):
        os.makedirs(pasta_relatorios)
        print(f"Pasta '{pasta_relatorios}' criada! Coloque seus relatórios lá dentro e rode o script novamente.")

    else:
        # Cria os nomes dos arquivos finais com data e hora
        data_hoje = datetime.now().strftime(FORMATO_DATA_ARQUIVO)
        arquivo_saida_total = os.path.join(diretorio_atual, f"{PREFIXO_ARQUIVO_TOTAL}_{data_hoje}{EXTENSAO_SAIDA}")
        arquivo_saida_limpo = os.path.join(diretorio_atual, f"{PREFIXO_ARQUIVO_LIMPO}_{data_hoje}{EXTENSAO_SAIDA}")
        consolidar_relatorios(pasta_relatorios, arquivo_saida_total, arquivo_saida_limpo)

    input("\nPressione Enter para fechar...")
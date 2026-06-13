import pandas as pd
import os
import sys

def carregar_planilha(caminho_arquivo):
    """Função auxiliar para carregar arquivos .xlsx ou .csv."""
    if not os.path.exists(caminho_arquivo):
        print(f"ERRO: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return None
    
    print(f"Lendo arquivo: {caminho_arquivo}")
    try:
        if caminho_arquivo.endswith('.csv'):
            # Tenta ler com separador ';' e encoding 'utf-8-sig' primeiro
            try:
                df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8-sig', dtype=str)
            except:
                # Fallback para separador ',' e outros encodings
                df = pd.read_csv(caminho_arquivo, dtype=str)
        else:
            df = pd.read_excel(caminho_arquivo, dtype=str)
        
        # Limpa espaços nos nomes das colunas
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        print(f"ERRO: Falha ao ler o arquivo '{caminho_arquivo}'. Detalhes: {e}")
        return None

def realizar_procv_multiplo(caminho_entrada, caminho_base_dados, caminho_saida):
    """
    Realiza uma operação de VLOOKUP de múltiplas colunas entre duas planilhas.

    Args:
        caminho_entrada (str): Caminho para a planilha de entrada com CPFs.
        caminho_base_dados (str): Caminho para a planilha que serve como base de dados.
        caminho_saida (str): Caminho onde o arquivo de resultado será salvo.
    """
    
    # 1. Carregar os arquivos em DataFrames do pandas
    df_entrada = carregar_planilha(caminho_entrada)
    if df_entrada is None:
        return

    df_base = carregar_planilha(caminho_base_dados)
    if df_base is None:
        return

    # 2. Definir as colunas
    coluna_chave = 'nome_paciente'
    colunas_para_preencher = [
        'status_tratamento',
        'etapa_tratamento',
        'n_sessoes_sugerido',
        'ultimo_atendimento'
    ]
    
    # 3. Validar se as colunas necessárias existem nos arquivos
    if coluna_chave not in df_entrada.columns:
        print(f"ERRO: A coluna chave '{coluna_chave}' não foi encontrada no arquivo de entrada.")
        print(f"Colunas disponíveis: {df_entrada.columns.tolist()}")
        return
        
    if coluna_chave not in df_base.columns:
        print(f"ERRO: A coluna chave '{coluna_chave}' não foi encontrada na base de dados.")
        print(f"Colunas disponíveis: {df_base.columns.tolist()}")
        return

    # Normaliza o Nome: remove acentos, converte para minúsculas e remove espaços extras nas pontas
    df_entrada[coluna_chave] = df_entrada[coluna_chave].astype(str).str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8').str.lower().str.strip()
    df_base[coluna_chave] = df_base[coluna_chave].astype(str).str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8').str.lower().str.strip()

    # Preenche o Nome com '#ND' caso ele estivesse vazio ou fosse inválido na entrada
    df_entrada[coluna_chave] = df_entrada[coluna_chave].replace('', '#ND')
    # Previne que linhas sem Nome na base cruzem indevidamente com os '#ND' da entrada
    df_base[coluna_chave] = df_base[coluna_chave].replace('', 'SEM_NOME_BASE')

    # Garante que as colunas a preencher existam na base de dados
    colunas_lookup_validas = [col for col in colunas_para_preencher if col in df_base.columns]
    
    # 4. Preparar os DataFrames para a junção (merge)
    df_entrada_chave = df_entrada[[coluna_chave]].copy()
    df_lookup = df_base[[coluna_chave] + colunas_lookup_validas].copy()
    
    # Lógica para duplicatas: manter a mais recente baseada em 'ultimo_atendimento'
    coluna_data = 'ultimo_atendimento'
    if coluna_data in df_lookup.columns:
        print(f"Ordenando a base pela data mais recente em '{coluna_data}' para tratar duplicatas...")
        # Converte a coluna para datetime, tratando erros. Datas inválidas se tornam NaT.
        df_lookup[coluna_data] = pd.to_datetime(df_lookup[coluna_data], errors='coerce')
        # Ordena pela data (mais recentes primeiro). NaT (datas inválidas) vão para o final.
        df_lookup.sort_values(by=coluna_data, ascending=False, inplace=True, na_position='last')

    # Remove duplicatas na base de dados pela chave, mantendo a primeira ocorrência (que agora é a mais recente)
    df_lookup.drop_duplicates(subset=[coluna_chave], keep='first', inplace=True)
    
    print(f"\nRealizando o PROCV para {len(df_entrada_chave)} Nomes...")
    
    # 5. Realizar o "PROCV" usando a função merge do pandas
    df_resultado = pd.merge(
        df_entrada_chave,
        df_lookup,
        on=coluna_chave,
        how='left'
    )
    
    # Formata a data do último atendimento para o padrão DD/MM/YYYY
    coluna_data = 'ultimo_atendimento'
    if coluna_data in df_resultado.columns:
        # .dt acessa propriedades de datetime. strftime formata a data. NaT (datas que não puderam ser convertidas) viram NaN.
        df_resultado[coluna_data] = df_resultado[coluna_data].dt.strftime('%d/%m/%Y')

    # Preenche com '#ND' os valores de Nomes não encontrados ou células vazias
    for col in colunas_lookup_validas:
        df_resultado[col] = df_resultado[col].fillna('#ND')
        df_resultado[col] = df_resultado[col].replace(['', 'nan', 'NaN', 'None'], '#ND')
    
    # 6. Salvar o resultado em um novo arquivo Excel
    try:
        df_resultado.to_excel(caminho_saida, index=False)
        print("-" * 50)
        print("SUCESSO! O processo foi concluído.")
        print(f"O arquivo com os dados preenchidos foi salvo em:\n{caminho_saida}")
        print("-" * 50)
    except Exception as e:
        print(f"ERRO: Falha ao salvar o arquivo de saída. Detalhes: {e}")


if __name__ == "__main__":
    # --- CONFIGURAÇÃO ---
    # Os arquivos devem estar na mesma pasta que este script.
    
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))

    # Nome do arquivo de ENTRADA (com a lista de Nomes)
    arquivo_de_entrada = "entrada.xlsx"
    
    # Nome do arquivo da BASE DE DADOS (onde os dados completos estão)
    arquivo_base_dados = "base.xlsx"
    
    # Nome do arquivo de SAÍDA que será gerado
    arquivo_de_saida = "resultado_procv.xlsx"

    # --- EXECUÇÃO ---
    caminho_entrada = os.path.join(diretorio_script, arquivo_de_entrada)
    caminho_base = os.path.join(diretorio_script, arquivo_base_dados)
    caminho_saida = os.path.join(diretorio_script, arquivo_de_saida)
    
    realizar_procv_multiplo(caminho_entrada, caminho_base, caminho_saida)
    
    input("\nPressione Enter para fechar a tela...")

import pandas as pd
import os
import glob
import sys
import re
from datetime import datetime

def consolidar_relatorios(pasta_entrada, caminho_saida):
    print(f"Buscando arquivos na pasta: {pasta_entrada}")
    
    # Busca todos os arquivos Excel e CSV na pasta de entrada
    arquivos_excel = glob.glob(os.path.join(pasta_entrada, "*.xlsx"))
    arquivos_csv = glob.glob(os.path.join(pasta_entrada, "*.csv"))
    todos_arquivos = arquivos_excel + arquivos_csv
    
    if not todos_arquivos:
        print("Nenhum arquivo encontrado. Verifique se os relatórios estão na pasta correta.")
        return
        
    lista_dataframes = []
    nomes_processados = set()
    
    for arquivo in todos_arquivos:
        nome_base = os.path.splitext(os.path.basename(arquivo))[0]
        # Identifica se é uma cópia baixada (ex: "relatorio (1)") e remove esse sufixo
        nome_limpo = re.sub(r'\s*\(\d+\)$', '', nome_base)
        
        # Evita empilhar o mesmo arquivo duas vezes (ou um .csv e .xlsx com o mesmo nome)
        if nome_limpo in nomes_processados:
            print(f"Ignorando arquivo duplicado: {os.path.basename(arquivo)}")
            continue
            
        nomes_processados.add(nome_limpo)
        print(f"Lendo: {os.path.basename(arquivo)}")
        try:
            # Identifica a extensão para usar o motor de leitura correto
            if arquivo.endswith('.csv'):
                try:
                    df = pd.read_csv(arquivo, sep=';', encoding='utf-8-sig', na_filter=False)
                except UnicodeDecodeError:
                    df = pd.read_csv(arquivo, sep=';', encoding='latin1', na_filter=False)
            else:
                df = pd.read_excel(arquivo, na_filter=False)
            
            # Remove espaços em branco e normaliza para minúsculo para evitar duplicação de colunas (ex: "Nome" e "nome")
            df.columns = df.columns.str.strip().str.lower()
            
            # Extrai a data do nome do arquivo
            nome_arquivo = os.path.basename(arquivo)
            match = re.search(r'(\d{8})', nome_arquivo)
            if match:
                data_str = match.group(1)
                # Converte a string DDMMYYYY para um objeto real de Data
                df['data de envio'] = pd.to_datetime(data_str, format='%d%m%Y').date()
            else:
                df['data de envio'] = pd.NaT
            
            lista_dataframes.append(df)
        except Exception as e:
            print(f"Erro ao ler o arquivo {arquivo}: {e}")
            
    if lista_dataframes:
        print("\nEmpilhando todos os relatórios...")
        # Concatena todos os arquivos. Colunas novas de alguns arquivos serão adicionadas automaticamente.
        df_consolidado = pd.concat(lista_dataframes, ignore_index=True)
        
        # Preenche os valores vazios gerados pela junção de arquivos com colunas diferentes
        for col in df_consolidado.columns:
            if col != 'data de envio':
                df_consolidado[col] = df_consolidado[col].fillna("")
                
        print(f"Total de linhas na base final (sem filtros): {len(df_consolidado)}")
        
        # Salva o arquivo final consolidado
        try:
            df_consolidado.to_excel(caminho_saida, index=False)
            print("-" * 50)
            print(f"SUCESSO! Relatório consolidado gerado em:")
            print(caminho_saida)
            print("-" * 50)
        except Exception as e:
            print(f"Erro ao salvar o arquivo final: {e}")

if __name__ == "__main__":
    # === CONFIGURAÇÕES DE PASTAS ===
    # Diretório onde o script está rodando (com suporte para o executável .exe)
    if getattr(sys, 'frozen', False):
        # Se estiver rodando como arquivo compilado (.exe)
        diretorio_atual = os.path.dirname(sys.executable)
    else:
        # Se estiver rodando como script normal (.py)
        diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    
    # Pasta onde você deve colocar os relatórios baixados (cria a pasta se não existir)
    pasta_relatorios = os.path.join(diretorio_atual, "relatorios_entrada")
    if not os.path.exists(pasta_relatorios):
        os.makedirs(pasta_relatorios)
        print(f"Pasta '{pasta_relatorios}' criada! Coloque seus relatórios lá dentro e rode o script novamente.")
    else:
        # Nome do arquivo final consolidado com a data de hoje
        data_hoje = datetime.now().strftime('%Y%m%d_%H%M')
        arquivo_saida = os.path.join(diretorio_atual, f"relatorio_consolidado_{data_hoje}.xlsx")
        
        consolidar_relatorios(pasta_relatorios, arquivo_saida)
        
    input("\nPressione Enter para fechar...")
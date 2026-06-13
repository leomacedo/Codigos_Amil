import pandas as pd
from datetime import datetime
import os
import sys
import traceback

def processar_sm_ocupacional(caminho_arquivo):
    """
    Lê uma planilha de Saúde Mental Ocupacional, processa os dados
    e gera um arquivo Excel com duas abas: 'report' e 'call center'.
    """
    print(f"Lendo arquivo: {caminho_arquivo}")
    
    if not os.path.exists(caminho_arquivo):
        print(f"ERRO: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return

    try:
        # 1. Carrega o arquivo de origem, suporta csv e excel
        if caminho_arquivo.lower().endswith('.csv'):
            df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8-sig')
        else:
            df = pd.read_excel(caminho_arquivo, engine='openpyxl')
            
        # Limpa espaços nos nomes das colunas
        df.columns = df.columns.str.strip()
        
        # --- VERIFICAÇÃO DE COLUNAS ---
        col_nome = 'Nome Funcionário'
        col_nasc = 'Data de Nascimento'
        col_cpf = 'CPF'
        col_tel = 'Telefone de Contato do Colaborador (atualizado e obrigatório):'
        col_data = 'Data Ficha Clínica'

        colunas_necessarias = [col_nome, col_nasc, col_cpf, col_tel, col_data]
        colunas_faltantes = [c for c in colunas_necessarias if c not in df.columns]

        if colunas_faltantes:
            print("\nERRO: As seguintes colunas obrigatórias não foram encontradas no arquivo:")
            for col in colunas_faltantes:
                print(f"- {col}")
            print(f"\nColunas disponíveis no seu arquivo:")
            for col in df.columns.tolist():
                print(f"- {col}")
            return

        # --- TRATAMENTO DE DADOS ---
        def formatar_data(val):
            if pd.isna(val) or str(val).strip() in ['', 'NaT', 'nan', 'NaN']:
                return ''
            try:
                dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
                return dt.strftime('%d/%m/%Y') if pd.notna(dt) else str(val)
            except Exception:
                return str(val)

        def limpar_texto(val):
            if pd.isna(val):
                return ''
            return str(val).strip()

        def formatar_cpf(val):
            if pd.isna(val) or str(val).strip() == '':
                return ''
            
            cpf = str(val).strip()
            cpf = cpf.replace('.0', '')
            cpf = cpf.replace('.', '').replace('-', '').replace('/', '')
            
            if cpf.isdigit():
                cpf = cpf.zfill(11)
            
            return cpf

        def formatar_telefone(val):
            if pd.isna(val):
                return ''
            
            tel = str(val).strip()
            tel = tel.replace('.0', '')
            return tel

        # Formata campos principais
        df[col_nasc] = df[col_nasc].apply(formatar_data)
        df[col_data] = df[col_data].apply(formatar_data)
        df[col_nome] = df[col_nome].apply(limpar_texto)
        df[col_cpf] = df[col_cpf].apply(formatar_cpf)
        df[col_tel] = df[col_tel].apply(formatar_telefone)

        data_hoje = datetime.now().strftime('%d/%m/%Y')
        tamanho = len(df)

        if tamanho == 0:
            print("Nenhum registro encontrado na base.")
            return

        # --- CRIAÇÃO DA ABA 1: report ---
        df_report = pd.DataFrame(index=df.index)

        df_report['Origem'] = "Base Periódicos"
        df_report['Data de Envio'] = df[col_data].values
        df_report['Nome Completo'] = df[col_nome].values
        df_report['Data de Nascimento'] = df[col_nasc].values
        df_report['CPF'] = df[col_cpf].values
        df_report['Linha'] = "Saúde Mental"
        df_report['Data de direcionamento à captação'] = data_hoje

        df_report = df_report.reset_index(drop=True)

        # --- CRIAÇÃO DA ABA 2: call center ---
        df_call_center = pd.DataFrame(index=df.index)

        df_call_center['FONTE'] = "Saúde Ocupacional"
        df_call_center['CAMINHO'] = "Saúde Ocupacional"
        df_call_center['DATA INSERÇÃO'] = data_hoje
        df_call_center['LINHA DE CUIDADO'] = "Saúde Mental Ocupacional"
        df_call_center['Nome'] = df[col_nome].values
        df_call_center['MO'] = ""
        df_call_center['CPF'] = df[col_cpf].values
        df_call_center['DATA DE NASCIMENTO'] = df[col_nasc].values
        df_call_center['NÍVEL DE REDE'] = ""
        df_call_center['UF'] = ""
        df_call_center['MUNICÍPIO'] = ""
        df_call_center['TELEFONE'] = df[col_tel].values

        df_call_center = df_call_center.reset_index(drop=True)

        # --- SALVAR NA BASE CONSOLIDADA MANTENDO HISTÓRICO ---
        diretorio_saida = os.path.dirname(caminho_arquivo)
        caminho_saida = os.path.join(diretorio_saida, "saida_sm_ocupacional.xlsx")
        
        # Lê a base antiga se ela já existir para não sobrescrever os dados
        if os.path.exists(caminho_saida):
            print("Arquivo de saída já existe. Anexando novos dados...")
            
            try:
                base_report_antiga = pd.read_excel(caminho_saida, sheet_name='report', engine='openpyxl')
                df_report = pd.concat([base_report_antiga, df_report], ignore_index=True)
            except ValueError:
                print("Aviso: Aba 'report' não encontrada no arquivo existente. Será criada uma nova.")
            except Exception as e:
                print(f"Aviso: Não foi possível ler a aba 'report': {e}")
                
            try:
                base_call_antiga = pd.read_excel(caminho_saida, sheet_name='call center', engine='openpyxl')
                df_call_center = pd.concat([base_call_antiga, df_call_center], ignore_index=True)
            except ValueError:
                print("Aviso: Aba 'call center' não encontrada no arquivo existente. Será criada uma nova.")
            except Exception as e:
                print(f"Aviso: Não foi possível ler a aba 'call center': {e}")
        
        # Salva o arquivo com as duas abas
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            df_report.to_excel(writer, sheet_name='report', index=False)
            df_call_center.to_excel(writer, sheet_name='call center', index=False)
            
        print("-" * 50)
        print(f"SUCESSO! Dados consolidados e atualizados em:\n{caminho_saida}")
        print(f"Total de registros processados: {len(df)}")
        print("-" * 50)
        
    except PermissionError:
        print("ERRO: Não foi possível salvar o arquivo.")
        print("Verifique se o arquivo 'saida_sm_ocupacional.xlsx' está aberto no Excel.")
        print("Feche o arquivo e tente novamente.")
        
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {str(e)}")
        print("--- DETALHES DO ERRO ---")
        traceback.print_exc()


if __name__ == "__main__":
    # Obtém o diretório correto dependendo se está rodando como script .py ou executável .exe
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))
        
    caminho_xlsx = os.path.join(diretorio_script, "base.xlsx")
    caminho_csv = os.path.join(diretorio_script, "base.csv")
    
    if os.path.exists(caminho_xlsx):
        processar_sm_ocupacional(caminho_xlsx)
    elif os.path.exists(caminho_csv):
        processar_sm_ocupacional(caminho_csv)
    else:
        print(f"ERRO: Arquivo 'base.xlsx' ou 'base.csv' não encontrado na pasta:\n{diretorio_script}")
        
    input("\nPressione Enter para fechar a tela...")
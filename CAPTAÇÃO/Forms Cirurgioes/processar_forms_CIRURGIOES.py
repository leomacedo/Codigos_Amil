import pandas as pd
from datetime import datetime
import os
import sys
import traceback

def processar_forms_cirurgioes(caminho_arquivo):
    print(f"Lendo arquivo: {caminho_arquivo}")
    
    if not os.path.exists(caminho_arquivo):
        print(f"ERRO: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return

    try:
        # 1. Carrega o arquivo (suporta csv e excel)
        if caminho_arquivo.endswith('.csv'):
            df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8-sig')
        else:
            df = pd.read_excel(caminho_arquivo)
            
        # Limpar espaços invisíveis nos nomes das colunas originais
        df.columns = df.columns.str.strip()
        
        # --- TRATAMENTOS ---
        # Formatar MO para 9 dígitos
        col_mo = 'Número da Carteirinha (Marca ótica)'
        if col_mo in df.columns:
            # Força a coluna a aceitar textos antes de substituir para não dar erro de tipagem
            df[col_mo] = df[col_mo].astype(object)
            df[col_mo] = df[col_mo].apply(lambda x: 0 if pd.isna(x) or str(x).strip().lower() in ['nan', 'nat', ''] else x)
            df[col_mo] = df[col_mo].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(9)
            
        # Formatar TODAS as datas para DD/MM/AAAA (Extraindo apenas a data para campos como Hora de início)
        def formatar_data(val):
            if pd.isna(val) or str(val).strip() in ['', 'NaT', 'nan', 'NaN']:
                return ''
            try:
                dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    return dt.strftime('%d/%m/%Y')
                else:
                    # Fallback caso venha como texto com hora
                    return str(val).split(' ')[0]
            except Exception:
                return str(val).split(' ')[0]
                
        # Aplica a formatação criando colunas auxiliares
        if 'Hora de início' in df.columns:
            df['Hora de início_fmt'] = df['Hora de início'].apply(formatar_data)
        else:
            df['Hora de início_fmt'] = ""
            
        if 'Data de nascimento' in df.columns:
            df['Data de nascimento_fmt'] = df['Data de nascimento'].apply(formatar_data)
        else:
            df['Data de nascimento_fmt'] = ""
            
        # Tratar a coluna Origem (Juntar SP e RJ)
        def unificar_origem(row):
            sp = str(row.get('Origem SP', '')).strip()
            rj = str(row.get('Origem RJ', '')).strip()
            
            if sp and sp.lower() not in ['nan', 'nat', 'none']: return sp
            if rj and rj.lower() not in ['nan', 'nat', 'none']: return rj
            return ''
            
        df['Origem Consolidada'] = df.apply(unificar_origem, axis=1)

        # --- ATUALIZAÇÃO DO STATUS NO ARQUIVO DE ORIGEM ---
        # Usamos uma coluna de controle para não reprocessar as mesmas linhas num próximo uso
        col_status = 'STATUS'
        if col_status not in df.columns:
            df[col_status] = pd.NA
        df[col_status] = df[col_status].fillna('').astype(str)

        # Define a máscara de status vazio (só processa se não tiver sido processado antes)
        mask_status_vazio = df[col_status].str.strip() == ''

        # VERIFICAÇÃO DE PARADA
        if not mask_status_vazio.any():
            print("Todos os registros já foram processados. Parando a execução para não duplicar dados na base consolidada.")
            return

        # Separa as linhas que serão processadas agora
        df_processar = df[mask_status_vazio].copy()
        print(f"Novas linhas para processamento encontradas: {len(df_processar)}")

        # Atualiza o status no dataframe original
        df.loc[mask_status_vazio, col_status] = 'ENVIADO CAPTAÇÃO E NAVEGAÇÃO'

        # Salva o arquivo de origem de volta com as alterações de status
        print("Atualizando a coluna de controle no arquivo de origem...")
        try:
            if caminho_arquivo.endswith('.csv'):
                df.to_csv(caminho_arquivo, sep=';', index=False, encoding='utf-8-sig')
            else:
                df.to_excel(caminho_arquivo, index=False, engine='openpyxl')
            print(f"Arquivo de origem '{os.path.basename(caminho_arquivo)}' foi atualizado.")
        except Exception as e:
            print(f"AVISO: Falha ao salvar as alterações no arquivo de origem: {e}")

        # --- MAPEAMENTO PARA AS ABAS CAPTAÇÃO E NAVEGAÇÃO ---
        # A mesma base (df_processar) alimenta ambas as planilhas
        
        # 1. ABA CAPTAÇÃO
        df_captacao = pd.DataFrame(index=df_processar.index)
        
        df_captacao['FONTE'] = "Forms Bariátrica (Cirurgiões)"
        df_captacao['CAMINHO'] = "Dom Pedro"
        df_captacao['DATA INSERÇÃO'] = datetime.now().strftime('%d/%m/%Y')
        df_captacao['LINHA DE CUIDADO'] = "Emagrecimento"
        df_captacao['Nome'] = df_processar['Nome completo do paciente (Sem abreviações)'] if 'Nome completo do paciente (Sem abreviações)' in df_processar.columns else ""
        df_captacao['MO'] = df_processar['Número da Carteirinha (Marca ótica)'] if 'Número da Carteirinha (Marca ótica)' in df_processar.columns else ""
        df_captacao['CPF'] = df_processar['CPF do Paciente'] if 'CPF do Paciente' in df_processar.columns else ""
        df_captacao['DATA DE NASCIMENTO'] = df_processar['Data de nascimento_fmt']
        df_captacao['NÍVEL DE REDE'] = df_processar['nivel de rede'] if 'nivel de rede' in df_processar.columns else ""
        df_captacao['UF'] = df_processar['Regional'] if 'Regional' in df_processar.columns else ""
        df_captacao['MUNICÍPIO'] = ""
        df_captacao['TELEFONE'] = df_processar['Telefone atualizado com DDD'] if 'Telefone atualizado com DDD' in df_processar.columns else ""
        df_captacao['STATUS'] = " "
        df_captacao['Enviado para'] = "Dom Pedro"

        
        # 2. ABA NAVEGAÇÃO
        df_navegacao = pd.DataFrame(index=df_processar.index)
        
        df_navegacao['Prioridade no Contato?'] = "Sim"
        df_navegacao['Data Envio'] = datetime.now().strftime('%d/%m/%Y')
        df_navegacao['Motivo envio'] = "Captação + Navegação"
        df_navegacao['Origem'] = df_processar['Origem Consolidada']
        df_navegacao['MO'] = df_processar['Número da Carteirinha (Marca ótica)'] if 'Número da Carteirinha (Marca ótica)' in df_processar.columns else ""
        df_navegacao['Nome do paciente'] = df_processar['Nome completo do paciente (Sem abreviações)'] if 'Nome completo do paciente (Sem abreviações)' in df_processar.columns else ""
        df_navegacao['CPF'] = df_processar['CPF do Paciente'] if 'CPF do Paciente' in df_processar.columns else ""
        df_navegacao['Data de Nascimento'] = df_processar['Data de nascimento_fmt']
        df_navegacao['Idade'] = df_processar['Idade do paciente'] if 'Idade do paciente' in df_processar.columns else ""
        df_navegacao['Genero'] = ""
        df_navegacao['Estado'] = df_processar['Regional'] if 'Regional' in df_processar.columns else ""
        df_navegacao['MUNICÍPIO'] = ""
        df_navegacao['bairro'] = ""
        df_navegacao['Nível de Rede'] = df_processar['nivel de rede'] if 'nivel de rede' in df_processar.columns else ""
        df_navegacao['Telefone atualizado'] = df_processar['Telefone atualizado com DDD'] if 'Telefone atualizado com DDD' in df_processar.columns else ""
    
        
        # Salva na Base Consolidada Mantendo o Histórico
        diretorio_saida = os.path.dirname(caminho_arquivo)
        caminho_saida = os.path.join(diretorio_saida, "base_consolidada_cirurgioes.xlsx")
        
        # Lê a base antiga se ela já existir para não sobrescrever os dados anteriores
        if os.path.exists(caminho_saida):
            try:
                base_cap = pd.read_excel(caminho_saida, sheet_name='base captação')
                df_captacao = pd.concat([base_cap, df_captacao], ignore_index=True)
            except Exception: pass
                
            try:
                base_nav = pd.read_excel(caminho_saida, sheet_name='planilha dom pedro')
                df_navegacao = pd.concat([base_nav, df_navegacao], ignore_index=True)
            except Exception: pass
        
        # Salva o arquivo com as duas abas
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            df_captacao.to_excel(writer, sheet_name='base captação', index=False)
            df_navegacao.to_excel(writer, sheet_name='planilha dom pedro', index=False)
            
        print("-" * 50)
        print(f"SUCESSO! {len(df_processar)} registros copiados para Captação e Navegação.")
        print(f"Dados consolidados salvos e atualizados em:\n{caminho_saida}")
        print("-" * 50)
        
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        print("--- DETALHES DO ERRO ---")
        traceback.print_exc()

if __name__ == "__main__":
    # Obtém o diretório correto dependendo se está rodando como script (.py) ou executável (.exe)
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))
        
    caminho_xlsx = os.path.join(diretorio_script, "Forms_Cirurgioes.xlsx")
    caminho_csv = os.path.join(diretorio_script, "Forms_Cirurgioes.csv")
    
    if os.path.exists(caminho_xlsx):
        processar_forms_cirurgioes(caminho_xlsx)
    elif os.path.exists(caminho_csv):
        processar_forms_cirurgioes(caminho_csv)
    else:
        print(f"ERRO: Arquivo 'Forms_Cirurgioes.xlsx' ou '.csv' não encontrado na pasta: {diretorio_script}")
        
    input("\nPressione Enter para fechar a tela...")

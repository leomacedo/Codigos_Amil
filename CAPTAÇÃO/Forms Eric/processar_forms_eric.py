import pandas as pd
from datetime import datetime
import os
import sys
import traceback

def processar_forms_eric(caminho_arquivo):
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
        
        col_filtro = 'Data do agendamento realizado'
        
        if col_filtro not in df.columns:
            print(f"ERRO: Coluna '{col_filtro}' não encontrada no arquivo.")
            print(f"Colunas disponíveis: {df.columns.tolist()}")
            return
            
        # --- TRATAMENTOS ---
        # Formatar MO para 9 dígitos
        col_mo = 'Marca ótica (MO)'
        if col_mo in df.columns:
            # Força a coluna a aceitar textos antes de substituir para não dar erro de tipagem
            df[col_mo] = df[col_mo].astype(object)
            df[col_mo] = df[col_mo].apply(lambda x: 0 if pd.isna(x) or str(x).strip().lower() in ['nan', 'nat', ''] else x)
            df[col_mo] = df[col_mo].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(9)
            
        # Formatar TODAS as datas para DD/MM/AAAA
        def formatar_data(val):
            if pd.isna(val) or str(val).strip() in ['', 'NaT', 'nan', 'NaN']:
                return ''
            try:
                dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    return dt.strftime('%d/%m/%Y')
                else:
                    return str(val)
            except Exception:
                return str(val)
                
        colunas_datas = [col for col in df.columns if 'data' in str(col).lower()]
        for col in colunas_datas:
            # Força a coluna a aceitar textos antes de formatar a data
            df[col] = df[col].astype(object)
            df[col] = df[col].apply(formatar_data)

        # --- ATUALIZAÇÃO DO STATUS NO ARQUIVO DE ORIGEM ---
        col_status = 'STATUS'
        if col_status not in df.columns:
            df[col_status] = pd.NA
        df[col_status] = df[col_status].fillna('').astype(str)

        # Define as máscaras de captação/navegação
        mask_captacao = df[col_filtro].isna() | (df[col_filtro].astype(str).str.strip() == '') | (df[col_filtro].astype(str).str.strip().str.lower() == 'nan')
        mask_navegacao = ~mask_captacao

        # Define a máscara de status vazio (só atualiza se não tiver sido preenchido antes)
        mask_status_vazio = df[col_status].str.strip() == ''

        # VERIFICAÇÃO DE PARADA: Se não houver nenhum status vazio, PARA a execução.
        if not mask_status_vazio.any():
            print("Todos os registros já possuem status. Parando a execução para não duplicar dados na base consolidada.")
            return

        # Aplica o status no dataframe principal nas linhas com status vazio
        df.loc[mask_captacao & mask_status_vazio, col_status] = 'ENVIADO CAPTAÇÃO'
        df.loc[mask_navegacao & mask_status_vazio, col_status] = 'ENVIADO NAVEGAÇÃO'

        # Salva o arquivo de origem de volta, se algo foi alterado
        if (mask_captacao & mask_status_vazio).any() or (mask_navegacao & mask_status_vazio).any():
            print("Atualizando a coluna 'status' no arquivo de origem...")
            try:
                if caminho_arquivo.endswith('.csv'):
                    df.to_csv(caminho_arquivo, sep=';', index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(caminho_arquivo, index=False, engine='openpyxl')
                print(f"Arquivo de origem '{os.path.basename(caminho_arquivo)}' foi atualizado.")
            except Exception as e:
                print(f"AVISO: Falha ao salvar as alterações no arquivo de origem: {e}")

        # --- SEPARAÇÃO ENTRE CAPTAÇÃO E NAVEGAÇÃO ---
        # Filtra APENAS os registros que acabaram de ser preenchidos (que estavam vazios) para não duplicar na base final
        df_vazios = df[mask_captacao & mask_status_vazio].copy()
        df_com_data = df[mask_navegacao & mask_status_vazio].copy()
        
        print(f"Total de linhas na planilha: {len(df)}")
        print(f"  -> Irão para 'captação': {len(df_vazios)}")
        print(f"  -> Irão para 'navegação': {len(df_com_data)}")
        
        if df_vazios.empty and df_com_data.empty:
            print("Nenhum registro encontrado para organizar.")
            return
            
        # Mapeia as colunas CAPTAÇÃO (Sem data)
        df_captacao = pd.DataFrame(index=df_vazios.index)
        
        df_captacao['FONTE'] = "Forms TotalCare"
        df_captacao['CAMINHO'] = "Forms TotalCare"
        df_captacao['DATA INSERÇÃO'] = datetime.now().strftime('%d/%m/%Y')
        df_captacao['LINHA DE CUIDADO'] = df_vazios['Para qual programa o beneficiário será encaminhado?'] if 'Para qual programa o beneficiário será encaminhado?' in df_vazios.columns else ""
        df_captacao['Nome'] = df_vazios['Nome completo do beneficiário (a)'] if 'Nome completo do beneficiário (a)' in df_vazios.columns else ""
        df_captacao['MO'] = df_vazios['Marca ótica (MO)'] if 'Marca ótica (MO)' in df_vazios.columns else ""
        df_captacao['CPF'] = ""
        df_captacao['DATA DE NASCIMENTO'] = ""
        df_captacao['NÍVEL DE REDE'] = ""
        df_captacao['UF'] = df_vazios['Estado em que reside'] if 'Estado em que reside' in df_vazios.columns else ""
        df_captacao['MUNICÍPIO'] = ""
        df_captacao['TELEFONE'] = df_vazios['Número de telefone'] if 'Número de telefone' in df_vazios.columns else ""
        df_captacao['STATUS'] = ""
        
        # Mapeia as colunas NAVEGAÇÃO (Com data)
        df_navegacao = pd.DataFrame(index=df_com_data.index)
        
        df_navegacao['Origem'] = "Forms Eric (Desosp Total Care)"
        df_navegacao['Jornada'] = df_com_data['Para qual programa o beneficiário será encaminhado?'] if 'Para qual programa o beneficiário será encaminhado?' in df_com_data.columns else ""
        df_navegacao['Flag'] = ""
        df_navegacao['Data Input'] = datetime.now().strftime('%d/%m/%Y')
        df_navegacao['MO'] = df_com_data['Marca ótica (MO)'] if 'Marca ótica (MO)' in df_com_data.columns else ""
        df_navegacao['Nome do paciente'] = df_com_data['Nome completo do beneficiário (a)'] if 'Nome completo do beneficiário (a)' in df_com_data.columns else ""
        df_navegacao['CPF'] = ""
        df_navegacao['Data de Nascimento'] = ""
        df_navegacao['Idade'] = ""
        df_navegacao['Genero'] = ""
        df_navegacao['Estado'] = df_com_data['Estado em que reside'] if 'Estado em que reside' in df_com_data.columns else ""
        df_navegacao['Município'] = ""
        df_navegacao['Bairro'] = ""
        df_navegacao['Nível de Rede'] = ""
        df_navegacao['Telefone atualizado'] = df_com_data['Número de telefone'] if 'Número de telefone' in df_com_data.columns else ""
        df_navegacao['Hub nvg'] = "CMDP"
        df_navegacao['Enviado nvg?'] = "Não"
        df_navegacao['data envio nvg'] = ""
        df_navegacao['Status nvg'] = ""
        df_navegacao['Data Atualização Status nvg'] = ""
        df_navegacao['Data de Entrada na Linha'] = df_com_data['Data de inscrição'] if 'Data de inscrição' in df_com_data.columns else ""
        df_navegacao['Data da Consulta'] = df_com_data['Data do agendamento realizado'] if 'Data do agendamento realizado' in df_com_data.columns else ""
        df_navegacao['Porta de Entrada'] = ""
        
        # Salva na Base Consolidada Mantendo o Histórico
        diretorio_saida = os.path.dirname(caminho_arquivo)
        caminho_saida = os.path.join(diretorio_saida, "base_consolidada_eric.xlsx")
        
        # Lê a base antiga se ela já existir para não sobrescrever os dados anteriores
        if os.path.exists(caminho_saida):
            try:
                base_cap = pd.read_excel(caminho_saida, sheet_name='captação')
                df_captacao = pd.concat([base_cap, df_captacao], ignore_index=True)
            except Exception: pass
                
            try:
                base_nav = pd.read_excel(caminho_saida, sheet_name='navegação')
                df_navegacao = pd.concat([base_nav, df_navegacao], ignore_index=True)
            except Exception: pass
        
        # Salva o arquivo com as duas abas
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            df_captacao.to_excel(writer, sheet_name='captação', index=False)
            df_navegacao.to_excel(writer, sheet_name='navegação', index=False)
            
        print("-" * 50)
        print(f"SUCESSO! Dados consolidados salvos e atualizados em:\n{caminho_saida}")
        print("-" * 50)
        
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        print("--- DETALHES DO ERRO PARA O CHAT ---")
        traceback.print_exc()

if __name__ == "__main__":
    # Obtém o diretório correto dependendo se está rodando como script (.py) ou executável (.exe)
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))
        
    caminho_xlsx = os.path.join(diretorio_script, "Forms_Cru.xlsx")
    caminho_csv = os.path.join(diretorio_script, "Forms_Cru.csv")
    
    if os.path.exists(caminho_xlsx):
        processar_forms_eric(caminho_xlsx)
    elif os.path.exists(caminho_csv):
        processar_forms_eric(caminho_csv)
    else:
        print(f"ERRO: Arquivo 'Forms_Cru.xlsx' ou 'Forms_Cru.csv' não encontrado na pasta: {diretorio_script}")
        
    input("\nPressione Enter para fechar a tela...")
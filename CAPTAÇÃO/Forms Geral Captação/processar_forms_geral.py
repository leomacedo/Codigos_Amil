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
        # 1. Carrega o arquivo, suporta CSV e Excel
        if caminho_arquivo.endswith('.csv'):
            df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8-sig')
        else:
            df = pd.read_excel(caminho_arquivo)
            
        # Limpar espaços invisíveis nos nomes das colunas originais
        df.columns = df.columns.str.strip()
        
        # --- FUNÇÃO PARA CONSOLIDAR COLUNAS REPETIDAS ---
        def consolidar_colunas(df_alvo, prefixos):
            cols = []
            for pref in prefixos:
                cols.extend([c for c in df_alvo.columns if pref.lower() in str(c).lower()])
            
            cols = list(dict.fromkeys(cols))
            
            s = pd.Series(None, index=df_alvo.index, dtype=object)
            for c in cols:
                valida = df_alvo[c].apply(
                    lambda x: None if pd.isna(x) or str(x).strip().lower() in ['', 'nan', 'nat'] else x
                )
                s = s.combine_first(valida)
            return s
            
        # --- MAPEAMENTO DAS COLUNAS ORIGINAIS ---
        df['__MO__'] = consolidar_colunas(df, ['Marca ótica (MO)'])
        df['__NOME__'] = consolidar_colunas(df, ['Nome completo do beneficiário'])
        df['__CPF__'] = consolidar_colunas(df, ['CPF'])
        df['__NASCIMENTO__'] = consolidar_colunas(df, ['Data de nascimento'])
        df['__GENERO__'] = consolidar_colunas(df, ['Gênero', 'Genero'])
        df['__ESTADO__'] = consolidar_colunas(df, ['Estado de origem', 'Estado em que reside'])
        df['__MUNICIPIO__'] = consolidar_colunas(df, ['Município', 'Municipio'])
        df['__REDE__'] = consolidar_colunas(df, ['Nível de rede', 'Nivel de rede'])
        df['__TELEFONE__'] = consolidar_colunas(df, ['Telefone'])
        df['__PROGRAMA__'] = consolidar_colunas(df, ['Para qual programa'])
        df['__AREA__'] = consolidar_colunas(df, ['Área de origem', 'Area de origem'])
        df['__EMPRESA__'] = consolidar_colunas(df, ['Empresa'])
        df['__RESPONSAVEL_GSC__'] = consolidar_colunas(df, ['Analista responsável (GSC):'])
        df['__DESCRICAO_BREVE__'] = consolidar_colunas(df, ['Descrição breve do caso'])
        df['__DATA_AGENDAMENTO__'] = consolidar_colunas(df, ['Data do agendamento realizado:'])
        df['__HORA_INICIO__'] = consolidar_colunas(df, ['Hora de início', 'Hora de inicio'])
        df['__SAUDE_MENTAL_1__'] = consolidar_colunas(df, ['Descrição da origem da solicitação e motivo do direcionamento'])
        df['__SAUDE_MENTAL_2__'] = consolidar_colunas(df, ['Descrição da queixa/solicitação principal'])
        df['__SAUDE_MENTAL_3__'] = consolidar_colunas(df, ['Descrição da ciência ou não do direcionamento'])

        # --- TRATAMENTOS ---
        # Formatar MO para 9 dígitos
        df['__MO__'] = df['__MO__'].apply(lambda x: 0 if pd.isna(x) else x)
        df['__MO__'] = df['__MO__'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(9)
        df['__MO__'] = df['__MO__'].replace('000000000', '')
            
        # Formatar datas para DD/MM/AAAA
        def formatar_data(val):
            if pd.isna(val) or str(val).strip() in ['', 'NaT', 'nan', 'NaN']:
                return ''
            try:
                val_str = str(val).split(' ')[0]
                dt = pd.to_datetime(val_str, dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    return dt.strftime('%d/%m/%Y')
                else:
                    return val_str
            except Exception:
                return str(val).split(' ')[0]
                
        colunas_datas = ['__NASCIMENTO__', '__DATA_AGENDAMENTO__']
        for col in colunas_datas:
            df[col] = df[col].apply(formatar_data)

        def formatar_data_us(val):
            if pd.isna(val) or str(val).strip() in ['', 'NaT', 'nan', 'NaN']:
                return ''
            try:
                val_str = str(val).split(' ')[0]
                dt = pd.to_datetime(val_str, dayfirst=False, errors='coerce')
                if pd.notna(dt):
                    return dt.strftime('%d/%m/%Y')
                else:
                    return val_str
            except Exception:
                return str(val).split(' ')[0]
                
        df['__HORA_INICIO__'] = df['__HORA_INICIO__'].apply(formatar_data_us)

        # --- NORMALIZAÇÃO DOS PROGRAMAS ---
        def normalizar_programa(nome_bruto):
            if pd.isna(nome_bruto):
                return ""
            
            val = str(nome_bruto).upper()
            
            if 'ANTICOAGULANTE' in val:
                return 'Anticoagulante Seguro'
            if 'VALVOPATIAS' in val:
                return 'Cuidado Cardíaco Valvar'
            if 'MAMA' in val:
                return 'Cuidado Integral da Mama'
            if 'PRÓSTATA' in val or 'PROSTATA' in val:
                return 'Cuidado Integral da Próstata'
            if 'COLORRETAL' in val:
                return 'Cuidado Oncológico Colorretal'
            if 'PULMONAR' in val:
                return 'Cuidado Oncológico Pulmonar'
            if 'ENDOMETRIOSE' in val:
                return 'Cuidados para Endometriose'
            if 'INFARTO' in val or 'IAM' in val:
                return 'Cuidados Pós Infarto'
            if 'EMAGRECIMENTO' in val:
                return 'Emagrecimento'
            if 'FUMO' in val:
                return 'Fumo Zero'
            if 'GESTA' in val or 'GESTANTE' in val:
                return 'Gestação Segura'
            if 'INSUFICI' in val or 'ICC' in val:
                return 'Insuficiência Cardíaca Controlada'
            if 'AVC' in val:
                return 'Pós AVC'
            if 'ARRITMIA' in val:
                return 'Ritmo Certo'
            if 'COLUNA' in val:
                return 'Saúde da Coluna'
            if 'MENTAL' in val or 'PSIQUIATRIA' in val:
                return 'Saúde Mental'
            if 'RENAL' in val:
                return 'Saúde Renal'
            
            return str(nome_bruto).strip()
            
        df['__PROGRAMA__'] = df['__PROGRAMA__'].apply(normalizar_programa)

        # --- ATUALIZAÇÃO DO STATUS NO ARQUIVO DE ORIGEM ---
        col_status = 'STATUS DO LEO MACEDO'
        if col_status not in df.columns:
            df[col_status] = pd.NA
        
        df[col_status] = df[col_status].fillna('').astype(str)

        # Máscaras principais
        mask_eric = df['__AREA__'].astype(str).str.contains('TIME ERIC|ERIC', case=False, na=False)
        mask_com_data = df['__DATA_AGENDAMENTO__'].notna() & (df['__DATA_AGENDAMENTO__'].astype(str).str.strip() != '')
        
        # Linhas a serem descartadas
        mask_remover_programa = df['__PROGRAMA__'].astype(str).str.contains('Viva Bem|Bem Cuidado', case=False, na=False)
        mask_remover_santa_helena = df['__REDE__'].astype(str).str.contains('santa helena', case=False, na=False)
        mask_remover_ana_costa = df['__REDE__'].astype(str).str.contains('ana costa', case=False, na=False)
        mask_remover_sobam = df['__REDE__'].astype(str).str.contains('sobam', case=False, na=False)
        
        mask_remover_redes_especificas = mask_remover_santa_helena | mask_remover_ana_costa | mask_remover_sobam
        mask_remover = mask_remover_programa | mask_remover_redes_especificas
        
        # Saúde Mental
        mask_saude_mental_cond = (
            (df['__SAUDE_MENTAL_1__'].notna() & (df['__SAUDE_MENTAL_1__'].astype(str).str.strip() != '')) |
            (df['__SAUDE_MENTAL_2__'].notna() & (df['__SAUDE_MENTAL_2__'].astype(str).str.strip() != '')) |
            (df['__SAUDE_MENTAL_3__'].notna() & (df['__SAUDE_MENTAL_3__'].astype(str).str.strip() != ''))
        )
        
        # Saúde Ocupacional
        mask_saude_ocupacional_cond = df['__AREA__'].astype(str).str.contains(
            'SAÚDE OCUPACIONAL|SAUDE OCUPACIONAL',
            case=False,
            na=False
        )
        
        # Captação Dom Pedro
        mask_captacao_dp_cond = df['__PROGRAMA__'].astype(str).str.contains(
            r'Mama|Próstata|Pulmonar|Colorretal|Anticoagulante Seguro|Pós AVC',
            case=False,
            na=False
        )
        
        # GSC
        mask_gsc_cond = df['__AREA__'].astype(str).str.contains('GSC', case=False, na=False)
        
        # Máscaras finais
        mask_saude_mental = mask_saude_mental_cond & ~mask_remover

        mask_saude_ocupacional = (
            mask_saude_ocupacional_cond
            & ~mask_saude_mental_cond
            & ~mask_remover
        )

        # Saúde Ocupacional NÃO vai para navegação
        mask_navegacao = (
            mask_eric
            & mask_com_data
            & ~mask_saude_mental_cond
            & ~mask_saude_ocupacional_cond
            & ~mask_remover
        )

        # Saúde Ocupacional VAI para captação também
        mask_captacao = (
            ~mask_navegacao
            & ~mask_saude_mental_cond
            & ~mask_remover
        )

        # Só atualiza se STATUS estiver vazio
        mask_status_vazio = df[col_status].str.strip() == ''

        # Se não houver nenhum status vazio, para
        if not mask_status_vazio.any():
            print("Todos os registros já possuem status. Parando a execução para não duplicar dados na base consolidada.")
            return

        # Aplica status
        df.loc[mask_remover & mask_status_vazio, col_status] = 'DESCARTADO'
        df.loc[mask_captacao & mask_status_vazio, col_status] = 'ENVIADO CAPTAÇÃO'
        df.loc[mask_navegacao & mask_status_vazio, col_status] = 'ENVIADO NAVEGAÇÃO'
        df.loc[mask_saude_mental & mask_status_vazio, col_status] = 'ENVIADO SAÚDE MENTAL'
        df.loc[mask_saude_ocupacional & mask_status_vazio, col_status] = 'ENVIADO SAÚDE OCUPACIONAL'

        # Salva o arquivo de origem atualizado
        if (
            (mask_captacao & mask_status_vazio).any()
            or (mask_navegacao & mask_status_vazio).any()
            or (mask_saude_mental & mask_status_vazio).any()
            or (mask_saude_ocupacional & mask_status_vazio).any()
            or (mask_remover & mask_status_vazio).any()
        ):
            print("Atualizando a coluna 'STATUS' no arquivo de origem...")
            
            cols_temp = [c for c in df.columns if c.startswith('__') and c.endswith('__')]
            df_salvar = df.drop(columns=cols_temp)
            
            try:
                if caminho_arquivo.endswith('.csv'):
                    df_salvar.to_csv(caminho_arquivo, sep=';', index=False, encoding='utf-8-sig')
                else:
                    df_salvar.to_excel(caminho_arquivo, index=False, engine='openpyxl')
                
                print(f"Arquivo de origem '{os.path.basename(caminho_arquivo)}' foi atualizado.")
            except Exception as e:
                print(f"AVISO: Falha ao salvar as alterações no arquivo de origem: {e}")

        # --- SEPARAÇÃO ENTRE AS ABAS ---
        df_vazios = df[mask_captacao & mask_status_vazio].copy()
        df_com_data = df[mask_navegacao & mask_status_vazio].copy()
        df_saude_mental_linhas = df[mask_saude_mental & mask_status_vazio].copy()
        df_saude_ocupacional_linhas = df[mask_saude_ocupacional & mask_status_vazio].copy()
        df_gsc_linhas = df[mask_gsc_cond & mask_status_vazio & ~mask_remover].copy()
        df_captacao_dp_linhas = df[mask_captacao_dp_cond & mask_status_vazio & ~mask_remover].copy()
        
        print(f"Total de linhas na planilha: {len(df)}")
        print(f"Linhas com STATUS vazio: {mask_status_vazio.sum()}")  # conta quantas linhas estão sem status
        print(f"  -> Irão para 'captação': {len(df_vazios)}")
        print(f"  -> Irão para 'navegação': {len(df_com_data)}")
        print(f"  -> Irão para 'Saúde Mental': {len(df_saude_mental_linhas)}")
        print(f"  -> Irão para 'saude ocupacional': {len(df_saude_ocupacional_linhas)}")
        print(f"  -> Irão para 'GSC': {len(df_gsc_linhas)}")
        print(f"  -> Irão para 'captação dom pedro': {len(df_captacao_dp_linhas)}")
        print(f"  -> Descartados (Viva Bem/Bem Cuidado): {(mask_remover_programa & mask_status_vazio).sum()}")
        print(f"  -> Descartados (Redes: Santa Helena/Ana Costa/Sobam): {(mask_remover_redes_especificas & mask_status_vazio).sum()}")
        
        if (
            df_vazios.empty
            and df_com_data.empty
            and df_saude_mental_linhas.empty
            and df_saude_ocupacional_linhas.empty
            and df_gsc_linhas.empty
            and df_captacao_dp_linhas.empty
        ):
            print("Nenhum registro encontrado para organizar.")
            return
            
        # --- ABA CAPTAÇÃO ---
        df_captacao = pd.DataFrame(index=df_vazios.index)
        
        df_captacao['FONTE'] = df_vazios['__AREA__'].fillna("Forms Geral")
        
        mask_onco_cap = df_vazios['__PROGRAMA__'].astype(str).str.contains(
            r'Mama|Próstata|Pulmonar|Colorretal|Anticoagulante Seguro|Pós AVC',
            case=False,
            na=False
        )
        
        df_captacao['CAMINHO'] = "Call Center"
        df_captacao.loc[mask_onco_cap, 'CAMINHO'] = "Dom Pedro"
        
        df_captacao['DATA INSERÇÃO'] = datetime.now().strftime('%d/%m/%Y')
        df_captacao['LINHA DE CUIDADO'] = df_vazios['__PROGRAMA__']
        df_captacao['Nome'] = df_vazios['__NOME__']
        df_captacao['MO'] = df_vazios['__MO__']
        df_captacao['CPF'] = df_vazios['__CPF__']
        df_captacao['DATA DE NASCIMENTO'] = df_vazios['__NASCIMENTO__']
        df_captacao['NÍVEL DE REDE'] = df_vazios['__REDE__']
        df_captacao['UF'] = df_vazios['__ESTADO__']
        df_captacao['MUNICÍPIO'] = df_vazios['__MUNICIPIO__']
        df_captacao['TELEFONE'] = df_vazios['__TELEFONE__']
        df_captacao['STATUS'] = ""
        
        # --- ABA SAÚDE MENTAL ---
        df_saude_mental = pd.DataFrame(index=df_saude_mental_linhas.index)
        
        df_saude_mental['FONTE'] = df_saude_mental_linhas['__AREA__'].fillna("Forms Geral")
        df_saude_mental['CAMINHO'] = "Call Center"
        df_saude_mental['DATA INSERÇÃO'] = datetime.now().strftime('%d/%m/%Y')
        df_saude_mental['LINHA DE CUIDADO'] = df_saude_mental_linhas['__PROGRAMA__']
        df_saude_mental['Nome'] = df_saude_mental_linhas['__NOME__']
        df_saude_mental['MO'] = df_saude_mental_linhas['__MO__']
        df_saude_mental['CPF'] = df_saude_mental_linhas['__CPF__']
        df_saude_mental['DATA DE NASCIMENTO'] = df_saude_mental_linhas['__NASCIMENTO__']
        df_saude_mental['NÍVEL DE REDE'] = df_saude_mental_linhas['__REDE__']
        df_saude_mental['UF'] = df_saude_mental_linhas['__ESTADO__']
        df_saude_mental['MUNICÍPIO'] = df_saude_mental_linhas['__MUNICIPIO__']
        df_saude_mental['TELEFONE'] = df_saude_mental_linhas['__TELEFONE__']
        df_saude_mental['STATUS'] = ""
        df_saude_mental['Descrição da origem da solicitação e motivo do direcionamento'] = df_saude_mental_linhas['__SAUDE_MENTAL_1__']
        df_saude_mental['Descrição da queixa/solicitação principal'] = df_saude_mental_linhas['__SAUDE_MENTAL_2__']
        df_saude_mental['Descrição da ciência ou não do direcionamento'] = df_saude_mental_linhas['__SAUDE_MENTAL_3__']

        # --- ABA SAÚDE OCUPACIONAL ---
        df_saude_ocupacional = pd.DataFrame(index=df_saude_ocupacional_linhas.index)

        df_saude_ocupacional['ORIGEM'] = df_saude_ocupacional_linhas['__AREA__']
        df_saude_ocupacional['DATA DE ENVIO PARA'] = df_saude_ocupacional_linhas['__HORA_INICIO__']
        df_saude_ocupacional['NOME COMPLETO'] = df_saude_ocupacional_linhas['__NOME__']
        df_saude_ocupacional['DATA DE NASCIMENTO'] = df_saude_ocupacional_linhas['__NASCIMENTO__']
        df_saude_ocupacional['CPF'] = df_saude_ocupacional_linhas['__CPF__']
        df_saude_ocupacional['LINHA'] = df_saude_ocupacional_linhas['__PROGRAMA__']
        df_saude_ocupacional['DATA DE ENVIO DE ENVIO PARA CAPTAÇÃO'] = datetime.now().strftime('%d/%m/%Y')
        
        # --- ABA CAPTAÇÃO DOM PEDRO ---
        df_captacao_dp = pd.DataFrame(index=df_captacao_dp_linhas.index)
        
        df_captacao_dp['Programa'] = df_captacao_dp_linhas['__PROGRAMA__']
        df_captacao_dp['Prioridade no Contato?'] = "Sim"
        df_captacao_dp['Data Envio'] = datetime.now().strftime('%d/%m/%Y')
        df_captacao_dp['Motivo envio'] = "Captação + Navegação"
        
        df_captacao_dp['Origem'] = df_captacao_dp_linhas['__AREA__'].fillna("Forms Geral")
        
        mask_eric_dp = df_captacao_dp_linhas['__AREA__'].astype(str).str.contains(
            'TIME ERIC|ERIC',
            case=False,
            na=False
        )
        
        df_captacao_dp.loc[mask_eric_dp, 'Origem'] = "Forms"
        
        df_captacao_dp['MO'] = df_captacao_dp_linhas['__MO__']
        df_captacao_dp['Nome do paciente'] = df_captacao_dp_linhas['__NOME__']
        df_captacao_dp['CPF'] = df_captacao_dp_linhas['__CPF__']
        df_captacao_dp['Data de Nascimento'] = df_captacao_dp_linhas['__NASCIMENTO__']
        df_captacao_dp['Idade'] = ""
        df_captacao_dp['Genero'] = df_captacao_dp_linhas['__GENERO__']
        df_captacao_dp['Estado'] = df_captacao_dp_linhas['__ESTADO__']
        df_captacao_dp['Município'] = df_captacao_dp_linhas['__MUNICIPIO__']
        df_captacao_dp['Bairro'] = ""
        df_captacao_dp['Nível de Rede'] = df_captacao_dp_linhas['__REDE__']
        df_captacao_dp['Telefone atualizado'] = df_captacao_dp_linhas['__TELEFONE__']
        
        # --- ABA GSC ---
        df_gsc = pd.DataFrame(index=df_gsc_linhas.index)
        
        df_gsc['Data de envio'] = df_gsc_linhas['__HORA_INICIO__']
        df_gsc['Linha De Cuidado'] = df_gsc_linhas['__PROGRAMA__']
        df_gsc['MO'] = df_gsc_linhas['__MO__']
        df_gsc['Nome do paciente'] = df_gsc_linhas['__NOME__']
        df_gsc['CPF'] = df_gsc_linhas['__CPF__']
        df_gsc['Data de Nascimento'] = df_gsc_linhas['__NASCIMENTO__']
        df_gsc['Genero'] = df_gsc_linhas['__GENERO__']
        df_gsc['Estado'] = df_gsc_linhas['__ESTADO__']
        df_gsc['Município'] = df_gsc_linhas['__MUNICIPIO__']
        df_gsc['Nível de Rede'] = df_gsc_linhas['__REDE__']
        df_gsc['Telefone atualizado'] = df_gsc_linhas['__TELEFONE__']
        df_gsc['Empresa'] = df_gsc_linhas['__EMPRESA__']
        df_gsc['Responsável pelo caso (GSC)'] = df_gsc_linhas['__RESPONSAVEL_GSC__']
        df_gsc['Descrição breve do caso'] = df_gsc_linhas['__DESCRICAO_BREVE__']
        df_gsc['Enviado p captação'] = "Sim"
        df_gsc['Data Input'] = datetime.now().strftime('%d/%m/%Y')
        
        # --- ABA NAVEGAÇÃO ---
        df_navegacao = pd.DataFrame(index=df_com_data.index)
        
        df_navegacao['Origem'] = "Forms Eric (Desosp Total Care)"
        df_navegacao['Jornada'] = df_com_data['__PROGRAMA__']
        df_navegacao['Flag'] = ""
        df_navegacao['Data Input'] = datetime.now().strftime('%d/%m/%Y')
        df_navegacao['MO'] = df_com_data['__MO__']
        df_navegacao['Nome do paciente'] = df_com_data['__NOME__']
        df_navegacao['CPF'] = df_com_data['__CPF__']
        df_navegacao['Data de Nascimento'] = df_com_data['__NASCIMENTO__']
        df_navegacao['Idade'] = ""
        df_navegacao['Genero'] = df_com_data['__GENERO__']
        df_navegacao['Estado'] = df_com_data['__ESTADO__']
        df_navegacao['Município'] = df_com_data['__MUNICIPIO__']
        df_navegacao['Bairro'] = ""
        df_navegacao['Nível de Rede'] = df_com_data['__REDE__']
        df_navegacao['Telefone atualizado'] = df_com_data['__TELEFONE__']
        df_navegacao['Hub nvg'] = "CMDP"
        df_navegacao['Enviado nvg?'] = "Não"
        df_navegacao['data envio nvg'] = ""
        df_navegacao['Status nvg'] = ""
        df_navegacao['Data Atualização Status nvg'] = ""
        df_navegacao['Data de Entrada na Linha'] = df_com_data['__HORA_INICIO__']
        df_navegacao['Data da Consulta'] = df_com_data['__DATA_AGENDAMENTO__']
        df_navegacao['Porta de Entrada'] = ""
        
        # --- SALVA NA BASE CONSOLIDADA MANTENDO HISTÓRICO ---
        diretorio_saida = os.path.dirname(caminho_arquivo)
        caminho_saida = os.path.join(diretorio_saida, "base_consolidada_geral.xlsx")
        
        # Lê a base antiga, se existir
        if os.path.exists(caminho_saida):
            try:
                base_cap = pd.read_excel(caminho_saida, sheet_name='captação')
                df_captacao = pd.concat([base_cap, df_captacao], ignore_index=True)
            except Exception:
                pass
                
            try:
                base_nav = pd.read_excel(caminho_saida, sheet_name='navegação')
                df_navegacao = pd.concat([base_nav, df_navegacao], ignore_index=True)
            except Exception:
                pass
            
            try:
                base_saude_mental = pd.read_excel(caminho_saida, sheet_name='Saúde Mental')
                df_saude_mental = pd.concat([base_saude_mental, df_saude_mental], ignore_index=True)
            except Exception:
                pass

            try:
                base_saude_ocupacional = pd.read_excel(caminho_saida, sheet_name='saude ocupacional')
                df_saude_ocupacional = pd.concat([base_saude_ocupacional, df_saude_ocupacional], ignore_index=True)
            except Exception:
                pass
            
            try:
                base_captacao_dp = pd.read_excel(caminho_saida, sheet_name='captação dom pedro')
                df_captacao_dp = pd.concat([base_captacao_dp, df_captacao_dp], ignore_index=True)
            except Exception:
                pass
            
            try:
                base_gsc = pd.read_excel(caminho_saida, sheet_name='GSC')
                df_gsc = pd.concat([base_gsc, df_gsc], ignore_index=True)
            except Exception:
                pass
        
        # Salva o arquivo com todas as abas
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            df_captacao.to_excel(writer, sheet_name='captação', index=False)
            df_navegacao.to_excel(writer, sheet_name='navegação', index=False)
            df_saude_mental.to_excel(writer, sheet_name='Saúde Mental', index=False)
            df_saude_ocupacional.to_excel(writer, sheet_name='saude ocupacional', index=False)
            df_captacao_dp.to_excel(writer, sheet_name='captação dom pedro', index=False)
            df_gsc.to_excel(writer, sheet_name='GSC', index=False)
            
        print("-" * 50)
        print(f"SUCESSO! Dados consolidados salvos e atualizados em:\n{caminho_saida}")
        print("-" * 50)
        
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        print("--- DETALHES DO ERRO PARA O CHAT ---")
        traceback.print_exc()


if __name__ == "__main__":
    # Obtém o diretório correto dependendo se está rodando como script .py ou executável .exe
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))
        
    caminho_xlsx = os.path.join(diretorio_script, "Forms unificado captação CuidadosMil.xlsx")
    caminho_csv = os.path.join(diretorio_script, "Forms unificado captação CuidadosMil.csv")
    
    if os.path.exists(caminho_xlsx):
        processar_forms_eric(caminho_xlsx)
    elif os.path.exists(caminho_csv):
        processar_forms_eric(caminho_csv)
    else:
        print(f"ERRO: Arquivo 'Forms_Cru.xlsx' ou 'Forms_Cru.csv' não encontrado na pasta: {diretorio_script}")
    
    #finalizar
    input("\nPressione Enter para fechar a tela...")
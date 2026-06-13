import pandas as pd
from datetime import datetime
import os
import sys
import unicodedata
import re

def processar_relatorio(caminho_arquivo, caminho_base_historica=None):
    print(f"Lendo arquivo: {caminho_arquivo}")
    
    # Variáveis para armazenar os logs solicitados ao final
    log_novo_programa = 0
    log_programa_igual = 0
    log_agendamento_nao = 0
    log_navegacao = 0
    log_navegacao_emagrecimento_dp = 0
    log_removidos_historico = 0

    try:
        # Carrega o arquivo (suporta .xlsx e .csv)
        if caminho_arquivo.endswith('.csv'):
            try:
                # Tenta ler com utf-8-sig (trata BOM e acentos comuns hoje em dia)
                df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8-sig')
            except UnicodeDecodeError:
                # Se falhar, tenta latin1 (fallback para arquivos mais antigos)
                df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')
        else:
            df = pd.read_excel(caminho_arquivo)
        
        # Remove espaços em branco antes e depois dos nomes das colunas originais
        df.columns = df.columns.str.strip()

        # --- 1. FILTRAGEM ---
        col_status = 'solicitacoes.statusAbertura'
        # Usamos regex para garantir que pegue independente do tipo de traço (- ou –)
        valor_filtro_regex = r'Captado.*Agendamento Realizado'
        
        if col_status not in df.columns:
            print(f"ERRO: Coluna '{col_status}' não encontrada.")
            print(f"Colunas disponíveis: {df.columns.tolist()}")
            return

        # Filtra usando 'contains' (mais flexível) em vez de igualdade exata
        df_filtrado = df[df[col_status].astype(str).str.contains(valor_filtro_regex, regex=True, na=False)].copy()
        print(f"Linhas filtradas: {len(df_filtrado)}")

        if len(df_filtrado) == 0:
            print("Nenhum registro encontrado com o status desejado.")
            return

        # --- 1.1 FILTRO ADICIONAL: AGENDAMENTO REALIZADO ---
        col_agendado = 'O agendamento foi realizado?'
        if col_agendado in df_filtrado.columns:
            # Conta quantas respostas foram "Não" antes de aplicar o filtro
            log_agendamento_nao = df_filtrado[col_agendado].astype(str).str.strip().str.lower().isin(['não', 'nao']).sum()
            # Filtra apenas onde a coluna é 'Sim' (remove espaços extras por segurança)
            df_filtrado = df_filtrado[df_filtrado[col_agendado].astype(str).str.strip() == 'Sim'].copy()
            print(f"Linhas após filtro '{col_agendado} = Sim': {len(df_filtrado)}")
        
        # --- 1.2 LÓGICA: DUPLICAR LINHAS COM SEGUNDO AGENDAMENTO ---
        col_outro_agendamento = 'Agendou em outro programa/linha de cuidado?'
        col_programa_orig = 'O paciente entrou em contato para mais informações/agendamento de qual Programa?'
        col_programa_novo = 'Qual programa?'

        cols_duplicacao = [col_outro_agendamento, col_programa_orig, col_programa_novo, 'Data agendamento', 'Unidade']
        cols_faltantes = [c for c in cols_duplicacao if c not in df_filtrado.columns]

        # Garante que as colunas existam antes de prosseguir
        if not cols_faltantes:
            # Normaliza a coluna de confirmação para 'sim'
            cond_sim = df_filtrado[col_outro_agendamento].astype(str).str.strip().str.lower() == 'sim'
            
            # Garante que o novo programa não seja nulo/vazio
            cond_prog_valido = df_filtrado[col_programa_novo].notna() & (df_filtrado[col_programa_novo].astype(str).str.strip() != '')
            
            # Compara os nomes dos programas (normalizando para evitar falsos negativos)
            cond_prog_diferente = df_filtrado[col_programa_orig].astype(str).str.strip() != df_filtrado[col_programa_novo].astype(str).str.strip()

            # Log estatístico solicitado
            mask_igual = cond_sim & cond_prog_valido & ~cond_prog_diferente
            
            log_novo_programa = (cond_sim & cond_prog_valido & cond_prog_diferente).sum()
            log_programa_igual = mask_igual.sum()
            print(f"\n--- LOG DE SEGUNDO AGENDAMENTO ---")
            print(f"Total 'Agendou em outro programa' = SIM: {cond_sim.sum()}")
            print(f"  > Programas Diferentes (Serão duplicados): {log_novo_programa}")
            print(f"  > Programas Iguais (Ignorados): {mask_igual.sum()}")
            print("-" * 35 + "\n")

            # Máscara final para identificar as linhas que precisam ser duplicadas
            mask_duplicar = cond_sim & cond_prog_valido & cond_prog_diferente
            
            linhas_para_duplicar = df_filtrado[mask_duplicar]
            
            if not linhas_para_duplicar.empty:
                print(f"Encontradas {len(linhas_para_duplicar)} linhas com segundo agendamento para duplicar.")
                
                novas_linhas = []
                for index, linha_original in linhas_para_duplicar.iterrows():
                    nova_linha = linha_original.copy()
                    
                    # Atualiza a nova linha com os dados do segundo agendamento
                    nova_linha[col_programa_orig] = linha_original[col_programa_novo]
                    nova_linha['Qual a data da consulta?'] = linha_original['Data agendamento']
                    nova_linha['Em qual unidade o agendamento foi feito?'] = linha_original['Unidade']
                    
                    novas_linhas.append(nova_linha)
                
                # Adiciona as novas linhas ao dataframe principal
                df_filtrado = pd.concat([df_filtrado, pd.DataFrame(novas_linhas)], ignore_index=True)
                print(f"Total de linhas após duplicação: {len(df_filtrado)}")
        else:
            print("\n" + "!"*50)
            print(f"AVISO: Não foi possível gerar o log de segundo agendamento.\nAs seguintes colunas não foram encontradas no arquivo:\n{cols_faltantes}")
            print("!"*50 + "\n")

        # --- 1.2 LÓGICA: RESPONSÁVEL PELA NAVEGAÇÃO ---
        # Cria a coluna nova no dataframe filtrado (para sair no relatório original também)
        df_filtrado['Hub nvg'] = ""

        # Identifica linhas que NÃO devem ser Dom Pedro (Critérios de exclusão)
        # Excluir se Unidade for Dom Pedro (case=False ignora maiúsculas/minúsculas)
        mask_unidade_excluir = df_filtrado['Em qual unidade o agendamento foi feito?'].astype(str).str.contains(r'CENTRO MEDICO DOM PEDRO|Dom Pedro', case=False, regex=True, na=False)
        # Excluir se Programa for Saúde Mental, Endometriose, Gestante, Fumo Zero, Cefaleia
        mask_programa_excluir = df_filtrado['O paciente entrou em contato para mais informações/agendamento de qual Programa?'].astype(str).str.contains(r'Saúde Mental|Endometriose|Gestante|Fumo Zero|Cefaleia', case=False, regex=True, na=False)

        # Aplica 'Dom Pedro' nas linhas que sobraram (onde NÃO é unidade excluída E NÃO é programa excluído)
        df_filtrado.loc[~(mask_unidade_excluir | mask_programa_excluir), 'Hub nvg'] = 'CMDP'

        # --- 2. TRATAMENTO DE DADOS ---

        # Função auxiliar para formatar datas
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

        # Tratamento: Marca Ótica (9 dígitos)
        # Converte para string, remove decimais (.0) se houver, e preenche com zeros
        df_filtrado['Marca Ótica'] = df_filtrado['Marca Ótica'].fillna(0).astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(9)

        # Tratamento: Datas
        # Data de Criação (remove hora)
        df_filtrado['Data de Criação_Fmt'] = df_filtrado['Data de Criação'].apply(formatar_data)
        # BirthDate (remove hora)
        df_filtrado['beneficiario.birthDate_Fmt'] = df_filtrado['beneficiario.birthDate'].apply(formatar_data)
        # Data da Consulta (formato ISO para dd/mm/aaaa)
        df_filtrado['Qual a data da consulta?_Fmt'] = df_filtrado['Qual a data da consulta?'].apply(formatar_data)

        # Data de hoje para consolidação
        data_hoje = datetime.now().strftime('%d/%m/%Y')

        # --- 1.3 LÓGICA: PREENCHIMENTO DE COLUNAS ADICIONAIS ---
        # (MOVIDO PARA CÁ PARA USAR AS COLUNAS _Fmt JÁ CRIADAS)
        # Inicializa as colunas no dataframe principal
        df_filtrado['data envio ncg'] = ""
        df_filtrado['Enviado nvg?'] = ""

        # REGRA 1: Para as linhas que JÁ receberam "Dom Pedro" pela lógica de exclusão acima
        mask_ja_dom_pedro = df_filtrado['Hub nvg'] == 'CMDP'
        log_navegacao = mask_ja_dom_pedro.sum()
        df_filtrado.loc[mask_ja_dom_pedro, 'Enviado nvg?'] = 'Não'

        # REGRA 2: Específica para Emagrecimento E Unidade Dom Pedro
        mask_emagrecimento = df_filtrado['O paciente entrou em contato para mais informações/agendamento de qual Programa?'].astype(str).str.contains('Emagrecimento', case=False, na=False)
        mask_unidade_dp = df_filtrado['Em qual unidade o agendamento foi feito?'].astype(str).str.contains(r'CENTRO MEDICO DOM PEDRO|Dom Pedro', case=False, regex=True, na=False)
        mask_regra_especifica = mask_emagrecimento & mask_unidade_dp
        log_navegacao_emagrecimento_dp = mask_regra_especifica.sum()

        df_filtrado.loc[mask_regra_especifica, 'Hub nvg'] = 'CMDP'
        df_filtrado.loc[mask_regra_especifica, 'Enviado nvg?'] = 'Não'

        # Preenche os restantes que não receberam 'Dom Pedro' com 'Sem Navegação'
        # Garante que campos vazios ou nulos recebam 'Sem Navegação'
        df_filtrado.loc[(df_filtrado['Hub nvg'] == '') | (df_filtrado['Hub nvg'].isna()), 'Hub nvg'] = 'Sem Navegação'

        # --- 2.1 REMOÇÃO DE DUPLICATAS (LÓGICA FINAL) ---
        col_mo_dup = 'Marca Ótica'
        col_prog_dup = 'O paciente entrou em contato para mais informações/agendamento de qual Programa?'

        # Identifica duplicatas baseadas APENAS em MO e Programa (mantém a primeira ocorrência)
        mask_duplicadas_remover = df_filtrado.duplicated(subset=[col_mo_dup, col_prog_dup], keep='first')
        qtd_removidos = mask_duplicadas_remover.sum()
        
        # Remove as duplicatas
        df_filtrado = df_filtrado[~mask_duplicadas_remover].copy()
        
        # Verifica quantos pacientes restaram com linhas de cuidado diferentes (MOs repetidos)
        qtd_nomes_multiplos = (df_filtrado[col_mo_dup].value_counts() > 1).sum()

        print(f"\n--- LOG DE DUPLICATAS FINAL ---")
        print(f"Registros duplicados removidos (Mesmo MO + Mesmo Programa): {qtd_removidos}")
        print(f"Pacientes com Programas Diferentes (Mantidos): {qtd_nomes_multiplos}")
        print(f"Total de linhas final: {len(df_filtrado)}\n")

        def normalizar_linha_cuidado(nome_bruto):
            if pd.isna(nome_bruto): return ""
            val = str(nome_bruto).upper()
            if 'ANTICOAGULANTE' in val: return 'Anticoagulante Seguro'
            if 'VALVULOPATIA' in val: return 'Cuidado Cardíaco Valvar'
            if 'MAMA' in val: return 'Cuidado Integral da Mama'
            if 'ENDOMETRIOSE' in val: return 'Cuidados para Endometriose'
            if 'INFARTO' in val or 'IAM' in val: return 'Cuidados Pós Infarto'
            if 'EMAGRECIMENTO' in val: return 'Emagrecimento'
            if 'FUMO' in val: return 'Fumo Zero'
            if 'GESTA' in val or 'GESTANTE' in val: return 'Gestação Segura'
            if 'INSUFICI' in val or 'ICC' in val: return 'Insuficiência Cardíaca Controlada'
            if 'AVC' in val: return 'Pós AVC'
            if 'ARRITMIA' in val: return 'Ritmo Certo'
            if 'COLUNA' in val: return 'Saúde da Coluna'
            if 'MENTAL' in val: return 'Saúde Mental'
            if 'RENAL' in val: return 'Saúde Renal'
            return nome_bruto

        # Prepara a Jornada normalizada para usar no relatório final e na validação histórica
        df_filtrado['Jornada_norm'] = df_filtrado['O paciente entrou em contato para mais informações/agendamento de qual Programa?'].apply(normalizar_linha_cuidado)

        def normalizar_unidade(nome_bruto):
            if pd.isna(nome_bruto): return ""
            
            # Remove acentos e joga para maiúsculo
            val = unicodedata.normalize('NFD', str(nome_bruto))
            val = "".join([c for c in val if not unicodedata.combining(c)]).upper().strip()
            
            # Lixo ou preenchimento inválido
            if re.search(r'^X+$', val) or 'MARIANA PENNA' in val: return 'Não Preenchido/Inválido'
            
            # Digital / Telemedicina
            if re.search(r'TELEMEDICINA|VIRTUAL|FUMO ZERO', val): return 'Telemedicina'
            if 'DOM PEDRO' in val: return 'Centro Médico Dom Pedro'
            if 'CONEXA' in val: return 'Conexa'
            if 'ONLINE' in val: return 'Online'
            # Amil Espaço Saúde (AES)
            if 'CAMPO GRANDE' in val: return 'Amil Espaço Saúde - Campo Grande'
            if 'NOVA IGUACU' in val: return 'Amil Espaço Saúde - Nova Iguaçu'
            if 'TIJUCA' in val: return 'Amil Espaço Saúde - Tijuca'
            if 'BOTAFOGO' in val: return 'Amil Espaço Saúde - Botafogo'
            if 'CAXIAS' in val: return 'Amil Espaço Saúde - Caxias'
            if 'GUARULHOS' in val: return 'Amil Espaço Saúde - Guarulhos'
            if 'NITEROI' in val: return 'Amil Espaço Saúde - Niterói'
            if 'OSASCO' in val: return 'Amil Espaço Saúde - Osasco'
            if 'SANTANA' in val: return 'Amil Espaço Saúde - Santana'
            if 'TATUAPE' in val: return 'Amil Espaço Saúde - Tatuapé'
            if 'ANA ROSA' in val: return 'Amil Espaço Saúde - Ana Rosa'
            
            # Unidades Avançadas
            if 'BUTANTA' in val: return 'Unidade Avançada Luz Butantã'
            if 'JOAO DIAS' in val or 'LUZ JOAO' in val: return 'Unidade Avançada Luz João Dias'
            if 'CARLOS CHAGAS' in val: return 'Unidade Avançada Carlos Chagas'
            if 'VITORIA' in val: return 'Unidade Avançada Vitória'
            if 'AVANCADA LUZ' in val: return 'Unidade Avançada Luz'
            
            # Hospitais e Centros Médicos
            if 'IPIRANGA MOGI' in val: return 'Hospital Ipiranga Mogi'
            if 'PAULISTANO' in val: return 'Hospital Paulistano'
            if 'CUBATAO' in val: return 'Hospital da Luz - Cubatão'
            if 'DA LUZ' in val: return 'Hospital da Luz'
            if 'MEDICO JARDIM' in val: return 'AP Centro Médico Jardim'
            if 'JOAO AZEVEDO' in val: return 'AP Centro Médico João Azevedo (SBC)'
            if 'LUIZ FERREIRA' in val: return 'AP Centro Médico Luiz Ferreira (SBC)'
            if 'SAO LUCAS' in val: return 'Hospital São Lucas'
            if 'SAMARITANO HIGIENOPOLIS' in val: return 'Hospital Samaritano Higienópolis'
            if 'SAMARITANO PAULISTA' in val: return 'Hospital Samaritano Paulista'
            if 'ALVORADA' in val: return 'Hospital Alvorada'
            if 'AMERICAS' in val: return 'Hospital Américas'
            if 'SANTA JOANA' in val: return 'Hospital Santa Joana (Recife)'
            
            # Fallback de segurança (Limpa lixo com a palavra UNIDADE caso apareça unidade nova não mapeada)
            fallback = re.sub(r'^(DADOS UNIDADES )?UNIDADE:\s*', '', val)
            return str(nome_bruto).strip() if fallback == val else fallback.title()

        # Cria a coluna Chave (Jornada + MO)
        df_filtrado['Chave'] = df_filtrado['Jornada_norm'].astype(str).str.strip().str.replace(' ', '') + df_filtrado['Marca Ótica']

        # --- 2.2 VERIFICAÇÃO NA BASE HISTÓRICA ---
        if caminho_base_historica and os.path.exists(caminho_base_historica):
            print(f"\n--- VERIFICAÇÃO NA BASE HISTÓRICA ---")
            print(f"Lendo base histórica: {caminho_base_historica}")
            try:
                if caminho_base_historica.endswith('.csv'):
                    df_hist = pd.read_csv(caminho_base_historica, sep=';', encoding='latin1')
                else:
                    df_hist = pd.read_excel(caminho_base_historica)
                
                df_hist.columns = df_hist.columns.str.strip()
                
                mapa_cols_lower = {str(c).strip().lower(): c for c in df_hist.columns}
                
                possiveis_mo = ['mo', 'marca ótica', 'marca otica', 'marca_otica']
                col_mo_hist = next((mapa_cols_lower[k] for k in possiveis_mo if k in mapa_cols_lower), None)
                
                possiveis_programas = ['jornada', 'programa', 'linha de cuidado', 'o paciente entrou em contato para mais informações/agendamento de qual programa?']
                col_prog_hist = next((mapa_cols_lower[k] for k in possiveis_programas if k in mapa_cols_lower), None)
                
                if col_mo_hist and col_prog_hist:
                    # Garante que a MO da base histórica também esteja formatada como 9 dígitos
                    df_hist[col_mo_hist] = df_hist[col_mo_hist].astype(object)
                    df_hist[col_mo_hist] = df_hist[col_mo_hist].apply(lambda x: 0 if pd.isna(x) or str(x).strip().lower() in ['nan', 'nat', ''] else x)
                    df_hist[col_mo_hist] = df_hist[col_mo_hist].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(9)
                    
                    # Chave no Histórico: Jornada (sem espaços) + MO
                    chave_hist = df_hist[col_prog_hist].astype(str).str.strip().str.replace(' ', '') + df_hist[col_mo_hist]
                    chaves_existentes = set(chave_hist)
                    
                    mask_ja_existe = df_filtrado['Chave'].isin(chaves_existentes)
                    log_removidos_historico = mask_ja_existe.sum()
                    
                    if log_removidos_historico > 0:
                        print(f"ATENÇÃO: {log_removidos_historico} pacientes já existem na base histórica e serão removidos deste processamento.")
                        df_filtrado = df_filtrado[~mask_ja_existe].copy()
                    else:
                        print("Nenhum paciente duplicado encontrado na base histórica.")
                        
                else:
                    print("AVISO: Colunas 'MO' ou 'Jornada' não encontradas na base histórica.")
                    print(f"Colunas encontradas no arquivo: {df_hist.columns.tolist()}")
            except Exception as e:
                print(f"Erro ao ler/processar base histórica: {e}")
            print("-" * 35 + "\n")

        # --- 3. MAPEAMENTO DE COLUNAS (NOVA ESTRUTURA) ---
        novo_df = pd.DataFrame(index=df_filtrado.index)

        novo_df['Origem'] = "Amil 360"
        novo_df['Jornada'] = df_filtrado['Jornada_norm']
        novo_df['Flag'] = ""
        novo_df['Data Input'] = data_hoje
        novo_df['MO'] = df_filtrado['Marca Ótica']
        novo_df['Nome do paciente'] = df_filtrado['beneficiario.nome']
        novo_df['CPF'] = df_filtrado['beneficiario.cpf']
        novo_df['Data de Nascimento'] = df_filtrado['beneficiario.birthDate_Fmt']
        novo_df['Idade'] = ""
        novo_df['Genero'] = df_filtrado['beneficiario.sexo.nome']
        novo_df['Estado'] = df_filtrado['beneficiario.homeState'] if 'beneficiario.homeState' in df_filtrado.columns else ""
        novo_df['Município'] = df_filtrado['beneficiario.homeCity'] if 'beneficiario.homeCity' in df_filtrado.columns else ""
        novo_df['Bairro'] = df_filtrado['beneficiario.homeNeighborhood'] if 'beneficiario.homeNeighborhood' in df_filtrado.columns else ""
        novo_df['Nível de Rede'] = ""
        novo_df['Telefone atualizado'] = df_filtrado['beneficiario.contatos.telefoneCompleto']
        novo_df['Hub nvg'] = df_filtrado['Hub nvg']
        novo_df['Enviado nvg?'] = df_filtrado['Enviado nvg?']
        novo_df['data envio ncg'] = df_filtrado['data envio ncg']
        novo_df['Status nvg'] = ""
        novo_df['Data Atualização Status nvg'] = ""
        novo_df['Data de Entrada na Linha'] = df_filtrado['Data de Criação_Fmt']
        novo_df['Data da Consulta'] = df_filtrado['Qual a data da consulta?_Fmt']
        novo_df['Porta de Entrada'] = df_filtrado['Em qual unidade o agendamento foi feito?'].apply(normalizar_unidade)
        novo_df['Chave'] = df_filtrado['Chave']

        # --- 4. EXPORTAÇÃO ---
        data_str = datetime.now().strftime('%Y%m%d')
        diretorio_saida = os.path.dirname(caminho_arquivo)

        # 1. Relatório com Layout Novo (Colunas organizadas)
        nome_saida_novo = f"relatorio_formatado_{data_str}.xlsx"
        caminho_saida_novo = os.path.join(diretorio_saida, nome_saida_novo)
        novo_df.to_excel(caminho_saida_novo, index=False)

        print("-" * 50)
        print(f"SUCESSO! Arquivo gerado na pasta:")
        print(f"Layout Novo: {nome_saida_novo}")
        print(f"Total de registros: {len(novo_df)}")
        print("-" * 50)

        # --- 5. ATUALIZAR BASE HISTÓRICA CONSOLIDADA ---
        if caminho_base_historica and os.path.exists(caminho_base_historica) and not novo_df.empty:
            print(f"\nAtualizando base histórica: {caminho_base_historica}")
            try:
                if caminho_base_historica.endswith('.csv'):
                    df_hist_atual = pd.read_csv(caminho_base_historica, sep=';', encoding='latin1')
                else:
                    df_hist_atual = pd.read_excel(caminho_base_historica)
                
                # Normaliza a coluna MO da base histórica atual para 9 dígitos antes de salvar
                mapa_cols_lower_atual = {str(c).strip().lower(): c for c in df_hist_atual.columns}
                possiveis_mo = ['mo', 'marca ótica', 'marca otica', 'marca_otica']
                col_mo_hist_atual = next((mapa_cols_lower_atual[k] for k in possiveis_mo if k in mapa_cols_lower_atual), None)
                
                if col_mo_hist_atual:
                    df_hist_atual[col_mo_hist_atual] = df_hist_atual[col_mo_hist_atual].astype(object)
                    df_hist_atual[col_mo_hist_atual] = df_hist_atual[col_mo_hist_atual].apply(lambda x: 0 if pd.isna(x) or str(x).strip().lower() in ['nan', 'nat', ''] else x)
                    df_hist_atual[col_mo_hist_atual] = df_hist_atual[col_mo_hist_atual].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(9)
                
                # Concatena os novos registros diretamente abaixo do histórico atual
                df_hist_novo = pd.concat([df_hist_atual, novo_df], ignore_index=True)
                
                # Garante que 'Origem' seja a primeira coluna na base histórica
                cols_hist = df_hist_novo.columns.tolist()
                if 'Origem' in cols_hist:
                    cols_hist.insert(0, cols_hist.pop(cols_hist.index('Origem')))
                    df_hist_novo = df_hist_novo[cols_hist]
                    
                # Preenche a origem para os pacientes antigos da base histórica que estavam com essa coluna vazia
                df_hist_novo['Origem'] = df_hist_novo['Origem'].fillna("Amil 360")
                
                # Garante que todas as colunas de data da base histórica fiquem estritamente em DD/MM/AAAA
                colunas_data_historico = [
                    'Data Input', 
                    'Data de Nascimento', 
                    'data envio ncg', 
                    'Data Atualização Status nvg', 
                    'Data de Entrada na Linha', 
                    'Data da Consulta'
                ]
                for col in colunas_data_historico:
                    if col in df_hist_novo.columns:
                        df_hist_novo[col] = df_hist_novo[col].apply(formatar_data)
                
                # Salva o arquivo consolidado
                if caminho_base_historica.endswith('.csv'):
                    df_hist_novo.to_csv(caminho_base_historica, sep=';', index=False, encoding='latin1')
                else:
                    df_hist_novo.to_excel(caminho_base_historica, index=False)
                
                print(f"SUCESSO! Novos pacientes adicionados na base histórica. Total agora: {len(df_hist_novo)}")
            except Exception as e:
                print(f"Erro ao atualizar base histórica: {e}")
            print("-" * 50)

        # --- LOG FINAL SOLICITADO ---
        print("\n" + "="*50)
        print("RESUMO DOS LOGS SOLICITADOS:")
        print(f"novo programa: {log_novo_programa}")
        print(f"programa igual: {log_programa_igual}")
        print(f"resposta não no agendamento foi realizado: {log_agendamento_nao}")
        print(f"após filtrar programas que tem navegação: {log_navegacao}")
        print(f"após filtrar programas que tem navegação mas é emagrecimento dom pedro: {log_navegacao_emagrecimento_dp}")
        print(f"remoção de duplicados na base histórica: {log_removidos_historico}")
        print("="*50 + "\n")

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")

if __name__ == "__main__":
    # Obtém o diretório correto dependendo se está rodando como script (.py) ou executável (.exe)
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))
    
    caminho_xlsx = os.path.join(diretorio_script, "relatorio.xlsx")
    caminho_csv = os.path.join(diretorio_script, "relatorio.csv")
    
    # Define o caminho da base histórica (agora busca na mesma pasta do programa)
    caminho_historico = os.path.join(diretorio_script, "base.xlsx")

    if os.path.exists(caminho_xlsx):
        processar_relatorio(caminho_xlsx, caminho_historico)
    elif os.path.exists(caminho_csv):
        processar_relatorio(caminho_csv, caminho_historico)
    else:
        print(f"ERRO: Arquivo 'relatorio.xlsx' ou 'relatorio.csv' não encontrado na pasta: {diretorio_script}")
        
    input("\nPressione Enter para fechar a tela...")
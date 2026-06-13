import pandas as pd
import os
import glob
import sys
import re
from datetime import datetime


def normalizar_texto(valor):
    """
    Padroniza textos para comparação:
    - remove espaços antes/depois
    - deixa maiúsculo
    - remove espaços duplicados
    """
    if pd.isna(valor):
        return ""

    valor = str(valor).strip().upper()
    valor = re.sub(r"\s+", " ", valor)

    return valor


def preparar_relatorio_outro_programa(df):
    """
    Regra:
    Se 'Agendou em outro programa/linha de cuidado?' = SIM
    e o programa principal for diferente de 'Qual programa?',
    duplica a linha.

    Na linha duplicada, troca:
    'O paciente entrou em contato para mais informações/agendamento de qual Programa?'
    pelo valor de:
    'Qual programa?'
    """

    coluna_agendou_outro = "Agendou em outro programa/linha de cuidado?"
    coluna_programa = "O paciente entrou em contato para mais informações/agendamento de qual Programa?"
    coluna_qual_programa = "Qual programa?"

    colunas_necessarias = [
        coluna_agendou_outro,
        coluna_programa,
        coluna_qual_programa
    ]

    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            print(f"ATENÇÃO: A coluna '{coluna}' não foi encontrada.")
            print("A regra de duplicar linhas por outro programa será ignorada.")
            return df

    print("\nAplicando regra de outro programa/linha de cuidado...")

    df = df.copy()

    # Colunas auxiliares para auditoria
    df["PROGRAMA_ORIGINAL_ANTES_DUPLICAR"] = df[coluna_programa]
    df["COMPARACAO_PROGRAMA_X_QUAL_PROGRAMA"] = ""
    df["LINHA_DUPLICADA_OUTRO_PROGRAMA"] = "NÃO"

    # Identifica linhas com SIM
    mascara_sim = (
        df[coluna_agendou_outro]
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("SIM")
    )

    linhas_sim = df[mascara_sim].copy()

    if linhas_sim.empty:
        print("Nenhuma linha com 'Agendou em outro programa/linha de cuidado?' = SIM.")
        return df

    # Normaliza os dois programas para comparar corretamente
    programa_original_normalizado = linhas_sim[coluna_programa].apply(normalizar_texto)
    qual_programa_normalizado = linhas_sim[coluna_qual_programa].apply(normalizar_texto)

    # Marca IGUAL ou DIFERENTE
    linhas_sim["COMPARACAO_PROGRAMA_X_QUAL_PROGRAMA"] = [
        "IGUAL" if prog == qual else "DIFERENTE"
        for prog, qual in zip(programa_original_normalizado, qual_programa_normalizado)
    ]

    # Atualiza essa informação nas linhas originais
    df.loc[
        linhas_sim.index,
        "COMPARACAO_PROGRAMA_X_QUAL_PROGRAMA"
    ] = linhas_sim["COMPARACAO_PROGRAMA_X_QUAL_PROGRAMA"]

    # Pega somente os casos diferentes
    programas_diferentes = linhas_sim[
        linhas_sim["COMPARACAO_PROGRAMA_X_QUAL_PROGRAMA"] == "DIFERENTE"
    ].copy()

    # Não duplica se "Qual programa?" estiver vazio
    programas_diferentes = programas_diferentes[
        programas_diferentes[coluna_qual_programa].notna()
        & (programas_diferentes[coluna_qual_programa].astype(str).str.strip() != "")
    ].copy()

    # Cria as linhas duplicadas
    duplicadas = programas_diferentes.copy()

    # Marca como linha duplicada
    duplicadas["LINHA_DUPLICADA_OUTRO_PROGRAMA"] = "SIM"

    # Na linha duplicada, o programa principal vira o "Qual programa?"
    duplicadas[coluna_programa] = duplicadas[coluna_qual_programa]

    # Junta original + duplicadas
    df_expandido = pd.concat(
        [df, duplicadas],
        ignore_index=True
    )

    print(f"Linhas com SIM encontradas: {len(linhas_sim)}")
    print(f"Linhas com programas diferentes: {len(programas_diferentes)}")
    print(f"Linhas duplicadas criadas: {len(duplicadas)}")

    return df_expandido


def consolidar_relatorios(pasta_entrada, caminho_saida_total, caminho_saida_limpo):
    print(f"Buscando arquivos na pasta: {pasta_entrada}")
    
    # Busca todos os arquivos Excel e CSV na pasta de entrada
    arquivos_excel = glob.glob(os.path.join(pasta_entrada, "*.xlsx"))
    arquivos_csv = glob.glob(os.path.join(pasta_entrada, "*.csv"))
    todos_arquivos = arquivos_excel + arquivos_csv
    
    if not todos_arquivos:
        print("Nenhum arquivo encontrado. Verifique se os relatórios estão na pasta correta.")
        return
        
    lista_dataframes = []
    
    for arquivo in todos_arquivos:
        print(f"Lendo: {os.path.basename(arquivo)}")
        try:
            # Identifica a extensão para usar o motor de leitura correto
            if arquivo.endswith('.csv'):
                try:
                    df = pd.read_csv(arquivo, sep=';', encoding='utf-8-sig')
                except UnicodeDecodeError:
                    df = pd.read_csv(arquivo, sep=';', encoding='latin1')
            else:
                df = pd.read_excel(arquivo)
            
            # Remove espaços em branco do início e fim dos nomes das colunas
            df.columns = df.columns.str.strip()
            
            lista_dataframes.append(df)

        except Exception as e:
            print(f"Erro ao ler o arquivo {arquivo}: {e}")
            
    if lista_dataframes:
        print("\nEmpilhando todos os relatórios...")

        # Junta todos os relatórios em uma única base
        df_consolidado = pd.concat(lista_dataframes, ignore_index=True)

        print(f"Total de linhas após empilhar os relatórios: {len(df_consolidado)}")

        # ============================================================
        # REGRA NOVA — ANTES DE TUDO
        # Duplica linha somente quando:
        # 1. Agendou em outro programa/linha de cuidado? = SIM
        # 2. Programa principal diferente de Qual programa?
        # ============================================================
        df_consolidado = preparar_relatorio_outro_programa(df_consolidado)

        print(f"Total de linhas após aplicar regra de outro programa: {len(df_consolidado)}")
        
        # Formata a coluna 'Data de Criação'
        if 'Data de Criação' in df_consolidado.columns:
            datas_texto = (
                df_consolidado['Data de Criação']
                .astype(str)
                .str.split(',')
                .str[0]
                .str.split(' ')
                .str[0]
                .str.strip()
            )

            df_consolidado['Data de Criação'] = pd.to_datetime(
                datas_texto,
                errors='coerce',
                dayfirst=True
            ).dt.date
        
        linhas_antes = len(df_consolidado)
        print(f"Total de linhas antes da remoção de duplicatas exatas: {linhas_antes}")
        
        # Remove somente linhas completamente duplicadas
        df_total = df_consolidado.drop_duplicates().copy()
        
        duplicatas_exatas = linhas_antes - len(df_total)
        print(f"Total de duplicatas exatas removidas: {duplicatas_exatas}")
        print(f"Total de linhas no relatório TOTAL: {len(df_total)}")

        # Formata a Marca Ótica para 9 dígitos
        col_mo = 'Marca Ótica'

        if col_mo in df_total.columns:
            df_total[col_mo] = df_total[col_mo].astype(object)

            df_total[col_mo] = df_total[col_mo].apply(
                lambda x: 0 if pd.isna(x) or str(x).strip().lower() in ['nan', 'nat', ''] else x
            )

            df_total[col_mo] = (
                df_total[col_mo]
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .str.zfill(9)
            )

        # Cria a Chave Primária usando Programa + Marca Ótica
        col_programa = 'O paciente entrou em contato para mais informações/agendamento de qual Programa?'

        if col_mo in df_total.columns and col_programa in df_total.columns:
            df_total['Chave Primária'] = (
                df_total[col_programa]
                .astype(str)
                .str.strip()
                .str.upper()
                + "_"
                + df_total[col_mo]
            )

        # Salva o relatório TOTAL
        # Esse relatório remove apenas duplicatas exatas.
        # Ele mantém registros com mesmo código/Marca Ótica e programas diferentes.
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

        # Ordena pela 'Data de Criação'
        # A mais antiga fica no topo
        if 'Data de Criação' in df_limpo.columns:
            df_limpo = df_limpo.sort_values(
                by='Data de Criação',
                ascending=True
            )
            
        # Remove duplicadas pela Chave Primária, mantendo a mais antiga
        if 'Chave Primária' in df_limpo.columns:
            linhas_antes_chave = len(df_limpo)

            df_limpo.drop_duplicates(
                subset=['Chave Primária'],
                keep='first',
                inplace=True
            )

            duplicatas_chave = linhas_antes_chave - len(df_limpo)

            print(
                f"Total de duplicatas de Chave Primária removidas "
                f"no relatório LIMPO, mantendo a mais antiga: {duplicatas_chave}"
            )
            
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


if __name__ == "__main__":
    # === CONFIGURAÇÕES DE PASTAS ===

    # Diretório onde o script está rodando
    # com suporte para o executável .exe
    if getattr(sys, 'frozen', False):
        # Se estiver rodando como arquivo compilado .exe
        diretorio_atual = os.path.dirname(sys.executable)
    else:
        # Se estiver rodando como script normal .py
        diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    
    # Pasta onde você deve colocar os relatórios baixados
    pasta_relatorios = os.path.join(diretorio_atual, "relatorios_entrada")

    if not os.path.exists(pasta_relatorios):
        os.makedirs(pasta_relatorios)

        print(
            f"Pasta '{pasta_relatorios}' criada! "
            f"Coloque seus relatórios lá dentro e rode o script novamente."
        )

    else:
        # Nome dos arquivos finais com data e hora atuais
        data_hoje = datetime.now().strftime('%Y%m%d_%H%M')

        arquivo_saida_total = os.path.join(
            diretorio_atual,
            f"relatorioTOTAL_{data_hoje}.xlsx"
        )

        arquivo_saida_limpo = os.path.join(
            diretorio_atual,
            f"relatorioLIMPO_{data_hoje}.xlsx"
        )
        
        consolidar_relatorios(
            pasta_relatorios,
            arquivo_saida_total,
            arquivo_saida_limpo
        )
        
    input("\nPressione Enter para fechar...")
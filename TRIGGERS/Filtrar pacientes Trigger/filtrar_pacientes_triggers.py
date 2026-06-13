import pandas as pd
import unicodedata
import os
import sys
from datetime import datetime

# Função para remover acentos e normalizar texto
def remover_acentos(texto):
    if not isinstance(texto, str):
        return texto
    # Normaliza para 'NFD' que separa acentos dos caracteres base
    nfkd_form = unicodedata.normalize('NFD', texto)
    # Retorna a string apenas com caracteres que não são de combinação
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

# --- DICIONÁRIO DE REGRAS (LINHAS DE CUIDADO) ---
# Estas regras definem cidades, idade mínima e rede máxima para cada linha.
REGRAS = {
    "Pos IAM": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "RIO DE JANEIRO", "BRASILIA"],
        "idademin": 18, "idademax": 95, "rede": 700
    },
    "Emagrecimento": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "RIO DE JANEIRO", "NITEROI", "NOVA IGUAÇU"],
        "idademin": 18, "idademax": 95, "rede": 700
    },
    "Ritmo Certo": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "RIO DE JANEIRO", "BRASILIA"],
        "idademin": 18, "idademax": 95, "rede": 700
    },
    "Cuidado Cardíaco Valvar": {
        "cidades": ["SÃO PAULO", "RIO DE JANEIRO", "BRASILIA"],
        "idademin": 18, "idademax": 95, "rede": 700
    },
    "Insuficiência Cardíaca Controlada": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "RIO DE JANEIRO", "BRASILIA"],
        "idademin": 18, "idademax": 95, "rede": 700
    },
    "Saúde da Coluna": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "MOGI DAS CRUZES", 
                    "RIO DE JANEIRO", "NITEROI", "NOVA IGUAÇU", 
                    "RECIFE", "OLINDA", "PAULISTA", "CAMARAGIBE", "JABOATÃO DOS GUARARAPES"],
        "idademin": 18, "idademax": 95, "rede": 700
    },
    "Gestação Segura": {
        "cidades": [], # Lista vazia = Brasil Todo (Sem filtro de cidade)
        "idademin": 14, "idademax": 95, "rede": 700
    },
    "Cuidados para Endometriose": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL"],
        "idademin": 12, "idademax": 50, "rede": 700
    },
    "Saúde Mental": {
        "cidades": [], # BR Todo
        "idademin": 12, "idademax": 95, "rede": 700
    },
    "Saúde Renal": {
        "cidades": ["SÃO PAULO"],
        "idademin": 18, "idademax": 95, "rede": 700
    },
    # Oncologia
    "Cuidado Integral da Mama": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "DIADEMA", "EMBU DAS ARTES", "TABOÃO DA SERRA", "ITAPECERICA", "ITAPECERICA DA SERRA", "OSASCO"],
        "idademin": 18, "idademax": 95, "rede": 750
    },
    "Cuidado Integral da Próstata": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "DIADEMA", "EMBU DAS ARTES", "TABOÃO DA SERRA", "ITAPECERICA", "ITAPECERICA DA SERRA", "OSASCO"],
        "idademin": 18, "idademax": 95, "rede": 750
    },
    "Cuidado Oncológico Colorretal": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "DIADEMA", "EMBU DAS ARTES", "TABOÃO DA SERRA", "ITAPECERICA", "ITAPECERICA DA SERRA", "OSASCO"],
        "idademin": 18, "idademax": 95, "rede": 750
    },
    "Cuidado Oncológico Pulmonar": {
        "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "DIADEMA", "EMBU DAS ARTES", "TABOÃO DA SERRA", "ITAPECERICA", "ITAPECERICA DA SERRA", "OSASCO"],
        "idademin": 18, "idademax": 95, "rede": 750
    }
}

def filtrar_pacientes_total(caminho_arquivo):
    print(f"Processando arquivo: {caminho_arquivo}")
    
    try:
        # Carregar arquivo
        if caminho_arquivo.endswith('.csv'):
            df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')
        else:
            df = pd.read_excel(caminho_arquivo)
        
        # Normalizar nomes das colunas
        df.columns = df.columns.str.strip().str.lower()
        
        colunas_necessarias = ['municipio', 'idade', 'nivel_rede', 'linha_cuidado']
        for col in colunas_necessarias:
            if col not in df.columns:
                print(f"ERRO: Coluna obrigatória '{col}' não encontrada no arquivo.")
                return

        # Normalizar dados auxiliares para comparação
        df['municipio_norm'] = df['municipio'].astype(str).str.upper().str.strip().apply(remover_acentos)
        # Garantir que linha_cuidado seja string e sem espaços extras nas pontas
        df['linha_cuidado_clean'] = df['linha_cuidado'].astype(str).str.strip()
        
        # Converter idade e nivel_rede para números para evitar erros de comparação com textos
        df['idade'] = pd.to_numeric(df['idade'], errors='coerce').fillna(0)
        df['nivel_rede'] = pd.to_numeric(df['nivel_rede'], errors='coerce').fillna(9999)
        
        # Criar coluna de elegibilidade inicializada como False
        df['elegivel'] = False
        
        print("Aplicando regras por Linha de Cuidado...")
        print("-" * 50)

        # Iterar sobre cada regra definida
        for nome_linha, regra in REGRAS.items():
            # Identificar linhas no DataFrame que pertencem a esta regra
            mask_linha = df['linha_cuidado_clean'] == nome_linha
            
            if not mask_linha.any():
                continue # Nenhum paciente desta linha encontrado, pula
            
            # 1. Filtro de Cidades
            if regra["cidades"]:
                cidades_validas = [remover_acentos(c) for c in regra["cidades"]]
                mask_cidade = df.loc[mask_linha, 'municipio_norm'].isin(cidades_validas)
            else:
                mask_cidade = True # Todos aceitos se lista vazia
            
            # 2. Filtro de Idade
            mask_idade = (df.loc[mask_linha, 'idade'] >= regra["idademin"]) & (df.loc[mask_linha, 'idade'] <= regra["idademax"])
            
            # 3. Filtro de Rede
            mask_rede = df.loc[mask_linha, 'nivel_rede'] <= regra["rede"]
            
            # Combina filtros para este grupo
            mask_aprovados = mask_cidade & mask_idade & mask_rede
            
            # Contagem para o log
            total_linha = mask_linha.sum()
            total_elegiveis = mask_aprovados.sum()
            print(f"Linha: {nome_linha:<35} | Total: {total_linha:<5} | Elegíveis: {total_elegiveis}")
            
            # Atualiza a coluna 'elegivel' apenas para os aprovados desta linha
            df.loc[mask_linha, 'elegivel'] = mask_aprovados

        print("-" * 50)
        # Filtrar o DataFrame final
        df_elegiveis = df[df['elegivel'] == True].copy()
        
        print(f"Sucesso! {len(df_elegiveis)} pacientes elegíveis encontrados.")
        
        if df_elegiveis.empty:
            print("Nenhum paciente elegível para exportar.")
            return
            
        # --- MAPEAMENTO PARA O NOVO FORMATO PADRÃO DE CAPTAÇÃO ---
        df_captacao = pd.DataFrame(index=df_elegiveis.index)
        
        # Função auxiliar para buscar colunas no df original (que estão em minúsculo)
        def obter_coluna(nomes_possiveis):
            for nome in nomes_possiveis:
                if nome in df_elegiveis.columns:
                    return df_elegiveis[nome]
            return ""

        df_captacao['FONTE'] = "Base MCO Triggers"
        df_captacao['CAMINHO'] = "Base MCO Triggers"
        df_captacao['DATA INSERÇÃO'] = datetime.now().strftime('%d/%m/%Y')
        df_captacao['LINHA DE CUIDADO'] = df_elegiveis['linha_cuidado']
        df_captacao['Nome'] = obter_coluna(['nome_beneficiario'])
        df_captacao['MO'] = obter_coluna(['marca otica'])
        df_captacao['CPF'] = obter_coluna(['cpf'])
        df_captacao['DATA DE NASCIMENTO'] = obter_coluna(['data_nascimento'])
        df_captacao['NÍVEL DE REDE'] = df_elegiveis['nivel_rede']
        df_captacao['UF'] = obter_coluna(['uf'])
        df_captacao['MUNICÍPIO'] = df_elegiveis['municipio']
        df_captacao['TELEFONE'] = obter_coluna(['telefone beneficiário'])
        df_captacao['STATUS'] = ""
        
        diretorio_saida = os.path.dirname(caminho_arquivo)
        
        # 1. Exportar o arquivo diário formatado (Arquivo 1)
        arquivo_saida = os.path.join(diretorio_saida, "pacientes_elegiveis_TOTAL.xlsx")
        df_captacao.to_excel(arquivo_saida, index=False)
        print(f"Arquivo 1 (Diário Formato Captação) salvo: {arquivo_saida}")
        
        # 2. Atualizar/Criar o Tabelão Histórico Formatado (Arquivo 2)
        arquivo_tabelao = os.path.join(diretorio_saida, "tabelao_pacientes_elegiveis.xlsx")
        if os.path.exists(arquivo_tabelao):
            df_tabelao = pd.read_excel(arquivo_tabelao)
            # Concatena os novos registros ao tabelão mantendo o histórico
            df_tabelao = pd.concat([df_tabelao, df_captacao], ignore_index=True)
        else:
            df_tabelao = df_captacao.copy()
            
        df_tabelao.to_excel(arquivo_tabelao, index=False)
        print(f"Arquivo 2 (Tabelão) atualizado e salvo: {arquivo_tabelao} | Total Acumulado: {len(df_tabelao)}")
        
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    # Obtém o diretório correto dependendo se está rodando como script (.py) ou executável (.exe)
    if getattr(sys, 'frozen', False):
        diretorio_script = os.path.dirname(sys.executable)
    else:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))
        
    caminho_xlsx = os.path.join(diretorio_script, "data.xlsx")
    caminho_csv = os.path.join(diretorio_script, "data.csv")
    
    if os.path.exists(caminho_xlsx):
        filtrar_pacientes_total(caminho_xlsx)
    elif os.path.exists(caminho_csv):
        filtrar_pacientes_total(caminho_csv)
    else:
        print(f"ERRO: Arquivo 'data.xlsx' ou 'data.csv' não encontrado na pasta:\n{diretorio_script}")
        
    input("\nPressione Enter para fechar a tela...")

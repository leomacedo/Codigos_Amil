import pandas as pd
import unicodedata
import os
import sys
from datetime import datetime

def remover_acentos(texto):
    if not isinstance(texto, str):
        return texto
    nfkd_form = unicodedata.normalize('NFD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

REGRA_ONCOLOGIA = {
    "cidades": ["SÃO PAULO", "GUARULHOS", "SANTO ANDRÉ", "SÃO BERNARDO DO CAMPO", "SÃO CAETANO DO SUL", "DIADEMA", "EMBU DAS ARTES", "TABOÃO DA SERRA", "ITAPECERICA", "ITAPECERICA DA SERRA", "OSASCO"],
    "idade": 18,
    "rede": 750
}

REGRAS = {
    "Cuidado Integral da Mama": REGRA_ONCOLOGIA,
    "Cuidado Integral da Próstata": REGRA_ONCOLOGIA,
    "Cuidado Oncológico Colorretal": REGRA_ONCOLOGIA,
    "Cuidado Oncológico Pulmonar": REGRA_ONCOLOGIA
}

def filtrar_pacientes_total(caminho_arquivo):
    print(f"Processando arquivo: {caminho_arquivo}")

    try:
        if caminho_arquivo.endswith('.csv'):
            df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')
        else:
            df = pd.read_excel(caminho_arquivo)

        df.columns = df.columns.str.strip().str.lower()

        colunas_necessarias = ['municipio', 'idade', 'nivel_rede', 'linha_cuidado']
        for col in colunas_necessarias:
            if col not in df.columns:
                print(f"ERRO: Coluna obrigatória '{col}' não encontrada no arquivo.")
                return

        df['municipio_norm'] = df['municipio'].astype(str).str.upper().str.strip().apply(remover_acentos)
        df['linha_cuidado_clean'] = df['linha_cuidado'].astype(str).str.strip()
        df['elegivel'] = False

        print("Aplicando regras por Linha de Cuidado...")
        print("-" * 50)

        for nome_linha, regra in REGRAS.items():
            mask_linha = df['linha_cuidado_clean'] == nome_linha

            if not mask_linha.any():
                continue

            if regra["cidades"]:
                cidades_validas = [remover_acentos(c) for c in regra["cidades"]]
                mask_cidade = df.loc[mask_linha, 'municipio_norm'].isin(cidades_validas)
            else:
                mask_cidade = True

            mask_idade = df.loc[mask_linha, 'idade'] >= regra["idade"]
            mask_rede = df.loc[mask_linha, 'nivel_rede'] <= regra["rede"]
            mask_aprovados = mask_cidade & mask_idade & mask_rede

            total_linha = mask_linha.sum()
            total_elegiveis = mask_aprovados.sum()

            print(f"Linha: {nome_linha:<35} | Total: {total_linha:<5} | Elegíveis: {total_elegiveis}")

            df.loc[mask_linha, 'elegivel'] = mask_aprovados

        print("-" * 50)

        df_elegiveis = df[df['elegivel'] == True].copy()

        print(f"Sucesso! {len(df_elegiveis)} pacientes elegíveis encontrados.")

        if df_elegiveis.empty:
            print("Nenhum paciente elegível para exportar.")
            return

        def obter_coluna(nomes_possiveis):
            for nome in nomes_possiveis:
                nome = nome.lower().strip()
                if nome in df_elegiveis.columns:
                    return df_elegiveis[nome]
            return ""

        data_hoje = datetime.now().strftime('%d/%m/%Y')

        # ABA DOM PEDRO
        df_dom_pedro = pd.DataFrame(index=df_elegiveis.index)

        df_dom_pedro['Linha de Cuidado'] = df_elegiveis['linha_cuidado']
        df_dom_pedro['Prioridade no Contato?'] = "Sim"
        df_dom_pedro['Data Envio Paciente'] = data_hoje
        df_dom_pedro['Motivo Envio'] = "Captação + Navegação"
        df_dom_pedro['Origem'] = "Triggers"
        df_dom_pedro['MO'] = obter_coluna(['marca otica', 'marca ótica', 'mo', 'numero da carteirinha', 'número da carteirinha'])
        df_dom_pedro['Nome do paciente'] = obter_coluna(['nome_beneficiario', 'nome beneficiario', 'nome beneficiário', 'nome', 'nome do paciente'])
        df_dom_pedro['CPF'] = obter_coluna(['cpf', 'cpf do paciente'])
        df_dom_pedro['Data de Nascimento'] = obter_coluna(['data_nascimento', 'data nascimento', 'data de nascimento'])
        df_dom_pedro['Idade'] = " "
        df_dom_pedro['Genero'] = obter_coluna(['genero', 'gênero', 'sexo'])
        df_dom_pedro['Estado'] = obter_coluna(['uf', 'estado', 'regional'])
        df_dom_pedro['Município'] = df_elegiveis['municipio']
        df_dom_pedro['Bairro'] = obter_coluna(['bairro'])
        df_dom_pedro['Nível de Rede'] = df_elegiveis['nivel_rede']
        df_dom_pedro['Telefone atualizado'] = obter_coluna(['telefone beneficiário', 'telefone beneficiario', 'telefone atualizado', 'telefone atualizado com ddd', 'telefone', 'celular'])

        # ABA CAPTAÇÃO
        df_captacao = pd.DataFrame(index=df_elegiveis.index)

        df_captacao['FONTE'] = "FTriggers"
        df_captacao['CAMINHO'] = "Dom Pedro"
        df_captacao['DATA INSERÇÃO'] = data_hoje
        df_captacao['LINHA DE CUIDADO'] = df_elegiveis['linha_cuidado']
        df_captacao['Nome'] = obter_coluna(['nome completo do paciente (sem abreviações)', 'nome completo do paciente', 'nome_beneficiario', 'nome beneficiario', 'nome beneficiário', 'nome do paciente', 'nome'])
        df_captacao['MO'] = obter_coluna(['número da carteirinha (marca ótica)', 'numero da carteirinha (marca otica)', 'marca otica', 'marca ótica', 'mo', 'numero da carteirinha', 'número da carteirinha'])
        df_captacao['CPF'] = obter_coluna(['cpf do paciente', 'cpf'])
        df_captacao['DATA DE NASCIMENTO'] = obter_coluna(['data de nascimento_fmt', 'data_nascimento', 'data nascimento', 'data de nascimento'])
        df_captacao['NÍVEL DE REDE'] = obter_coluna(['nivel de rede', 'nível de rede', 'nivel_rede'])
        df_captacao['UF'] = obter_coluna(['regional', 'uf', 'estado'])
        df_captacao['MUNICÍPIO'] = df_elegiveis['municipio']
        df_captacao['TELEFONE'] = obter_coluna(['telefone atualizado com ddd', 'telefone atualizado', 'telefone beneficiário', 'telefone beneficiario', 'telefone', 'celular'])
        df_captacao['STATUS'] = " "
        df_captacao['Enviado para'] = "Dom Pedro"

        diretorio_saida = os.path.dirname(caminho_arquivo)

        arquivo_saida = os.path.join(diretorio_saida, "pacientes_elegiveis_TOTAL.xlsx")

        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            df_dom_pedro.to_excel(writer, sheet_name='Dom Pedro', index=False)
            df_captacao.to_excel(writer, sheet_name='Captação', index=False)

        print(f"Arquivo diário salvo com as abas 'Dom Pedro' e 'Captação': {arquivo_saida}")

        arquivo_tabelao = os.path.join(diretorio_saida, "tabelao_pacientes_elegiveis.xlsx")

        if os.path.exists(arquivo_tabelao):
            try:
                df_tabelao_dom_pedro = pd.read_excel(arquivo_tabelao, sheet_name='Dom Pedro', engine='openpyxl')
            except Exception:
                df_tabelao_dom_pedro = pd.DataFrame()

            try:
                df_tabelao_captacao = pd.read_excel(arquivo_tabelao, sheet_name='Captação', engine='openpyxl')
            except Exception:
                df_tabelao_captacao = pd.DataFrame()

            df_tabelao_dom_pedro = pd.concat([df_tabelao_dom_pedro, df_dom_pedro], ignore_index=True)
            df_tabelao_captacao = pd.concat([df_tabelao_captacao, df_captacao], ignore_index=True)
        else:
            df_tabelao_dom_pedro = df_dom_pedro.copy()
            df_tabelao_captacao = df_captacao.copy()

        with pd.ExcelWriter(arquivo_tabelao, engine='openpyxl') as writer:
            df_tabelao_dom_pedro.to_excel(writer, sheet_name='Dom Pedro', index=False)
            df_tabelao_captacao.to_excel(writer, sheet_name='Captação', index=False)

        print(f"Arquivo tabelão atualizado: {arquivo_tabelao} | Total Dom Pedro: {len(df_tabelao_dom_pedro)} | Total Captação: {len(df_tabelao_captacao)}")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
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
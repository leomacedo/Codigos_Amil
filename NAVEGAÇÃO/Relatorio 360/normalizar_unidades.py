import os
import sys
import re
import unicodedata
import pandas as pd


NOME_ARQUIVO_ENTRADA = "leo.xlsx"
SUFIXO_ARQUIVO_SAIDA = "_formatado"
NOME_DA_COLUNA = "Local de Acomapanhamento atual"

REGRAS_UNIDADES = [
    (["TELEMEDICINA", "VIRTUAL", "FUMO ZERO"], "Telemedicina"),
    (["DOM PEDRO"], "Centro Médico Dom Pedro"),
    (["CONEXA"], "Conexa"),
    (["ONLINE"], "Online"),

    (["CAMPO GRANDE"], "Amil Espaço Saúde - Campo Grande"),
    (["NOVA IGUACU"], "Amil Espaço Saúde - Nova Iguaçu"),
    (["TIJUCA"], "Amil Espaço Saúde - Tijuca"),
    (["BOTAFOGO"], "Amil Espaço Saúde - Botafogo"),
    (["CAXIAS"], "Amil Espaço Saúde - Caxias"),
    (["GUARULHOS"], "Amil Espaço Saúde - Guarulhos"),
    (["NITEROI"], "Amil Espaço Saúde - Niterói"),
    (["OSASCO"], "Amil Espaço Saúde - Osasco"),
    (["SANTANA"], "Amil Espaço Saúde - Santana"),
    (["TATUAPE"], "Amil Espaço Saúde - Tatuapé"),
    (["ANA ROSA"], "Amil Espaço Saúde - Ana Rosa"),
    (["SANTO AMARO"], "Amil Espaço Saúde - Santo Amaro"),

    (["BUTANTA"], "Unidade Avançada Luz Butantã"),
    (["JOAO DIAS", "LUZ JOAO"], "Unidade Avançada Luz João Dias"),
    (["CARLOS CHAGAS"], "Unidade Avançada Carlos Chagas"),
    (["VITORIA"], "Unidade Avançada Vitória"),
    (["AVANCADA LUZ"], "Unidade Avançada Luz"),

    (["IPIRANGA MOGI"], "Hospital Ipiranga Mogi"),
    (["PAULISTANO"], "Hospital Paulistano"),
    (["CUBATAO"], "Hospital da Luz - Cubatão"),
    (["DA LUZ"], "Hospital da Luz"),
    (["MEDICO JARDIM"], "AP Centro Médico Jardim"),
    (["JOAO AZEVEDO"], "AP Centro Médico João Azevedo (SBC)"),
    (["LUIZ FERREIRA"], "AP Centro Médico Luiz Ferreira (SBC)"),
    (["SAO LUCAS"], "Hospital São Lucas"),
    (["SAMARITANO HIGIENOPOLIS"], "Hospital Samaritano Higienópolis"),
    (["SAMARITANO PAULISTA"], "Hospital Samaritano Paulista"),
    (["ALVORADA"], "Hospital Alvorada"),
    (["AMERICAS"], "Hospital Américas"),
    (["SANTA JOANA"], "Hospital Santa Joana (Recife)"),
    (["SEM DADOS"], "Sem Dados"),
    (["SANTO ANDRE"], "Amil Espaço Saúde - Santo André"),
]


if getattr(sys, "frozen", False):
    diretorio_atual = os.path.dirname(sys.executable)
else:
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))


# Normaliza textos para comparação
def limpar_texto(valor):
    valor = unicodedata.normalize("NFD", str(valor))
    valor = "".join([c for c in valor if not unicodedata.combining(c)])
    return valor.upper().strip()


# Normaliza nome da unidade
def normalizar_unidade(nome_bruto, log_nao_mapeadas=None):
    if pd.isna(nome_bruto):
        return ""

    val = limpar_texto(nome_bruto)

    if re.search(r"^X+$", val) or "MARIANA PENNA" in val:
        return "Não Preenchido/Inválido"

    for palavras_chave, nome_padrao in REGRAS_UNIDADES:
        if any(palavra in val for palavra in palavras_chave):
            return nome_padrao

    fallback = re.sub(r"^(DADOS UNIDADES )?UNIDADE:\s*", "", val)
    resultado_final = str(nome_bruto).strip() if fallback == val else fallback.title()

    if log_nao_mapeadas is not None and resultado_final:
        log_nao_mapeadas.add(str(nome_bruto).strip())

    return resultado_final


# Lê arquivo CSV ou Excel
def carregar_arquivo(caminho_arquivo):
    if caminho_arquivo.lower().endswith(".csv"):
        try:
            return pd.read_csv(caminho_arquivo, sep=";", encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(caminho_arquivo, sep=";", encoding="latin1")

    return pd.read_excel(caminho_arquivo)


# Salva arquivo CSV ou Excel
def salvar_arquivo(df, caminho_saida):
    if caminho_saida.lower().endswith(".csv"):
        df.to_csv(caminho_saida, sep=";", index=False, encoding="utf-8-sig")
    else:
        df.to_excel(caminho_saida, index=False)


# Processa a coluna de unidade
def processar_arquivo(caminho_arquivo, nome_coluna):
    print(f"Lendo arquivo: {caminho_arquivo}")

    try:
        df = carregar_arquivo(caminho_arquivo)
        df.columns = df.columns.str.strip()

        if nome_coluna not in df.columns:
            print(f"ERRO: A coluna '{nome_coluna}' não foi encontrada na planilha.")
            print(f"Colunas disponíveis: {df.columns.tolist()}")
            return

        print(f"Formatando a coluna '{nome_coluna}'...")

        unidades_desconhecidas = set()
        df[nome_coluna] = df[nome_coluna].apply(normalizar_unidade, log_nao_mapeadas=unidades_desconhecidas)

        nome_base, ext = os.path.splitext(os.path.basename(caminho_arquivo))
        caminho_saida = os.path.join(diretorio_atual, f"{nome_base}{SUFIXO_ARQUIVO_SAIDA}{ext}")

        salvar_arquivo(df, caminho_saida)

        print("-" * 50)
        print(f"SUCESSO! Arquivo salvo em:\n{caminho_saida}")

        if unidades_desconhecidas:
            print("\nATENÇÃO: Foram encontradas unidades que não estão no dicionário de regras:")
            for unidade in sorted(unidades_desconhecidas):
                print(f" - {unidade}")
            print(f"Total de unidades não mapeadas: {len(unidades_desconhecidas)}")

        print("-" * 50)

    except Exception as erro:
        print(f"Ocorreu um erro: {erro}")


if __name__ == "__main__":
    caminho_planilha = os.path.join(diretorio_atual, NOME_ARQUIVO_ENTRADA)

    if os.path.exists(caminho_planilha):
        processar_arquivo(caminho_planilha, NOME_DA_COLUNA)
    else:
        print(f"Arquivo não encontrado: {caminho_planilha}")
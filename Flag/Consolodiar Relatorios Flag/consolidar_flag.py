from pathlib import Path
import pandas as pd
import csv
import unicodedata
import re

# =========================
# CONFIGURAÇÕES
# =========================
NOME_PASTA = "relatorios"
NOME_ARQUIVO_SAIDA = "base_empilhada.xlsx"
SEPARADOR_CSV = ";"
ENCODINGS = ["utf-8-sig", "utf-8", "latin1", "cp1252"]

# De / Para por palavra-chave
# O código procura a palavra-chave dentro da coluna PROGRAMA
PROGRAMAS_POR_PALAVRA_CHAVE = {
    "renal": "Saúde Renal",
    "mental": "Saúde Mental",
    "emagrecimento": "Emagrecimento",
    "pos infarto": "Cuidados Pós Infarto",
    "insuficiencia cardiaca": "Insuficiência Cardíaca Controlada",
    "anticoagulante": "Anticoagulante Seguro",
    "gestacao": "Gestação Segura",
    "mama": "Cuidado Integral da Mama",
    "coluna": "Saúde da Coluna",
    "fumo": "Fumo Zero",
    "endometriose": "Cuidados para Endometriose",
    "valvar": "Cuidado Cardíaco Valvar",
    "ritmo": "Ritmo Certo",
    "pos avc": "Pós AVC",
    "prostata": "Cuidado Oncológico Próstata",
    "pulmonar": "Cuidado Oncológico Pulmonar",
    "colorretal": "Cuidado Oncológico Colorretal"
}

# Cabeçalho oficial
COLUNAS_PADRAO = [
    "MO",
    "BENEFICIÁRIO",
    "SITUAÇÃO DO BENEFICIÁRIO",
    "OPERADORA",
    "FILIAL",
    "NUMERO DO CONTRATO",
    "CONTRATO",
    "REDE",
    "PROGRAMA",
    "INÍCIO DA VIGÊNCIA",
    "FIM DA VIGÊNCIA"
]


def normalizar_texto(texto):
    texto = "" if pd.isna(texto) else str(texto)
    texto = texto.strip().lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def detectar_encoding(arquivo):
    for encoding in ENCODINGS:
        try:
            with open(arquivo, "r", encoding=encoding) as f:
                f.readline()
            return encoding
        except UnicodeDecodeError:
            continue

    raise ValueError(f"Não foi possível identificar o encoding do arquivo: {arquivo.name}")


def linha_parece_cabecalho(linha):
    linha_normalizada = [str(campo).strip().upper() for campo in linha]
    colunas_normalizadas = [coluna.upper() for coluna in COLUNAS_PADRAO]

    acertos = 0
    for coluna in colunas_normalizadas:
        if coluna in linha_normalizada:
            acertos += 1

    return acertos >= 3


def padronizar_programa(valor):
    if pd.isna(valor):
        return valor

    programa_original = str(valor).strip()
    programa_normalizado = normalizar_texto(programa_original)

    # Procura pelas palavras-chave dentro do texto da coluna PROGRAMA
    for palavra_chave, nome_fantasia in PROGRAMAS_POR_PALAVRA_CHAVE.items():
        palavra_chave_normalizada = normalizar_texto(palavra_chave)

        if palavra_chave_normalizada in programa_normalizado:
            return nome_fantasia

    # Se não encontrar nenhuma palavra-chave, mantém o valor original
    return programa_original


def ler_csv_com_colunas_padrao(arquivo):
    encoding = detectar_encoding(arquivo)
    registros = []
    maior_qtd_auxiliar = 0

    with open(arquivo, "r", encoding=encoding, newline="") as f:
        leitor = csv.reader(f, delimiter=SEPARADOR_CSV, quotechar='"')
        linhas = list(leitor)

    if not linhas:
        return pd.DataFrame(columns=COLUNAS_PADRAO)

    # Se a primeira linha parecer cabeçalho, ignora ela
    primeira_linha = linhas[0]
    if linha_parece_cabecalho(primeira_linha):
        linhas_dados = linhas[1:]
    else:
        linhas_dados = linhas

    for numero_linha, linha in enumerate(linhas_dados, start=2):
        linha = [campo.strip() if isinstance(campo, str) else campo for campo in linha]
        registro = {}

        # Preenche as colunas padrão
        for i, coluna in enumerate(COLUNAS_PADRAO):
            registro[coluna] = linha[i] if i < len(linha) else ""

        # Se tiver campos extras, cria AUXILIAR_1, AUXILIAR_2...
        extras = linha[len(COLUNAS_PADRAO):]

        for i, valor_extra in enumerate(extras, start=1):
            registro[f"AUXILIAR_{i}"] = valor_extra

        if len(extras) > maior_qtd_auxiliar:
            maior_qtd_auxiliar = len(extras)

        registros.append(registro)

    colunas_finais = COLUNAS_PADRAO + [f"AUXILIAR_{i}" for i in range(1, maior_qtd_auxiliar + 1)]
    df = pd.DataFrame(registros)

    for coluna in colunas_finais:
        if coluna not in df.columns:
            df[coluna] = ""

    return df[colunas_finais]


def empilhar_csvs():
    raiz = Path(__file__).resolve().parent
    pasta = raiz if NOME_PASTA == "." else raiz / NOME_PASTA
    caminho_saida = raiz / NOME_ARQUIVO_SAIDA

    # Valida a pasta
    if not pasta.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {pasta}")

    # Busca os CSVs
    arquivos = sorted(
        [a for a in pasta.iterdir() if a.is_file() and a.suffix.lower() == ".csv"],
        key=lambda a: a.name.lower()
    )

    if not arquivos:
        raise FileNotFoundError(f"Nenhum CSV encontrado em: {pasta}")

    dfs = []

    for arquivo in arquivos:
        print(f"Lendo: {arquivo.name}")
        df = ler_csv_com_colunas_padrao(arquivo)
        dfs.append(df)

    # Empilha tudo
    base_final = pd.concat(dfs, ignore_index=True, sort=False)

    # Padroniza a coluna PROGRAMA usando palavra-chave
    if "PROGRAMA" in base_final.columns:
        base_final["PROGRAMA"] = base_final["PROGRAMA"].apply(padronizar_programa)

    # Garante ordem das colunas padrão primeiro e auxiliares depois
    colunas_auxiliares = sorted(
        [coluna for coluna in base_final.columns if coluna.startswith("AUXILIAR_")],
        key=lambda x: int(x.replace("AUXILIAR_", ""))
    )

    colunas_finais = COLUNAS_PADRAO + colunas_auxiliares
    base_final = base_final[colunas_finais]

    # Gera Excel
    base_final.to_excel(caminho_saida, index=False, engine="openpyxl")

    print("Concluído com sucesso!")
    print(f"Pasta lida: {pasta}")
    print(f"Arquivos empilhados: {len(arquivos)}")
    print(f"Linhas finais: {len(base_final)}")
    print(f"Colunas finais: {len(base_final.columns)}")
    print(f"Arquivo gerado: {caminho_saida}")


if __name__ == "__main__":
    empilhar_csvs()
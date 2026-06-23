import pandas as pd
import os
import sys
import glob
import unicodedata

# ================= CONFIGURAÇÕES =================
NOME_ARQUIVO_ENTRADA = "Base_CID.xlsx"  

NOME_ARQUIVO_SAIDA = "Base_Cid_Filtrada.xlsx"

LINHAS_DE_CUIDADO = [
    "Asma Sob Controle",
    "Autoimunidade",
    "Diabetes Sob Controle",
    "DPOC",
    "Saúde Renal"
]

UFS_PERMITIDAS = [
    "RJ",
    "SP"
]

LIMITE_NIVEL_REDE = 750

NOMES_COLUNA_LINHA_CUIDADO = ["Linha de Cuidado", "linha de cuidado", "Linha Cuidado"]
NOMES_COLUNA_UF = ["UF Prestador", "uf", "Estado", "estado"]
NOMES_COLUNA_NIVEL_REDE = ["nivel_rede", "Nível de Rede", "Nivel de Rede", "nível_rede", "Nivel_Rede"]
# ==================================================

def remover_acentos(texto):
    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    return "".join([c for c in texto if not unicodedata.combining(c)])

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    return remover_acentos(str(texto).strip().lower())

def encontrar_coluna(df, nomes_possiveis):
    colunas_normalizadas = {normalizar_texto(col): col for col in df.columns}
    for nome in nomes_possiveis:
        nome_normalizado = normalizar_texto(nome)
        if nome_normalizado in colunas_normalizadas:
            return colunas_normalizadas[nome_normalizado]
    return None

def converter_nivel_rede(valor):
    if pd.isna(valor):
        return None
    texto = str(valor).strip().replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except:
        return None

def obter_diretorio_script():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def obter_arquivo_entrada(diretorio_script):
    if NOME_ARQUIVO_ENTRADA.strip() != "":
        caminho = os.path.join(diretorio_script, NOME_ARQUIVO_ENTRADA)
        if os.path.exists(caminho):
            return caminho
        print(f"Arquivo informado não encontrado: {caminho}")
        return None

    arquivos_excel = glob.glob(os.path.join(diretorio_script, "*.xlsx"))
    arquivos_excel = [arquivo for arquivo in arquivos_excel if not os.path.basename(arquivo).startswith("~$") and os.path.basename(arquivo) != NOME_ARQUIVO_SAIDA]

    if len(arquivos_excel) == 0:
        print("Nenhuma planilha .xlsx encontrada na mesma pasta do script.")
        return None

    if len(arquivos_excel) > 1:
        print("Mais de uma planilha encontrada. Informe o nome exato em NOME_ARQUIVO_ENTRADA.")
        print("\nPlanilhas encontradas:")
        for arquivo in arquivos_excel:
            print(f"- {os.path.basename(arquivo)}")
        return None

    return arquivos_excel[0]

def filtrar_planilha():
    diretorio_script = obter_diretorio_script()
    caminho_entrada = obter_arquivo_entrada(diretorio_script)

    if caminho_entrada is None:
        return

    caminho_saida = os.path.join(diretorio_script, NOME_ARQUIVO_SAIDA)

    try:
        df = pd.read_excel(caminho_entrada, sheet_name=0, engine="openpyxl")
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return

    coluna_linha_cuidado = encontrar_coluna(df, NOMES_COLUNA_LINHA_CUIDADO)
    coluna_uf = encontrar_coluna(df, NOMES_COLUNA_UF)
    coluna_nivel_rede = encontrar_coluna(df, NOMES_COLUNA_NIVEL_REDE)

    if coluna_linha_cuidado is None:
        print("Coluna de Linha de Cuidado não encontrada.")
        return

    if coluna_uf is None:
        print("Coluna de UF não encontrada.")
        return

    if coluna_nivel_rede is None:
        print("Coluna de nivel_rede não encontrada.")
        return

    linhas_normalizadas = [normalizar_texto(valor) for valor in LINHAS_DE_CUIDADO]
    ufs_normalizadas = [normalizar_texto(valor) for valor in UFS_PERMITIDAS]

    filtro_linha_cuidado = df[coluna_linha_cuidado].apply(normalizar_texto).isin(linhas_normalizadas)
    filtro_uf = df[coluna_uf].apply(normalizar_texto).isin(ufs_normalizadas)
    filtro_nivel_rede = df[coluna_nivel_rede].apply(converter_nivel_rede) < LIMITE_NIVEL_REDE

    df_filtrado = df[filtro_linha_cuidado & filtro_uf & filtro_nivel_rede].copy()

    print(f"\nArquivo lido: {os.path.basename(caminho_entrada)}")
    print(f"Total de linhas na planilha: {len(df)}")
    print(f"Filtro Linha de Cuidado: {filtro_linha_cuidado.sum()}")
    print(f"Filtro UF RJ/SP: {filtro_uf.sum()}")
    print(f"Filtro nivel_rede menor que {LIMITE_NIVEL_REDE}: {filtro_nivel_rede.sum()}")
    print(f"Total após todos os filtros: {len(df_filtrado)}")

    if df_filtrado.empty:
        print("\nNenhum registro encontrado com os filtros informados.")
        return

    try:
        df_filtrado.to_excel(caminho_saida, index=False, engine="openpyxl")
        print(f"\nSucesso! Planilha filtrada salva em: {caminho_saida}")
    except Exception as e:
        print(f"Erro ao salvar a planilha filtrada: {e}")

if __name__ == "__main__":
    filtrar_planilha()
    input("\nPressione Enter para fechar...")
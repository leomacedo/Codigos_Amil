# OBSERVAÇÕES
# MANTER ZOOM EM 150% E NAO DEIXAR TELA CHEIA

import os
import pandas as pd
import unicodedata
from datetime import datetime

# =========================================================
# CAMINHOS
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # Pasta base do projeto

PLANILHA_ENTRADA = os.path.join(BASE_DIR, "lista_mos.xlsx")  # Planilha de entrada
PLANILHA_SAIDA = os.path.join(BASE_DIR, "resultado_desflag.xlsx")  # Planilha de saída

# =========================================================
# COLUNAS DA PLANILHA
# =========================================================

COLUNA_MO = "mo"  # Nome normalizado da coluna MO
COLUNA_PROGRAMA = "programa"  # Nome normalizado da coluna programa
COLUNA_STATUS = "status"  # Status detalhado
COLUNA_STATUS_FLAG = "status_flag"  # Ativo/Excluído/Erro
COLUNA_STATUS_ATIVO = "status_ativo"  # Sim/Não/Erro

# =========================================================
# CONFIGURAÇÕES GERAIS
# =========================================================

DATA = datetime.now().strftime("%d/%m/%Y") # Data que será colada no campo Data fim

CONF_GERAL = 0.98  # Confiança para lupa, finalizar, limpar
CONF_SUCESSO = 0.97  # Confiança para sucesso
CONF_ERRO = 0.95  # Confiança para mensagem de erro
CONF_EXCLUIDO = 0.92

DEBUG = False  # True mostra logs de match no terminal

REGION = (0, 0, 1920, 1080)  # Região da tela onde procura imagens

# =========================================================
# IMAGENS FIXAS
# =========================================================

IMG_LUPA = os.path.join(BASE_DIR, "img", "lupa.png")  # Imagem da lupa
IMG_FINALIZAR = os.path.join(BASE_DIR, "img", "finalizar.png")  # Botão finalizar
IMG_LIMPAR = os.path.join(BASE_DIR, "img", "limpar.png")  # Botão limpar
IMG_SUCESSO = os.path.join(BASE_DIR, "img", "sucesso.png")  # Confirmação de sucesso
IMG_ERRO = os.path.join(BASE_DIR, "img", "erro.png")  # Mensagem de erro após finalizar
IMG_EXCLUIDO = os.path.join(BASE_DIR, "img", "excluido.png")

# =========================================================
# PROGRAMAS
# =========================================================
# As chaves devem estar SEM ACENTO.
# A planilha pode ter acento, porque o código normaliza.
# Exemplo: "Saúde Mental" vira "saude mental".

PROGRAMAS = {
    "emagrecimento": {
        "img": os.path.join(BASE_DIR, "programas", "emagrecimento.png"),
        "conf": 0.90
    },
    "saude da coluna": {
        "img": os.path.join(BASE_DIR, "programas", "coluna.png"),
        "conf": 0.90
    },
    "cuidados pos infarto": {
        "img": os.path.join(BASE_DIR, "programas", "posiam.png"),
        "conf": 0.90
    },
    "gestacao segura": {
        "img": os.path.join(BASE_DIR, "programas", "gestacao.png"),
        "conf": 0.90
    },
    "insuficiencia cardiaca controlada": {
        "img": os.path.join(BASE_DIR, "programas", "ic.png"),
        "conf": 0.90
    },
    "ritmo certo": {
        "img": os.path.join(BASE_DIR, "programas", "ritmo.png"),
        "conf": 0.90
    },
    "saude mental": {
        "img": os.path.join(BASE_DIR, "programas", "saude_mental.png"),
        "conf": 0.90
    },
    "saude renal": {
        "img": os.path.join(BASE_DIR, "programas", "renal.png"),
        "conf": 0.90
    },
    "anticoagulante seguro": {
        "img": os.path.join(BASE_DIR, "programas", "anticoagulante.png"),
        "conf": 0.90
    },
    "cuidado cardiaco valvar": {
        "img": os.path.join(BASE_DIR, "programas", "valvar.png"),
        "conf": 0.90
    },
    "cuidado integral da mama": {
        "img": os.path.join(BASE_DIR, "programas", "mama.png"),
        "conf": 0.90
    },
    "cuidado oncologico colorretal": {
        "img": os.path.join(BASE_DIR, "programas", "oncologico_colorretal.png"),
        "conf": 0.90
    },
    "cuidado oncologico prostata": {
        "img": os.path.join(BASE_DIR, "programas", "oncologico_prostata.png"),
        "conf": 0.90
    },
    "cuidado oncologico pulmonar": {
        "img": os.path.join(BASE_DIR, "programas", "oncologico_pulmonar.png"),
        "conf": 0.90
    },
    "cuidados para endometriose": {
        "img": os.path.join(BASE_DIR, "programas", "endometriose.png"),
        "conf": 0.90
    },
    "fumo zero": {
        "img": os.path.join(BASE_DIR, "programas", "fumo_zero.png"),
        "conf": 0.90
    },
    "pos avc": {
        "img": os.path.join(BASE_DIR, "programas", "pos_avc.png"),
        "conf": 0.90
    },
    "viva bem - dom pedro": {
        "imgs": [os.path.join(BASE_DIR, "programas", "vivabemdp.png"),os.path.join(BASE_DIR, "programas", "vivabemdp1.png")],
        "conf": 0.90
    }
}

# =========================================================
# OFFSETS
# =========================================================

OFFSET_MO_X = -120  # Distância horizontal da lupa até o campo da MO
OFFSET_DATA_X = 500  # Distância horizontal do programa até o campo Data fim



# =========================================================
# TEMPOS
# =========================================================

TEMPO_MO = 0.08
TEMPO_DATA = 0.1
TEMPO_ENTER = 0

TEMPO_ESTABILIZACAO = 0.25
TEMPO_BUSCA_DATA = 0.5
TEMPO_BUSCA_DATA_RETRY = 1

TEMPO_FINALIZAR = 0
TEMPO_RESULTADO_FINALIZACAO = 3

TEMPO_LIMPAR = 0.1
TEMPO_CONFIRMA_LIMPAR = 5
TEMPO_CARREGAR_OU_EXCLUIDO = 10


MAX_FALHAS = 3  # Máximo de falhas na mesma MO

# =========================================================
# FUNÇÕES DE NORMALIZAÇÃO
# =========================================================

def normalizar_texto(valor):
    texto = str(valor).strip().lower()

    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(char for char in texto if unicodedata.category(char) != "Mn")

    return texto

def normalizar_mo(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if texto.lower() == "nan":
        return ""

    # Remove .0 caso Excel leia número como texto decimal
    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto.zfill(9)

# =========================================================
# CARREGAR PLANILHA
# =========================================================

def carregar_planilha():
    if not os.path.exists(PLANILHA_ENTRADA):
        print("[ERRO] Planilha não encontrada:", PLANILHA_ENTRADA)
        return pd.DataFrame(), []

    df = pd.read_excel(PLANILHA_ENTRADA, dtype=str, engine="openpyxl")

    # Normaliza nomes das colunas
    df.columns = [normalizar_texto(col) for col in df.columns]

    print("[INFO] Colunas encontradas na planilha:", list(df.columns))

    if COLUNA_MO not in df.columns:
        raise Exception(f"Coluna obrigatória não encontrada: {COLUNA_MO}")

    if COLUNA_PROGRAMA not in df.columns:
        raise Exception(f"Coluna obrigatória não encontrada: {COLUNA_PROGRAMA}")

    if COLUNA_STATUS not in df.columns:
        df[COLUNA_STATUS] = ""

    if COLUNA_STATUS_FLAG not in df.columns:
        df[COLUNA_STATUS_FLAG] = ""

    if COLUNA_STATUS_ATIVO not in df.columns:
        df[COLUNA_STATUS_ATIVO] = ""

    lista = []

    for idx, linha in df.iterrows():
        mo = normalizar_mo(linha[COLUNA_MO])
        programa = normalizar_texto(linha[COLUNA_PROGRAMA])

        if mo:
            lista.append({
                "idx": idx,
                "mo": mo,
                "programa": programa
            })

    print(f"[INFO] Total de linhas carregadas: {len(lista)}")

    return df, lista

DF_PLANILHA, LISTA = carregar_planilha()
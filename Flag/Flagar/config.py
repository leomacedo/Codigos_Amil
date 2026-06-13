# OBSERVAÇÕES
# MANTER ZOOM EM 150% E NAO DEIXAR TELA CHEIA

from datetime import datetime
import os
import pandas as pd
import unicodedata

# =========================================================
# PLANILHA
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # Pasta base do projeto

PLANILHA_ENTRADA = os.path.join(BASE_DIR, "lista_mos.xlsx")  # Planilha de entrada
PLANILHA_SAIDA = os.path.join(BASE_DIR, "resultado_flag.xlsx")  # Planilha gerada com status

# =========================================================
# CONFIGURAÇÕES GERAIS
# =========================================================

DATA = datetime.now().strftime("%d/%m/%Y") # Data que será colada no campo Data fim
    
CONF_GERAL = 0.98  # Confiança para imagens fixas
CONF_ADICIONAR = 0.92  # Confiança para botão adicionar
CONF_ADICIONAR_DEPOIS = 0.92  # Confiança para botão adicionar depois
CONF_SUCESSO = 0.97  # Confiança para confirmação de sucesso
CONF_ERRO_ATIVO = 0.95  # Confiança para mensagem de programa já ativo
CONF_MENSAGEM_ENTER = 0.95  # Confiança da mensagem após ENTER

DEBUG = True  # True mostra logs de precisão no terminal

REGION = (0, 0, 1920, 1080)  # Região da tela onde o bot procura imagens

# =========================================================
# IMAGENS FIXAS
# =========================================================

IMG_LUPA = os.path.join(BASE_DIR, "img", "lupa.png")  # Imagem da lupa
IMG_ADICIONAR = os.path.join(BASE_DIR, "img", "adicionar.png")  # Botão adicionar antes/depois
IMG_ADICIONAR_DEPOIS = os.path.join(BASE_DIR, "img", "adicionar.png")  # Botão adicionar após selecionar programa
IMG_FINALIZAR = os.path.join(BASE_DIR, "img", "finalizar.png")  # Botão finalizar
IMG_SUCESSO = os.path.join(BASE_DIR, "img", "sucesso.png")  # Confirmação de sucesso
IMG_LIMPAR = os.path.join(BASE_DIR, "img", "limpar.png")  # Botão limpar
IMG_ERRO_ATIVO = os.path.join(BASE_DIR, "img", "erro.png")  # Mensagem: Já possui programa ativo
IMG_MENSAGEM_ENTER = os.path.join(BASE_DIR, "img", "mensagem.png")  # Mensagem que aparece após ENTER

# =========================================================
# PROGRAMAS
# =========================================================

PROGRAMAS = {
    "anticoagulante seguro": {"setas": 15},  
    "cuidado cardiaco valvar": {"setas": 18},  
    "cuidado integral da mama": {"setas": 19},  
    "cuidado oncologico colorretal": {"setas": 21},  
    "cuidado oncologico prostata": {"setas": 20},  
    "cuidado oncologico pulmonar": {"setas": 22},  
    "cuidados para endometriose": {"setas": 29},  
    "cuidados pos infarto": {"setas": 23},
    "emagrecimento": {"setas": 12},
    "fumo zero": {"setas": 34},  
    "gestacao segura": {"setas": 30},
    "insuficiencia cardiaca controlada": {"setas": 13},
    "pos avc": {"setas": 26},  
    "ritmo certo": {"setas": 31},
    "saude da coluna": {"setas": 32},
    "saude mental": {"setas": 14},
    "saude renal": {"setas": 33},
    "alta dependencia - bem cuidado - dom pedro": {"setas": 2},
    "melhores cuidados - dom pedro": {"setas": 44},
    "melhores cuidados - dot": {"setas": 45},
    "melhores cuidados - lacos": {"setas": 46},
    "melhores cuidados - valsa": {"setas": 47},
    "melhores cuidados - vivaz": {"setas": 48},
    "vitalidade": {"setas": 62},
    "viva bem - dom pedro": {"setas": 64},
    "viva bem - dot": {"setas": 65}
  
}

# =========================================================
# OFFSETS
# =========================================================

OFFSET_MO_X = -120  # Distância da lupa até o campo da MO
OFFSET_DATA_INICIO_X = 80  # Distância do adicionar até o campo data início
OFFSET_DATA_INICIO_Y = -65  # Altura do adicionar até o campo data início


# =========================================================
# TEMPOS - SUPER RÁPIDO SEGURO
# =========================================================

TEMPO_TESTE_DATA_FIM = 0.05

TEMPO_MO = 0.05
TEMPO_ENTER = 0
TEMPO_ESTABILIZACAO = 0.15

TEMPO_DROPDOWN = 0.08
TEMPO_SETA = 0
TEMPO_SELECIONAR_PROGRAMA = 0.08

TEMPO_DATA_INICIO = 0.08
TEMPO_FINALIZAR = 0

TEMPO_BUSCA_PROGRAMA_EXISTENTE = 0.2
TEMPO_BUSCA_ADICIONAR = 1.2

TENTATIVAS_ADICIONAR = 2
TEMPO_RETRY_ADICIONAR = 0.5


TEMPO_LIMPAR = 0
TEMPO_CONFIRMA_LIMPAR = 5

TEMPO_RESULTADO_FINALIZACAO = 2.5
TEMPO_CARREGAR_OU_MENSAGEM = 8

INTERVALO_CHECAGEM_RESULTADO = 0.03  # Intervalo entre buscas de sucesso/erro
TEMPO_ENTRE_MOS = 0  # Sem pausa entre uma MO e outra
SALVAR_A_CADA = 10  # Salva planilha a cada 10 MOs

MAX_FALHAS = 3  # Máximo de falhas permitidas na mesma MO

# =========================================================
# CARREGAR PLANILHA
# =========================================================


COLUNA_MO = "mo"  # Nome da coluna de MO normalizado
COLUNA_PROGRAMA = "programa"  # Nome da coluna de programa normalizado
COLUNA_STATUS = "status"
COLUNA_STATUS_FLAG = "status_flag"  
COLUNA_STATUS_ATIVO = "status_ativo"  # Sim/Não/Erro

def normalizar_programa(valor):
    texto = str(valor).strip().lower()

    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(char for char in texto if unicodedata.category(char) != "Mn")

    return texto

def carregar_planilha():
    if not os.path.exists(PLANILHA_ENTRADA):
        print("[ERRO] Planilha não encontrada:", PLANILHA_ENTRADA)
        return pd.DataFrame(), []

    df = pd.read_excel(PLANILHA_ENTRADA, dtype=str, engine="openpyxl")

    # Normaliza nomes das colunas
    df.columns = [str(col).strip().lower() for col in df.columns]

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
        mo = str(linha[COLUNA_MO]).strip().zfill(9)
        programa = normalizar_programa(linha[COLUNA_PROGRAMA])

        if mo and mo.lower() != "nan":
            lista.append({
                "idx": idx,
                "mo": mo,
                "programa": programa
            })

    return df, lista

DF_PLANILHA, LISTA = carregar_planilha()

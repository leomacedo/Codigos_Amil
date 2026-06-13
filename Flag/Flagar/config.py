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
# PROGRAMAS (ATUALIZADO PARA PESQUISA POR TEXTO)
# =========================================================

PROGRAMAS = {
    "anticoagulante seguro": {"texto_pesquisa": "CuidadosMil - Anticoagulante Seguro"},  
    "cuidado cardiaco valvar": {"texto_pesquisa": "CuidadosMil - Cuidado Cardíaco Valvar"},  
    "cuidado integral da mama": {"texto_pesquisa": "CuidadosMil - Cuidado Integral da Mama"},  
    "cuidado oncologico colorretal": {"texto_pesquisa": "CuidadosMil - Cuidado Oncológico Colorretal"},  
    "cuidado oncologico prostata": {"texto_pesquisa": "CuidadosMil - Cuidado Oncológico Próstata"},  
    "cuidado oncologico pulmonar": {"texto_pesquisa": "CuidadosMil - Cuidado Oncológico Pulmonar"},  
    "cuidados para endometriose": {"texto_pesquisa": "CuidadosMil - Endometriose"},  
    "cuidados pos infarto": {"texto_pesquisa": "CuidadosMil - Cuidados Pós Infarto"},
    "emagrecimento": {"texto_pesquisa": "CuidadosMil - Emagrecimento"},
    "fumo zero": {"texto_pesquisa": "CuidadosMil - Fumo Zero"},  
    "gestacao segura": {"texto_pesquisa": "CuidadosMil - Gestação Segura"},
    "insuficiencia cardiaca controlada": {"texto_pesquisa": "CuidadosMil - Insuficiência Cardíaca Controlada"},
    "pos avc": {"texto_pesquisa": "CuidadosMil - Cuidados Pós-AVC"},  
    "ritmo certo": {"texto_pesquisa": "CuidadosMil - Ritmo Certo"},
    "saude da coluna": {"texto_pesquisa": "CuidadosMil - Saúde da Coluna"},
    "saude mental": {"texto_pesquisa": "CuidadosMil - Saúde Mental"},
    "saude renal": {"texto_pesquisa": "CuidadosMil - Saúde Renal"},
    
    # Mantidos do seu código original por segurança
    "alta dependencia - bem cuidado - dom pedro": {"texto_pesquisa": "Alta Dependência - Bem Cuidado - Dom Pedro"},
    "melhores cuidados - dom pedro": {"texto_pesquisa": "Melhores Cuidados - Dom Pedro"},
    "melhores cuidados - dot": {"texto_pesquisa": "Melhores Cuidados - Dot"},
    "melhores cuidados - lacos": {"texto_pesquisa": "Melhores Cuidados - Laços"},
    "melhores cuidados - valsa": {"texto_pesquisa": "Melhores Cuidados - Valsa"},
    "melhores cuidados - vivaz": {"texto_pesquisa": "Melhores Cuidados - Vivaz"},
    "vitalidade": {"texto_pesquisa": "Vitalidade"},
    "viva bem - dom pedro": {"texto_pesquisa": "Viva Bem - Dom Pedro"},
    "viva bem - dot": {"texto_pesquisa": "Viva Bem - Dot"}
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

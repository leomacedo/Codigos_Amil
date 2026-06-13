import cv2
import numpy as np
import pyautogui
import time
import os

from config import REGION, CONF_GERAL, DEBUG

# =========================================================
# BUSCA DE IMAGEM COM OPENCV
# =========================================================

def encontrar_imagem(img, conf=None):
    # Se nenhuma confiança for informada, usa a confiança geral
    if conf is None:
        conf = CONF_GERAL

    # Tira print da região configurada
    tela = pyautogui.screenshot(region=REGION)

    # Converte o print para escala de cinza
    tela = cv2.cvtColor(np.array(tela), cv2.COLOR_BGR2GRAY)

    # Lê a imagem modelo em escala de cinza
    template = cv2.imread(img, 0)

    # Se a imagem não existir no caminho informado, retorna erro
    if template is None:
        print("[ERRO] Imagem não encontrada:", img)
        return None

    # Compara a imagem modelo com a tela
    res = cv2.matchTemplate(tela, template, cv2.TM_CCOEFF_NORMED)

    # Pega o melhor nível de correspondência encontrado
    _, max_val, _, max_loc = cv2.minMaxLoc(res)

    # Mostra o log de precisão, se DEBUG estiver ativo
    if DEBUG:
        print(f"[MATCH] {os.path.basename(img)} -> {round(max_val, 4)}")

    # Se a precisão for suficiente, retorna o centro da imagem encontrada
    if max_val >= conf:
        h, w = template.shape
        return (max_loc[0] + w // 2, max_loc[1] + h // 2)

    return None

# =========================================================
# ESPERA INTELIGENTE POR IMAGEM
# =========================================================

def esperar_imagem(img, timeout=5, conf=None):
    inicio = time.time()

    while True:
        # Tenta encontrar a imagem
        pos = encontrar_imagem(img, conf)

        # Se encontrou, retorna a posição
        if pos:
            return pos

        # Se passou do tempo limite, retorna None
        if time.time() - inicio > timeout:
            return None

        # Pequena pausa entre tentativas para não pesar a CPU
        time.sleep(0.1)
        
def esperar_imagem_sumir(img, timeout=5, conf=None):
    inicio = time.time()

    while True:
        pos = encontrar_imagem(img, conf)

        if not pos:
            return True

        if time.time() - inicio > timeout:
            return False

        time.sleep(0.1)       
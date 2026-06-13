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
    if conf is None:
        conf = CONF_GERAL

    tela = pyautogui.screenshot(region=REGION)
    tela = cv2.cvtColor(np.array(tela), cv2.COLOR_BGR2GRAY)

    template = cv2.imread(img, 0)

    if template is None:
        print("[ERRO] Imagem não encontrada:", img)
        return None

    res = cv2.matchTemplate(tela, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)

    if DEBUG:
        print(f"[MATCH] {os.path.basename(img)} -> {round(max_val, 4)}")

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
        pos = encontrar_imagem(img, conf)

        if pos:
            return pos

        if time.time() - inicio > timeout:
            return None

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
        
def encontrar_todas_imagens(img, conf=None):
    if conf is None:
        conf = CONF_GERAL

    tela = pyautogui.screenshot(region=REGION)
    tela = cv2.cvtColor(np.array(tela), cv2.COLOR_BGR2GRAY)

    template = cv2.imread(img, 0)

    if template is None:
        print("[ERRO] Imagem não encontrada:", img)
        return []

    res = cv2.matchTemplate(tela, template, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= conf)

    h, w = template.shape
    pontos = []

    for pt in zip(*loc[::-1]):
        cx = pt[0] + w // 2
        cy = pt[1] + h // 2

        duplicado = False

        for p in pontos:
            if abs(cx - p[0]) < w // 2 and abs(cy - p[1]) < h // 2:
                duplicado = True
                break

        if not duplicado:
            pontos.append((cx, cy))

    pontos.sort(key=lambda p: p[1])

    if DEBUG:
        print(f"[MATCH_ALL] {os.path.basename(img)} -> {len(pontos)} ocorrência(s)")

    return pontos
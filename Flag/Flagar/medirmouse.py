import cv2
import numpy as np
import pyautogui
import keyboard
import os

# =========================================================
# CONFIG
# =========================================================

CONF = 0.90
REGION = (0, 0, 1920, 1080)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Troque aqui a imagem do programa que quer testar
IMG_PROGRAMA = os.path.join(BASE_DIR, "programas", "ritmo.png")

# Ajuste aqui até o mouse parar exatamente na Data fim
OFFSET_DATA_FIM_X = 520
OFFSET_DATA_FIM_Y = 0

# =========================================================
# OPENCV
# =========================================================

def encontrar_imagem(img, conf=CONF):
    tela = pyautogui.screenshot(region=REGION)
    tela = cv2.cvtColor(np.array(tela), cv2.COLOR_RGB2GRAY)

    template = cv2.imread(img, 0)

    if template is None:
        print("[ERRO] Imagem não encontrada:", img)
        return None

    res = cv2.matchTemplate(tela, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)

    print(f"[MATCH] {os.path.basename(img)} -> {round(max_val, 4)}")

    if max_val >= conf:
        h, w = template.shape
        x = max_loc[0] + w // 2
        y = max_loc[1] + h // 2
        return (x, y)

    return None

# =========================================================
# TESTE DO PONTO
# =========================================================

def testar_ponto():
    pos_programa = encontrar_imagem(IMG_PROGRAMA)

    if not pos_programa:
        print("[ERRO] Programa não encontrado")
        return

    x_programa = pos_programa[0]
    y_programa = pos_programa[1]

    x_data_fim = x_programa + OFFSET_DATA_FIM_X
    y_data_fim = y_programa + OFFSET_DATA_FIM_Y

    print("[OK] Programa encontrado em:", x_programa, y_programa)
    print("[INFO] Movendo mouse para Data fim:", x_data_fim, y_data_fim)

    pyautogui.moveTo(x_data_fim, y_data_fim, duration=0.2)

def mostrar_posicao():
    x, y = pyautogui.position()
    print(f"Mouse em: X={x} | Y={y}")

# =========================================================
# HOTKEYS
# =========================================================

print("J = testar ponto Data fim")
print("P = mostrar posição atual do mouse")
print("ESC = sair")

keyboard.add_hotkey('j', testar_ponto)
keyboard.add_hotkey('p', mostrar_posicao)

keyboard.wait('esc')
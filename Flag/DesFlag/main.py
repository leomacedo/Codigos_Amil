import keyboard

from automacao import alternar_loop, fluxo

# =========================================================
# HOTKEYS
# =========================================================

print("F9 = ligar/desligar loop")
print("J = executar 1 vez")
print("ESC = sair")

keyboard.add_hotkey('f9', alternar_loop)
keyboard.add_hotkey('j', fluxo)

keyboard.wait('esc')
import pyperclip
import keyboard
import time
import mouse
import threading

# --- CONFIGURAÇÃO ---
DATA_DESEJADA = "16/06/2026 14:00"
TEXTO_ESPACO = "leonardo"

TEXTO_T = """Verificamos o caso e esclarecemos que o pedido mencionado refere-se a uma captação para a linha de cuidados Amil.
Ressaltamos que todos os contatos realizados pela equipe de enfermagem do Centro Médico Dom Pedro geram um token, com o objetivo de validar a comunicação e a identificação do prestador, não implicando qualquer custo para o beneficiário."""

# Cole seus números aqui entre as três aspas, um por linha:
NUMEROS_CRUS = """
096737478
094418641
092857092
097420124
079649693
"""

# Converte o bloco de texto acima em uma lista, removendo espaços e linhas vazias.
# Se o item for um número, preenche com zeros à esquerda até atingir 9 dígitos.
LISTA_DE_NUMEROS = [
    num.strip().zfill(9) if num.strip().isdigit() else num.strip()
    for num in NUMEROS_CRUS.strip().split("\n")
    if num.strip()
]

indice_atual = 0
pausado = False
executando_triplo_clique = False


def colar_texto(texto):
    """Cola um texto preservando o clipboard antigo."""
    texto_antigo_clipboard = pyperclip.paste()

    pyperclip.copy(texto)
    keyboard.send("ctrl+v")
    time.sleep(0.1)

    pyperclip.copy(texto_antigo_clipboard)


def colar_proximo_numero():
    """Pega o próximo número da lista, cola e avança o índice."""
    global indice_atual

    if indice_atual < len(LISTA_DE_NUMEROS):
        numero_para_colar = LISTA_DE_NUMEROS[indice_atual]
        indice_atual += 1
    else:
        numero_para_colar = "acabou"

    colar_texto(numero_para_colar)

    keyboard.send("enter")
    time.sleep(0.1)

    mouse.wheel(100)

    print(f"Colado n°{indice_atual}: {numero_para_colar}")

    if indice_atual == len(LISTA_DE_NUMEROS) and numero_para_colar != "acabou":
        print("--- Fim da lista. O próximo aperto em 'F' colará 'acabou'. ---")


def colar_data_configurada():
    """Cola a data configurada."""
    colar_texto(DATA_DESEJADA)
    print(f"Colado: {DATA_DESEJADA}")


def digitar_leonardo():
    """Digita/cola Leonardo ao apertar espaço."""
    colar_texto(TEXTO_ESPACO)
    print(f"Colado: {TEXTO_ESPACO}")


def digitar_texto_t():
    """Cola o texto configurado ao apertar T."""
    colar_texto(TEXTO_T)
    print("Texto colado com quebra de linha.")


def voltar_numero():
    """Volta um número na lista caso o usuário tenha apertado 'F' por engano."""
    global indice_atual

    if indice_atual > 0:
        indice_atual -= 1
        numero_voltado = LISTA_DE_NUMEROS[indice_atual]
        print(f"Voltado 1 posição! O próximo 'F' colará novamente: {numero_voltado}")
    else:
        print("Já estamos no início da lista. Não é possível voltar mais.")


def copiar():
    keyboard.send("ctrl+c")


def colar():
    keyboard.send("ctrl+v")


def cadastrar_atalhos():
    """Cadastra os atalhos principais."""
    keyboard.add_hotkey("f", lambda: threading.Thread(target=colar_proximo_numero, daemon=True).start(), suppress=True)
    keyboard.add_hotkey("g", lambda: threading.Thread(target=colar_data_configurada, daemon=True).start(), suppress=True)
    keyboard.add_hotkey("h", voltar_numero, suppress=True)
    keyboard.add_hotkey("c", copiar, suppress=True)
    keyboard.add_hotkey("v", colar, suppress=True)
    keyboard.add_hotkey("space", lambda: threading.Thread(target=digitar_leonardo, daemon=True).start(), suppress=True)
    keyboard.add_hotkey("t", lambda: threading.Thread(target=digitar_texto_t, daemon=True).start(), suppress=True)


def remover_atalhos():
    """Remove os atalhos principais para liberar teclado."""
    for tecla in ["f", "g", "h", "c", "v", "space", "t"]:
        try:
            keyboard.remove_hotkey(tecla)
        except KeyError:
            pass


def alternar_pause():
    """Pausa ou retoma a execução do robô."""
    global pausado

    pausado = not pausado

    if pausado:
        print("\n[PAUSADO] O robô está pausado. Suas teclas e mouse funcionam normalmente.")
        print("Pressione 'F8' novamente para retomar.")

        remover_atalhos()

    else:
        print("\n[RETOMADO] O robô voltou a funcionar!")

        cadastrar_atalhos()


def triplo_clique():
    """Dá 3 cliques esquerdos rápidos e copia."""
    global executando_triplo_clique

    if executando_triplo_clique:
        return

    executando_triplo_clique = True

    try:
        time.sleep(0.10)

        # Fecha o menu do botão direito, caso tenha aberto
        keyboard.send("esc")
        time.sleep(0.08)

        # Triplo clique esquerdo
        mouse.click(button="left")
        time.sleep(0.07)

        mouse.click(button="left")
        time.sleep(0.07)

        mouse.click(button="left")
        time.sleep(0.15)

        # Copia o texto selecionado
        keyboard.send("ctrl+c")
        time.sleep(0.10)

        print("Triplo clique + copiar executado.")

    finally:
        executando_triplo_clique = False


def iniciar_triplo_clique_thread():
    """Inicia o triplo clique em uma thread separada."""
    threading.Thread(target=triplo_clique, daemon=True).start()


def handle_mouse_events(event):
    """
    Intercepta eventos do mouse.
    Se pausado, libera o mouse normalmente.
    Se clicar com botão direito, faz triplo clique + copia.
    """
    global pausado

    if pausado:
        return True

    if isinstance(event, mouse.ButtonEvent):
        if event.button == "right" and event.event_type == "down":
            iniciar_triplo_clique_thread()
            return False

        # Bloqueio extra no soltar do botão direito
        if event.button == "right" and event.event_type == "up":
            return False

    return True


print("Robô de atalhos iniciado!")
print(f"A lista carregada contém {len(LISTA_DE_NUMEROS)} números.")
print("Pressione 'F' para colar o próximo número da lista.")
print("Pressione 'H' para VOLTAR um número.")
print(f"Pressione 'G' para colar a data: {DATA_DESEJADA}")
print(f"Pressione 'ESPAÇO' para colar: {TEXTO_ESPACO}")
print("Pressione 'T' para colar o texto configurado com quebra de linha.")
print("Pressione o 'Botão Direito' do mouse para dar 3 cliques e copiar.")
print("Pressione 'F8' para PAUSAR / RETOMAR o robô.")
print("Pressione 'ESC' para desligar o robô.")

# Configura os atalhos
cadastrar_atalhos()
keyboard.add_hotkey("f8", alternar_pause, suppress=False)

# Hook do mouse
mouse.hook(handle_mouse_events)

# Mantém rodando até ESC
keyboard.wait("esc")
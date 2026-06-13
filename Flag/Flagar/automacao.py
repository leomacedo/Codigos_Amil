import pyautogui
import pyperclip
import mouse
import time
import threading
import os

from config import *
from visao import encontrar_imagem, esperar_imagem, esperar_imagem_sumir

# =========================================================
# ESTADO GLOBAL DA AUTOMAÇÃO
# =========================================================

indice = 0
LOOP_ATIVO = False
item_atual = None
falhas_mo = 0

tempo_inicio_loop = None
mos_processadas_loop = 0
falhas_loop = []

alteracoes_pendentes = 0

# =========================================================
# FUNÇÕES UTILITÁRIAS
# =========================================================

def clicar_com_offset(img, offset_x=0, offset_y=0, delay=0.5, conf=None):
    pos = esperar_imagem(img, 5, conf)

    if not pos:
        return False

    x = pos[0] + offset_x
    y = pos[1] + offset_y

    pyautogui.click(x, y, clicks=2, interval=0.1)

    if delay > 0:
        time.sleep(delay)

    return True

def clicar_imagem(img, delay=0.5, conf=None, timeout=5):
    pos = esperar_imagem(img, timeout, conf)

    if not pos:
        return False

    pyautogui.click(pos[0], pos[1])

    if delay > 0:
        time.sleep(delay)

    return True

def colar_texto(texto):
    antigo = pyperclip.paste()

    pyperclip.copy(texto)
    time.sleep(0.05)

    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.1)

    pyperclip.copy(antigo)

def limpar_e_confirmar():
    if not clicar_imagem(IMG_LIMPAR, TEMPO_LIMPAR, CONF_GERAL):
        print("[ERRO] Não encontrou limpar")
        return False

    print("[INFO] Aguardando tela limpar...")

    if not esperar_imagem_sumir(IMG_FINALIZAR, TEMPO_CONFIRMA_LIMPAR, CONF_GERAL):
        print("[ERRO] Tela não limpou")
        return False

    print("[OK] Tela limpa confirmada")

    return True
def gravar_planilha():
    global alteracoes_pendentes

    try:
        DF_PLANILHA.to_excel(PLANILHA_SAIDA, index=False, engine="openpyxl")
        alteracoes_pendentes = 0

    except PermissionError:
        nome_backup = f"resultado_flag_backup_{int(time.time())}.xlsx"
        caminho_backup = os.path.join(BASE_DIR, nome_backup)

        DF_PLANILHA.to_excel(caminho_backup, index=False, engine="openpyxl")
        alteracoes_pendentes = 0

        print("[AVISO] Não consegui salvar em resultado_flag.xlsx")
        print("[AVISO] Provavelmente o arquivo está aberto no Excel ou travado pelo OneDrive")
        print(f"[AVISO] Salvei backup em: {nome_backup}")
        
def salvar_status(item, status, status_flag):
    global alteracoes_pendentes

    idx = item["idx"]

    if status_flag == "Ativo":
        status_ativo = "Sim"
    elif status_flag == "Excluído":
        status_ativo = "Não"
    else:
        status_ativo = "Erro"

    DF_PLANILHA.loc[idx, COLUNA_STATUS] = status
    DF_PLANILHA.loc[idx, COLUNA_STATUS_FLAG] = status_flag
    DF_PLANILHA.loc[idx, COLUNA_STATUS_ATIVO] = status_ativo

    alteracoes_pendentes += 1

    print(f"[STATUS] MO {item['mo']} | {item['programa']} | {status} | {status_flag} | {status_ativo}")

    if SALVAR_A_CADA <= 1 or alteracoes_pendentes >= SALVAR_A_CADA or status_flag == "Erro":
        gravar_planilha()
# =========================================================
# CONTROLE DE FALHAS
# =========================================================

def registrar_log_falha(motivo):
    if item_atual:
        falhas_loop.append({
            "mo": item_atual["mo"],
            "programa": item_atual["programa"],
            "tentativa": falhas_mo + 1,
            "motivo": motivo
        })

def registrar_falha(motivo="falha não especificada"):
    global falhas_mo
    global LOOP_ATIVO
    global item_atual
    global indice
    global mos_processadas_loop

    registrar_log_falha(motivo)

    falhas_mo += 1

    print(f"[FALHA] Tentativa {falhas_mo}")

    if falhas_mo >= MAX_FALHAS:
        print("[ERRO CRÍTICO] Muitas falhas na mesma MO")

        if item_atual:
            print("[ERRO CRÍTICO] MO problemática:")
            print(item_atual["mo"])

            # Registra no resumo final que essa MO teve falha crítica
            falhas_loop.append({
                "mo": item_atual["mo"],
                "programa": item_atual["programa"],
                "tentativa": falhas_mo,
                "motivo": "FALHA CRÍTICA - MO pulada após atingir máximo de tentativas"
            })

            # Salva na planilha como erro
            salvar_status(item_atual, "falha crítica", "Erro")

            # Pula para a próxima MO
            indice += 1
            mos_processadas_loop += 1
            item_atual = None

        # Reseta falhas para a próxima MO
        falhas_mo = 0

        print("[INFO] Falha crítica registrada - seguindo para próxima MO")

        # Importante: NÃO para mais o loop
        return
    
def resetar_falhas():
    global falhas_mo

    falhas_mo = 0

# =========================================================
# CONTROLE DA LISTA DE MOS
# =========================================================

def pegar_item():
    global indice
    global LOOP_ATIVO
    global item_atual

    if item_atual:
        return item_atual

    if indice >= len(LISTA):
        print("[FIM] Lista acabou")

        LOOP_ATIVO = False

        return None

    item_atual = LISTA[indice]

    return item_atual

def confirmar_item_processado(status, status_flag):
    global indice
    global item_atual
    global mos_processadas_loop

    salvar_status(item_atual, status, status_flag)

    indice += 1
    mos_processadas_loop += 1
    item_atual = None

# =========================================================
# ETAPA 1 - INSERIR MO
# =========================================================

def etapa_inserir_mo():
    item = pegar_item()

    if not item:
        return False

    mo = item["mo"]

    if not clicar_com_offset(IMG_LUPA, OFFSET_MO_X, 0, TEMPO_MO, CONF_GERAL):
        print("[ERRO] Não encontrou lupa")
        return False

    colar_texto(mo)

    pyautogui.press("enter")

    if TEMPO_ENTER > 0:
        time.sleep(TEMPO_ENTER)

    mouse.wheel(100)

    print("[OK] MO:", mo)
    print("[INFO] Programa da linha:", item["programa"])

    return True

# =========================================================
# ETAPA 2 - ESPERAR TELA OU MENSAGEM APÓS ENTER
# =========================================================

def esperar_tela_ou_mensagem_enter():
    inicio = time.time()

    while time.time() - inicio < TEMPO_CARREGAR_OU_MENSAGEM:
        if encontrar_imagem(IMG_MENSAGEM_ENTER, CONF_MENSAGEM_ENTER):
            print("[INFO] Mensagem após ENTER detectada")
            return "mensagem_enter"

        if encontrar_imagem(IMG_FINALIZAR, CONF_GERAL):
            print("[OK] Tela carregada")
            return "tela_carregada"

        time.sleep(0.1)

    print("[ERRO] Tela não carregou e nenhuma mensagem apareceu")
    return "timeout"

# =========================================================
# ETAPA 3 - CLICAR ADICIONAR E SELECIONAR PROGRAMA
# =========================================================

def etapa_selecionar_programa(programa_cfg, nome_programa):
    for tentativa in range(1, TENTATIVAS_ADICIONAR + 1):
        print(f"[INFO] Procurando adicionar - tentativa {tentativa}/{TENTATIVAS_ADICIONAR}")

        pos = esperar_imagem(IMG_ADICIONAR, TEMPO_BUSCA_ADICIONAR, CONF_ADICIONAR)

        if pos:
            pyautogui.click(pos[0], pos[1])

            if TEMPO_DROPDOWN > 0:
                time.sleep(TEMPO_DROPDOWN)

            # Aqui ocorre a substituição das setas pela colagem do texto
            texto_para_digitar = programa_cfg.get("texto_pesquisa", nome_programa)
            colar_texto(texto_para_digitar)
            
            # Uma pequena pausa para dar tempo de a interface do sistema filtrar a lista
            time.sleep(0.3) 

            pyautogui.press("enter")

            if TEMPO_SELECIONAR_PROGRAMA > 0:
                time.sleep(TEMPO_SELECIONAR_PROGRAMA)

            print(f"[OK] Programa selecionado via texto: {texto_para_digitar}")
            return True

        print("[INFO] Adicionar não encontrado nesta tentativa")

        if tentativa < TENTATIVAS_ADICIONAR:
            time.sleep(TEMPO_RETRY_ADICIONAR)

    print("[ERRO] Não encontrou adicionar antes")
    return False

# =========================================================
# ETAPA 4 - COLAR DATA INÍCIO
# =========================================================

def etapa_data_inicio():
    pos = esperar_imagem(IMG_ADICIONAR_DEPOIS, 5, CONF_ADICIONAR_DEPOIS)

    if not pos:
        print("[ERRO] Não encontrou adicionar depois para localizar data início")
        return False

    x = pos[0] + OFFSET_DATA_INICIO_X
    y = pos[1] + OFFSET_DATA_INICIO_Y

    pyautogui.click(x, y)

    if TEMPO_DATA_INICIO > 0:
        time.sleep(TEMPO_DATA_INICIO)

    colar_texto(DATA)

    print("[OK] Data início colada")

    return True

# =========================================================
# ETAPA 5 - FINALIZAR
# =========================================================

def etapa_finalizar():
    if not clicar_com_offset(IMG_FINALIZAR, 0, 0, TEMPO_FINALIZAR, CONF_GERAL):
        print("[ERRO] Não encontrou finalizar")
        return False

    print("[OK] Finalizado")

    return True

# =========================================================
# RESULTADO DA FINALIZAÇÃO
# =========================================================

def esperar_resultado_finalizacao():
    inicio = time.time()

    while time.time() - inicio < TEMPO_RESULTADO_FINALIZACAO:
        if encontrar_imagem(IMG_SUCESSO, CONF_SUCESSO):
            print("[OK] Sucesso confirmado")
            return "sucesso"

        if encontrar_imagem(IMG_ERRO_ATIVO, CONF_ERRO_ATIVO):
            print("[INFO] Programa já ativo detectado")
            return "ja_existe"

        time.sleep(INTERVALO_CHECAGEM_RESULTADO)

    print("[ERRO] Não apareceu sucesso nem erro de programa ativo")
    return "sem_resultado"

# =========================================================
# TEMPO E PROGRESSO
# =========================================================

def formatar_tempo(segundos):
    segundos = int(segundos)
    minutos = segundos // 60
    segundos = segundos % 60
    horas = minutos // 60
    minutos = minutos % 60

    if horas > 0:
        return f"{horas}h {minutos}min {segundos}s"

    if minutos > 0:
        return f"{minutos}min {segundos}s"

    return f"{segundos}s"

def mostrar_progresso():
    if tempo_inicio_loop:
        tempo_decorrido = time.time() - tempo_inicio_loop
        media = tempo_decorrido / mos_processadas_loop if mos_processadas_loop > 0 else 0
        restante = len(LISTA) - indice
        estimado_restante = media * restante

        print(f"[PROGRESSO] Processadas: {mos_processadas_loop} | Restantes: {restante}")
        print(f"[PROGRESSO] Média: {round(media, 2)}s/MO | Estimado restante: {formatar_tempo(estimado_restante)}")

def mostrar_resumo_falhas():
    if falhas_loop:
        print("[RESUMO FALHAS]")

        for falha in falhas_loop:
            print(
                f"[FALHA LOG] MO {falha['mo']} | "
                f"{falha['programa']} | "
                f"Tentativa {falha['tentativa']} | "
                f"{falha['motivo']}"
            )
    else:
        print("[RESUMO FALHAS] Nenhuma falha registrada")

# =========================================================
# FLUXO COMPLETO
# =========================================================

def fluxo():
    global LOOP_ATIVO

    item = pegar_item()

    if not item:
        return

    nome_programa = item["programa"]

    if nome_programa not in PROGRAMAS:
        print(f"[ERRO] Programa não encontrado no dicionário: {nome_programa}")

        if limpar_e_confirmar():
            confirmar_item_processado("programa inválido", "Excluído")
            resetar_falhas()
            mostrar_progresso()
            return

        registrar_falha("programa inválido e falha ao limpar")
        return

    programa_cfg = PROGRAMAS[nome_programa]

    if not etapa_inserir_mo():
        if LOOP_ATIVO:
            registrar_falha("falha ao inserir MO")
        return

    print("[INFO] Esperando tela carregar ou mensagem após ENTER...")

    resultado_carregamento = esperar_tela_ou_mensagem_enter()

    if resultado_carregamento == "mensagem_enter":
        print("[INFO] MO com mensagem após ENTER - limpando e seguindo")

        if not limpar_e_confirmar():
            registrar_falha("mensagem após enter e falha ao limpar")
            return

        confirmar_item_processado("mensagem após enter", "Excluído")
        resetar_falhas()
        mostrar_progresso()
        return

    if resultado_carregamento == "timeout":
        print("[ERRO] Tela não carregou")
        print("[RETRY] Mesma MO será repetida")

        if not limpar_e_confirmar():
            registrar_falha("tela não carregou e falha ao limpar")
            return

        registrar_falha("tela não carregou")
        return

    if TEMPO_ESTABILIZACAO > 0:
        time.sleep(TEMPO_ESTABILIZACAO)

    print("[INFO] Tentando flagar programa...")

    if not etapa_selecionar_programa(programa_cfg, nome_programa):
        print("[INFO] Não encontrou adicionar - limpando e seguindo")

        if not limpar_e_confirmar():
            registrar_falha("não encontrou adicionar e falha ao limpar")
            return

        confirmar_item_processado("sem adicionar", "Excluído")
        resetar_falhas()
        mostrar_progresso()
        return

    if not etapa_data_inicio():
        print("[ERRO] Falha ao colar data início - limpando antes de tentar novamente")

        if not limpar_e_confirmar():
            registrar_falha("falha ao colar data início e falha ao limpar")
            return

        registrar_falha("falha ao colar data início")
        return

    if not etapa_finalizar():
        print("[ERRO] Falha ao finalizar - limpando antes de tentar novamente")

        if not limpar_e_confirmar():
            registrar_falha("falha ao finalizar e falha ao limpar")
            return

        registrar_falha("falha ao finalizar")
        return

    resultado = esperar_resultado_finalizacao()

    if resultado == "sucesso":
        confirmar_item_processado("flagado", "Ativo")
        resetar_falhas()
        mostrar_progresso()
        return

    if resultado == "ja_existe":
        print("[INFO] Limpando tela após erro de programa ativo")

        if not limpar_e_confirmar():
            registrar_falha("erro ativo detectado e falha ao limpar")
            return

        confirmar_item_processado("já existe", "Ativo")
        resetar_falhas()
        mostrar_progresso()
        return

    print("[ERRO] Sem sucesso e sem erro ativo - limpando antes de tentar novamente")

    if not limpar_e_confirmar():
        registrar_falha("sem sucesso/erro ativo e falha ao limpar")
        return

    registrar_falha("sem sucesso e sem erro ativo")
    return

# =========================================================
# LOOP AUTOMÁTICO
# =========================================================

def loop():
    global LOOP_ATIVO
    global tempo_inicio_loop
    global mos_processadas_loop
    global falhas_loop
    global alteracoes_pendentes

    tempo_inicio_loop = time.time()
    mos_processadas_loop = 0
    falhas_loop = []
    alteracoes_pendentes = 0


    total_mos = len(LISTA)
    restantes = total_mos - indice

    print("[LOOP] Iniciado")
    print(f"[LOOP] Total de MOs na lista: {total_mos}")
    print(f"[LOOP] MOs restantes para processar: {restantes}")
    print(f"[LOOP] Começando no índice: {indice + 1}")

    while LOOP_ATIVO:
        fluxo()

        if TEMPO_ENTRE_MOS > 0:
            time.sleep(TEMPO_ENTRE_MOS)

    tempo_total = time.time() - tempo_inicio_loop

    print("[LOOP] Finalizado")
    print(f"[RESUMO] MOs processadas neste loop: {mos_processadas_loop}")
    print(f"[RESUMO] Tempo total: {formatar_tempo(tempo_total)}")

    if mos_processadas_loop > 0:
        tempo_medio = tempo_total / mos_processadas_loop
        print(f"[RESUMO] Tempo médio por MO: {round(tempo_medio, 2)}s")
    else:
        print("[RESUMO] Nenhuma MO processada neste loop")

    if alteracoes_pendentes > 0:
        print("[INFO] Salvando alterações pendentes na planilha...")
        gravar_planilha()
    mostrar_resumo_falhas()

# =========================================================
# CONTROLE DO LOOP
# =========================================================

def alternar_loop():
    global LOOP_ATIVO

    LOOP_ATIVO = not LOOP_ATIVO

    if LOOP_ATIVO:
        threading.Thread(target=loop, daemon=True).start()
    else:
        print("[LOOP] Parando...")
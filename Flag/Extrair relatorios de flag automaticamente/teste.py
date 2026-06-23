from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time

# 1. Inicializa o navegador
driver = webdriver.Chrome()
driver.maximize_window()

# 2. Abre a página de login
driver.get("https://sisamil.amil.com.br/ace/ace001a.asp?tp_cliente=")

try:
    # 3. Preenche o campo Usuário
    campo_usuario = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "Login"))
    )
    campo_usuario.send_keys("le37118890")

    # 4. Preenche o campo Senha
    campo_senha = driver.find_element(By.NAME, "Senha")
    campo_senha.send_keys("Letimao18*")

    # 5. Clica no botão Entrar
    botao_entrar = driver.find_element(By.ID, "Submit1")
    botao_entrar.click()
   
    print("Sucesso: Logado")

except Exception as e:
    print(f"Ocorreu um erro no processo de login: {e}")

# =======================================================================
# PASSO INICIAL: ABRE O RELATÓRIO PELA PRIMEIRA E ÚNICA VEZ
# =======================================================================
try:
    driver.switch_to.default_content()
    
    # 1. Localiza e abre o menu
    botao_menu = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.ID, "mostraMenu"))
    )
    botao_menu.click()
    print("Menu aberto!")
    time.sleep(1.5)

    # 2. Clica na barra de pesquisa (70px para baixo)
    actions_pesquisa = ActionChains(driver)
    actions_pesquisa.move_to_element_with_offset(botao_menu, xoffset=120, yoffset=70).click().perform()
    time.sleep(0.5)
   
    # 3. Digita o nome do relatório desejado e dá ENTER
    texto_relatorio = "Relatório de Beneficiários por Programa de Gestão"
    actions_digitar = ActionChains(driver)
    actions_digitar.send_keys(texto_relatorio).send_keys(Keys.ENTER).perform()
    time.sleep(1.5)

    # 4. Clica no resultado da pesquisa (140px x 140px)
    print("Entrando na página do relatório...")
    actions_resultado = ActionChains(driver)
    actions_resultado.move_to_element_with_offset(botao_menu, xoffset=140, yoffset=140).click().perform()
    time.sleep(4.0) 

except Exception as e:
    print(f"Erro ao abrir o relatório inicialmente: {e}")

# =======================================================================
# LOOP OTIMIZADO COM SISTEMA DE RE-TENTATIVA DE FRAMES
# =======================================================================
#lista_codigos = ["1001", "1002", "1003", "1004", "1005", "1006", "1007", "1008", "1009", "1010", "1011", "1012", "1013", "1014", "1015", "1016", "1017"]
lista_codigos = ["1013", "1014"]

def focar_no_conteudo_do_relatorio_com_retry(drv):
    # Tenta em loop até 5 vezes encontrar os elementos, esperando o site carregar
    for tentativa in range(5):
        try:
            drv.switch_to.default_content()
            if len(drv.find_elements(By.NAME, "ind_ativo")) > 0:
                return True
            
            frames = drv.find_elements(By.XPATH, "//iframe | //frame")
            for index, frame in enumerate(frames):
                try:
                    drv.switch_to.default_content()
                    drv.switch_to.frame(index)
                    if len(drv.find_elements(By.NAME, "ind_ativo")) > 0:
                        return True
                except:
                    continue
        except:
            pass
        print(f"Aguardando estabilização da página (Tentativa {tentativa+1}/5)...")
        time.sleep(2.0) # Pausa antes de tentar ler os frames de novo
    return False

for i, codigo_atual in enumerate(lista_codigos):
    print(f"\n====================================================")
    print(f"PROCESSANDO CÓDIGO ({i+1}/{len(lista_codigos)}): {codigo_atual}")
    print(f"====================================================")

    try:
        # Força o reset do foco e aguarda
        driver.switch_to.default_content()
        time.sleep(1.5)

        # Garante o foco usando o novo sistema inteligente de tentativas
        if not focar_no_conteudo_do_relatorio_com_retry(driver):
            print(f"Aviso: Não conseguiu estabilizar os frames para o código {codigo_atual}. Tentando seguir na página principal...")
            driver.switch_to.default_content()

        # 1. Desmarca o checkbox 'ind_ativo'
        print("Ajustando checkbox...")
        checkbox = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.NAME, "ind_ativo"))
        )
        checkbox.click()
        time.sleep(0.5)

        # 2. Clica no campo de Código e insere o código atual do loop
        print(f"Inserindo código {codigo_atual}...")
        campo_codigo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "cod_programa"))
        )
        campo_codigo.click()
        campo_codigo.clear()
        campo_codigo.send_keys(codigo_atual)
        campo_codigo.send_keys(Keys.ENTER)
        time.sleep(1.0)

        # 3. Busca o botão Continuar (Run) nos frames pelo ID
        botao_run = None
        if len(driver.find_elements(By.ID, "btn_acao_continuar")) > 0:
            botao_run = driver.find_element(By.ID, "btn_acao_continuar")
        else:
            driver.switch_to.default_content() 
            total_frames = len(driver.find_elements(By.XPATH, "//iframe | //frame"))
            for index in range(total_frames):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(index)
                    if len(driver.find_elements(By.ID, "btn_acao_continuar")) > 0:
                        botao_run = driver.find_element(By.ID, "btn_acao_continuar")
                        break
                except:
                    continue

        if not botao_run:
            driver.switch_to.default_content()
            botao_run = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "btn_acao_continuar"))
            )

        botao_run.click()
        print("Botão Continuar clicado. Aguardando pop-ups...")
        time.sleep(2.0)

        # 4. Lida com os 2 Pop-ups (alertas)
        try:
            WebDriverWait(driver, 15).until(EC.alert_is_present())
            alerta1 = driver.switch_to.alert
            alerta1.accept()
            time.sleep(1.0)

            WebDriverWait(driver, 15).until(EC.alert_is_present())
            alerta2 = driver.switch_to.alert
            alerta2.accept()
            print("Pop-ups confirmados com sucesso!")
        except Exception as alert_error:
            print(f"Pop-up não detectado ou erro: {alert_error}")

        # 5. Busca e clica no botão Voltar (btn_acao_voltar) para retornar aos filtros
        time.sleep(3.0) 
        print("Buscando botão Voltar...")
        botao_voltar = None
        
        if len(driver.find_elements(By.ID, "btn_acao_voltar")) > 0:
            botao_voltar = driver.find_element(By.ID, "btn_acao_voltar")
        else:
            driver.switch_to.default_content()
            total_frames = len(driver.find_elements(By.XPATH, "//iframe | //frame"))
            for index in range(total_frames):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(index)
                    if len(driver.find_elements(By.ID, "btn_acao_voltar")) > 0:
                        botao_voltar = driver.find_element(By.ID, "btn_acao_voltar")
                        break
                except:
                    continue

        if not botao_voltar:
            driver.switch_to.default_content()
            botao_voltar = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "btn_acao_voltar"))
            )

        if botao_voltar:
            actions_voltar = ActionChains(driver)
            actions_voltar.move_to_element(botao_voltar).click().perform()
            print(f"Retornando para a tela de filtros.")

        # 6. ESPERA OBRIGATÓRIA DE 61 SEGUNDOS (Apenas se não for o último da lista)
        if i < len(lista_codigos) - 1:
            print("Aguardando 61 segundos para evitar duplicidade no agendamento...")
            for segundos_restantes in range(61, 0, -1):
                print(f"Próximo código em {segundos_restantes} segundos...", end="\r")
                time.sleep(1)
            print("\nTempo esgotado! Avançando para o próximo código...")

    except Exception as e:
        print(f"Erro ao processar o código {codigo_atual}: {e}")
        driver.switch_to.default_content()

print("\n====================================================")
print("Todos os relatórios foram disparados com sucesso!")
print("====================================================")

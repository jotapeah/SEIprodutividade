import tkinter as tk
from tkinter import ttk, Toplevel, Button, messagebox, filedialog, Label, Entry
import pyperclip
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import time
import re
import os
import pandas as pd
from openpyxl import load_workbook
import pyarrow
import sv_ttk  # Importando a biblioteca sv_ttk

# Variáveis globais para o driver e lista de processos com erro
driver = None
processos_nao_baixados = []  # Renomeada de processos_com_erro

# Caminho para o executável do Excel no seu sistema
excel_path = "C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXE"

# Função para processar um processo e classificá-lo
def processar_processo(driver, numero_processo, abortar):
    global processos_nao_baixados

    if abortar.is_set():
        return

    try:
        # Realizar pesquisa de processo/documento
        campo_pesquisa = driver.find_element(By.ID, "txtPesquisaRapida")
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(numero_processo)
        campo_pesquisa.send_keys(Keys.ENTER)

        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "#ifrVisualizacao")))

        # Verifique se o elemento com o CSS Selector "#frmProcedimentoPdf > label" está presente no segundo iframe
        elementos_com_erro = driver.find_elements(By.CSS_SELECTOR, "#ifrVisualizacao #frmProcedimentoPdf > label")

        # Usando o novo seletor CSS específico para o botão de download de PDF
        elementos_possiveis = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#divArvoreAcoes > a:nth-child(8) > img:nth-child(1)"))
        )

        encontrado = False

        # Como estamos usando o seletor específico, podemos simplificar esta parte
        if len(elementos_possiveis) > 0:
            elementos_possiveis[0].click()
            encontrado = True

        driver.switch_to.default_content()

        if not encontrado or elementos_com_erro:
            # Processo não encontrado ou com erro
            processos_nao_baixados.append({"Processo não baixado": numero_processo, "LOTE": ""})
        else:
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "#ifrVisualizacao")))

            # Usando o novo seletor CSS para o botão "Gerar"
            botoes_gerar = driver.find_elements(By.CSS_SELECTOR, "button.infraButton:nth-child(2)")

            if not botoes_gerar:
                # Não tem botão para gerar, então não foi possível baixar
                processos_nao_baixados.append({"Processo não baixado": numero_processo, "LOTE": ""})
            else:
                # Tem botão para gerar, tenta clicar nele
                segundo_botao = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.infraButton:nth-child(2)"))
                )
                segundo_botao.click()
                # O processo foi baixado com sucesso, não adiciona à lista de não baixados

        driver.switch_to.default_content()

    except Exception as e:
        print(f"Ocorreu um erro com o processo {numero_processo}: {e}")
        # Se ocorreu algum erro, adiciona à lista de não baixados
        processos_nao_baixados.append({"Processo não baixado": numero_processo, "LOTE": ""})

# Função para exportar a planilha
def exportar_planilha():
    global processos_nao_baixados

    # Crie DataFrame com a lista de processos não baixados
    df = pd.DataFrame(processos_nao_baixados)

    # Carregue a planilha de produção
    planilha_producao_path = planilha_entry.get()

    if not planilha_producao_path:
        messagebox.showwarning("Aviso", "Por favor, selecione uma planilha de produção.")
        return

    try:
        df_producao = pd.read_excel(planilha_producao_path)

        # Preencha a coluna "LOTE" na planilha de erro com base no mapeamento
        df["LOTE"] = df["Processo não baixado"].map(dict(zip(df_producao["PROCESSO INDIVIDUAL"], df_producao["LOTE"])))

        # Salve a planilha de erro
        df.to_excel("processos_nao_baixados.xlsx", index=False)
        os.system(f'start "{excel_path}" "processos_nao_baixados.xlsx"')

    except Exception as e:
        messagebox.showerror("Erro", str(e))

# Função para realizar login com retentativas
def realizar_login(driver, username, password, max_tentativas=3):
    for tentativa in range(max_tentativas):
        try:
            # Aguardar até que a página carregue e os campos estejam disponíveis
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "txtUsuario")))

            # Limpar e preencher o campo de usuário
            campo_usuario = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "txtUsuario"))
            )
            campo_usuario.clear()
            campo_usuario.send_keys(username)

            # Esperar um pouco para garantir que a página não está atualizando
            time.sleep(1)

            # Limpar e preencher o campo de senha
            campo_senha = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "pwdSenha"))
            )
            campo_senha.clear()
            campo_senha.send_keys(password)

            # Tentar várias abordagens para clicar no botão de login
            try:
                # Procurar botão pelo tipo e classe
                botao_login = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
                )
                botao_login.click()
            except:
                try:
                    # Procurar por qualquer botão que possa ser o de login
                    botoes = driver.find_elements(By.TAG_NAME, "button")
                    for botao in botoes:
                        if botao.is_displayed() and botao.is_enabled():
                            botao.click()
                            break
                except:
                    # Como último recurso, enviar ENTER no campo de senha
                    campo_senha.send_keys(Keys.ENTER)

            # Aguardar até que o login seja concluído e a página inicial seja carregada
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "txtPesquisaRapida")))
            return True  # Login bem-sucedido

        except Exception as e:
            print(f"Tentativa {tentativa+1} de login falhou: {e}")
            if tentativa == max_tentativas - 1:
                # Se for a última tentativa, relançar a exceção
                raise
            # Recarregar a página para uma nova tentativa
            driver.refresh()
            time.sleep(2)  # Dar tempo para a página recarregar

    return False  # Não deve chegar aqui, mas por segurança

# Função para iniciar a automação
def iniciar_automacao(username, password, lista_processos, abortar, finished_callback):
    global driver, processos_nao_baixados

    # Limpar a lista de processos não baixados no início da execução
    processos_nao_baixados = []

    if driver is None:
        driver = webdriver.Edge(service=Service('C:/WebDrivers/msedgedriver.exe'))

    try:
        driver.get("https://sei.incra.gov.br/sip/login.php?sigla_orgao_sistema=INCRA&sigla_sistema=SEI&infra_url=L3NlaS8=")

        # Usar a nova função de login com retentativas
        realizar_login(driver, username, password)

        for numero_processo in lista_processos:
            processar_processo(driver, numero_processo, abortar)
            if abortar.is_set():
                break
            time.sleep(2)

        print("Automação finalizada.")

    except Exception as e:
        messagebox.showerror("Erro", str(e))
    finally:
        finished_callback()

# Função para iniciar a thread de automação
def iniciar_thread(username, password, lista_processos, abortar, finished_callback):
    thread = threading.Thread(target=iniciar_automacao, args=(username, password, lista_processos, abortar, finished_callback))
    thread.start()

# Função chamada quando a automação é finalizada
def on_finished():
    global processos_nao_baixados
    print("Automação finalizada.")
    messagebox.showwarning("Execução concluída!", "Abra a janela do programa para abrir a planilha.")

# Função para abortar o programa
def on_abortar_pressed(abortar, root):
    global driver
    abortar.set()
    if driver is not None:
        driver.quit()  # Fecha o navegador quando o botão Abortar é pressionado
    root.destroy()

# Função para abrir a janela de seleção de arquivo e preencher o campo de planilha
def selecionar_planilha_producao():
    planilha_producao = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
    if planilha_producao:
        planilha_entry.delete(0, "end")
        planilha_entry.insert(0, planilha_producao)

# Função para iniciar a automação a partir do botão
def iniciar_automatico():
    username = username_entry.get()
    password = password_entry.get()
    entrada_processos = processos_entry.get()
    lista_processos = re.findall(r'\d{5}\.\d{6}/\d{4}-\d{2}', entrada_processos)

    if username and password:
        iniciar_thread(username, password, lista_processos, abortar, on_finished)
    else:
        messagebox.showwarning("Aviso", "Por favor, preencha todas as informações.")

# Função para criar, relacionar e abrir planilha
def criar_relacionar_abrir_planilha():
    exportar_planilha()

# Interface gráfica
root = tk.Tk()
root.title("Script download de processos SEI! v.2.5")

# Aplicar tema Sun Valley
sv_ttk.set_theme("dark")

# Frame principal com padding
main_frame = ttk.Frame(root, padding=20)
main_frame.pack(fill=tk.BOTH, expand=True)

# Título e subtítulo
title_frame = ttk.Frame(main_frame)
title_frame.pack(fill=tk.X, pady=(0, 20))

title_label = ttk.Label(
    title_frame,
    text="Download de Processos SEI",
    font=("Segoe UI", 16, "bold")
)
title_label.pack()

subtitle_label = ttk.Label(
    title_frame,
    text="Ferramenta para download automático de processos",
    font=("Segoe UI", 10)
)
subtitle_label.pack()

# Separador
ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)

# Carregar e redimensionar a imagem
logo_frame = ttk.Frame(main_frame)
logo_frame.pack(pady=10)

logo_path = "C:/imagem_processosSEI.png"
if os.path.exists(logo_path):
    logo_image = tk.PhotoImage(file=logo_path)
    logo_image = logo_image.subsample(3)  # Ajuste o valor do subsample conforme necessário
    logo_label = ttk.Label(logo_frame, image=logo_image)
    logo_label.pack()

# Adicionando o texto informativo mais próximo da imagem
info_label = ttk.Label(
    main_frame, 
    text="Crie a pasta 'WebDrivers' no disco C, baixe o Microsoft Edge WebDriver e salve nessa pasta antes de executar o programa.",
    wraplength=600,
    justify="center"
)
info_label.pack(pady=(0, 20))

# Variável para controle de abortar
abortar = threading.Event()

# Frame de configurações
config_frame = ttk.LabelFrame(main_frame, text="Configurações", padding=15)
config_frame.pack(fill=tk.X, pady=10)

# Usuário
user_frame = ttk.Frame(config_frame)
user_frame.pack(fill=tk.X, pady=5)

ttk.Label(user_frame, text="Usuário SEI:", width=15).pack(side=tk.LEFT, padx=(0, 5))
username_entry = ttk.Entry(user_frame)
username_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

# Senha
pass_frame = ttk.Frame(config_frame)
pass_frame.pack(fill=tk.X, pady=5)

ttk.Label(pass_frame, text="Senha SEI:", width=15).pack(side=tk.LEFT, padx=(0, 5))
password_entry = ttk.Entry(pass_frame, show='*')
password_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

# Lista de processos
proc_frame = ttk.Frame(config_frame)
proc_frame.pack(fill=tk.X, pady=5)

ttk.Label(proc_frame, text="Lista de processos:", width=15).pack(side=tk.LEFT, padx=(0, 5))
processos_entry = ttk.Entry(proc_frame)
processos_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

# Planilha de produção
planilha_frame = ttk.Frame(config_frame)
planilha_frame.pack(fill=tk.X, pady=5)

ttk.Label(planilha_frame, text="Planilha:", width=15).pack(side=tk.LEFT, padx=(0, 5))
planilha_entry = ttk.Entry(planilha_frame)
planilha_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

selecionar_planilha_button = ttk.Button(
    planilha_frame, 
    text="Selecionar", 
    command=selecionar_planilha_producao,
    style="Accent.TButton"
)
selecionar_planilha_button.pack(side=tk.LEFT)

# Frame para os botões de ação
action_frame = ttk.Frame(main_frame)
action_frame.pack(pady=20)

# Botões com estilos do Sun Valley
iniciar_button = ttk.Button(
    action_frame, 
    text="Iniciar Execução", 
    command=iniciar_automatico,
    style="Accent.TButton",
    width=20
)
iniciar_button.pack(side=tk.LEFT, padx=5)

criar_relacionar_abrir_button = ttk.Button(
    action_frame, 
    text="Cruzar Informações", 
    command=criar_relacionar_abrir_planilha,
    width=20
)
criar_relacionar_abrir_button.pack(side=tk.LEFT, padx=5)

abortar_button = ttk.Button(
    action_frame, 
    text="Abortar Programa", 
    command=lambda: on_abortar_pressed(abortar, root),
    width=15
)
abortar_button.pack(side=tk.LEFT, padx=5)

# Rodapé
footer_frame = ttk.Frame(main_frame)
footer_frame.pack(fill=tk.X, pady=(20, 0))

ttk.Label(
    footer_frame, 
    text="© 2025 Download Processos SEI - v.2.5",
    font=("Segoe UI", 8)
).pack(side=tk.RIGHT)

# Centralizar a janela
root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f'{width}x{height}+{x}+{y}')

root.mainloop()

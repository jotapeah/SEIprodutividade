import tkinter as tk
from tkinter import Toplevel, Button, messagebox, filedialog, Label, Entry
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

# Variáveis globais para o driver e lista de processos com erro
driver = None
processos_com_erro = []
processos_bloqueados = []

# Caminho para o executável do Excel no seu sistema
excel_path = "C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXE"

# Função para processar um processo e classificá-lo
def processar_processo(driver, numero_processo, abortar):
    global processos_com_erro, processos_bloqueados

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

        elementos_possiveis = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#divArvoreAcoes > a"))
        )

        encontrado = False

        for elemento in elementos_possiveis:
            img_elemento = elemento.find_element(By.TAG_NAME, "img")
            if "Gerar Arquivo PDF do Processo" in img_elemento.get_attribute("title"):
                elemento.click()
                encontrado = True
                break

        driver.switch_to.default_content()

        if not encontrado or elementos_com_erro:
            processos_com_erro.append({"Processo não encontrado": numero_processo, "LOTE": ""})
            processos_bloqueados.append({"Processos com download indisponível": numero_processo, "LOTE": ""})

        else:
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "#ifrVisualizacao")))

            # Verifique se o botão "Gerar" está presente, se não estiver, considere-o bloqueado
            if not driver.find_elements(By.CSS_SELECTOR, "button[name='btnGerar'][value='Gerar']"):
                processos_bloqueados.append({"Processos com download indisponível": numero_processo, "LOTE": ""})
            else:
                segundo_botao = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button[name='btnGerar'][value='Gerar']"))
                )
                segundo_botao.click()
                processos_com_erro.append({"Processo não encontrado": numero_processo, "LOTE": ""})

        driver.switch_to.default_content()

    except Exception as e:
        print(f"Ocorreu um erro com o processo {numero_processo}: {e}")
        processos_com_erro.append({"Processo não encontrado": numero_processo, "LOTE": ""})
        processos_bloqueados.append({"Processos com download indisponível": numero_processo, "LOTE": ""})

# Função para exportar a planilha
def exportar_planilha():
    global processos_com_erro, processos_bloqueados

    # Certifique-se de que ambas as listas tenham o mesmo comprimento
    while len(processos_com_erro) < len(processos_bloqueados):
        processos_com_erro.append({"Processo não encontrado": "", "LOTE": ""})  # Adicione dicionários vazios
    while len(processos_bloqueados) < len(processos_com_erro):
        processos_bloqueados.append({"Processos com download indisponível": "", "LOTE": ""})  # Adicione dicionários vazios

    # Crie DataFrames com as listas de dicionários
    df_com_erro = pd.DataFrame(processos_com_erro)
    df_bloqueados = pd.DataFrame(processos_bloqueados)

    # Concatene os DataFrames
    df = pd.concat([df_com_erro, df_bloqueados], axis=0, ignore_index=True)

    # Reordene as colunas na ordem desejada
    colunas = ["Processo não encontrado", "Processos com download indisponível", "LOTE"]
    df = df[colunas]

    # Carregue a planilha de produção
    planilha_producao_path = planilha_entry.get()

    if not planilha_producao_path:
        messagebox.showwarning("Aviso", "Por favor, selecione uma planilha de produção.")
        return

    try:
        df_producao = pd.read_excel(planilha_producao_path)

        # Preencha a coluna "LOTE" na planilha de erro com base no mapeamento
        df["LOTE"] = df["Processo não encontrado"].map(dict(zip(df_producao["PROCESSO INDIVIDUAL"], df_producao["LOTE"])))

        # Preencha a coluna "LOTE" na planilha de erro para processos com download indisponível
        df["LOTE"] = df["LOTE"].combine_first(df["Processos com download indisponível"].map(dict(zip(df_producao["PROCESSO INDIVIDUAL"], df_producao["LOTE"]))))

        # Salve a planilha de erro
        df.to_excel("processos_com_erro.xlsx", index=False)
        os.system(f'start "{excel_path}" "processos_com_erro.xlsx"')

    except Exception as e:
        messagebox.showerror("Erro", str(e))

# Função para iniciar a automação
def iniciar_automacao(username, password, lista_processos, abortar, finished_callback):
    global driver, processos_com_erro, processos_bloqueados

    if driver is None:
        driver = webdriver.Edge(service=Service('C:/WebDrivers/msedgedriver.exe'))

    try:
        driver.get("https://sei.incra.gov.br/sip/login.php?sigla_orgao_sistema=INCRA&sigla_sistema=SEI&infra_url=L3NlaS8=")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "txtUsuario")))

        campo_usuario = driver.find_element(By.ID, "txtUsuario")
        campo_usuario.click()
        campo_usuario.send_keys(username)

        time.sleep(1)

        campo_senha = driver.find_element(By.ID, "pwdSenha")
        campo_senha.click()
        campo_senha.send_keys(password)

        botao_login = driver.find_element(By.ID, "sbmLogin")
        botao_login.click()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "txtPesquisaRapida")))

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
    global processos_com_erro
    print("Automação finalizada.")
    messagebox.showwarning("Execução concluída!", "Abra a janela do programa para abrir a planilha.")

# Função para abortar o programa
def on_abortar_pressed(abortar, root):
    global driver
    abortar.set()
    if driver is not None:
        driver.quit()  # Fecha o navegador quando o botão Abortar é pressionado
    root.destroy()

# Interface gráfica
root = tk.Tk()
root.title("Script download de processos SEI! v.2.4")

# Criando um frame para conter a imagem e os widgets
main_frame = tk.Frame(root)
main_frame.pack()

# Carregar e redimensionar a imagem
logo_path = "C:/imagem_processosSEI.png"
if os.path.exists(logo_path):
    logo_image = tk.PhotoImage(file=logo_path)
    logo_image = logo_image.subsample(3)  # Ajuste o valor do subsample conforme necessário
    logo_label = tk.Label(main_frame, image=logo_image)
    logo_label.pack()

# Adicionando o texto informativo mais próximo da imagem
info_label = tk.Label(main_frame, text="Crie a pasta 'WebDrivers' no disco C, baixe o Microsoft Edge WebDriver e salve nessa pasta antes de executar o programa.")
info_label.pack(pady=(10, 20))  # Aumente o espaço após o texto informativo

abortar = threading.Event()

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

username_label = tk.Label(frame, text="Usuário SEI:")
username_label.pack()
username_entry = tk.Entry(frame)
username_entry.pack()

password_label = tk.Label(frame, text="Senha SEI:")
password_label.pack()
password_entry = tk.Entry(frame, show='*')
password_entry.pack()

processos_label = tk.Label(frame, text="Lista de processos:")
processos_label.pack()
processos_entry = tk.Entry(frame)
processos_entry.pack()

# Campo para selecionar a planilha de produção
planilha_label = Label(frame, text="Selecione a planilha de produção:")
planilha_label.pack()
planilha_entry = Entry(frame, width=40)
planilha_entry.pack()

# Função para abrir a janela de seleção de arquivo e preencher o campo de planilha
def selecionar_planilha_producao():
    planilha_producao = filedialog.askopenfilename()
    if planilha_producao:
        planilha_entry.delete(0, "end")
        planilha_entry.insert(0, planilha_producao)

# Botão para selecionar a planilha de produção
selecionar_planilha_button = tk.Button(frame, text="Selecionar planilha", command=selecionar_planilha_producao)
selecionar_planilha_button.pack()

# Botão Iniciar
def iniciar_automatico():
    username = username_entry.get()
    password = password_entry.get()
    entrada_processos = processos_entry.get()
    lista_processos = re.findall(r'\d{5}\.\d{6}/\d{4}-\d{2}', entrada_processos)

    if username and password:
        iniciar_thread(username, password, lista_processos, abortar, on_finished)
    else:
        messagebox.showwarning("Aviso", "Por favor, preencha todas as informações.")

iniciar_button = tk.Button(frame, text="Iniciar execução", command=iniciar_automatico)
iniciar_button.pack(pady=(10, 0))

# Botão para criar, relacionar e abrir planilha
def criar_relacionar_abrir_planilha():
    exportar_planilha()

criar_relacionar_abrir_button = tk.Button(main_frame, text="Cruzar informações e abrir planilha", command=criar_relacionar_abrir_planilha)
criar_relacionar_abrir_button.pack()

# Botão para abortar o programa
abortar_button = tk.Button(frame, text="Abortar programa", command=lambda: on_abortar_pressed(abortar, root))
abortar_button.pack(pady=10)

root.mainloop()

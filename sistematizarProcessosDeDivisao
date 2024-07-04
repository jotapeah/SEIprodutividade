import csv
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def selecionar_webdriver():
    arquivo_webdriver = filedialog.askopenfilename(filetypes=[("Arquivos Executáveis", "*.exe")])
    if arquivo_webdriver:
        webdriver_path_entry.delete(0, tk.END)
        webdriver_path_entry.insert(0, arquivo_webdriver)

def selecionar_local_salvar():
    arquivo_salvar = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Arquivo CSV", "*.csv")])
    if arquivo_salvar:
        arquivo_salvar_entry.delete(0, tk.END)
        arquivo_salvar_entry.insert(0, arquivo_salvar)

def iniciar_raspagem():
    webdriver_path = webdriver_path_entry.get()
    service = webdriver.EdgeService(executable_path=webdriver_path)
    driver = webdriver.Edge(service=service)

    url_sei = url_sei_entry.get()
    if not url_sei:
        resultado_label.config(text="A URL do SEI! não foi fornecida.")
        return

    driver.get(url_sei)

    campo_usuario = driver.find_element(By.ID, 'txtUsuario')
    campo_senha = driver.find_element(By.ID, 'pwdSenha')
    usuario = usuario_entry.get()
    senha = senha_entry.get()

    if not usuario or not senha:
        resultado_label.config(text="Usuário e senha são obrigatórios.")
        driver.quit()
        return

    campo_usuario.send_keys(usuario)
    campo_senha.send_keys(senha)

    botao_login = driver.find_element(By.ID, 'sbmLogin')
    botao_login.click()

    campo_pesquisa_rapida = driver.find_element(By.ID, "txtPesquisaRapida")
    campo_pesquisa_rapida.send_keys("")
    campo_pesquisa_rapida.send_keys(Keys.RETURN)
    opcao_pesquisar_processos = driver.find_element(By.CSS_SELECTOR, "input#optProcessos")
    opcao_pesquisar_processos.click()

    campo_unidade_geradora = driver.find_element(By.CSS_SELECTOR, "input#txtUnidade")
    campo_unidade_geradora.click()
    nome_unidade_geradora = unidade_geradora_entry.get()

    if not nome_unidade_geradora:
        resultado_label.config(text="O nome da unidade geradora é obrigatório.")
        driver.quit()
        return

    for char in nome_unidade_geradora:
        campo_unidade_geradora.send_keys(char)
        time.sleep(0.5)

    time.sleep(2)
    campo_unidade_geradora.send_keys(Keys.ARROW_DOWN)
    campo_unidade_geradora.send_keys(Keys.ENTER)

    botao_pesquisar = driver.find_element(By.CSS_SELECTOR, '#sbmPesquisar')
    botao_pesquisar.click()

    caminho_arquivo = arquivo_salvar_entry.get()

    def realizar_scraping():
        while True:
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#conteudo > table > tbody > tr.resTituloRegistro > td.resTituloEsquerda"))
                )

                elementos = driver.find_elements(By.CSS_SELECTOR, "#conteudo > table > tbody > tr.resTituloRegistro > td.resTituloEsquerda")

                with open(caminho_arquivo, mode='a', newline='', encoding='utf-8') as arquivo:
                    escrever = csv.writer(arquivo)
                    for elemento in elementos:
                        escrever.writerow([elemento.text])

                next_link = driver.find_element(By.XPATH, "//a[contains(text(), 'Próxima')]")
                next_link_href = next_link.get_attribute("href")

                if next_link_href:
                    driver.execute_script(next_link_href)
                    WebDriverWait(driver, 10).until(
                        EC.staleness_of(next_link)
                    )
                else:
                    break

            except TimeoutException:
                print("Fim das páginas ou página demorou muito para carregar.")
                break

    realizar_scraping()
    driver.quit()
    resultado_label.config(text="Sucesso!")

app = tk.Tk()
app.title("Raspagem de Dados do SEI!")

webdriver_path_label = tk.Label(app, text="Caminho do Webdriver:")
url_sei_label = tk.Label(app, text="URL do SEI!:")
usuario_label = tk.Label(app, text="Usuário:")
senha_label = tk.Label(app, text="Senha:")
unidade_geradora_label = tk.Label(app, text="Nome da Unidade Geradora:")
arquivo_salvar_label = tk.Label(app, text="Arquivo CSV para Salvar:")

resultado_label = tk.Label(app, text="", fg="green")

webdriver_path_entry = tk.Entry(app)
url_sei_entry = tk.Entry(app)
usuario_entry = tk.Entry(app)
senha_entry = tk.Entry(app, show="*")
unidade_geradora_entry = tk.Entry(app)
arquivo_salvar_entry = tk.Entry(app)

selecionar_webdriver_button = tk.Button(app, text="Selecionar Webdriver", command=selecionar_webdriver)
selecionar_local_salvar_button = tk.Button(app, text="Selecionar Arquivo para Salvar", command=selecionar_local_salvar)

iniciar_button = tk.Button(app, text="Iniciar Raspagem", command=iniciar_raspagem)

webdriver_path_label.grid(row=0, column=0, sticky="e")
url_sei_label.grid(row=1, column=0, sticky="e")
usuario_label.grid(row=2, column=0, sticky="e")
senha_label.grid(row=3, column=0, sticky="e")
unidade_geradora_label.grid(row=4, column=0, sticky="e")
arquivo_salvar_label.grid(row=5, column=0, sticky="e")

webdriver_path_entry.grid(row=0, column=1)
url_sei_entry.grid(row=1, column=1)
usuario_entry.grid(row=2, column=1)
senha_entry.grid(row=3, column=1)
unidade_geradora_entry.grid(row=4, column=1)
arquivo_salvar_entry.grid(row=5, column=1)

selecionar_webdriver_button.grid(row=0, column=2)
selecionar_local_salvar_button.grid(row=5, column=2)

iniciar_button.grid(row=6, columnspan=3)
resultado_label.grid(row=7, columnspan=3)

app.mainloop()

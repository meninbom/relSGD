import tkinter as tk
from tkinter import messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
import requests
from urllib.parse import urljoin
import threading
import json
import os
from datetime import datetime

CACHE_FILE = "cache.json"
LOG_FILE = "log.txt"
ICON_PATH = None  # Coloque o caminho do seu .ico aqui, se desejar

# --------- Tema ---------
THEMES = {
    "claro": {
        "bg": "#f4f6fa",
        "fg": "#2d415a",
        "entry_bg": "white",
        "entry_fg": "#2d415a",
        "btn_bg": "#2d415a",
        "btn_fg": "white",
        "status_ok": "#2e7d32",
        "status_warn": "#c62828",
        "rodape": "#a0a4ad"
    },
    "escuro": {
        "bg": "#23272e",
        "fg": "#e0e6ed",
        "entry_bg": "#2d333b",
        "entry_fg": "#e0e6ed",
        "btn_bg": "#4f5b93",
        "btn_fg": "white",
        "status_ok": "#81c784",
        "status_warn": "#e57373",
        "rodape": "#6c6f7a"
    }
}
current_theme = "claro"

def get_theme():
    return THEMES[current_theme]

# --------- Utilitários ---------
def salvar_cache(login, senha, data_inicial, data_final):
    cache = {
        "login": login,
        "senha": senha,
        "data_inicial": data_inicial,
        "data_final": data_final
    }
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f)

def carregar_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            try:
                cache = json.load(f)
                return cache
            except Exception:
                return {}
    return {}

def registrar_log(login, data_inicial, data_final, status, mensagem=""):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"[{agora}] Login: {login} | Data Inicial: {data_inicial} | Data Final: {data_final} | Status: {status} | {mensagem}\n")

def preencher_campo_data(driver, xpath, valor):
    campo = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    campo.clear()
    campo.click()
    for char in valor:
        campo.send_keys(char)
        time.sleep(0.05)
    assert campo.get_attribute("value") == valor

def validar_datas(data_inicial, data_final):
    formatos = ["%d/%m/%y", "%d/%m/%Y"]
    dt_ini = dt_fim = None
    for fmt in formatos:
        try:
            dt_ini = datetime.strptime(data_inicial, fmt)
            break
        except ValueError:
            continue
    for fmt in formatos:
        try:
            dt_fim = datetime.strptime(data_final, fmt)
            break
        except ValueError:
            continue
    if not dt_ini or not dt_fim:
        return False, "As datas devem estar no formato dd/mm/aa ou dd/mm/aaaa (ex: 01/06/24 ou 01/06/2024)."
    if dt_ini > dt_fim:
        return False, "A data inicial não pode ser maior que a data final."
    return True, ""

def centralizar_janela(janela, largura, altura):
    janela.update_idletasks()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")

def limpar_campos():
    entry_login.delete(0, tk.END)
    entry_senha.delete(0, tk.END)
    entry_data_inicial.delete(0, tk.END)
    entry_data_final.delete(0, tk.END)
    status_label.config(text="", fg=get_theme()["fg"])
    for e in [entry_login, entry_senha, entry_data_inicial, entry_data_final]:
        e.config(highlightbackground=get_theme()["entry_bg"], highlightcolor=get_theme()["entry_bg"])

def toggle_senha():
    if entry_senha.cget('show') == '':
        entry_senha.config(show='*')
        btn_toggle_senha.config(text='👁')
    else:
        entry_senha.config(show='')
        btn_toggle_senha.config(text='🔒')

def toggle_tema():
    global current_theme
    current_theme = "escuro" if current_theme == "claro" else "claro"
    aplicar_tema()

def aplicar_tema():
    th = get_theme()
    root.configure(bg=th["bg"])
    frame.configure(bg=th["bg"])
    lbl_titulo.configure(bg=th["bg"], fg=th["fg"])
    lbl_login.configure(bg=th["bg"], fg=th["fg"])
    lbl_senha.configure(bg=th["bg"], fg=th["fg"])
    lbl_data_inicial.configure(bg=th["bg"], fg=th["fg"])
    lbl_data_final.configure(bg=th["bg"], fg=th["fg"])
    senha_frame.configure(bg=th["bg"])
    entry_login.configure(bg=th["entry_bg"], fg=th["entry_fg"], insertbackground=th["entry_fg"])
    entry_senha.configure(bg=th["entry_bg"], fg=th["entry_fg"], insertbackground=th["entry_fg"])
    entry_data_inicial.configure(bg=th["entry_bg"], fg=th["entry_fg"], insertbackground=th["entry_fg"])
    entry_data_final.configure(bg=th["entry_bg"], fg=th["entry_fg"], insertbackground=th["entry_fg"])
    btn.configure(bg=th["btn_bg"], fg=th["btn_fg"], activebackground=th["btn_bg"], activeforeground=th["btn_fg"])
    btn_toggle_senha.configure(bg=th["bg"], fg=th["fg"], activebackground=th["bg"])
    btn_limpar.configure(bg=th["btn_bg"], fg=th["btn_fg"], activebackground=th["btn_bg"], activeforeground=th["btn_fg"])
    btn_tema.configure(bg=th["btn_bg"], fg=th["btn_fg"], activebackground=th["btn_bg"], activeforeground=th["btn_fg"])
    status_label.configure(bg=th["bg"])
    lbl_rodape.configure(bg=th["bg"], fg=th["rodape"])

def marcar_invalido(entry):
    entry.config(highlightbackground=get_theme()["status_warn"], highlightcolor=get_theme()["status_warn"], highlightthickness=2)

def marcar_valido(entry):
    entry.config(highlightbackground=get_theme()["entry_bg"], highlightcolor=get_theme()["entry_bg"], highlightthickness=1)

def escolher_arquivo_sugestao(login, data_inicial, data_final):
    nome = f"relatorio_{login}_{data_inicial.replace('/','-')}_{data_final.replace('/','-')}.pdf"
    return filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("Todos os arquivos", "*.*")],
        initialfile=nome,
        title="Salvar relatório como"
    )

def baixar_relatorio(login, senha, data_inicial, data_final, status_label, caminho_arquivo):
    inicio_exec = datetime.now()
    try:
        status_label.config(text="Iniciando navegador...", fg=get_theme()["fg"])
        driver = webdriver.Chrome()  # ou o caminho do chromedriver

        # 1. Login
        driver.get("https://sgd.dominiosistemas.com.br/login.html")
        driver.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys(login)
        driver.find_element(By.XPATH, '/html/body/div/form/input[3]').send_keys(senha)
        driver.find_element(By.XPATH, '/html/body/div/form/input[3]').submit()
        try:
            WebDriverWait(driver, 10).until(EC.url_changes("https://sgd.dominiosistemas.com.br/login.html"))
        except Exception:
            status_label.config(text="Login inválido ou site indisponível.", fg=get_theme()["status_warn"])
            registrar_log(login, data_inicial, data_final, "ERRO", "Login inválido ou site indisponível.")
            return

        # 2. Relatório
        driver.get("https://sgd.dominiosistemas.com.br/sgsc/faces/rel-satisfacao-externa-funcionario.html")
        preencher_campo_data(driver, '//*[@id="formFiltroRelatorio:dataInicial"]', data_inicial)
        preencher_campo_data(driver, '//*[@id="formFiltroRelatorio:dataFinal"]', data_final)

        wait = WebDriverWait(driver, 10)
        select_element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="formFiltroRelatorio:ordem"]')))
        select = Select(select_element)
        select.select_by_visible_text("Atendimentos Realizados")

        botao_gerar = driver.find_element(By.XPATH, '/html/body/div[2]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr/td/table/tbody/tr[3]/td/input')
        driver.execute_script("arguments[0].click();", botao_gerar)

        status_label.config(text="Aguardando relatório...", fg=get_theme()["fg"])
        time.sleep(10)

        # 3. Download
        try:
            link_download = driver.find_element(By.XPATH, '/html/body/div[2]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[3]/td/table/tbody/tr[5]/td/a')
        except Exception:
            status_label.config(text="Relatório não encontrado.", fg=get_theme()["status_warn"])
            registrar_log(login, data_inicial, data_final, "ERRO", "Relatório não encontrado.")
            return

        href_relativo = link_download.get_attribute('href')
        url_base = "https://sgd.dominiosistemas.com.br"
        url_arquivo = urljoin(url_base, href_relativo)

        cookies = driver.get_cookies()
        s = requests.Session()
        for cookie in cookies:
            s.cookies.set(cookie['name'], cookie['value'])

        status_label.config(text="Baixando arquivo...", fg=get_theme()["fg"])
        response = s.get(url_arquivo)
        with open(caminho_arquivo, 'wb') as f:
            f.write(response.content)

        status_label.config(text="Download concluído!", fg=get_theme()["status_ok"])
        messagebox.showinfo("Sucesso", f"Relatório baixado como\n{os.path.basename(caminho_arquivo)}")
        registrar_log(login, data_inicial, data_final, "SUCESSO", f"Arquivo: {caminho_arquivo}")
        salvar_cache(login, senha, data_inicial, data_final)
    except Exception as e:
        status_label.config(text="Erro!", fg=get_theme()["status_warn"])
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        registrar_log(login, data_inicial, data_final, "ERRO", str(e))
    finally:
        try:
            driver.quit()
        except:
            pass
        fim_exec = datetime.now()
        duracao = (fim_exec - inicio_exec).total_seconds()
        print(f"Tempo de execução: {duracao:.1f}s")

def iniciar_download():
    # Limpa marcações anteriores
    for e in [entry_login, entry_senha, entry_data_inicial, entry_data_final]:
        marcar_valido(e)

    login = entry_login.get().strip()
    senha = entry_senha.get().strip()
    data_inicial = entry_data_inicial.get().strip()
    data_final = entry_data_final.get().strip()
    campos = [login, senha, data_inicial, data_final]
    entradas = [entry_login, entry_senha, entry_data_inicial, entry_data_final]
    validos = True

    for i, valor in enumerate(campos):
        if not valor:
            marcar_invalido(entradas[i])
            validos = False
    if not validos:
        status_label.config(text="Preencha todos os campos!", fg=get_theme()["status_warn"])
        return

    valido, msg = validar_datas(data_inicial, data_final)
    if not valido:
        status_label.config(text=msg, fg=get_theme()["status_warn"])
        marcar_invalido(entry_data_inicial)
        marcar_invalido(entry_data_final)
        return

    # Escolher local e nome do arquivo
    caminho_arquivo = escolher_arquivo_sugestao(login, data_inicial, data_final)
    if not caminho_arquivo:
        status_label.config(text="Download cancelado pelo usuário.", fg=get_theme()["status_warn"])
        return

    status_label.config(text="Iniciando download...", fg=get_theme()["fg"])
    threading.Thread(target=baixar_relatorio, args=(login, senha, data_inicial, data_final, status_label, caminho_arquivo), daemon=True).start()

# --------- Interface Tkinter aprimorada ---------
root = tk.Tk()
root.title("Relatório Domínio")
if ICON_PATH and os.path.exists(ICON_PATH):
    root.iconbitmap(ICON_PATH)
LARGURA, ALTURA = 300, 300
root.geometry(f"{LARGURA}x{ALTURA}")
root.resizable(False, False)
centralizar_janela(root, LARGURA, ALTURA)

frame = tk.Frame(root, bg=get_theme()["bg"])
frame.pack(expand=True, fill="both", padx=16, pady=10)

lbl_titulo = tk.Label(frame, text="Download de Relatório", font=("Segoe UI", 14, "bold"), bg=get_theme()["bg"], fg=get_theme()["fg"])
lbl_titulo.pack(pady=(0, 10))

lbl_login = tk.Label(frame, text="Login:", font=("Segoe UI", 10), bg=get_theme()["bg"], fg=get_theme()["fg"])
lbl_login.pack(anchor="w")
entry_login = tk.Entry(frame, width=28, font=("Segoe UI", 10), bg=get_theme()["entry_bg"], fg=get_theme()["entry_fg"], highlightthickness=1)
entry_login.pack(pady=(0, 6))
entry_login.insert(0, "")

lbl_senha = tk.Label(frame, text="Senha:", font=("Segoe UI", 10), bg=get_theme()["bg"], fg=get_theme()["fg"])
lbl_senha.pack(anchor="w")
senha_frame = tk.Frame(frame, bg=get_theme()["bg"])
senha_frame.pack(pady=(0, 6), fill="x")
entry_senha = tk.Entry(senha_frame, width=20, font=("Segoe UI", 10), show="*", bg=get_theme()["entry_bg"], fg=get_theme()["entry_fg"], highlightthickness=1)
entry_senha.pack(side="left", fill="x", expand=True)
btn_toggle_senha = tk.Button(senha_frame, text="👁", font=("Segoe UI", 8), command=toggle_senha, width=3, relief="flat", bg=get_theme()["bg"], fg=get_theme()["fg"])
btn_toggle_senha.pack(side="left", padx=(4,0))

lbl_data_inicial = tk.Label(frame, text="Data Inicial (dd/mm/aa ou aaaa):", font=("Segoe UI", 10), bg=get_theme()["bg"], fg=get_theme()["fg"])
lbl_data_inicial.pack(anchor="w")
entry_data_inicial = tk.Entry(frame, width=15, font=("Segoe UI", 10), bg=get_theme()["entry_bg"], fg=get_theme()["entry_fg"], highlightthickness=1)
entry_data_inicial.pack(pady=(0, 6))

lbl_data_final = tk.Label(frame, text="Data Final (dd/mm/aa ou aaaa):", font=("Segoe UI", 10), bg=get_theme()["bg"], fg=get_theme()["fg"])
lbl_data_final.pack(anchor="w")
entry_data_final = tk.Entry(frame, width=15, font=("Segoe UI", 10), bg=get_theme()["entry_bg"], fg=get_theme()["entry_fg"], highlightthickness=1)
entry_data_final.pack(pady=(0, 10))

btn = tk.Button(frame, text="Baixar Relatório", font=("Segoe UI", 11, "bold"), bg=get_theme()["btn_bg"], fg=get_theme()["btn_fg"], command=iniciar_download, width=22, height=1)
btn.pack(pady=(0, 4))

btn_limpar = tk.Button(frame, text="Limpar Campos", font=("Segoe UI", 9), bg=get_theme()["btn_bg"], fg=get_theme()["btn_fg"], command=limpar_campos, width=22, height=1)
btn_limpar.pack(pady=(0, 4))

btn_tema = tk.Button(frame, text="Alternar Tema", font=("Segoe UI", 9), bg=get_theme()["btn_bg"], fg=get_theme()["btn_fg"], command=toggle_tema, width=22, height=1)
btn_tema.pack(pady=(0, 4))

status_label = tk.Label(frame, text="", font=("Segoe UI", 9), bg=get_theme()["bg"], fg=get_theme()["fg"])
status_label.pack(pady=(0, 2))

lbl_rodape = tk.Label(root, text="Domínio Sistemas • 2024", font=("Segoe UI", 8), bg=get_theme()["bg"], fg=get_theme()["rodape"])
lbl_rodape.pack(side="bottom", pady=(0, 2))

# Carrega cache ao iniciar
cache = carregar_cache()
if cache:
    entry_login.delete(0, tk.END)
    entry_login.insert(0, cache.get("login", ""))
    entry_senha.delete(0, tk.END)
    entry_senha.insert(0, cache.get("senha", ""))
    entry_data_inicial.delete(0, tk.END)
    entry_data_inicial.insert(0, cache.get("data_inicial", ""))
    entry_data_final.delete(0, tk.END)
    entry_data_final.insert(0, cache.get("data_final", ""))

entry_login.focus_set()
aplicar_tema()

root.mainloop()
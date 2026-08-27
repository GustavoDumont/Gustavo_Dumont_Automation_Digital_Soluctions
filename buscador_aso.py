"""Buscador de ASO em lote.

Automatiza consultas autorizadas em portal ocupacional a partir de uma lista de CPFs.
Credenciais nunca são gravadas. A URL e os seletores devem ser configurados localmente.
"""
from __future__ import annotations

import csv
import os
import re
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from queue import Queue, Empty
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

APP_NAME = "Buscador de ASO"
DEFAULT_URL = os.getenv("ASO_PORTAL_URL", "")
CPF_RE = re.compile(r"\D")

@dataclass
class Result:
    cpf: str
    status: str
    detail: str = ""


def normalize_cpf(value: str) -> str:
    return CPF_RE.sub("", value or "")


def valid_cpf_shape(cpf: str) -> bool:
    return len(cpf) == 11 and cpf != cpf[0] * 11


class AsoAutomation:
    def __init__(self, url: str, username: str, password: str, download_dir: Path, log):
        self.url = url
        self.username = username
        self.password = password
        self.download_dir = download_dir
        self.log = log
        self.driver = None
        self.wait = None

    def start(self):
        options = Options()
        options.add_experimental_option("prefs", {
            "download.default_directory": str(self.download_dir.resolve()),
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
        })
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 30)

    def login_and_open_aso(self):
        self.driver.get(self.url)
        user = self.wait.until(EC.visibility_of_element_located((By.ID, "signInSafeDoc_txtUserName")))
        pwd = self.wait.until(EC.visibility_of_element_located((By.ID, "signInSafeDoc_txtPassword")))
        user.send_keys(self.username)
        pwd.send_keys(self.password)
        pwd.send_keys(Keys.RETURN)
        self.wait.until(EC.url_contains("Search.aspx"))
        self._open_tree()

    def _open_tree(self):
        actions = ActionChains(self.driver)
        for label in ("CONCREMAT", "Administrativo", "Aso"):
            element = self.wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[contains(normalize-space(),'{label}')]")))
            actions.double_click(element).perform()
            time.sleep(0.6)
        selector = self.wait.until(EC.visibility_of_element_located((By.ID, "CP_59")))
        Select(selector).select_by_value("7")

    def download_one(self, cpf: str) -> Result:
        self.driver.get(self.url)
        self.wait.until(EC.url_contains("Search.aspx"))
        try:
            aso = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(normalize-space(),'Aso')]")))
            ActionChains(self.driver).double_click(aso).perform()
            field = self.wait.until(EC.visibility_of_element_located((By.ID, "IDX_59")))
            field.clear(); field.send_keys(cpf)
            self.driver.find_element(By.ID, "_ctl0_ContentPlaceHolder_btnSearch").click()
            time.sleep(1)
            if "Não foram encontrados registros utilizando os critérios acima." in self.driver.page_source:
                return Result(cpf, "NAO_ENCONTRADO")
            buttons = self.driver.find_elements(By.CSS_SELECTOR, "i.fa-download[title='Download']")
            if not buttons:
                links = self.driver.find_elements(By.PARTIAL_LINK_TEXT, "Aso - Nome")
                if not links:
                    return Result(cpf, "SEM_DOCUMENTO", "Resultado sem link de ASO")
                links[0].click()
                button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "i.fa-download[title='Download']")))
            else:
                button = buttons[0]
            button.click()
            return Result(cpf, "DOWNLOAD_SOLICITADO")
        except TimeoutException as exc:
            return Result(cpf, "TIMEOUT", str(exc).splitlines()[0])
        except WebDriverException as exc:
            return Result(cpf, "ERRO_NAVEGADOR", str(exc).splitlines()[0])

    def close(self):
        if self.driver:
            self.driver.quit()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("680x620")
        self.minsize(620, 560)
        self.queue = Queue()
        self.running = False
        self.results = []
        self._build()
        self.after(150, self._drain_queue)

    def _build(self):
        pad = {"padx": 12, "pady": 5}
        tk.Label(self, text="Buscador de ASO em lote", font=("Segoe UI", 18, "bold")).pack(pady=(15, 3))
        tk.Label(self, text="Uso restrito a usuários autorizados. Um CPF por linha.", fg="#555").pack()
        form = tk.Frame(self); form.pack(fill="x", **pad)
        fields = [("URL do portal", "url"), ("Usuário", "user"), ("Senha", "password")]
        self.entries = {}
        for row, (label, key) in enumerate(fields):
            tk.Label(form, text=label).grid(row=row, column=0, sticky="e", padx=5, pady=4)
            entry = tk.Entry(form, width=62, show="*" if key == "password" else "")
            entry.grid(row=row, column=1, sticky="ew", padx=5, pady=4)
            self.entries[key] = entry
        form.columnconfigure(1, weight=1)
        self.entries["url"].insert(0, DEFAULT_URL)
        outrow = tk.Frame(self); outrow.pack(fill="x", **pad)
        tk.Label(outrow, text="Pasta de downloads").pack(side="left")
        self.download = tk.StringVar(value=str((Path.home()/"Downloads"/"ASOs").resolve()))
        tk.Entry(outrow, textvariable=self.download).pack(side="left", fill="x", expand=True, padx=8)
        tk.Button(outrow, text="Selecionar", command=self._select_dir).pack(side="left")
        tk.Label(self, text="CPFs").pack(anchor="w", padx=17)
        self.cpfs = scrolledtext.ScrolledText(self, height=10, font=("Consolas", 10)); self.cpfs.pack(fill="both", expand=True, padx=15, pady=5)
        actions = tk.Frame(self); actions.pack(fill="x", padx=15, pady=5)
        tk.Button(actions, text="Importar TXT/CSV", command=self._import_cpfs).pack(side="left")
        self.start_btn = tk.Button(actions, text="Iniciar processamento", command=self._start, bg="#F18700", fg="white")
        self.start_btn.pack(side="right")
        self.status = tk.StringVar(value="Pronto")
        tk.Label(self, textvariable=self.status, anchor="w").pack(fill="x", padx=15)
        self.log = scrolledtext.ScrolledText(self, height=9, state="disabled", font=("Consolas", 9)); self.log.pack(fill="both", expand=True, padx=15, pady=(5, 15))

    def _select_dir(self):
        value = filedialog.askdirectory()
        if value: self.download.set(value)

    def _import_cpfs(self):
        name = filedialog.askopenfilename(filetypes=[("TXT ou CSV", "*.txt *.csv"), ("Todos", "*.*")])
        if not name: return
        text = Path(name).read_text(encoding="utf-8-sig", errors="replace")
        candidates = re.findall(r"\d[\d.\- ]{9,16}\d", text)
        normalized = [normalize_cpf(x) for x in candidates]
        self.cpfs.delete("1.0", tk.END); self.cpfs.insert("1.0", "\n".join(dict.fromkeys(normalized)))

    def _start(self):
        if self.running: return
        url = self.entries["url"].get().strip()
        username = self.entries["user"].get().strip()
        password = self.entries["password"].get()
        cpfs = list(dict.fromkeys(normalize_cpf(x) for x in self.cpfs.get("1.0", tk.END).splitlines() if normalize_cpf(x)))
        invalid = [x for x in cpfs if not valid_cpf_shape(x)]
        if not url or not username or not password or not cpfs:
            messagebox.showwarning("Dados obrigatórios", "Informe URL, credenciais e ao menos um CPF."); return
        if invalid:
            messagebox.showwarning("CPF inválido", f"Revise {len(invalid)} entrada(s) que não possuem 11 dígitos válidos."); return
        folder = Path(self.download.get()); folder.mkdir(parents=True, exist_ok=True)
        self.running = True; self.results = []; self.start_btn.config(state="disabled")
        threading.Thread(target=self._worker, args=(url, username, password, cpfs, folder), daemon=True).start()

    def _worker(self, url, username, password, cpfs, folder):
        bot = AsoAutomation(url, username, password, folder, self.queue.put)
        try:
            self.queue.put(("status", "Abrindo navegador e autenticando..."))
            bot.start(); bot.login_and_open_aso()
            for i, cpf in enumerate(cpfs, 1):
                self.queue.put(("status", f"Processando {i}/{len(cpfs)}"))
                result = bot.download_one(cpf); self.results.append(result)
                self.queue.put(("log", f"{cpf[:3]}.***.***-{cpf[-2:]} | {result.status} | {result.detail}"))
            report = folder / f"resultado_asos_{datetime.now():%Y%m%d_%H%M%S}.csv"
            with report.open("w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";"); writer.writerow(["cpf", "status", "detalhe"])
                writer.writerows((r.cpf, r.status, r.detail) for r in self.results)
            self.queue.put(("done", str(report)))
        except Exception as exc:
            self.queue.put(("error", f"Processamento interrompido: {exc}"))
        finally:
            bot.close()

    def _drain_queue(self):
        try:
            while True:
                kind, value = self.queue.get_nowait()
                if kind == "status": self.status.set(value)
                elif kind == "log":
                    self.log.config(state="normal"); self.log.insert(tk.END, value + "\n"); self.log.see(tk.END); self.log.config(state="disabled")
                elif kind == "done":
                    self.running = False; self.start_btn.config(state="normal"); self.status.set("Concluído")
                    messagebox.showinfo("Concluído", f"Relatório salvo em:\n{value}")
                elif kind == "error":
                    self.running = False; self.start_btn.config(state="normal"); self.status.set("Erro")
                    messagebox.showerror("Erro", value)
        except Empty: pass
        self.after(150, self._drain_queue)

if __name__ == "__main__":
    App().mainloop()

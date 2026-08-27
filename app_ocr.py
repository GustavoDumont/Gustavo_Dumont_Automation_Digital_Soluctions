"""Versao publica e anonimizada para portfolio.
Nao contem credenciais, URLs corporativas ou dados reais.
Use apenas em ambiente proprio ou formalmente autorizado.
"""

import os
import threading
import re
import sys
import time
import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import undetected_chromedriver as uc
from pdf2image import convert_from_path
import pytesseract
import cv2
import numpy as np
from PIL import Image
from selenium.common.exceptions import (
    StaleElementReferenceException,
    TimeoutException,
    ElementClickInterceptedException,
    NoSuchElementException,
)

# Caminho da pasta base (funciona com PyInstaller --onefile)
BASE_DIR = getattr(sys, '_MEIPASS', os.path.abspath("."))

PORTAL_BASE_URL = os.getenv("APP_BASE_URL", "https://example.invalid").rstrip("/")


# Configura Tesseract
tesseract_path = os.path.join(BASE_DIR, "bin", "tesseract.exe")
pytesseract.pytesseract.tesseract_cmd = tesseract_path
os.environ["TESSDATA_PREFIX"] = os.path.join(BASE_DIR, "bin", "tessdata")

# Configura caminho do Poppler (você usará isso no pdf2image)
POPPLER_PATH = os.path.join(BASE_DIR, "poppler", "Library", "bin")
    

# ─── Configuração da interface ───────────────────────────────────────────────
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ─── Funções utilitárias ────────────────────────────────────────────────────
def sanitize_number(value: str) -> str:
    return re.sub(r'[\.\-/]', '', value or '')

def set_zoom_minimo(driver):
    try:
        driver.execute_script("document.body.style.zoom='25%'")
    except Exception as e:
        print(f"⚠ Erro ao ajustar zoom: {e}")

def extract_text_from_pdf_via_ocr(pdf_path: str, dpi: int = 300, lang: str = 'por') -> str:
    images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH, dpi=dpi)
    texts = []

    for img in images:
        # Converter a imagem PIL para NumPy array (OpenCV usa BGR)
        img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)

        # Aplicar binarização (threshold)
        _, img_thresh = cv2.threshold(img_cv, 150, 255, cv2.THRESH_BINARY)

        # (Opcional) filtro morfológico
        # kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 1))
        # img_thresh = cv2.morphologyEx(img_thresh, cv2.MORPH_CLOSE, kernel)

        # Converter de volta para PIL
        processed_img = Image.fromarray(img_thresh)

        # OCR com Tesseract
        text = pytesseract.image_to_string(processed_img, lang=lang, config="--psm 6 --oem 3")
        texts.append(text)

    return "\n".join(texts)

import subprocess
import re

def obter_versao_chrome():
    try:
        output = subprocess.check_output(
            r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
            shell=True
        ).decode()
        versao = re.search(r"(\d+\.\d+\.\d+\.\d+)", output).group(1)
        major = int(versao.split(".")[0])
        return major
    except Exception:
        return None

def parse_ocr_text(text: str) -> dict:
    data = {}
    data['prestadora_nome'] = 'EMPRESA_PRESTADORA_DEMO'
    data['tomadora_nome']   = 'EMPRESA_TOMADORA_DEMO'

    # Regex genérica para CNPJ no formato xx.xxx.xxx/xxxx-xx
    cnpj_pattern = r"([\d]{2}\.[\d]{3}\.[\d]{3}/[\d]{4}-[\d]{2})"

    # 1) Prestadora (Empresa prestadora demo) – procura "Empresa prestadora demo" em qualquer variação de caixa, limite a ~200 caracteres antes do "CNPJ"
    m1 = re.search(
        rf"(?i)EMPRESA_PRESTADORA_DEMO[\s\S]{{0,200}}CNPJ:\s*{cnpj_pattern}",
        text
    )
    data['prestadora_cnpj'] = sanitize_number(m1.group(1)) if m1 else ''

    # 2) Tomadora (EMPRESA_TOMADORA) – procura "empresa tomadora" em qualquer variação de caixa, limite a ~200 caracteres antes do "CNPJ"
    m2 = re.search(
        rf"(?i)Empresa[\s\S]{{0,200}}CNPJ:\s*{cnpj_pattern}",
        text
    )
    data['tomadora_cnpj'] = sanitize_number(m2.group(1)) if m2 else ''

    # 3) Inscrição municipal da prestadora
    m3 = re.search(r"Ins\.Municipal:\s*([\d\-/]+)", text)
    data['prestadora_inscricao_municipal'] = sanitize_number(m3.group(1)) if m3 else ''

    # 4) Número e data de emissão do documento
    m4 = re.search(r"(\d{6,})\s*/\s*(\d{2}\.\d{2}\.\d{4})", text)
    data['numero_documento'] = m4.group(1) if m4 else ''
    data['data_emissao']     = m4.group(2) if m4 else ''

    # 5) Data de pagamento
    m5 = re.search(r"Pagamento[\s\S]*?Até dia\s*(\d{2}\.\d{2}\.\d{4})", text)
    data['data_pagamento'] = m5.group(1) if m5 else ''

    # 6) Valor total
    m6 = re.search(r"Montante de fatura[\s\S]*?([\d\.]+,[\d]{2})", text)
    data['valor_total'] = m6.group(1) if m6 else ''

    # 7) Contrato, pedido, item, FRS, RF e centro
    m7 = re.search(
        r"CONTRATO:\s*(\d+)[\s\S]*?PEDIDO:\s*(\d+)[\s\S]*?ITEM:\s*(\d+)"
        r"[\s\S]*?FRS:\s*(\d+)[\s\S]*?RF:\s*(\d+)[\s\S]*?CENTRO:\s*(\d+)",
        text
    )
    keys = ['contrato', 'pedido', 'item', 'frs', 'rf', 'centro']
    if m7:
        for i, key in enumerate(keys, start=1):
            data[key] = m7.group(i)
    else:
        for key in keys:
            data[key] = ''

    return data

# ─── Função robusta que tenta parsing até 3 vezes e retorna campos com lista de faltantes ───────────
def parse_ocr_with_retry(pdf_path: str, max_attempts: int = 3, lang: str = 'por') -> dict:
    expected_keys = [
        'prestadora_cnpj', 'tomadora_cnpj', 'prestadora_inscricao_municipal',
        'numero_documento', 'data_emissao', 'data_pagamento',
        'valor_total', 'contrato', 'pedido', 'item', 'frs', 'rf', 'centro'
    ]
    last_data = {}
    for attempt in range(1, max_attempts + 1):
        text = extract_text_from_pdf_via_ocr(pdf_path, lang=lang)
        data = parse_ocr_text(text)
        missing = [key for key in expected_keys if not data.get(key)]
        if not missing:
            data['_missing'] = []
            return data
        last_data = data
    last_data['_missing'] = [key for key in expected_keys if not last_data.get(key)]
    return last_data


# Função auxiliar ── Conversor de data para formato ISO (aaaa-mm-dd)
def converter_para_iso_date(data_str: str) -> str:
    import re
    m = re.match(r"(\d{2})[./](\d{2})[./](\d{4})", data_str)
    if not m:
        return ""
    dia, mes, ano = m.group(1), m.group(2), m.group(3)
    return f"{ano}-{mes}-{dia}"

def clicar_iss_retention(driver, wait):
    elem = wait.until(EC.element_to_be_clickable((By.ID, 'tax_document_iss_retention_0')))
    # Scroll para garantir visibilidade
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
    try:
        elem.click()
    except ElementClickInterceptedException:
        # Tentar clicar via JS se o clique normal falhar
        driver.execute_script("arguments[0].click();", elem)
        

def tryuntil_success(etapa, func, tentativas=10, delay=2):
    for tentativa in range(1, tentativas + 1):
        try:
            func()
            print(f"   ✅ {etapa} concluída com sucesso.")
            return
        except (StaleElementReferenceException, ElementClickInterceptedException) as e:
            print(f"   ⚠️ Tentativa {tentativa}/{tentativas} — {etapa}: {type(e).__name__} — tentando novamente...")
        except (TimeoutException, NoSuchElementException) as e:
            print(f"   ⚠️ Tentativa {tentativa}/{tentativas} — {etapa}: {type(e).__name__} — elemento ainda não disponível...")
        except Exception as e:
            print(f"   ❌ Tentativa {tentativa}/{tentativas} — Erro inesperado em '{etapa}': {type(e).__name__} — {e}")
        time.sleep(delay)
    
    raise RuntimeError(f"❌ Erro crítico: '{etapa}' falhou após {tentativas} tentativas.")


# ─── Aplicação Principal ───────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Upload de Notas de Débito")
        self.geometry("600x400")

        self.select_button = ctk.CTkButton(self, text="Selecionar Notas (PDF)", command=self.select_files)
        self.select_button.pack(pady=20)

        # Variáveis de controle
        self.email_var = tk.StringVar(value=os.getenv("APP_EMAIL", ""))
        self.token_var = tk.StringVar(value=os.getenv("APP_TOKEN", ""))

        # Widgets
        tk.Label(self, text="Email:").pack()
        tk.Entry(self, textvariable=self.email_var).pack()

        tk.Label(self, text="Token (Senha):").pack()
        tk.Entry(self, textvariable=self.token_var, show='*').pack()

        self.start_button = ctk.CTkButton(self, text="Iniciar Processamento", command=self.start_process)
        self.start_button.pack(pady=20)
        self.start_button.configure(state="disabled")

        self.log_text = ctk.CTkTextbox(self, width=750, height=200)
        self.log_text.pack(pady=20)

        self.selected_files = []

    def log(self, message: str):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(ctk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(ctk.END)

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.selected_files = files
            self.log(f"Arquivos selecionados: {files}")
            self.start_button.configure(state="normal")

    def start_process(self):
        threading.Thread(target=self.run_automation, daemon=True).start()
        self.start_button.configure(state="disabled")

        
# ─── helper genérico para Select2 robusto ────────────────────────────────
    def selecionar_select2_robusto(self, driver, container_css, search_css, nome_item):
        try:
            # desativa header interceptador
            driver.execute_script("document.getElementById('header').style.pointerEvents = 'none';")

            # 1) dispara mousedown no container para abrir dropdown
            campo = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, container_css))
            )
            # scroll até o elemento e aguarda estar visível
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
            WebDriverWait(driver, 5).until(EC.visibility_of(campo))
            driver.execute_script(
                "arguments[0].dispatchEvent(new MouseEvent('mousedown',{bubbles:true,cancelable:true}));",
                campo
            )

            # 2) aguarda campo de busca e digita o texto
            busca = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, search_css))
            )
            busca.clear()
            busca.send_keys(nome_item)
            driver.execute_script("arguments[0].dispatchEvent(new Event('input',{ bubbles:true }));", busca)
            time.sleep(1.0)

            # 3) coleta todas as <li> e seleciona a exata
            opcoes = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".select2-results__option"))
            )
            for item in opcoes:
                texto = item.text.strip()
                if nome_item in texto:
                    driver.execute_script("arguments[0].scrollIntoView(true);", item)
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new MouseEvent('mouseover',{bubbles:true}));"
                        "arguments[0].dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));",
                        item
                    )
                    self.log(f"   ✅ '{texto}' (contém '{nome_item}') selecionado com sucesso.")
                    return
            raise Exception(f"Item '{nome_item}' não encontrado nas {len(opcoes)} opções.")

        except Exception as e:
            self.log(f"❌ Erro ao selecionar '{nome_item}': {e}")
            raise

        finally:
            # restaura header
            driver.execute_script("document.getElementById('header').style.pointerEvents = '';")

    def aguardar_modal(self, driver, timeout=30):
        """Espera até que o overlay de carregamento saia da tela."""
        try:
            WebDriverWait(driver, timeout).until(
                EC.invisibility_of_element_located((By.ID, "loadingModal"))
            )
            self.log("   ✔️ Modal de carregamento fechado.")
        except:
            self.log("   ❌ Timeout esperando modal de carregamento desaparecer.")
            raise
        
    def run_automation(self):
        try:
            options = uc.ChromeOptions()
            options.add_argument("--window-size=800,600")  # Largura x Altura

            major_version = obter_versao_chrome()

            if not major_version:
                raise RuntimeError("Não foi possível detectar a versão do Chrome instalada.")

            driver = uc.Chrome(
                options=options,
                version_main=major_version
            )
            wait = WebDriverWait(driver, 60)
            EMAIL = self.email_var.get().strip()
            TOKEN = self.token_var.get().strip()


            # ─── Login ────────────────────────────────────────────────────────
            self.log("1) Acessando portal e fazendo login...")
            driver.get(PORTAL_BASE_URL)
            wait.until(EC.element_to_be_clickable((By.ID, "login_portal"))).click()
            self.log(" Clicou em Login Fornecedor.")
            wait.until(EC.element_to_be_clickable((By.ID, "user_login"))).send_keys(EMAIL)
            self.log(" Preencheu e-mail.")
            driver.find_element(By.ID, "user_password").send_keys(TOKEN)
            self.log(" Preencheu senha.")
            driver.find_element(By.CSS_SELECTOR, "button.submit-button").click()
            wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href,'tax_documents')]")))
            self.log(" Login efetuado com sucesso.")

            # ─── Funções de espera e retry ─────────────────────────────────────
            def aguardar_carregamento():
                # espera spinners e loading divs desaparecerem
                try:
                    WebDriverWait(driver, 30).until_not(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div.loading[role='status']"))
                    )
                except:
                    pass
                try:
                    WebDriverWait(driver, 60).until_not(
                        EC.presence_of_element_located((By.CLASS_NAME, "spinner-border"))
                    )
                except:
                    pass

            def try_until_success(description, func, attempts=5, delay=2):
                for i in range(attempts):
                    try:
                        result = func()
                        self.log(f"   ✔️ {description}")
                        return result
                    except Exception as e:
                        self.log(f"   ❌ Tentativa {i+1}/{attempts} de '{description}' falhou: {e}")
                        aguardar_carregamento()
                        time.sleep(delay)
                raise Exception(f"Falha ao executar: {description}")

            def preencher_campo_robusto(field_id, value, max_tentativas_campo, atraso_campo):
                sucesso = False
                for tentativa in range(1, max_tentativas_campo + 1):
                    try:
                        el = wait.until(EC.element_to_be_clickable((By.ID, field_id)))
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                        time.sleep(0.2)
                        driver.execute_script("arguments[0].removeAttribute('data-reload-form');", el)
                        driver.execute_script("""
                            var el = arguments[0];
                            el.value = arguments[1];
                            el.dispatchEvent(new Event('input',{ bubbles:true }));
                            el.dispatchEvent(new Event('change',{ bubbles:true }));
                        """, el, value)
                        sucesso = True
                        break

                    except StaleElementReferenceException:
                        time.sleep(0.5)

                    except Exception as e:
                        if tentativa < max_tentativas_campo:
                            time.sleep(atraso_campo)
                        else:
                            driver.execute_script("window.open('');")
                            driver.switch_to.window(driver.window_handles[-1])
                            driver.get(driver.current_url)
                            time.sleep(1) 

            # função auxiliar para recarregar a página 
            def execute_or_skip(driver, descrição, func, neutral_selector="body"):
                try:
                    func()
                    self.log(f"✅ {descrição} concluída com sucesso.")
                except Exception as e:
                    self.log(f"⚠️ Falha em '{descrição}': {type(e).__name__} — pulando. Tentando clique neutro...")
                    try:
                        neutral = driver.find_element(By.CSS_SELECTOR, neutral_selector)
                        driver.execute_script("arguments[0].click();", neutral)
                        self.log("   ⚙️ Clique neutro realizado.")
                    except Exception:
                        self.log("   ⚠️ Clique neutro falhou; ignorando.")

            # Função para checar e re-preencher campos faltantes após o preenchimento
            def check_and_refill_fields(driver, data):
                # Mapeamento: {ID do campo HTML : chave no dicionário de dados}
                field_map = {
                    'tax_document_issue_date': 'data_emissao',
                    'tax_document_net_due_date': 'data_pagamento',
                    'tax_document_supplier_identification_number': 'prestadora_cnpj',
                    'tax_document_customer_identification_number': 'tomadora_cnpj',
                    # Adicione outros campos conforme necessário
                }

                missing = []
                for field_id, key in field_map.items():
                    try:
                        el = driver.find_element(By.ID, field_id)
                        if not el.get_attribute('value'):
                            missing.append((field_id, key))
                    except:
                        missing.append((field_id, key))

                # Repreenche os campos vazios
                for field_id, key in missing:
                    valor = data.get(key, '')
                    if valor:
                        driver.execute_script(f"""
                            var el = document.getElementById('{field_id}');
                            if (el) {{
                                el.value = '{valor}';
                                el.setAttribute('data-date', '{valor.replace(".", "/")}' );
                                el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                            }}
                        """)

                return missing


            # ─── Loop de processamento de cada PDF ────────────────────────────
            for pdf in self.selected_files:
                self.log(f"--- Iniciando processamento: {os.path.basename(pdf)} ---")

                # 2) OCR e parsing robusto com retry
                self.log("2) OCR e parsing robusto via retry...")
                data = parse_ocr_with_retry(pdf, max_attempts=3, lang='por')
                missing = data.get('_missing', [])
                if missing:
                    self.log(f"   ⚠️ Itens não encontrados ({os.path.basename(pdf)}): {', '.join(missing)}")
                else:
                    self.log(f"   ▶️ Todos os campos encontrados com sucesso em {os.path.basename(pdf)}")

                self.log(f"   ▶️ Parsed prestadora_cnpj: {data['prestadora_cnpj']}")
                self.log(f"   ▶️ Parsed tomadora_cnpj:   {data['tomadora_cnpj']}")


                # 1) Abrir nova aba e navegar
                self.log("3) Abrindo nova aba e navegando para página de nota...")
                driver.switch_to.new_window('tab')
                driver.switch_to.window(driver.window_handles[-1])
                def wait_page_ready(driver, timeout=30):
                    WebDriverWait(driver, timeout).until(
                        lambda d: d.execute_script("return document.readyState") == "complete"
                    )

                def wait_upload_complete(driver, timeout=60):
                    wait = WebDriverWait(driver, timeout)
                    try:
                        wait.until_not(EC.presence_of_element_located((By.CSS_SELECTOR, "div.loading[role='status']")))
                    except:
                        pass
                    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "i.fa.fa-pencil.btn")))

                # Navegar até a página de upload
                driver.get(PORTAL_BASE_URL + "/nf/tax_documents/other_invoice/new")
                self.log("3) Aguardando carregamento completo da página...")
                wait_page_ready(driver, 30)
                time.sleep(2)
                self.log(" Página carregada e estável.")

                # Upload do PDF
                self.log("4) Enviando PDF...")
                up = try_until_success("Localizar input de upload",
                    lambda: wait.until(EC.presence_of_element_located((By.ID, 'tax_document_document_pdf'))))
                driver.execute_script("arguments[0].scrollIntoView(true);", up)
                up.send_keys(pdf)
                self.log(" PDF enviado, aguardando confirmação visual do upload...")
                # Upload do PDF
                self.log("4) Enviando PDF...")
                up = try_until_success(
                    "Localizar input de upload",
                    lambda: wait.until(EC.presence_of_element_located((By.ID, 'tax_document_document_pdf')))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", up)
                up.send_keys(pdf)

                self.log("   PDF enviado, aguardando confirmação pelo link e tamanho exibido...")

                def check_file_size_displayed():
                    # localiza o container que envolve o link e o texto de tamanho
                    span = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "span[data-input-file]")))
                    text = span.text.strip()
                    if "Mb" in text or "MB" in text:
                        return span
                    raise Exception(f"Tamanho não encontrado ainda no texto: '{text}'")

                # tenta até 5 vezes, 2s de intervalo
                file_span = try_until_success(
                    "Aguardar link + tamanho do arquivo",
                    check_file_size_displayed,
                    attempts=20,
                    delay=1
                )

                self.log(f"   Upload confirmado — arquivo listado como: {file_span.text}")

                # ─── 19) Seleção de modelo Nota de Débito com tentativas ─────────────────────────────
                self.log("19) Selecionando modelo Nota de Débito…")

                max_tentativas_modelo = 5
                atraso_modelo = 1.0

                for tentativa in range(1, max_tentativas_modelo + 1):
                    try:
                        self.log(f"   🔄 Tentativa {tentativa}/{max_tentativas_modelo} para selecionar o modelo Nota de Débito")

                        # a) Scroll até o span para ficar visível
                        span_model = wait.until(EC.presence_of_element_located((
                            By.CSS_SELECTOR,
                            "span.select2-selection--single[aria-labelledby='select2-tax_document_model-container']"
                        )))
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", span_model)
                        time.sleep(0.3)

                        # b) Abrir e selecionar via helper robusto
                        self.selecionar_select2_robusto(
                            driver,
                            container_css="span.select2-selection--single[aria-labelledby='select2-tax_document_model-container']",
                            search_css="input.select2-search__field",
                            nome_item="Nota de Débito"
                        )

                        # c) Forçar o <select> oculto
                        options = driver.find_elements(By.CSS_SELECTOR, "#tax_document_model option")
                        value = next(
                            (o.get_attribute("value") for o in options if o.text.strip() == "Nota de Débito"),
                            None
                        )
                        if not value:
                            raise Exception("Não achou value para Nota de Débito no <select> oculto")

                        driver.execute_script("""
                            var sel = document.getElementById('tax_document_model');
                            sel.value = arguments[0];
                            sel.dispatchEvent(new Event('change',{bubbles:true}));
                        """, value)

                        self.aguardar_modal(driver)
                        self.log(f"   ✅ <select> oculto atualizado para value='{value}'")
                        break  # sucesso, sai do loop

                    except Exception as e:
                        self.log(f"   ❌ Erro na tentativa {tentativa} ao selecionar modelo: {e}")
                        if tentativa < max_tentativas_modelo:
                            time.sleep(atraso_modelo)
                        else:
                            raise Exception(
                                f"Não foi possível selecionar o modelo Nota de Débito após {max_tentativas_modelo} tentativas."
                            )



                # 6) ISS Retention
                execute_or_skip(driver, "Selecionar ISS retention", lambda: clicar_iss_retention(driver, wait))

                # 7) Prestadora – preencher CNPJ e disparar reload via ISS Retention
                self.log("7) Preenchendo CNPJ da prestadora...")

                for tentativa in range(1, 6):
                    try:
                        self.log(f"   🔄 Tentativa {tentativa}/5 para preencher CNPJ da prestadora")
                        # 1) preenche o CNPJ
                        cnpj_field = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, 'tax_document_supplier_identification_number'))
                        )
                        cnpj_field.clear()
                        cnpj_field.send_keys(data['prestadora_cnpj'])

                        # 2) dispara reload via ISS retention
                        self.log("   Disparando atualização automática clicando no ISS Retention…")
                        clicar_iss_retention(driver, wait)

                        # 3) aguarda o modal e, em seguida, dá um clique neutro no <body>
                        self.aguardar_modal(driver)
                        try:
                            driver.find_element(By.TAG_NAME, 'body').click()
                            self.log("   ⚙️ Clique neutro realizado para liberar legal_name.")
                        except Exception:
                            self.log("   ⚠️ Clique neutro falhou; mas continuando igual.")

                        # 4) aguarda legal_name
                        self.log("   Aguardando legal_name ser populado automaticamente...")
                        WebDriverWait(driver, 15, poll_frequency=1).until(
                            lambda d: d.find_element(By.ID, 'tax_document_supplier_legal_name')
                                         .get_attribute('value').strip() != ''
                        )
                        self.log("   ✅ legal_name preenchido — atualização concluída.")
                        break  # sucesso, sai do for

                    except Exception as e:
                        self.log(f"   ❌ Erro ao processar prestadora (tentativa {tentativa}): {e}")
                        if tentativa == 5:
                            raise Exception("Não foi possível preencher a prestadora após 5 tentativas.")
                        time.sleep(1)

                # 8) Inscrição Municipal — aguardar até estar clicável e só então preencher
                self.log("8) Aguardando campo Inscrição Municipal estar clicável...")
                try:
                    im_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, 'tax_document_supplier_municipal_registration')))
                    im_field.clear()
                    im_field.send_keys(data['prestadora_inscricao_municipal'])
                    self.log("   ✅ Inscrição Municipal preenchida.")
                except Exception as e:
                    raise Exception(f"Erro ao preencher Inscrição Municipal: {e}")

                # 9) Seleção de cidade RJ – robusta com busca, MouseEvents e tentativas
                self.log("9) Selecionando cidade RJ - Rio de Janeiro via JS de MouseEvents com busca…")

                def selecionar_cidade_robusta(driver, nome_cidade="RJ - Rio de Janeiro", tentativas=7):
                    for tentativa in range(1, tentativas + 1):
                        try:
                            self.log(f"   🔄 Tentativa {tentativa}/{tentativas} para selecionar cidade...")
                            driver.execute_script("document.getElementById('header').style.pointerEvents = 'none';")

                            campo = WebDriverWait(driver, 15).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "span[aria-labelledby='select2-tax_document_supplier_city_id-container']"))
                            )
                            driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('mousedown',{bubbles:true,cancelable:true}));", campo)

                            busca = WebDriverWait(driver, 10).until(
                                EC.visibility_of_element_located((By.CSS_SELECTOR, "input.select2-search__field"))
                            )
                            busca.clear()
                            busca.send_keys(nome_cidade)
                            driver.execute_script("arguments[0].dispatchEvent(new Event('input',{ bubbles:true }));", busca)
                            time.sleep(1.0)

                            opcoes = WebDriverWait(driver, 10).until(
                                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".select2-results__option"))
                            )
                            self.log(f"   ⚙️ {len(opcoes)} opções carregadas pós-busca")

                            for item in opcoes:
                                if item.text.strip() == nome_cidade:
                                    driver.execute_script("arguments[0].scrollIntoView(true);", item)
                                    driver.execute_script(
                                        "arguments[0].dispatchEvent(new MouseEvent('mouseover',{bubbles:true}));"
                                        "arguments[0].dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));",
                                        item
                                    )
                                    self.log(f"✅ Cidade '{nome_cidade}' selecionada com sucesso.")
                                    return

                            raise Exception(f"Cidade '{nome_cidade}' não encontrada entre as opções.")
                        except Exception as e:
                            self.log(f"   ❌ Erro ao selecionar cidade (tentativa {tentativa}): {e}")
                            if tentativa == tentativas:
                                raise Exception("Falha ao selecionar cidade após múltiplas tentativas.")
                            time.sleep(1)
                        finally:
                            driver.execute_script("document.getElementById('header').style.pointerEvents = '';")

                selecionar_cidade_robusta(driver)




                # ─── 10) Tomadora – combobox customer_identification_number com tentativas ───────────────
                self.log("10) Preenchendo CNPJ da tomadora…")

                max_tentativas_tomadora = 7
                atraso_tomadora = 1.5

                for tentativa in range(1, max_tentativas_tomadora + 1):
                    try:
                        self.log(f"   🔄 Tentativa {tentativa}/{max_tentativas_tomadora} para selecionar a tomadora")
                        self.selecionar_select2_robusto(
                            driver,
                            container_css="span.select2-selection--single[aria-labelledby='select2-tax_document_customer_identification_number-container']",
                            search_css="input.select2-search__field",
                            nome_item=data['tomadora_cnpj']
                        )
                        self.log("   ⏳ Aguardando legal name da tomadora ser populado...")
                        self.aguardar_modal(driver)
                        WebDriverWait(driver, 30).until(
                            lambda d: d.find_element(By.ID, "tax_document_customer_legal_name").get_attribute("value").strip() != ""
                        )
                        legal_name = driver.find_element(By.ID, "tax_document_customer_legal_name").get_attribute("value")
                        self.log(f"   ✔️ Campo customer_legal_name preenchido com: {legal_name}")
                        break  # sucesso, sai do loop
                    except Exception as e:
                        self.log(f"   ❌ Erro ao preencher tomadora na tentativa {tentativa}: {e}")
                        if tentativa < max_tentativas_tomadora:
                            time.sleep(atraso_tomadora)


                # ─── 11–17) Campos numéricos e de texto (com otimização para RF e contrato) ───
                self.log("11–17) Preenchendo campos numéricos e textuais...")

                field_map = [
                    ('tax_document_number',                            data['numero_documento']),
                    ('tax_document_total_value',                       data['valor_total']),
                    ('tax_document_invoice_items_attributes_0_purchase_order', data['pedido']),
                    ('tax_document_invoice_items_attributes_0_line_number',   data['item']),
                    ('tax_document_invoice_items_attributes_0_frs',           data['frs']),
                    ('tax_document_invoice_items_attributes_0_billing_report_code', data['rf']),
                    ('tax_document_invoice_items_attributes_0_contract_number',       data['contrato']),
                ]

                max_tentativas_campo = 6
                atraso_campo = 1.0

                for field_id, value in field_map:
                    label = field_id.replace("tax_document_", "")
                    if not value:
                        self.log(f"   ⚠️ Campo {label}: sem valor detectado no OCR, pulando preenchimento.")
                        continue

                    for tentativa in range(1, max_tentativas_campo + 1):
                        try:
                            # localiza de forma mais robusta: por ID ou name
                            try:
                                el = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.ID, field_id))
                                )
                            except Exception:
                                el = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, f"input[name*='{field_id.split('_',1)[-1]}']"))
                                )

                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                            driver.execute_script("arguments[0].removeAttribute('data-reload-form');", el)
                            driver.execute_script("""
                                var el = arguments[0];
                                el.value = arguments[1];
                                el.dispatchEvent(new Event('input', { bubbles:true }));
                                el.dispatchEvent(new Event('change', { bubbles:true }));
                            """, el, value)

                            self.log(f"   ✅ {label} preenchido com '{value}'")
                            break
                        except Exception as e:
                            self.log(f"   ❌ Tentativa {tentativa}/{max_tentativas_campo} falhou para {label}: {e}")
                            time.sleep(atraso_campo)
                            if tentativa == max_tentativas_campo:
                                self.log(f"   ❌ Erro persistente: campo {label} não pôde ser preenchido.")



                # ─── 18) Selecionar método de pagamento (fix final para empresa tomadora) ───
                self.log("18) Selecionando método de pagamento Crédito em Conta…")

                try:
                    select_id = "tax_document_cf_payment_method"
                    opcao_texto = "Crédito em Conta"
                    opcao_valor = "D"  # valor real usado pelo backend da empresa tomadora

                    # Espera o campo aparecer no DOM e ficar estável
                    select_el = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.ID, select_id))
                    )

                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_el)
                    driver.execute_script("arguments[0].removeAttribute('disabled');", select_el)

                    # Define explicitamente o value correto
                    driver.execute_script("""
                        const sel = arguments[0];
                        const value = arguments[1];
                        const texto = arguments[2];
                        const opt = Array.from(sel.options).find(o => o.value === value || o.text.includes(texto));
                        if (opt) {
                            sel.value = opt.value;
                            sel.dispatchEvent(new Event('input', { bubbles:true }));
                            sel.dispatchEvent(new Event('change', { bubbles:true }));
                        }
                    """, select_el, opcao_valor, opcao_texto)

                    time.sleep(1.0)
                    value_final = driver.execute_script("return arguments[0].value;", select_el)
                    text_final = driver.execute_script("return arguments[0].selectedOptions[0].text;", select_el)
                    self.log(f"   ✅ Método de pagamento definido -> value='{value_final}' texto='{text_final}'")

                    if value_final != opcao_valor and opcao_texto not in text_final:
                        raise Exception(f"Campo não persistiu corretamente (value='{value_final}', texto='{text_final}')")

                except Exception as e:
                    self.log(f"   ❌ Falha ao setar método de pagamento: {e}")
                    # fallback absoluto: reexecuta JS direto no DOM
                    driver.execute_script("""
                        const sel = document.getElementById('tax_document_cf_payment_method');
                        if (sel) {
                            const opt = Array.from(sel.options).find(o => o.text.includes('Crédito em Conta'));
                            if (opt) {
                                sel.value = opt.value;
                                sel.dispatchEvent(new Event('change', { bubbles:true }));
                            }
                        }
                    """)
                    self.log("   ⚙️ Fallback final aplicado: forçado 'Crédito em Conta' via JS.")


                iso_emissao = converter_para_iso_date(data['data_emissao'])
                iso_pagamento = converter_para_iso_date(data['data_pagamento'])

                # Converte para formato BR (dd/mm/aaaa) caso esteja no formato dd.mm.aaaa
                data_emissao_br = data['data_emissao'].replace(".", "/")
                data_pagamento_br = data['data_pagamento'].replace(".", "/")

                # ─── 19) Preencher datas no formato ISO + data-date em atributo ───────────────────────────────────────
                try:
                    self.log("19a) Preenchendo data de emissão (formato ISO + data-date)…")
                    issue_field = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.ID, "tax_document_issue_date"))
                    )
                    if not iso_emissao:
                        raise ValueError(f"Data de emissão inválida: '{data['data_emissao']}'")
                    driver.execute_script(f"""
                        var el = document.getElementById('tax_document_issue_date');
                        el.value = '{iso_emissao}';
                        el.setAttribute('data-date', '{data_emissao_br}');
                        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    """)
                    self.log(f"   ✅ Campo issue_date setado para (ISO): {iso_emissao} e data-date: {data_emissao_br}")
                except Exception as e:
                    self.log(f"   ❌ Erro ao preencher issue_date: {e}")
                    raise

                try:
                    self.log("19b) Preenchendo data de pagamento (formato ISO + data-date)…")
                    due_field = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.ID, "tax_document_net_due_date"))
                    )
                    if not iso_pagamento:
                        raise ValueError(f"Data de pagamento inválida: '{data['data_pagamento']}'")
                    driver.execute_script(f"""
                        var el = document.getElementById('tax_document_net_due_date');
                        el.value = '{iso_pagamento}';
                        el.setAttribute('data-date', '{data_pagamento_br}');
                        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    """)
                    self.log(f"   ✅ Campo net_due_date setado para (ISO): {iso_pagamento} e data-date: {data_pagamento_br}")
                except Exception as e:
                    self.log(f"   ❌ Erro ao preencher net_due_date: {e}")
                    raise
                
                # Verifica se campos críticos ficaram vazios e tenta repreencher
                campos_corrigidos = check_and_refill_fields(driver, data)
                if campos_corrigidos:
                    self.log(f"   ⚠️ Repreenchidos automaticamente: {[campo for campo, _ in campos_corrigidos]}")
    
            self.log("=== Processamento concluído para todos os arquivos ===")

        except Exception as e:
            self.log(f"❌ Erro crítico: {e}")

        finally:
            self.start_button.configure(state="normal")

    def select_model_nota_debito(self, driver, wait):
        # mesma lógica de buscas e cliques detalhada antes
        aguardar = WebDriverWait(driver, 20)
        for tentativa in range(3):
            try:
                self.log(f"   Tentativa {tentativa+1}/5 para selecionar Nota de Débito...")
                aguardar_carregamento = lambda: WebDriverWait(driver, 20).until_not(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.loading[role='status']"))
                )
                aguardar_carregamento()

                combo = aguardar.until(EC.element_to_be_clickable((By.XPATH,
                    "//span[@aria-labelledby='select2-tax_document_model-container']")))
                driver.execute_script("arguments[0].click();", combo)

                search = aguardar.until(EC.element_to_be_clickable((By.XPATH,
                    "//input[contains(@class,'select2-search__field')]")))
                search.clear(); search.send_keys("Nota de Débito")
                time.sleep(1)

                opt = aguardar.until(EC.element_to_be_clickable((By.XPATH,
                    "//li[contains(@class,'select2-results__option--highlighted') and text()='Nota de Débito']")))
                driver.execute_script("arguments[0].click();", opt)

                self.log("   ✔️ Nota de Débito selecionada.")
                return
            except Exception as err:
                self.log(f"   ❌ Falha na seleção Nota de Débito: {err}")
                time.sleep(2)
        raise Exception("Falha ao selecionar 'Nota de Débito' após 5 tentativas.")

if __name__ == '__main__':
    app = App()
    app.mainloop()

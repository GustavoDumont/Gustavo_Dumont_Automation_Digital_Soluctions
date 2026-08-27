"""Versao publica e anonimizada para portfolio.
Nao contem credenciais, URLs corporativas ou dados reais.
Use apenas em ambiente proprio ou formalmente autorizado.
"""

import os
import threading
import re
import sys
import time
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
from selenium.common.exceptions import (
    StaleElementReferenceException,
    TimeoutException,
    ElementClickInterceptedException,
    NoSuchElementException,
)

# Caminho da pasta base (funciona com PyInstaller --onefile)
BASE_DIR = getattr(sys, '_MEIPASS', os.path.abspath("."))

PORTAL_BASE_URL = os.getenv("APP_BASE_URL", "https://example.invalid").rstrip("/")


# ─── Configuração da interface ───────────────────────────────────────────────
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ─── Funções utilitárias e auxiliares ────────────────────────────────────────
def sanitize_number(value) -> str:
    if pd.isna(value):
        return ''
    val_str = str(value).split('.')[0] if isinstance(value, float) else str(value)
    return re.sub(r'[\.\-/]', '', val_str).strip()

def format_currency(value) -> str:
    if pd.isna(value):
        return ''
    if isinstance(value, (int, float)):
        return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(value).strip()

def set_zoom_minimo(driver):
    try:
        driver.execute_script("document.body.style.zoom='25%'")
    except Exception as e:
        print(f"⚠ Erro ao ajustar zoom: {e}")

def obter_versao_chrome():
    import subprocess
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

def converter_para_iso_date(data_str: str) -> str:
    if not data_str or pd.isna(data_str):
        return ""
    if re.match(r"\d{4}-\d{2}-\d{2}", str(data_str)):
        return str(data_str)
    m = re.match(r"(\d{2})[./](\d{2})[./](\d{4})", str(data_str))
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return str(data_str)

def converter_para_br_date(data_str: str) -> str:
    if not data_str or pd.isna(data_str):
        return ""
    if re.match(r"(\d{4})-(\d{2})-(\d{2})", str(data_str)):
        ano, mes, dia = str(data_str).split('-')
        return f"{dia}/{mes}/{ano}"
    return str(data_str).replace(".", "/")

def extrair_numero_do_nome_arquivo(nome_arquivo: str) -> str:
    numeros = re.findall(r'\d+', nome_arquivo)
    return "".join(numeros) if numeros else ""

def clicar_iss_retention(driver, wait):
    elem = wait.until(EC.element_to_be_clickable((By.ID, 'tax_document_iss_retention_0')))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
    try:
        elem.click()
    except ElementClickInterceptedException:
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
        self.title("Upload de Notas de Débito (Via Planilha + Fallbacks Otimizados)")
        self.geometry("750x550")

        self.excel_button = ctk.CTkButton(self, text="1. Selecionar Planilha (Excel/CSV)", command=self.select_excel)
        self.excel_button.pack(pady=10)

        self.select_button = ctk.CTkButton(self, text="2. Selecionar Notas (PDF)", command=self.select_files)
        self.select_button.pack(pady=10)
        self.select_button.configure(state="disabled")

        self.email_var = tk.StringVar(value=os.getenv("APP_EMAIL", ""))
        self.token_var = tk.StringVar(value=os.getenv("APP_TOKEN", ""))

        tk.Label(self, text="Email:").pack()
        tk.Entry(self, textvariable=self.email_var, width=40).pack()

        tk.Label(self, text="Token (Senha):").pack()
        tk.Entry(self, textvariable=self.token_var, show='*', width=40).pack()

        self.start_button = ctk.CTkButton(self, text="Iniciar Processamento", command=self.start_process)
        self.start_button.pack(pady=15)
        self.start_button.configure(state="disabled")

        self.log_text = ctk.CTkTextbox(self, width=700, height=200)
        self.log_text.pack(pady=10)

        self.selected_files = []
        self.df_dados = None

    def log(self, message: str):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(ctk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(ctk.END)

    def select_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos de Planilha", "*.xlsx *.xls *.csv")])
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    try:
                        self.df_dados = pd.read_csv(file_path, sep=',', skiprows=1)
                        if 'Número' not in self.df_dados.columns:
                            raise ValueError
                    except:
                        self.df_dados = pd.read_csv(file_path, sep=';', skiprows=1)
                else:
                    self.df_dados = pd.read_excel(file_path, skiprows=1)
                
                self.df_dados.columns = self.df_dados.columns.str.strip()
                self.log(f"✅ Planilha carregada! Colunas: {list(self.df_dados.columns)}")
                self.select_button.configure(state="normal")
            except Exception as e:
                self.log(f"❌ Erro ao ler a planilha: {e}.")

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.selected_files = files
            self.log(f"📂 {len(files)} arquivos PDF selecionados.")
            self.start_button.configure(state="normal")

    def start_process(self):
        if self.df_dados is None or not self.selected_files:
            self.log("❌ Erro: Selecione a planilha e os PDFs antes de iniciar.")
            return
        threading.Thread(target=self.run_automation, daemon=True).start()
        self.start_button.configure(state="disabled")

    def selecionar_select2_robusto(self, driver, container_css, search_css, nome_item):
        try:
            driver.execute_script("document.getElementById('header').style.pointerEvents = 'none';")

            campo = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, container_css)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
            WebDriverWait(driver, 5).until(EC.visibility_of(campo))
            driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('mousedown',{bubbles:true,cancelable:true}));", campo)

            busca = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, search_css)))
            busca.clear()
            busca.send_keys(nome_item)
            driver.execute_script("arguments[0].dispatchEvent(new Event('input',{ bubbles:true }));", busca)
            time.sleep(1.0)

            opcoes = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".select2-results__option")))
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
            driver.execute_script("document.getElementById('header').style.pointerEvents = '';")

    def aguardar_modal(self, driver, timeout=30):
        try:
            WebDriverWait(driver, timeout).until(EC.invisibility_of_element_located((By.ID, "loadingModal")))
            self.log("   ✔️ Modal de carregamento fechado.")
        except:
            self.log("   ❌ Timeout esperando modal de carregamento desaparecer.")
            raise

    def buscar_dados_planilha(self, numero_nd) -> dict:
        self.df_dados['Num_Busca'] = self.df_dados['Número'].apply(lambda x: sanitize_number(x))
        row = self.df_dados[self.df_dados['Num_Busca'] == sanitize_number(numero_nd)]
        if row.empty:
            return None
        
        item_dados = row.iloc[0]
        return {
            'prestadora_cnpj': sanitize_number(item_dados.get('CNPJ do Fornecedor')),
            'prestadora_inscricao_municipal': sanitize_number(item_dados.get('I M')),
            'cidade': str(item_dados.get('CIDADE', 'RJ - Rio de Janeiro')).strip(),
            'tomadora_cnpj': sanitize_number(item_dados.get('CNPJ do Tomador')),
            'numero_documento': sanitize_number(item_dados.get('Número')),
            'data_emissao': str(item_dados.get('Data de Emissão')),
            'data_pagamento': str(item_dados.get('Data de Vencimento')),
            'valor_total': format_currency(item_dados.get('Valor Total')),
            'pedido': sanitize_number(item_dados.get('Pedido de Compra')),
            'item': sanitize_number(item_dados.get('Item')),
            'frs': sanitize_number(item_dados.get('FRS')),
            'rf': sanitize_number(item_dados.get('RF')),
            'contrato': sanitize_number(item_dados.get('CONTRATO'))
        }

    def run_automation(self):
        driver = None
        try:
            options = uc.ChromeOptions()
            options.add_argument("--window-size=1024,768")
            major_version = obter_versao_chrome()

            if not major_version:
                raise RuntimeError("Não foi possível detectar a versão do Chrome.")

            driver = uc.Chrome(options=options, version_main=major_version)
            wait = WebDriverWait(driver, 60)
            EMAIL = self.email_var.get().strip()
            TOKEN = self.token_var.get().strip()

            self.log("1) Acessando portal e fazendo login...")
            driver.get(PORTAL_BASE_URL)
            wait.until(EC.element_to_be_clickable((By.ID, "login_portal"))).click()
            wait.until(EC.element_to_be_clickable((By.ID, "user_login"))).send_keys(EMAIL)
            driver.find_element(By.ID, "user_password").send_keys(TOKEN)
            driver.find_element(By.CSS_SELECTOR, "button.submit-button").click()
            wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href,'tax_documents')]")))
            self.log(" Login efetuado com sucesso.")

            # ─── Funções Internas de Espera e Retry ──────────────────────────────────
            def aguardar_carregamento():
                try:
                    WebDriverWait(driver, 30).until_not(EC.presence_of_element_located((By.CSS_SELECTOR, "div.loading[role='status']")))
                except: pass
                try:
                    WebDriverWait(driver, 60).until_not(EC.presence_of_element_located((By.CLASS_NAME, "spinner-border")))
                except: pass

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
                for tentativa in range(1, max_tentativas_campo + 1):
                    try:
                        try:
                            el = wait.until(EC.element_to_be_clickable((By.ID, field_id)))
                        except Exception:
                            el = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"input[name*='{field_id.split('_',1)[-1]}']")))
                            
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                        time.sleep(0.2)
                        driver.execute_script("arguments[0].removeAttribute('data-reload-form');", el)
                        driver.execute_script("""
                            var el = arguments[0];
                            el.value = arguments[1];
                            el.dispatchEvent(new Event('input',{ bubbles:true }));
                            el.dispatchEvent(new Event('change',{ bubbles:true }));
                        """, el, value)
                        return True
                    except StaleElementReferenceException:
                        time.sleep(0.5)
                    except Exception:
                        if tentativa < max_tentativas_campo:
                            time.sleep(atraso_campo)
                        else:
                            driver.execute_script("window.open('');")
                            driver.switch_to.window(driver.window_handles[-1])
                            driver.get(driver.current_url)
                            time.sleep(1)
                return False

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

            def check_and_refill_fields(driver, data):
                field_map = {
                    'tax_document_issue_date': 'data_emissao',
                    'tax_document_net_due_date': 'data_pagamento',
                    'tax_document_supplier_identification_number': 'prestadora_cnpj',
                    'tax_document_customer_identification_number': 'tomadora_cnpj'
                }
                missing = []
                for field_id, key in field_map.items():
                    try:
                        el = driver.find_element(By.ID, field_id)
                        if not el.get_attribute('value'):
                            missing.append((field_id, key))
                    except:
                        missing.append((field_id, key))

                for field_id, key in missing:
                    valor = data.get(key, '')
                    if valor:
                        if "date" in field_id:
                            val_iso = converter_para_iso_date(valor)
                            val_br = converter_para_br_date(valor)
                            driver.execute_script(f"""
                                var el = document.getElementById('{field_id}');
                                if (el) {{
                                    el.value = '{val_iso}';
                                    el.setAttribute('data-date', '{val_br}');
                                    el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                    el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                                }}
                            """)
                        else:
                            driver.execute_script(f"""
                                var el = document.getElementById('{field_id}');
                                if (el) {{
                                    el.value = '{valor}';
                                    el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                    el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                                }}
                            """)
                return missing

            # ─── Processamento dos Arquivos ──────────────────────────────────────────
            for pdf_path in self.selected_files:
                nome_arquivo = os.path.basename(pdf_path)
                numero_nd = extrair_numero_do_nome_arquivo(nome_arquivo)
                
                self.log(f"--- Processando: {nome_arquivo} (ND: {numero_nd}) ---")
                
                data = self.buscar_dados_planilha(numero_nd)
                if not data:
                    self.log(f"❌ ND '{numero_nd}' não encontrada na planilha. Pulando...")
                    continue
                
                self.log(f"   ▶️ Prestadora: {data['prestadora_cnpj']} | Tomadora: {data['tomadora_cnpj']}")

                driver.switch_to.new_window('tab')
                driver.switch_to.window(driver.window_handles[-1])
                
                def wait_page_ready(driver, timeout=30):
                    WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")

                driver.get(PORTAL_BASE_URL + "/nf/tax_documents/other_invoice/new")
                self.log("3) Aguardando carregamento completo da página...")
                wait_page_ready(driver, 30)
                time.sleep(2)

                self.log("4) Enviando PDF...")
                up = try_until_success("Localizar input de upload", lambda: wait.until(EC.presence_of_element_located((By.ID, 'tax_document_document_pdf'))))
                driver.execute_script("arguments[0].scrollIntoView(true);", up)
                up.send_keys(pdf_path)

                def check_file_size_displayed():
                    span = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "span[data-input-file]")))
                    text = span.text.strip()
                    if "Mb" in text or "MB" in text:
                        return span
                    raise Exception(f"Tamanho não encontrado ainda no texto: '{text}'")

                file_span = try_until_success("Aguardar link + tamanho do arquivo", check_file_size_displayed, attempts=20, delay=1)
                self.log(f"   Upload confirmado: {file_span.text}")

                self.log("19) Selecionando modelo Nota de Débito…")
                max_tentativas_modelo = 5
                for tentativa in range(1, max_tentativas_modelo + 1):
                    try:
                        span_model = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.select2-selection--single[aria-labelledby='select2-tax_document_model-container']")))
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", span_model)
                        time.sleep(0.3)

                        self.selecionar_select2_robusto(driver, container_css="span.select2-selection--single[aria-labelledby='select2-tax_document_model-container']", search_css="input.select2-search__field", nome_item="Nota de Débito")

                        options_mod = driver.find_elements(By.CSS_SELECTOR, "#tax_document_model option")
                        value = next((o.get_attribute("value") for o in options_mod if o.text.strip() == "Nota de Débito"), None)
                        if value:
                            driver.execute_script("var sel = document.getElementById('tax_document_model'); sel.value = arguments[0]; sel.dispatchEvent(new Event('change',{bubbles:true}));", value)

                        self.aguardar_modal(driver)
                        break
                    except Exception as e:
                        self.log(f"   ❌ Erro tentativa {tentativa} modelo: {e}")
                        if tentativa < max_tentativas_modelo:
                            time.sleep(1)

                execute_or_skip(driver, "Selecionar ISS retention", lambda: clicar_iss_retention(driver, wait))

                self.log("7) Preenchendo CNPJ da prestadora...")
                for tentativa in range(1, 6):
                    try:
                        cnpj_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'tax_document_supplier_identification_number')))
                        cnpj_field.clear()
                        cnpj_field.send_keys(data['prestadora_cnpj'])

                        clicar_iss_retention(driver, wait)
                        self.aguardar_modal(driver)

                        try:
                            driver.find_element(By.TAG_NAME, 'body').click()
                        except: pass

                        WebDriverWait(driver, 15, poll_frequency=1).until(lambda d: d.find_element(By.ID, 'tax_document_supplier_legal_name').get_attribute('value').strip() != '')
                        self.log("   ✅ legal_name preenchido.")
                        break
                    except Exception as e:
                        self.log(f"   ❌ Erro ao processar prestadora: {e}")
                        time.sleep(1)

                self.log("8) Aguardando Inscrição Municipal...")
                if data['prestadora_inscricao_municipal']:
                    try:
                        im_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, 'tax_document_supplier_municipal_registration')))
                        im_field.clear()
                        im_field.send_keys(data['prestadora_inscricao_municipal'])
                    except Exception as e:
                        self.log(f"Erro IM: {e}")

                self.log("9) Selecionando cidade via JS de MouseEvents…")
                def selecionar_cidade_robusta(driver, nome_cidade, tentativas=7):
                    if not nome_cidade: return
                    for tentativa in range(1, tentativas + 1):
                        try:
                            driver.execute_script("document.getElementById('header').style.pointerEvents = 'none';")
                            campo = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[aria-labelledby='select2-tax_document_supplier_city_id-container']")))
                            driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('mousedown',{bubbles:true,cancelable:true}));", campo)

                            busca = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
                            busca.clear()
                            busca.send_keys(nome_cidade)
                            driver.execute_script("arguments[0].dispatchEvent(new Event('input',{ bubbles:true }));", busca)
                            time.sleep(1.0)

                            opcoes = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".select2-results__option")))
                            for item in opcoes:
                                if item.text.strip() == nome_cidade:
                                    driver.execute_script("arguments[0].scrollIntoView(true);", item)
                                    driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('mouseover',{bubbles:true})); arguments[0].dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));", item)
                                    self.log(f"✅ Cidade '{nome_cidade}' selecionada.")
                                    return
                        except Exception as e:
                            time.sleep(1)
                        finally:
                            driver.execute_script("document.getElementById('header').style.pointerEvents = '';")

                selecionar_cidade_robusta(driver, data['cidade'])

                self.log("10) Preenchendo CNPJ da tomadora…")
                for tentativa in range(1, 8):
                    try:
                        self.selecionar_select2_robusto(driver, container_css="span.select2-selection--single[aria-labelledby='select2-tax_document_customer_identification_number-container']", search_css="input.select2-search__field", nome_item=data['tomadora_cnpj'])
                        self.aguardar_modal(driver)
                        WebDriverWait(driver, 30).until(lambda d: d.find_element(By.ID, "tax_document_customer_legal_name").get_attribute("value").strip() != "")
                        break
                    except Exception as e:
                        time.sleep(1.5)

                self.log("11–17) Preenchendo campos numéricos e textuais...")
                field_map = [
                    ('tax_document_number', data['numero_documento']),
                    ('tax_document_total_value', data['valor_total']),
                    ('tax_document_invoice_items_attributes_0_purchase_order', data['pedido']),
                    ('tax_document_invoice_items_attributes_0_line_number', data['item']),
                    ('tax_document_invoice_items_attributes_0_frs', data['frs']),
                    ('tax_document_invoice_items_attributes_0_billing_report_code', data['rf']),
                    ('tax_document_invoice_items_attributes_0_contract_number', data['contrato']),
                ]
                for field_id, value in field_map:
                    if value:
                        if preencher_campo_robusto(field_id, value, max_tentativas_campo=6, atraso_campo=1.0):
                            self.log(f"   ✅ {field_id.split('_')[-1]} preenchido com '{value}'")

                self.log("18) Selecionando método de pagamento Crédito em Conta…")
                try:
                    select_el = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "tax_document_cf_payment_method")))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_el)
                    driver.execute_script("arguments[0].removeAttribute('disabled');", select_el)
                    driver.execute_script("""
                        const sel = arguments[0];
                        const opt = Array.from(sel.options).find(o => o.value === 'D' || o.text.includes('Crédito em Conta'));
                        if (opt) {
                            sel.value = opt.value;
                            sel.dispatchEvent(new Event('input', { bubbles:true }));
                            sel.dispatchEvent(new Event('change', { bubbles:true }));
                        }
                    """, select_el)
                    time.sleep(1.0)
                except Exception:
                    driver.execute_script("""
                        const sel = document.getElementById('tax_document_cf_payment_method');
                        if (sel) {
                            const opt = Array.from(sel.options).find(o => o.text.includes('Crédito em Conta'));
                            if (opt) { sel.value = opt.value; sel.dispatchEvent(new Event('change', { bubbles:true })); }
                        }
                    """)
                    self.log("   ⚙️ Fallback final aplicado para Crédito em Conta.")

                self.log("19) Preenchendo Datas (ISO + data-date)…")
                iso_emissao = converter_para_iso_date(data['data_emissao'])
                iso_pagamento = converter_para_iso_date(data['data_pagamento'])
                br_emissao = converter_para_br_date(data['data_emissao'])
                br_pagamento = converter_para_br_date(data['data_pagamento'])

                try:
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "tax_document_issue_date")))
                    driver.execute_script(f"""
                        var el = document.getElementById('tax_document_issue_date');
                        el.value = '{iso_emissao}';
                        el.setAttribute('data-date', '{br_emissao}');
                        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    """)
                    self.log(f"   ✅ Emissão setada para (ISO): {iso_emissao}")
                except Exception as e:
                    self.log(f"   ❌ Erro ao preencher emissão: {e}")

                try:
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "tax_document_net_due_date")))
                    driver.execute_script(f"""
                        var el = document.getElementById('tax_document_net_due_date');
                        el.value = '{iso_pagamento}';
                        el.setAttribute('data-date', '{br_pagamento}');
                        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    """)
                    self.log(f"   ✅ Vencimento setado para (ISO): {iso_pagamento}")
                except Exception as e:
                    self.log(f"   ❌ Erro ao preencher vencimento: {e}")
                
                campos_corrigidos = check_and_refill_fields(driver, data)
                if campos_corrigidos:
                    self.log(f"   ⚠️ Repreenchidos automaticamente: {[campo for campo, _ in campos_corrigidos]}")
    
            self.log("=== Processamento concluído para todos os arquivos ===")

        except Exception as e:
            self.log(f"❌ Erro crítico: {e}")
        finally:
            if driver:
                self.start_button.configure(state="normal")

if __name__ == '__main__':
    app = App()
    app.mainloop()
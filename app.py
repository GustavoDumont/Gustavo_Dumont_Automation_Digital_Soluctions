"""
Versao publica e anonimizada de uma automacao de postagem de documentos fiscais.

IMPORTANTE:
- Nao contem credenciais, contratos, contatos ou URL corporativa reais.
- A URL padrao aponta para example.invalid e nao executa postagem real.
- Use somente em ambiente proprio, de demonstracao ou formalmente autorizado.
- Os seletores podem exigir adaptacao para uma pagina de demonstracao.
"""

import os
import time
import re
import threading
import tkinter as tk
from datetime import datetime
from tkinter import Tk, Label, OptionMenu, StringVar, Button, messagebox, filedialog, Toplevel, Text, Scrollbar, RIGHT, Y, END, Radiobutton, Entry, Checkbutton, Scale, HORIZONTAL, LEFT, BOTH, X, Frame
from tkinter.ttk import Progressbar

import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import (
    WebDriverException,
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
    ElementNotInteractableException,
)

# ======================
# Dados de acesso e Contratos (mantive seus dados originais)
# ======================
# Configuracao publica e anonimizada.
# Credenciais devem ser informadas pelo usuario em tempo de execucao.
EMAIL = os.getenv("APP_EMAIL", "")
TOKEN = os.getenv("APP_TOKEN", "")

# Contratos e contatos abaixo sao totalmente ficticios.
CONTRATOS = {
    "CONTRATO DEMO 001": "gestor.demo1@example.invalid",
    "CONTRATO DEMO 002": "gestor.demo2@example.invalid",
}

# URL deliberadamente invalida. Configure somente para ambiente autorizado.
BASE_URL = os.getenv("APP_BASE_URL", "https://example.invalid/").rstrip("/") + "/"

# ======================
# Helpers para empacotamento portátil
# ======================
def resource_path(relative_path: str) -> str:
    """
    Resolve caminho de recurso no modo normal e no modo PyInstaller.
    Permite carregar chromedriver e browser portátil ao lado do .exe.
    """
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

def find_chromedriver_path():
    candidates = [
        resource_path("chromedriver.exe"),
        os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "chromedriver.exe"),
        os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "driver", "chromedriver.exe"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

def find_chrome_binary():
    candidates = [
        os.environ.get("CHROME_BINARY"),
        resource_path(os.path.join("chrome", "chrome.exe")),
        os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "chrome", "chrome.exe"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Chromium\Application\chrome.exe",
        r"C:\Program Files (x86)\Chromium\Application\chrome.exe",
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return None

# ======================
# Globais para abas e processor
# ======================
processor_global = None

# ======================
# Helpers de espera dinâmica
# ======================
def wait_for_page_load(driver, timeout=20):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except Exception as e:
        print(f"[LOG] Erro esperando carregamento da página: {e}")

def sleep_delay():
    time.sleep(max(0.0, float(delay_var.get())))

# ======================
# Agrupamento e extração
# ======================
def agrupar_arquivos(pdf_files, xml_files):
    pares = []
    xml_list = list(xml_files)
    for pdf in pdf_files:
        pdf_basename = os.path.basename(pdf)
        numeros_pdf = re.findall(r'\d+', pdf_basename)
        for xml in list(xml_list):
            xml_basename = os.path.basename(xml)
            if any(num in xml_basename for num in numeros_pdf):
                pares.append((pdf, xml))
                xml_list.remove(xml)
                break
    return pares

def extrair_identificador(filename):
    numeros = re.findall(r'\d+', os.path.basename(filename))
    for num in numeros:
        if len(num) >= 4:
            return num
    return numeros[0] if numeros else os.path.basename(filename)

def scroll_down_page(driver, stop_elements_ids=None, pause=0.2, max_wait=10):
    """
    Desce a página continuamente.
    Para se encontrar algum dos elementos listados em stop_elements_ids.
    """
    stop_elements_ids = stop_elements_ids or []
    last_scroll = time.time()
    
    while True:
        driver.execute_script("window.scrollBy(0, 80);")  # scroll contínuo
        time.sleep(pause)
        
        # Verifica se algum elemento já apareceu
        for eid in stop_elements_ids:
            try:
                el = driver.find_element(By.ID, eid)
                if el.is_displayed():
                    return  # encontrou, para o scroll
            except Exception:
                continue
        
        # Se não encontrar nada por max_wait segundos, interrompe
        if time.time() - last_scroll > max_wait:
            break

# ======================
# Classe Selenium (com safe click, select2 e foco nos inputs)
# ======================
class NotaFiscalProcessor:
    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-popup-blocking')
        # options.add_argument('--headless=new')  # opcional

        chrome_binary = find_chrome_binary()
        if chrome_binary:
            options.binary_location = chrome_binary

        chromedriver_path = find_chromedriver_path()
        try:
            if chromedriver_path:
                service = Service(executable_path=chromedriver_path)
                self.nav = webdriver.Chrome(service=service, options=options)
            else:
                # Fallback: Selenium Manager tenta localizar/baixar o driver.
                # Útil em desenvolvimento, mas não ideal para distribuição offline.
                self.nav = webdriver.Chrome(options=options)
        except WebDriverException as e:
            extra = []
            if not chrome_binary:
                extra.append("Chrome/Chromium não encontrado.")
            if not chromedriver_path:
                extra.append("chromedriver.exe não encontrado ao lado do app.")
            hint = " ".join(extra)
            raise Exception(f"Erro ao iniciar o navegador. {hint} Detalhe original: {e}")
        self.wait = WebDriverWait(self.nav, 30)

    def _scroll_into_view_with_offset(self, element, offset=-120):
        try:
            self.nav.execute_script(
                """
                const el = arguments[0];
                const offset = arguments[1];
                const rect = el.getBoundingClientRect();
                window.scrollBy(0, rect.top + window.pageYOffset + offset - 100);
                """,
                element,
                offset,
            )
        except Exception:
            try:
                self.nav.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
            except Exception:
                pass

    def pagina_expirada(self):
        try:
            src = self.nav.page_source.lower()
            return (
                "expired" in src
                and (
                    "errormodal" in src
                    or "erro desconhecido" in src
                    or "ocorreu um erro desconhecido" in src
                )
            )
        except Exception:
            return False

    def reautenticar(self, email, token):
        try:
            self.nav.delete_all_cookies()
        except Exception:
            pass
        self.nav.get(BASE_URL)
        wait_for_page_load(self.nav)
        self.wait.until(EC.element_to_be_clickable((By.ID, "login_portal"))).click()
        wait_for_page_load(self.nav)
        inputEmail = self.wait.until(EC.presence_of_element_located((By.ID, "user_login")))
        inputToken = self.wait.until(EC.presence_of_element_located((By.ID, "user_password")))
        inputEmail.clear(); inputEmail.send_keys(email)
        inputToken.clear(); inputToken.send_keys(token)
        self.nav.find_element(By.CSS_SELECTOR, "button.submit-button").click()
        wait_for_page_load(self.nav)

    def login(self, email, token):
        try:
            self.nav.get(BASE_URL)
            wait_for_page_load(self.nav)
            self.wait.until(EC.element_to_be_clickable((By.ID, "login_portal"))).click()
            wait_for_page_load(self.nav)
            inputEmail = self.wait.until(EC.presence_of_element_located((By.ID, "user_login")))
            inputToken = self.wait.until(EC.presence_of_element_located((By.ID, "user_password")))
            inputEmail.clear(); inputEmail.send_keys(email)
            inputToken.clear(); inputToken.send_keys(token)
            self.nav.find_element(By.CSS_SELECTOR, "button.submit-button").click()
            wait_for_page_load(self.nav)
        except TimeoutException:
            raise Exception("Timeout no login. Verifique conexão e página.")
        except Exception as e:
            raise Exception(f"Falha no login: {e}")

    def open_new_tab(self):
        self.nav.execute_script("window.open('');")
        self.nav.switch_to.window(self.nav.window_handles[-1])

    # ---------- utilitários de robustez ----------
    def _try_remove_overlays(self):
        js = """
        const selectors = [
          '.modal-backdrop', '.overlay', '.ui-dialog', '.fancybox-overlay', '[role=\"dialog\"]',
          '.cookie-banner', '.cookie-consent', '.toast', '.fixed-header'
        ];
        selectors.forEach(s=>{
          document.querySelectorAll(s).forEach(el=>{
            try{ el.style.display='none'; }catch(e){}
          });
        });
        """
        try:
            self.nav.execute_script(js)
        except Exception:
            pass

    def _safe_click_element(self, el, attempts=3, timeout_between=0.4):
        if el is None:
            return False
        for attempt in range(attempts):
            try:
                # scroll to center
                try:
                    self.nav.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", el)
                except Exception:
                    pass
                time.sleep(0.12)
                try:
                    el.click()
                    return True
                except (ElementClickInterceptedException, ElementNotInteractableException, StaleElementReferenceException):
                    pass
                # ActionChains fallback
                try:
                    ActionChains(self.nav).move_to_element(el).pause(0.08).click(el).perform()
                    return True
                except Exception:
                    pass
                # JS click fallback
                try:
                    self.nav.execute_script("arguments[0].click();", el)
                    return True
                except Exception:
                    pass
                # label[for=id] fallback
                try:
                    eid = el.get_attribute("id")
                    if eid:
                        try:
                            label = self.nav.find_element(By.CSS_SELECTOR, f"label[for='{eid}']")
                            try:
                                label.click()
                                return True
                            except Exception:
                                try:
                                    self.nav.execute_script("arguments[0].click();", label)
                                    return True
                                except Exception:
                                    pass
                        except Exception:
                            pass
                except Exception:
                    pass

                # remover overlays e tentar de novo
                self._try_remove_overlays()
                time.sleep(timeout_between)
                # atualizar referência se stale
                try:
                    eid = el.get_attribute("id")
                    if eid:
                        el = self.wait.until(EC.presence_of_element_located((By.ID, eid)))
                except Exception:
                    pass
            except StaleElementReferenceException:
                try:
                    eid = el.get_attribute("id")
                    if eid:
                        el = self.wait.until(EC.presence_of_element_located((By.ID, eid)))
                except Exception:
                    pass
            except Exception:
                time.sleep(timeout_between)
        return False

    def scroll_until_element_or_timeout(driver, by, value, timeout=30, idle_limit=10):
        """
        Faz scroll na página até encontrar o elemento ou até passar do timeout total.
        Para se ficar mais de idle_limit segundos sem progresso.
        """

        start_time = time.time()
        last_scroll_time = time.time()

        while time.time() - start_time < timeout:
            try:
                element = driver.find_element(by, value)
                if element.is_displayed():
                    print("Elemento encontrado!")
                    return element
            except:
                pass

            # dá scroll
            driver.execute_script("window.scrollBy(0, 400);")
            time.sleep(1)  # evita loop rápido demais

            # verifica progresso: se o scroll não trouxe nada novo por muito tempo
            if time.time() - last_scroll_time > idle_limit:
                print(f"Nenhum progresso em {idle_limit}s, parando scroll.")
                break

            last_scroll_time = time.time()

        print("Elemento não encontrado dentro do tempo limite.")
        return None


    # ---------- interações específicas ----------
    def fill_email(self, email):
        email_input = self.wait.until(EC.presence_of_element_located((By.ID, "tax_document_requester_area")))
        email_input.clear()
        email_input.send_keys(email)
        sleep_delay()

    def upload_file(self, file_input_id, file_path, file_type):
        basename = os.path.basename(file_path)
        last_error = None
        for attempt in range(2):
            try:
                field = self.wait.until(EC.presence_of_element_located((By.ID, file_input_id)))
                field.send_keys(file_path)
                sleep_delay()
                self.wait.until(EC.presence_of_element_located((By.XPATH, f"//a[contains(., '{basename}')]")))
                return True
            except Exception as e:
                last_error = e
                sleep_delay()
        print(f"[LOG] Falha upload {file_type} '{basename}': {last_error}")
        return False

    def click_radio_simples(self):
        try:
            el = self.wait.until(EC.presence_of_element_located((By.ID, "tax_document_supplier_opting_for_simples_nacional_0")))
            time.sleep(0.35)
            ok = self._safe_click_element(el)
            if not ok:
                print("[LOG] Falha ao marcar Simples: todas as tentativas falharam.")
            return ok
        except (TimeoutException, NoSuchElementException) as e:
            print(f"[LOG] Falha ao marcar Simples (não encontrado): {e}")
            return False

    def click_checkbox_labor(self):
        try:
            el = self.wait.until(EC.presence_of_element_located((By.ID, "tax_document_labor_assignment")))
            time.sleep(0.35)
            try:
                if el.is_selected():
                    return True
            except Exception:
                pass
            ok = self._safe_click_element(el)
            if not ok:
                print("[LOG] Falha ao marcar Labor Assignment: todas as tentativas falharam.")
            return ok
        except (TimeoutException, NoSuchElementException) as e:
            print(f"[LOG] Falha ao marcar Labor Assignment (não encontrado): {e}")
            return False

    def click_submit(self):
        try:
            btn = self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, 'button[name="status_id"][value="11"]')
                )
            )
            self.nav.execute_script(
                "arguments[0].scrollIntoView({block:'center', inline:'center'});", btn
            )
            time.sleep(0.2)

            clicked = False
            try:
                btn.click()
                clicked = True
            except Exception:
                pass

            if not clicked:
                try:
                    ActionChains(self.nav).move_to_element(btn).pause(0.08).click(btn).perform()
                    clicked = True
                except Exception:
                    pass

            if not clicked:
                try:
                    self.nav.execute_script("arguments[0].click();", btn)
                    clicked = True
                except Exception:
                    pass

            if not clicked:
                raise Exception("Não foi possível clicar no botão Ingressar Nota.")

            sleep_delay()

            # 1) Confirmação nativa do navegador (window.confirm / alert)
            try:
                self.wait.until(EC.alert_is_present())
                alerta = self.nav.switch_to.alert
                alerta.accept()
                wait_for_page_load(self.nav)
                sleep_delay()
                return True
            except TimeoutException:
                pass

            # 2) Confirmação em modal HTML do portal
            modal_btn_xpaths = [
                "//button[normalize-space()='OK']",
                "//button[normalize-space()='Ok']",
                "//button[normalize-space()='OK, entendi']",
                "//button[contains(normalize-space(), 'OK')]",
                "//button[contains(@class,'btn') and contains(., 'OK')]",
                "//button[contains(@class,'btn') and contains(., 'Ok')]",
            ]
            for xpath in modal_btn_xpaths:
                try:
                    ok_btn = WebDriverWait(self.nav, 6).until(
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    )
                    self.nav.execute_script(
                        "arguments[0].scrollIntoView({block:'center', inline:'center'});", ok_btn
                    )
                    time.sleep(0.15)
                    try:
                        ok_btn.click()
                    except Exception:
                        try:
                            ActionChains(self.nav).move_to_element(ok_btn).pause(0.08).click(ok_btn).perform()
                        except Exception:
                            self.nav.execute_script("arguments[0].click();", ok_btn)
                    wait_for_page_load(self.nav)
                    sleep_delay()
                    return True
                except TimeoutException:
                    continue

            # 3) Último fallback: fecha qualquer modal visível com botão OK
            try:
                self.nav.execute_script("""
                    const dialogs = [...document.querySelectorAll('[role="dialog"], .modal, .swal2-container, .ui-dialog')];
                    for (const d of dialogs) {
                        if (getComputedStyle(d).display === 'none') continue;
                        const btn = d.querySelector('button');
                        if (btn && /ok/i.test(btn.textContent || '')) { btn.click(); break; }
                    }
                """)
                wait_for_page_load(self.nav)
                sleep_delay()
                return True
            except Exception:
                pass

            return True
        except Exception as e:
            print(f"[LOG] Falha ao enviar (submit): {e}")
            return False


    def clicar_ingressar_nota(self):
        return self.click_submit()

    def get_protocol(self):
        try:
            elemento = self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'v-h4')))
            match = re.search(r'#(\d+)', elemento.text)
            return match.group(1) if match else "N/A"
        except Exception as e:
            print(f"[LOG] Erro protocolo: {e}")
            return "Erro"

    # ---------- select2 e inputs currency ----------
    def _focus_and_click_input(self, input_id, attempts=3):
        try:
            inp = self.wait.until(EC.presence_of_element_located((By.ID, input_id)))
        except Exception as e:
            print(f"[LOG] input {input_id} não encontrado: {e}")
            return False
        try:
            self._scroll_into_view_with_offset(inp, offset=-140)
        except Exception:
            pass
        for _ in range(attempts):
            try:
                if self._safe_click_element(inp):
                    try:
                        self.nav.execute_script("arguments[0].focus();", inp)
                        time.sleep(0.06)
                    except Exception:
                        pass
                    return True
            except Exception:
                pass
            try:
                self.nav.execute_script("arguments[0].focus(); arguments[0].click();", inp)
                time.sleep(0.08)
                return True
            except Exception:
                pass
            try:
                self.nav.execute_script("window.scrollBy(0, 50);")
            except Exception:
                pass
            time.sleep(0.12)
        print(f"[LOG] Falha ao focar/clicar input {input_id}")
        return False

    def select_select2_option(self, container_id, option_text="0,00%", timeout=8):
        try:
            container = self.wait.until(EC.presence_of_element_located((By.ID, container_id)))
        except Exception as e:
            print(f"[LOG] container select2 {container_id} não encontrado: {e}")
            return False
        try:
            select_wrapper = container.find_element(By.XPATH, "./ancestor::span[contains(@class,'select2-selection')]")
        except Exception:
            select_wrapper = container
        try:
            self._scroll_into_view_with_offset(select_wrapper, offset=-140)
        except Exception:
            pass
        time.sleep(0.12)
        if not self._safe_click_element(select_wrapper):
            print(f"[LOG] Falha ao abrir select2 {container_id}")
            return False
        option_xpath = f"//li[contains(@class,'select2-results__option') and normalize-space(.)='{option_text}']"
        try:
            opt = WebDriverWait(self.nav, timeout).until(
                EC.presence_of_element_located((By.XPATH, option_xpath))
            )
        except Exception as e:
            time.sleep(0.3)
            try:
                opt = self.nav.find_element(By.XPATH, option_xpath)
            except Exception:
                print(f"[LOG] Opção '{option_text}' não encontrada para {container_id}: {e}")
                return False
        try:
            self.nav.execute_script("arguments[0].scrollIntoView({block:'center'});", opt)
            time.sleep(0.08)
        except Exception:
            pass
        if not self._safe_click_element(opt):
            print(f"[LOG] Falha ao clicar opção '{option_text}' em {container_id}")
            return False
        try:
            WebDriverWait(self.nav, 3).until(lambda d: option_text in d.find_element(By.ID, container_id).text)
        except Exception:
            print(f"[LOG] Atenção: container {container_id} não exibiu '{option_text}' imediatamente.")
        return True

    def select_tax_rates_zero_and_click_values(self):
        pairs = [
            ("select2-tax_document_pis_tax_rate-container", "tax_document_pis_value"),
            ("select2-tax_document_cofins_tax_rate-container", "tax_document_cofins_value"),
            ("select2-tax_document_csll_tax_rate-container", "tax_document_csll_value"),
            ("select2-tax_document_ir_tax_rate-container", "tax_document_ir_value"),
            ("select2-tax_document_inss_tax_rate-container", "tax_document_inss_value"),
        ]
        results = {}

        stop_ids = [select_id for select_id, _ in pairs]

        # Scroll contínuo em background enquanto tenta localizar os campos
        def scroll_continuous():
            last_scroll = time.time()
            while True:
                try:
                    self.nav.execute_script("window.scrollBy(0, 80);")  # scroll contínuo
                    time.sleep(0.15)
                    found_any = False
                    for eid in stop_ids:
                        try:
                            el = self.nav.find_element(By.ID, eid)
                            if el.is_displayed():
                                found_any = True
                        except Exception:
                            continue
                    if found_any:
                        break  # encontrou algum, pode parar
                    if time.time() - last_scroll > 10:  # máximo de 10s sem progresso
                        break
                except Exception:
                    break

        scroll_thread = threading.Thread(target=scroll_continuous, daemon=True)
        scroll_thread.start()

        # Processa cada par select2 -> input
        for select_id, input_id in pairs:
            print(f"[LOG] Iniciando seleção e clique: {select_id} -> {input_id}")
            time.sleep(0.18)
            ok_select = self.select_select2_option(select_id, option_text="0,00%", timeout=8)
            if not ok_select:
                time.sleep(0.25)
                ok_select = self.select_select2_option(select_id, option_text="0,00%", timeout=5)
            time.sleep(0.12)
            ok_input = self._focus_and_click_input(input_id)
            if not ok_input:
                try:
                    self.nav.execute_script("window.scrollBy(0, 120);")
                except Exception:
                    pass
                time.sleep(0.12)
                ok_input = self._focus_and_click_input(input_id)
            results[select_id] = {"select_ok": ok_select, "input_ok": ok_input}
            time.sleep(0.18)

        scroll_thread.join(timeout=1)  # garante que a thread de scroll finalize
        print("[LOG] select_tax_rates_zero_and_click_values results:", results)
        return results



def detectar_expirado_e_relogin(processor, email, token):
    if processor and processor.pagina_expirada():
        print("[LOG] Sessão expirada detectada. Reautenticando...")
        processor.reautenticar(email, token)
        return True
    return False

def postar_uma_nota(processor, contrato_email, pdf, xml, logica, simples, labor, email, token):
    """
    Processa uma única nota do início ao fim.
    Retorna True se a postagem foi concluída.
    """
    nota = extrair_identificador(pdf)
    max_retries = 2

    for tentativa in range(1, max_retries + 1):
        try:
            if tentativa == 1:
                processor.open_new_tab()
            processor.nav.get(BASE_URL + "nf/tax_documents/service_invoice/new")
            wait_for_page_load(processor.nav)

            if detectar_expirado_e_relogin(processor, email, token):
                processor.nav.get(BASE_URL + "nf/tax_documents/service_invoice/new")
                wait_for_page_load(processor.nav)

            ui_set_status(f"[Log] Preparando nota {nota} (tentativa {tentativa})")
            processor.fill_email(contrato_email)

            ok_pdf = processor.upload_file("tax_document_document_pdf", pdf, "PDF")
            if not ok_pdf:
                raise RuntimeError("Falha no upload do PDF.")

            ok_xml = True
            if logica == "Nota Carioca":
                ok_xml = processor.upload_file("tax_document_document_xml", xml, "XML")
                if not ok_xml:
                    raise RuntimeError("Falha no upload do XML.")

            if simples:
                ui_set_status(f"[Log] Marcando Simples Nacional {nota}")
                processor.click_radio_simples()

            if labor:
                ui_set_status(f"[Log] Marcando Labor Assignment {nota}")
                processor.click_checkbox_labor()

            if logica != "Nota Carioca":
                ui_set_status(f"[Log] Ajustando rates {nota}")
                processor.select_tax_rates_zero_and_click_values()

            ui_set_status(f"[Log] Postando nota {nota}")
            processor.click_submit()

            if detectar_expirado_e_relogin(processor, email, token):
                raise RuntimeError("Sessão expirada durante o envio.")

            return True

        except Exception as e:
            msg = str(e).lower()
            print(f"[LOG] Falha ao postar nota {nota} na tentativa {tentativa}: {e}")

            if "expired" in msg or "expirada" in msg or "session" in msg:
                try:
                    processor.reautenticar(email, token)
                except Exception as relogin_err:
                    print(f"[LOG] Falha ao reautenticar: {relogin_err}")

            if tentativa >= max_retries:
                ui_set_status(f"[Log] Falha ao postar {nota}: {e}")
                return False

            time.sleep(1.0)

    return False

# ======================
# Wrappers seguros de UI (thread-safe)
# ======================
def ui_set_status(msg: str):
    def _set():
        status_label.config(text=msg)
        print(msg)
    root.after(0, _set)

def ui_progress_set(maximum=None, value=None, step=None):
    def _set():
        if maximum is not None:
            progress_bar["maximum"] = maximum
        if value is not None:
            progress_bar["value"] = value
        if step is not None:
            progress_bar.step(step)
    root.after(0, _set)

def ui_message(kind: str, title: str, text: str):
    def _show():
        if kind == "error":
            messagebox.showerror(title, text)
        elif kind == "warning":
            messagebox.showwarning(title, text)
        else:
            messagebox.showinfo(title, text)
    root.after(0, _show)

def ui_append_preview(lines):
    def _append():
        file_preview_text.configure(state="normal")
        for ln in lines:
            file_preview_text.insert(END, ln + "\n")
        file_preview_text.configure(state="disabled")
    root.after(0, _append)

# ======================
# Processar abas (upload) - fluxo principal
# ======================

def processar_abas(contrato, pdfs, xmls, logica):
    global processor_global
    ui_set_status("Iniciando browser e login...")

    try:
        processor_global = NotaFiscalProcessor()
    except Exception as e:
        ui_message("error", "Erro", str(e))
        ui_set_status("Falha ao iniciar o navegador.")
        return

    try:
        processor_global.login(EMAIL, TOKEN)
    except Exception as e:
        ui_message("error", "Erro no Login", str(e))
        ui_set_status("Falha no login.")
        return

    ui_set_status("Login concluído.")
    sleep_delay()

    gestor = CONTRATOS.get(contrato)
    if not gestor:
        ui_message("error", "Erro", "Contrato inválido!")
        return

    if logica == "Nota Carioca":
        pares = agrupar_arquivos(pdfs, xmls)
        if not pares:
            ui_message("warning", "Atenção", "Nenhum par PDF/XML encontrado!")
            return
        ui_progress_set(maximum=len(pares), value=0)

        for idx, (pdf, xml) in enumerate(pares, start=1):
            nota = extrair_identificador(pdf)
            ui_set_status(f"[Log] Processando nota {nota}")
            ok = postar_uma_nota(
                processor_global,
                gestor,
                pdf,
                xml,
                logica,
                simples_var.get(),
                labor_var.get(),
                EMAIL,
                TOKEN,
            )
            if not ok:
                ui_message("warning", "Atenção", f"Falha ao postar a nota {nota}.")
            ui_progress_set(value=idx)
            sleep_delay()

    else:
        ui_progress_set(maximum=len(pdfs), value=0)

        for idx, pdf in enumerate(pdfs, start=1):
            nota = extrair_identificador(pdf)
            ui_set_status(f"[Log] Processando nota {nota}")
            ok = postar_uma_nota(
                processor_global,
                gestor,
                pdf,
                None,
                logica,
                simples_var.get(),
                labor_var.get(),
                EMAIL,
                TOKEN,
            )
            if not ok:
                ui_message("warning", "Atenção", f"Falha ao postar a nota {nota}.")
            ui_progress_set(value=idx)
            sleep_delay()

    ui_set_status("Processamento concluído.")

# ======================
# GUI
# ======================
root = Tk()
root.title("Processador de Notas Fiscais")
root.geometry("650x820")

# Campos EMAIL/TOKEN
email_var = tk.StringVar(value=EMAIL)
token_var = tk.StringVar(value=TOKEN)
Label(root, text="EMAIL:").pack(pady=(10,0), anchor="w")
Entry(root, textvariable=email_var, width=60).pack(fill=X, padx=6)
Label(root, text="TOKEN:").pack(pady=(5,0), anchor="w")
Entry(root, textvariable=token_var, width=60, show="*").pack(fill=X, padx=6)

# Checkbuttons para Simples e Labor
simples_var = tk.BooleanVar(value=True)
labor_var   = tk.BooleanVar(value=True)
Checkbutton(root, text="Cadastro no Simples Nacional", variable=simples_var).pack(pady=5, anchor="w", padx=6)
Checkbutton(root, text="Não Vinculado a CNO (Labor Assignment)", variable=labor_var).pack(pady=5, anchor="w", padx=6)

# Controle de velocidade
Label(root, text="Velocidade (segundos de delay):").pack(pady=(15,0), anchor="w", padx=6)
delay_var = tk.DoubleVar(value=1.0)
Scale(root, variable=delay_var, from_=0.1, to=5.0, resolution=0.1, orient=HORIZONTAL, length=520).pack(padx=6)

# Contrato e Lógica
Label(root, text="Selecione o contrato:").pack(pady=5, anchor="w", padx=6)
contrato_var = StringVar(value=list(CONTRATOS.keys())[0])
OptionMenu(root, contrato_var, *CONTRATOS.keys()).pack(padx=6, fill=X)

Label(root, text="Lógica de envio:").pack(pady=5, anchor="w", padx=6)
logica_envio = StringVar(value="Nota Carioca")
Radiobutton(root, text="Nota Carioca", variable=logica_envio, value="Nota Carioca").pack(anchor="w", padx=6)
Radiobutton(root, text="Nota do Milhão", variable=logica_envio, value="Nota do Milhão").pack(anchor="w", padx=6)

buttons_frame = Frame(root)
buttons_frame.pack(fill=X, padx=6, pady=10)

Button(
    buttons_frame,
    text="Selecionar Arquivos",
    command=lambda: selecionar_arquivos_interface()
).pack(side=LEFT, padx=(0, 8))

Button(
    buttons_frame,
    text="Processar Arquivos",
    command=lambda: iniciar_processamento()
).pack(side=LEFT)

# Pré-visualização com scrollbar correta
Label(root, text="Pré-visualização:").pack(anchor="w", padx=6)
preview_frame = Frame(root)
preview_frame.pack(fill=BOTH, expand=False, padx=6, pady=(0,6))
file_preview_text = Text(preview_frame, height=12, width=70, state="disabled")
file_preview_text.pack(side=LEFT, fill=BOTH, expand=True)
preview_scrollbar = Scrollbar(preview_frame, command=file_preview_text.yview)
preview_scrollbar.pack(side=RIGHT, fill=Y)
file_preview_text.config(yscrollcommand=preview_scrollbar.set)

progress_bar = Progressbar(root, orient="horizontal", length=580, mode="determinate")
progress_bar.pack(pady=15, padx=6)
status_label = Label(root, text="Status: Aguardando ação...")
status_label.pack(pady=5, padx=6, anchor="w")

# Funções de seleção e início
pdf_files_global = []
xml_files_global = []

def selecionar_arquivos_interface():
    global pdf_files_global, xml_files_global
    pdfs = filedialog.askopenfilenames(title="Selecione PDFs", filetypes=[("PDF", "*.pdf")])
    if not pdfs:
        return
    xmls = []
    if logica_envio.get() == "Nota Carioca":
        xmls = filedialog.askopenfilenames(title="Selecione XMLs", filetypes=[("XML", "*.xml")])
        if not xmls:
            ui_message("error", "Erro", "Selecione XMLs para Nota Carioca.")
            return
    pdf_files_global = list(pdfs)
    xml_files_global = list(xmls)
    lines = [os.path.basename(f) for f in pdf_files_global]
    if xml_files_global:
        lines.append("")
        lines.extend(os.path.basename(f) for f in xml_files_global)
    file_preview_text.configure(state="normal")
    file_preview_text.delete("1.0", END)
    file_preview_text.configure(state="disabled")
    ui_append_preview(lines)
    ui_message("info", "Arquivos", f"Selecionados {len(pdf_files_global)} PDF(s){' e '+str(len(xml_files_global))+' XML(s)' if xml_files_global else ''}.")

def iniciar_processamento():
    global EMAIL, TOKEN
    EMAIL = email_var.get().strip()
    TOKEN = token_var.get().strip()
    if not EMAIL or not TOKEN:
        ui_message("error", "Erro", "Preencha EMAIL e TOKEN antes de iniciar.")
        return
    contrato = contrato_var.get().strip()
    logica   = logica_envio.get()
    if logica == "Nota Carioca" and (not pdf_files_global or not xml_files_global):
        ui_message("error", "Erro", "Selecione PDFs e XMLs antes.")
        return
    if logica == "Nota do Milhão" and not pdf_files_global:
        ui_message("error", "Erro", "Selecione PDFs antes.")
        return
    t = threading.Thread(target=processar_abas, args=(contrato, pdf_files_global, xml_files_global, logica), daemon=True)
    t.start()

root.mainloop()

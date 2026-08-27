from __future__ import annotations
import csv, os, time
from datetime import date
from pathlib import Path
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from models import DateRange, DownloadResult

class TotvsVacationDownloader:
    """Adapter para uso autorizado no Portal Meu RH.

    Os seletores abaixo são placeholders seguros. Um responsável técnico deve
    mapeá-los na versão vigente do portal antes do uso. A aplicação não contorna
    MFA, CAPTCHA, bloqueios ou controles de acesso.
    """
    def __init__(self, portal_url: str, download_dir: Path, log=print):
        self.portal_url=portal_url; self.download_dir=download_dir; self.log=log
        self.driver=None; self.wait=None; self.stop_requested=False

    def start(self):
        self.download_dir.mkdir(parents=True,exist_ok=True)
        options=Options()
        options.add_experimental_option('prefs',{'download.default_directory':str(self.download_dir.resolve()),'download.prompt_for_download':False,'plugins.always_open_pdf_externally':True})
        self.driver=webdriver.Chrome(options=options); self.wait=WebDriverWait(self.driver,30)
        self.driver.get(self.portal_url)
        self.log('Portal aberto. Conclua a autenticação autorizada no navegador.')

    def wait_for_user_login(self):
        # Configure uma condição estável do ambiente autorizado.
        self.wait.until(lambda d: d.current_url and d.current_url != 'data:,')

    def list_records(self, period: DateRange) -> list[dict]:
        """Mapear no portal: aplicar período e retornar registros visíveis.

        Retorno esperado por item: {'reference': str, 'date': date,
        'notice_locator': tuple|None, 'receipt_locator': tuple|None}.
        """
        raise NotImplementedError('Configure os seletores do portal em list_records().')

    def download_locator(self, locator, timeout=30):
        WebDriverWait(self.driver,timeout).until(EC.element_to_be_clickable(locator)).click()
        time.sleep(0.8)

    def process(self, period: DateRange, types=('AVISO','RECIBO')) -> list[DownloadResult]:
        results=[]
        for item in self.list_records(period):
            if self.stop_requested: break
            for kind in types:
                locator=item.get('notice_locator' if kind=='AVISO' else 'receipt_locator')
                if not locator:
                    results.append(DownloadResult(item.get('reference',''),kind,'INDISPONIVEL')); continue
                try:
                    self.download_locator(locator); results.append(DownloadResult(item.get('reference',''),kind,'DOWNLOAD_SOLICITADO'))
                except TimeoutException as e:
                    results.append(DownloadResult(item.get('reference',''),kind,'TIMEOUT',str(e).splitlines()[0]))
                except WebDriverException as e:
                    results.append(DownloadResult(item.get('reference',''),kind,'ERRO_NAVEGADOR',str(e).splitlines()[0]))
        return results

    def close(self):
        if self.driver: self.driver.quit()

def save_report(results, path: Path):
    with path.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.writer(f,delimiter=';'); w.writerow(['referencia','tipo_documento','status','detalhe'])
        w.writerows((r.reference,r.document_type,r.status,r.detail) for r in results)

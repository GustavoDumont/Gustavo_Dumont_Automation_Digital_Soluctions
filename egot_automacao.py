from __future__ import annotations
import json, queue, threading, time
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait

class EgotBot:
    def __init__(self, config, output, emit):
        self.config=config; self.output=output; self.emit=emit; self.stop=threading.Event(); self.driver=None
    def run(self):
        self.output.mkdir(parents=True,exist_ok=True)
        self.driver=webdriver.Chrome(); self.driver.get(self.config['portal_url'])
        self.emit('Portal aberto. Conclua a autenticação autorizada.')
        self.process_contracts()
    def process_contracts(self):
        raise NotImplementedError('Configure os seletores do portal no ambiente autorizado.')
    def close(self):
        if self.driver:self.driver.quit()

class App(ctk.CTk):
    def __init__(self):
        super().__init__(); self.title('Download de Pendências do NACT'); self.geometry('760x520'); self.events=queue.Queue(); self.bot=None
        self.config_path=ctk.StringVar(value='egot_config.json'); self.output=ctk.StringVar(value=str(Path.home()/'Downloads'/'Pendencias_NACT')); self._build(); self.after(150,self.consume)
    def _build(self):
        ctk.CTkLabel(self,text='Download de Pendências do NACT',font=ctk.CTkFont(size=22,weight='bold')).pack(pady=18)
        for label,var,cmd in [('Configuração',self.config_path,lambda:self.config_path.set(filedialog.askopenfilename(filetypes=[('JSON','*.json')]) or self.config_path.get())),('Pasta de saída',self.output,lambda:self.output.set(filedialog.askdirectory() or self.output.get()))]:
            row=ctk.CTkFrame(self); row.pack(fill='x',padx=20,pady=6); ctk.CTkLabel(row,text=label,width=120).pack(side='left'); ctk.CTkEntry(row,textvariable=var).pack(side='left',fill='x',expand=True,padx=6); ctk.CTkButton(row,text='Selecionar',command=cmd,width=90).pack(side='left')
        self.status=ctk.StringVar(value='Pronto'); ctk.CTkLabel(self,textvariable=self.status).pack(pady=12)
        ctk.CTkButton(self,text='Iniciar',command=self.start,fg_color='#F58220').pack(); ctk.CTkButton(self,text='Solicitar parada',command=self.stop).pack(pady=8)
    def start(self):
        try:cfg=json.loads(Path(self.config_path.get()).read_text(encoding='utf-8'))
        except Exception as e:messagebox.showerror('Configuração',str(e));return
        self.bot=EgotBot(cfg,Path(self.output.get()),lambda m:self.events.put(m)); threading.Thread(target=self.worker,daemon=True).start()
    def worker(self):
        try:self.bot.run()
        except Exception as e:self.events.put(str(e))
        finally:self.bot.close()
    def stop(self):
        if self.bot:self.bot.stop.set();self.status.set('Parada solicitada')
    def consume(self):
        try:
            while True:self.status.set(self.events.get_nowait())
        except queue.Empty:pass
        self.after(150,self.consume)
if __name__=='__main__':App().mainloop()

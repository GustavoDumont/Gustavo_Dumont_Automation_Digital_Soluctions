from __future__ import annotations
import os, threading, tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from automation import TotvsVacationDownloader, save_report
from models import DateRange

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title('Avisos e Recibos de Férias'); self.geometry('720x520'); self.running=False; self.bot=None; self._build()
    def _build(self):
        ttk.Label(self,text='Download de Avisos e Recibos de Férias',font=('Segoe UI',18,'bold')).pack(pady=(18,3))
        ttk.Label(self,text='Automação autorizada do Portal Meu RH por intervalo de datas.').pack()
        form=ttk.Frame(self,padding=18); form.pack(fill='x'); form.columnconfigure(1,weight=1)
        self.url=tk.StringVar(value=os.getenv('TOTVS_PORTAL_URL','')); self.out=tk.StringVar(value=str(Path.home()/'Downloads'/'Ferias'))
        self.start_date=tk.StringVar(); self.end_date=tk.StringVar(); self.notice=tk.BooleanVar(value=True); self.receipt=tk.BooleanVar(value=True)
        fields=[('URL do portal',self.url),('Data inicial (dd/mm/aaaa)',self.start_date),('Data final (dd/mm/aaaa)',self.end_date),('Pasta de saída',self.out)]
        for i,(label,var) in enumerate(fields):
            ttk.Label(form,text=label).grid(row=i,column=0,sticky='e',padx=6,pady=7); ttk.Entry(form,textvariable=var).grid(row=i,column=1,sticky='ew',padx=6)
        ttk.Button(form,text='Selecionar',command=lambda:self.out.set(filedialog.askdirectory() or self.out.get())).grid(row=3,column=2,padx=6)
        kinds=ttk.Frame(self); kinds.pack(); ttk.Checkbutton(kinds,text='Avisos',variable=self.notice).pack(side='left',padx=10); ttk.Checkbutton(kinds,text='Recibos',variable=self.receipt).pack(side='left',padx=10)
        self.status=tk.StringVar(value='Pronto'); ttk.Label(self,textvariable=self.status).pack(pady=10)
        buttons=ttk.Frame(self); buttons.pack(); self.run_btn=ttk.Button(buttons,text='Abrir portal e processar',command=self.start); self.run_btn.pack(side='left',padx=5); ttk.Button(buttons,text='Solicitar parada',command=self.stop).pack(side='left',padx=5)
        ttk.Label(self,text='A versão GitHub exige configuração dos seletores do ambiente autorizado.',foreground='#8a4a00').pack(pady=18)
    def start(self):
        if self.running:return
        try:
            start=datetime.strptime(self.start_date.get(),'%d/%m/%Y').date(); end=datetime.strptime(self.end_date.get(),'%d/%m/%Y').date(); period=DateRange(start,end)
        except Exception as e: messagebox.showwarning('Período inválido',str(e)); return
        kinds=tuple(x for x,enabled in [('AVISO',self.notice.get()),('RECIBO',self.receipt.get())] if enabled)
        if not self.url.get().strip() or not self.out.get().strip() or not kinds: messagebox.showwarning('Campos obrigatórios','Informe URL, pasta e ao menos um tipo de documento.'); return
        self.running=True; self.run_btn.state(['disabled']); threading.Thread(target=self.worker,args=(period,kinds),daemon=True).start()
    def worker(self,period,kinds):
        folder=Path(self.out.get()); self.bot=TotvsVacationDownloader(self.url.get().strip(),folder,lambda m:self.after(0,lambda:self.status.set(m)))
        try:
            self.bot.start(); self.bot.wait_for_user_login(); results=self.bot.process(period,kinds)
            report=folder/f'resultado_ferias_{datetime.now():%Y%m%d_%H%M%S}.csv'; save_report(results,report)
            self.after(0,lambda:messagebox.showinfo('Concluído',f'Relatório salvo em:\n{report}'))
        except NotImplementedError as e: self.after(0,lambda:messagebox.showwarning('Configuração necessária',str(e)))
        except Exception as e: self.after(0,lambda:messagebox.showerror('Erro',str(e)))
        finally:
            self.bot.close(); self.running=False; self.after(0,lambda:self.run_btn.state(['!disabled'])); self.after(0,lambda:self.status.set('Pronto'))
    def stop(self):
        if self.bot:self.bot.stop_requested=True; self.status.set('Parada solicitada...')
if __name__=='__main__':App().mainloop()

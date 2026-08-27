from __future__ import annotations
import os, threading, tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from automation import UnicoBackupAdapter, write_results
from models import Candidate

CATEGORIES=['Documentos pessoais','Assinaturas','Ficha de registro','Etiquetas']
class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title('Backup UNICO People'); self.geometry('760x610'); self.bot=None; self.running=False; self._build()
    def _build(self):
        ttk.Label(self,text='Backup autorizado de registros',font=('Segoe UI',18,'bold')).pack(pady=(16,3))
        ttk.Label(self,text='Aplicação de contingência e preservação documental').pack()
        form=ttk.Frame(self,padding=14); form.pack(fill='x'); form.columnconfigure(1,weight=1)
        self.url=tk.StringVar(value=os.getenv('UNICO_PORTAL_URL','')); self.output=tk.StringVar(value=str(Path.home()/'Downloads'/'Backup_UNICO'))
        for row,(label,var) in enumerate([('URL do portal',self.url),('Pasta de saída',self.output)]):
            ttk.Label(form,text=label).grid(row=row,column=0,sticky='e',padx=6,pady=6); ttk.Entry(form,textvariable=var).grid(row=row,column=1,sticky='ew',padx=6)
        ttk.Button(form,text='Selecionar',command=lambda:self.output.set(filedialog.askdirectory() or self.output.get())).grid(row=1,column=2,padx=6)
        ttk.Label(self,text='Pessoas ou identificadores, um por linha').pack(anchor='w',padx=18)
        self.people=tk.Text(self,height=10); self.people.pack(fill='both',expand=True,padx=18,pady=5)
        kind=ttk.LabelFrame(self,text='Categorias',padding=10); kind.pack(fill='x',padx=18,pady=8); self.selected={}
        for c in CATEGORIES:
            v=tk.BooleanVar(value=True); self.selected[c]=v; ttk.Checkbutton(kind,text=c,variable=v).pack(side='left',padx=8)
        self.status=tk.StringVar(value='Pronto'); ttk.Label(self,textvariable=self.status).pack(pady=5)
        buttons=ttk.Frame(self); buttons.pack(pady=5); self.run_btn=ttk.Button(buttons,text='Abrir portal e iniciar',command=self.start); self.run_btn.pack(side='left',padx=5); ttk.Button(buttons,text='Solicitar parada',command=self.stop).pack(side='left',padx=5)
        ttk.Label(self,text='A versão pública exige configuração dos seletores no ambiente autorizado.',foreground='#8a4a00').pack(pady=12)
    def start(self):
        names=list(dict.fromkeys(x.strip() for x in self.people.get('1.0','end').splitlines() if x.strip())); cats=[k for k,v in self.selected.items() if v.get()]
        if not self.url.get().strip() or not names or not cats: messagebox.showwarning('Dados obrigatórios','Informe URL, pessoas e categorias.'); return
        if self.running:return
        self.running=True; self.run_btn.state(['disabled']); threading.Thread(target=self.worker,args=(names,cats),daemon=True).start()
    def worker(self,names,cats):
        root=Path(self.output.get()); self.bot=UnicoBackupAdapter(self.url.get().strip(),root,lambda m:self.after(0,lambda:self.status.set(m)))
        try:
            self.bot.start(); results=self.bot.process([Candidate(n) for n in names],cats)
            report=root/f'backup_resultado_{datetime.now():%Y%m%d_%H%M%S}.csv'; write_results(results,report)
            self.after(0,lambda:messagebox.showinfo('Finalizado','Relatório salvo na pasta de saída.'))
        except NotImplementedError as exc:self.after(0,lambda:messagebox.showwarning('Configuração necessária',str(exc)))
        except Exception as exc:self.after(0,lambda:messagebox.showerror('Erro',str(exc)))
        finally:
            self.bot.close(); self.running=False; self.after(0,lambda:self.run_btn.state(['!disabled'])); self.after(0,lambda:self.status.set('Pronto'))
    def stop(self):
        if self.bot:self.bot.stop_requested=True; self.status.set('Parada solicitada')
if __name__=='__main__':App().mainloop()

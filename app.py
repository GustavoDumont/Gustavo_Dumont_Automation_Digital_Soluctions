from __future__ import annotations
import threading, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from extractor import load_rule, process, configure_tesseract

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title('Buscador de Dados e Renomeador'); self.geometry('720x470')
        self.folder=tk.StringVar(); self.output=tk.StringVar(); self.rule=tk.StringVar(); self.rename=tk.BooleanVar(value=True); self._build()
    def _row(self,parent,label,var,cmd,row):
        ttk.Label(parent,text=label).grid(row=row,column=0,sticky='e',padx=6,pady=7)
        ttk.Entry(parent,textvariable=var).grid(row=row,column=1,sticky='ew',padx=6)
        ttk.Button(parent,text='Selecionar',command=cmd).grid(row=row,column=2,padx=6)
    def _build(self):
        ttk.Label(self,text='Buscador de Dados e Renomeador',font=('Segoe UI',18,'bold')).pack(pady=(18,4))
        ttk.Label(self,text='Extração configurável em PDFs e imagens, com OCR quando necessário.').pack()
        box=ttk.Frame(self,padding=16); box.pack(fill='x'); box.columnconfigure(1,weight=1)
        self._row(box,'Pasta de entrada',self.folder,lambda:self.folder.set(filedialog.askdirectory()),0)
        self._row(box,'Pasta de saída',self.output,lambda:self.output.set(filedialog.askdirectory()),1)
        self._row(box,'Regra JSON',self.rule,lambda:self.rule.set(filedialog.askopenfilename(filetypes=[('Configuração','*.json')])),2)
        ttk.Checkbutton(box,text='Copiar e renomear arquivos encontrados',variable=self.rename).grid(row=3,column=1,sticky='w',pady=8)
        self.status=tk.StringVar(value='Pronto'); ttk.Label(self,textvariable=self.status).pack(pady=8)
        self.btn=ttk.Button(self,text='Processar',command=self.start); self.btn.pack()
        ttk.Label(self,text='Use somente em documentos e pastas para os quais exista autorização.',foreground='#666').pack(side='bottom',pady=20)
    def start(self):
        if not all([self.folder.get(),self.output.get(),self.rule.get()]): messagebox.showwarning('Campos obrigatórios','Selecione entrada, saída e regra.'); return
        self.btn.state(['disabled']); self.status.set('Processando...')
        threading.Thread(target=self.run,daemon=True).start()
    def run(self):
        try:
            configure_tesseract(); report=process(Path(self.folder.get()),Path(self.output.get()),load_rule(Path(self.rule.get())),self.rename.get())
            self.after(0,lambda:messagebox.showinfo('Concluído',f'Relatório: {report}')); self.after(0,lambda:self.status.set('Concluído'))
        except Exception as e:
            self.after(0,lambda:messagebox.showerror('Erro',str(e))); self.after(0,lambda:self.status.set('Erro'))
        finally:self.after(0,lambda:self.btn.state(['!disabled']))
if __name__=='__main__': App().mainloop()

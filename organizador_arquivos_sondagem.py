from __future__ import annotations
import csv, difflib, re, shutil, threading, unicodedata
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox
CATEGORIES=['ATESTADO','FOLHA DE PONTO','HORAS EXTRAS','PERICULOSIDADE']
def norm(s):
    s=unicodedata.normalize('NFKD',s or '').encode('ascii','ignore').decode('ascii').lower();return re.sub(r'[^a-z0-9]+',' ',s).strip()
def score(a,b):return difflib.SequenceMatcher(None,norm(a),norm(b)).ratio()
def unique(p):
    if not p.exists():return p
    i=2
    while (p.parent/f'{p.stem}_{i}{p.suffix}').exists():i+=1
    return p.parent/f'{p.stem}_{i}{p.suffix}'
def organize(source,people_root,output,move=False,dry=True,threshold=.72):
    people=[p for p in people_root.iterdir() if p.is_dir()]; rows=[]
    for f in source.rglob('*'):
        if not f.is_file():continue
        ranked=sorted(((score(f.stem,p.name),p) for p in people),reverse=True,key=lambda x:x[0]); best=ranked[0] if ranked else (0,None)
        category=next((c for c in CATEGORIES if norm(c) in norm(f.stem)), 'NAO_CLASSIFICADO')
        if best[0]<threshold or not best[1]:rows.append([str(f),'','',best[0],'SEM_CORRESPONDENCIA']);continue
        dest=unique(output/best[1].name/category/f.name)
        if not dry:
            dest.parent.mkdir(parents=True,exist_ok=True); shutil.move(f,dest) if move else shutil.copy2(f,dest)
        rows.append([str(f),str(dest),category,best[0],'SIMULADO' if dry else 'CONCLUIDO'])
    output.mkdir(parents=True,exist_ok=True)
    with (output/'relatorio_organizacao.csv').open('w',newline='',encoding='utf-8-sig') as h:
        w=csv.writer(h,delimiter=';');w.writerow(['origem','destino','categoria','confianca','status']);w.writerows(rows)
class App(ctk.CTk):
    def __init__(self):
        super().__init__();self.title('Organizador de Arquivos NACT');self.geometry('760x540');self.vars={k:ctk.StringVar() for k in ['entrada','colaboradores','saida']};self.dry=ctk.BooleanVar(value=True);self.move=ctk.BooleanVar();self._build()
    def _build(self):
        ctk.CTkLabel(self,text='Organizador de Arquivos',font=ctk.CTkFont(size=22,weight='bold')).pack(pady=18)
        for label,key in [('Documentos recebidos','entrada'),('Pastas dos colaboradores','colaboradores'),('Pasta de saída','saida')]:
            r=ctk.CTkFrame(self);r.pack(fill='x',padx=20,pady=5);ctk.CTkLabel(r,text=label,width=170).pack(side='left');ctk.CTkEntry(r,textvariable=self.vars[key]).pack(side='left',fill='x',expand=True);ctk.CTkButton(r,text='Selecionar',command=lambda k=key:self.vars[k].set(filedialog.askdirectory())).pack(side='left',padx=5)
        ctk.CTkCheckBox(self,text='Modo de simulação',variable=self.dry).pack(pady=8);ctk.CTkCheckBox(self,text='Mover em vez de copiar',variable=self.move).pack();ctk.CTkButton(self,text='Processar',command=self.start,fg_color='#F28C28').pack(pady=18)
    def start(self):
        if not all(v.get() for v in self.vars.values()):messagebox.showwarning('Campos','Selecione as três pastas.');return
        threading.Thread(target=self.worker,daemon=True).start()
    def worker(self):
        try:organize(Path(self.vars['entrada'].get()),Path(self.vars['colaboradores'].get()),Path(self.vars['saida'].get()),self.move.get(),self.dry.get());self.after(0,lambda:messagebox.showinfo('Concluído','Relatório gerado.'))
        except Exception as e:self.after(0,lambda:messagebox.showerror('Erro',str(e)))
if __name__=='__main__':App().mainloop()

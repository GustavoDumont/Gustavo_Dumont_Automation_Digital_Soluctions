from __future__ import annotations
import csv, json, re, shutil, sys
from dataclasses import dataclass
from pathlib import Path
import fitz, pytesseract
from PIL import Image, ImageEnhance, ImageOps
from common import normalize, safe_name, unique_path

@dataclass
class Rule:
    name: str
    mode: str="between"
    start: str=""
    end: str=""
    pattern: str=""
    group: int=1
    pages: int=4

def configure_tesseract():
    local=Path(sys.executable).parent/'tesseract'/'tesseract.exe' if getattr(sys,'frozen',False) else Path(__file__).parent.parent/'tesseract'/'tesseract.exe'
    if local.exists(): pytesseract.pytesseract.tesseract_cmd=str(local)

def ocr(image: Image.Image) -> str:
    image=ImageOps.grayscale(image); image=ImageOps.autocontrast(image)
    image=ImageEnhance.Contrast(image).enhance(1.8)
    return pytesseract.image_to_string(image,lang='por',config='--oem 3 --psm 6') or ''

def pdf_text(path: Path, max_pages=4) -> str:
    chunks=[]
    with fitz.open(path) as doc:
        for page in list(doc)[:max_pages]:
            text=page.get_text('text') or ''
            if len(text.strip())<40:
                pix=page.get_pixmap(matrix=fitz.Matrix(300/72,300/72),alpha=False)
                text+='\n'+ocr(Image.frombytes('RGB',[pix.width,pix.height],pix.samples))
            chunks.append(text)
    return '\n'.join(chunks)

def image_text(path: Path) -> str: return ocr(Image.open(path))
def read_text(path: Path, pages=4) -> str:
    if path.suffix.lower()=='.pdf': return pdf_text(path,pages)
    if path.suffix.lower() in {'.png','.jpg','.jpeg','.tif','.tiff','.bmp'}: return image_text(path)
    if path.suffix.lower()=='.txt': return path.read_text(encoding='utf-8',errors='replace')
    return ''

def extract(text: str, rule: Rule) -> str:
    source=normalize(text)
    if rule.mode=='regex':
        m=re.search(rule.pattern,source,re.I|re.S)
        return normalize(m.group(rule.group)) if m else ''
    a=source.upper().find(normalize(rule.start).upper())
    if a<0:return ''
    a+=len(normalize(rule.start)); b=source.upper().find(normalize(rule.end).upper(),a) if rule.end else len(source)
    return normalize(source[a:b if b>=0 else len(source)])

def load_rule(path: Path) -> Rule: return Rule(**json.loads(path.read_text(encoding='utf-8')))
def process(folder: Path, output: Path, rule: Rule, rename=False):
    output.mkdir(parents=True,exist_ok=True); rows=[]
    for p in sorted(folder.rglob('*')):
        if not p.is_file() or p.suffix.lower() not in {'.pdf','.png','.jpg','.jpeg','.tif','.tiff','.bmp','.txt'}: continue
        try:
            value=extract(read_text(p,rule.pages),rule); status='ENCONTRADO' if value else 'NAO_ENCONTRADO'; dest=''
            if rename and value:
                target=unique_path(output,safe_name(value)+p.suffix.lower()); shutil.copy2(p,target); dest=str(target)
            rows.append([str(p),p.name,value,status,dest])
        except Exception as e: rows.append([str(p),p.name,'','ERRO',str(e)])
    report=output/'resultado_extracao.csv'
    with report.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.writer(f,delimiter=';'); w.writerow(['caminho','arquivo','valor','status','saida_ou_erro']); w.writerows(rows)
    return report

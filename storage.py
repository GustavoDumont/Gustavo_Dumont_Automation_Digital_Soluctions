from __future__ import annotations
import hashlib, json, re, unicodedata
from pathlib import Path

def normalize(value: str) -> str:
    value=unicodedata.normalize("NFKD",value or "").encode("ascii","ignore").decode("ascii").upper()
    return re.sub(r"\s+"," ",value).strip()

def safe_name(value: str, fallback="SEM_NOME") -> str:
    value=normalize(value); value=re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',value).strip(' .')
    return value[:100] or fallback

def candidate_folder(root: Path, name: str) -> Path:
    path=root/safe_name(name); path.mkdir(parents=True,exist_ok=True); return path

def sha256(path: Path, chunk=1024*1024) -> str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda:f.read(chunk),b''): h.update(block)
    return h.hexdigest()

def manifest(folder: Path) -> dict:
    files=[]
    for p in sorted(folder.rglob('*')):
        if p.is_file(): files.append({'path':str(p.relative_to(folder)),'bytes':p.stat().st_size,'sha256':sha256(p)})
    return {'folder':folder.name,'files':files,'file_count':len(files),'total_bytes':sum(x['bytes'] for x in files)}

def save_manifest(folder: Path) -> Path:
    target=folder/'_manifest.json'; target.write_text(json.dumps(manifest(folder),ensure_ascii=False,indent=2),encoding='utf-8'); return target

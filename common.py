from __future__ import annotations
import hashlib, mimetypes, os, re, unicodedata
from datetime import datetime
from pathlib import Path

def normalize(text: str) -> str:
    text=unicodedata.normalize("NFKD",text or "").encode("ascii","ignore").decode("ascii")
    return re.sub(r"\s+"," ",text).strip()

def safe_name(value: str, fallback="SEM_DADO") -> str:
    value=normalize(value).upper()
    value=re.sub(r"[^A-Z0-9._ -]","",value)
    value=re.sub(r"[ ._-]+","_",value).strip("_")
    return value or fallback

def unique_path(folder: Path, filename: str) -> Path:
    target=folder/filename
    if not target.exists(): return target
    stem,suffix=target.stem,target.suffix
    i=2
    while (folder/f"{stem}_{i}{suffix}").exists(): i+=1
    return folder/f"{stem}_{i}{suffix}"

def format_size(size: int) -> str:
    units=["B","KB","MB","GB","TB"]
    n=float(size)
    for unit in units:
        if n<1024 or unit==units[-1]: return f"{n:.1f} {unit}" if unit!="B" else f"{int(n)} B"
        n/=1024

def dt(ts: float) -> str:
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y %H:%M:%S")

def sha256(path: Path, chunk=1024*1024) -> str:
    h=hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda:f.read(chunk),b""): h.update(block)
    return h.hexdigest()

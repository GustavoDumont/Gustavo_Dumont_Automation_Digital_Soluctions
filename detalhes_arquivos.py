from pathlib import Path
import csv, mimetypes, os, sys
from common import dt, format_size, sha256

def inventory(root: Path, report: Path, hashes=False):
    rows=[]
    for p in root.rglob('*'):
        if not p.is_file(): continue
        try:
            s=p.stat(); mime=mimetypes.guess_type(p.name)[0] or ''
            rows.append([str(p.relative_to(root)),p.name,p.suffix.lower(),mime,s.st_size,format_size(s.st_size),dt(s.st_ctime),dt(s.st_mtime), 'R' if os.access(p,os.R_OK) else '-', 'W' if os.access(p,os.W_OK) else '-', sha256(p) if hashes else ''])
        except (OSError,PermissionError) as e: rows.append([str(p),p.name,p.suffix.lower(),'','','','','','','',f'ERRO: {e}'])
    with report.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.writer(f,delimiter=';'); w.writerow(['caminho','nome','extensao','mime','bytes','tamanho','criado','modificado','leitura','escrita','sha256']); w.writerows(rows)
if __name__=='__main__': inventory(Path(sys.argv[1]),Path(sys.argv[2]),'--hash' in sys.argv)

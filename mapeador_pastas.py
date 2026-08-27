from pathlib import Path
import csv, sys
from common import dt, format_size

def map_folder(root: Path, report: Path):
    rows=[]
    for p in [root,*root.rglob('*')]:
        try:
            s=p.stat(); rows.append([str(p.relative_to(root)) or '.', 'PASTA' if p.is_dir() else 'ARQUIVO', p.suffix.lower(), s.st_size if p.is_file() else '', format_size(s.st_size) if p.is_file() else '', dt(s.st_mtime), 'SIM' if p.is_dir() and not any(p.iterdir()) else 'NAO'])
        except (OSError,PermissionError) as e: rows.append([str(p),'ERRO','','','','',str(e)])
    with report.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.writer(f,delimiter=';'); w.writerow(['caminho_relativo','tipo','extensao','bytes','tamanho','modificado','pasta_vazia']); w.writerows(rows)
if __name__=='__main__': map_folder(Path(sys.argv[1]),Path(sys.argv[2]))

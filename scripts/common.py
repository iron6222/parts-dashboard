import json, math, re
from datetime import datetime, date
from pathlib import Path
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
ALIASES={'出貨誤品':'誤欠','出貨損品':'損品','包裝品質':'包裝','增援運費｜DLR(北區)':'DLR（北區）','增援運費｜SDTC':'SDTC','增援運費｜SDHN':'SDHN'}

def number(v):
    if v is None or isinstance(v,bool): return None
    if isinstance(v,(int,float)): return float(v) if math.isfinite(v) else None
    try: return float(str(v).replace(',','').replace('%','').strip())
    except (ValueError,TypeError): return None

def row(year,month,warehouse,category,kpi,actual,target=None,unit='',direction='INFO',source='',sheet='',cell=''):
    return dict(year=int(year),month=int(month),warehouse=warehouse,category=category,kpi=ALIASES.get(kpi,kpi),actual=number(actual),target=number(target),unit=unit,direction=direction,note='正式資料',source=source,sheet=sheet,cell=cell)

def files(area): return sorted((ROOT/'raw'/area).glob('*.xlsx'))

def workbook(path): return load_workbook(path,data_only=True,read_only=True)

def year_of_sheet(name,path):
    match=re.search(r'20\d{2}',str(name)) or re.search(r'20\d{2}',path.stem)
    return int(match.group()) if match else None

def canonical_key(r): return (r.get('year'),r.get('month'),r.get('warehouse'),r.get('category'),r.get('kpi'),r.get('item',''))

def dedupe(rows):
    output={}
    for r in rows:
        if r.get('actual') is not None: output[canonical_key(r)]=r
    return list(output.values())

def load_json(path,default):
    try: return json.loads(path.read_text('utf-8'))
    except FileNotFoundError: return default

def save_json(path,doc):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(doc,ensure_ascii=False,indent=2,allow_nan=False)+'\n',encoding='utf-8')

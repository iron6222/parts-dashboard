"""Extract official ※1全商品（月） monthly inventory months from KPI workbook."""
from common import *

def _norm(x): return re.sub(r'\s+','',str(x or '')).replace('　','')

def convert(path):
    path=Path(path); wb=workbook(path); out=[]
    for ws in wb.worksheets:
        m=re.fullmatch(r'\s*(20\d{2})年?\s*',ws.title)
        if not m:continue
        year=int(m.group(1));header=None;actual_row=None;target_row=None
        for cells in ws.iter_rows(min_row=1,max_row=min(ws.max_row,60),max_col=min(ws.max_column,36)):
            vals=[_norm(c.value) for c in cells]
            if '1月' in vals and '2月' in vals:header=cells[0].row
            if any('※1全商品(月)' in v or '※1全商品（月）' in v for v in vals) and any('庫存月數' in _norm(ws.cell(rr,2).value) for rr in range(max(1,cells[0].row-3),cells[0].row+1)):
                actual_row=cells[0].row
                if _norm(ws.cell(actual_row+1,9).value)=='目標':target_row=actual_row+1
                break
        if not header or not actual_row:continue
        months={}
        for col in range(1,ws.max_column+1):
            h=_norm(ws.cell(header,col).value)
            mt=re.fullmatch(r'(1[0-2]|[1-9])月',h)
            if mt:months[int(mt.group(1))]=col
        for month,col in sorted(months.items()):
            val=number(ws.cell(actual_row,col).value)
            if val is None:continue
            if not 0<val<5:raise ValueError(f'{path.name} {year}/{month} 庫存月數異常：{val}')
            target=number(ws.cell(target_row,col).value) if target_row else None
            out.append(dict(year=year,month=month,warehouse='訂貨管理',category='訂貨KPI',kpi='庫存月數',item='全商品',actual=val,target=target,last=None,unit='月',direction='LOW',note='※1全商品（月）正式月結',source=path.name,sheet=ws.title,cell=ws.cell(actual_row,col).coordinate))
    if not out:raise ValueError(f'{path.name}：找不到年度工作表「庫存月數／※1全商品（月）」')
    return out

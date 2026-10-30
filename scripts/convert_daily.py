from common import *
# 只轉換已明確辨識日期與指標欄位的日報資料；月結另行核對報表年月與資料年月。
KEYS={'庫存金額':'庫存金額日推移','全省供應率':'全省供應率日推移','BO點數':'BO點數日推移','期貨金額':'期貨金額日推移','A類庫存':'A類庫存日推移'}
def convert(path):
    wb=workbook(path); out=[]
    for ws in wb:
        for header_index,header in enumerate(ws.iter_rows(min_row=1,max_row=min(ws.max_row,30),values_only=True),start=1):
            h=[str(x or '').strip() for x in header]
            di=next((i for i,s in enumerate(h) if s in ('日期','資料日期','date')),None)
            if di is None:continue
            metrics=[(i,k) for i,s in enumerate(h) for label,k in KEYS.items() if s==label]
            if not metrics:continue
            for cells in ws.iter_rows(min_row=header_index+1,values_only=True):
                raw=cells[di] if di<len(cells) else None
                if isinstance(raw,(datetime,date)):dt=raw
                else:
                    try:dt=datetime.strptime(str(raw).split(' ')[0].replace('-','/'),'%Y/%m/%d')
                    except ValueError:continue
                for i,k in metrics:
                    v=number(cells[i]) if i<len(cells) else None
                    if v is None:continue
                    out.append(dict(year=dt.year,month=dt.month,kpi=k,item=f'{dt.year}/{dt.month}/{dt.day}',actual=v,target=None,last=None,unit='%' if '供應率' in k else '點' if '點數' in k else '元',source=path.name))
            break
    if not out:raise ValueError(f'{path.name}：找不到已確認的「日期＋日報指標」標準欄位；不得將空白或不明資料覆蓋正式日報。')
    return dedupe(out)

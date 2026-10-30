"""需給日報橫向日期版轉換器。放入 scripts/convert_daily.py 使用。"""
from common import *

# 以原始報表固定區塊與列名核對，避免抓錯庫存金額項目。
def _date(v):
    if isinstance(v,(datetime,date)):
        return v.date() if isinstance(v,datetime) else v
    if v is None: return None
    try: return datetime.strptime(str(v).strip()[:10].replace('-','/'),'%Y/%m/%d').date()
    except ValueError: return None

def _dates(ws,header_row,start_col):
    return [(col,_date(ws.cell(header_row,col).value)) for col in range(start_col,ws.max_column+1)
            if _date(ws.cell(header_row,col).value)]

def _metric(out,ws,header_row,start_col,data_row,kpi,unit,source,cutoff,validator=None):
    for col,dt in _dates(ws,header_row,start_col):
        if dt>cutoff: continue
        v=number(ws.cell(data_row,col).value)
        if v is None: continue
        if validator and not validator(v): continue
        out.append(dict(year=dt.year,month=dt.month,kpi=kpi,
            item=f'{dt.year}/{dt.month}/{dt.day}',actual=v,target=None,last=None,
            unit=unit,source=source))

def _check(ws,row,col,label):
    v=str(ws.cell(row,col).value or '').replace(' ','').replace('　','')
    if label not in v:
        raise ValueError(f'{ws.title}!{ws.cell(row,col).coordinate} 應為「{label}」，實際是「{v}」；報表格式可能已改版')

def convert(path):
    path=Path(path)
    wb=workbook(path)
    for name in ('當月份','附件一(A庫存金額)','附件二(期貨)'):
        if name not in wb.sheetnames: raise ValueError(f'{path.name}：缺少工作表「{name}」')
    main=wb['當月份']; a=wb['附件一(A庫存金額)']; future=wb['附件二(期貨)']
    _check(main,5,1,'管理全商品')
    _check(main,6,2,'主商品')
    _check(main,13,2,'其他')
    _check(main,18,2,'LEXUS零件')
    _check(main,26,3,'二通路零件')
    _check(main,47,1,'TOYOTA全商品')
    _check(main,96,1,'TOYOTA全商品')
    _check(a,9,1,'總計')
    _check(future,19,1,'合計')
    _check(future,26,1,'合計')
    # 以管理全商品最後一個正值日期決定正式截止日，排除尚未更新的未來零值。
    valid=[dt for col,dt in _dates(main,3,9) if (number(main.cell(5,col).value) or 0)>0]
    if not valid: raise ValueError(f'{path.name}：找不到管理全商品有效日期')
    cutoff=max(valid)
    out=[]
    # 庫存金額：依使用者指定口徑，只取「主商品＋其他」。
    # 不包含 LEXUS 零件、二通路零件，也不採管理全商品總額。
    for col,dt in _dates(main,3,9):
        if dt>cutoff: continue
        primary=number(main.cell(6,col).value)
        other=number(main.cell(13,col).value)
        if primary is None or other is None: continue
        if primary<0 or other<0: continue
        value=primary+other
        if value<=0: continue
        out.append(dict(year=dt.year,month=dt.month,kpi='庫存金額日推移',
            item=f'{dt.year}/{dt.month}/{dt.day}',actual=value,target=None,last=None,
            unit='元',source=path.name))
    _metric(out,main,38,9,47,'全省供應率日推移','%',path.name,cutoff,lambda x:0<=x<=100)
    _metric(out,main,87,9,96,'BO點數日推移','點',path.name,cutoff,lambda x:x>=0)
    _metric(out,a,4,4,9,'A類庫存日推移','元',path.name,cutoff,lambda x:x>0)
    # 期貨金額：TOYOTA 合計 + LEXUS 合計；兩區分開列示，不可只取第一個合計。
    for col,dt in _dates(future,4,3):
        if dt>cutoff: continue
        x=number(future.cell(19,col).value);y=number(future.cell(26,col).value)
        if x is None or y is None: continue
        v=x+y
        if v<=0: continue
        out.append(dict(year=dt.year,month=dt.month,kpi='期貨金額日推移',
            item=f'{dt.year}/{dt.month}/{dt.day}',actual=v,target=None,last=None,
            unit='元',source=path.name))
    wb.close()
    if not out: raise ValueError(f'{path.name}：無有效日報資料')
    # daily 資料每一日須保留，不可使用 common.dedupe()（其 key 未包含日期）。
    unique={}
    for r in out: unique[(r['year'],r['month'],r['kpi'],r['item'])]=r
    print(f'{path.name}：正式資料截止 {cutoff}，取得 {len(unique)} 筆日報 KPI')
    return list(unique.values())

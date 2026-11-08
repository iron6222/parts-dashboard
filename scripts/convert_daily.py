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
    dated=[(col,dt,number(ws.cell(data_row,col).value)) for col,dt in _dates(ws,header_row,start_col) if dt<=cutoff]
    # 月內尚未出現有效非零實績時，整月的預設零值不視為已發布資料。
    # 尚未有正值前的零通常是報表預設值；正式開工後的零才保留。
    first_valid={}
    for col,dt,v in dated:
        if v is not None and v>0 and (validator is None or validator(v)):
            key=(dt.year,dt.month)
            first_valid[key]=min(first_valid.get(key,dt),dt)
    for col,dt,v in dated:
        if dt<first_valid.get((dt.year,dt.month),date.max): continue
        if kpi=='全省供應率日推移' and dt.weekday()==6 and v==0: continue
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
    # 第二層每日商品別：與第一層使用相同工作表、日期及截止日。
    # 以第一層已發布的日期為準，避免 Excel 未來日期的預設 0 被誤認為實績。
    detail_specs=(
        ('全省供應率日推移','全省供應率日明細',38,range(40,47),'%'),
        ('BO點數日推移','BO點數日明細',87,range(89,96),'點'),
    )
    for summary_kpi,detail_kpi,header,product_rows,unit in detail_specs:
        published={(r['year'],r['month'],int(r['item'].split('/')[-1])) for r in out if r['kpi']==summary_kpi}
        for col,dt in _dates(main,header,9):
            if dt>cutoff or (dt.year,dt.month,dt.day) not in published:continue
            for row in product_rows:
                item=' '.join(str(main.cell(row,1).value or '').split())
                value=number(main.cell(row,col).value)
                if not item or value is None:continue
                if unit=='%' and not 0<=value<=100:continue
                if unit=='點' and value<0:continue
                out.append(dict(year=dt.year,month=dt.month,day=dt.day,
                    kpi=detail_kpi,item=item,actual=value,target=None,last=None,
                    unit=unit,source=path.name))

    # 庫存金額第二層：主商品及其他八個商品別，與總額相同的正式日期。
    inventory_rows=(7,8,9,10,12,15,16,17)  # 排除 C 日野及 E 噴射泵
    for col,dt in _dates(main,3,9):
        if dt>cutoff: continue
        for row in inventory_rows:
            v=number(main.cell(row,col).value)
            if v is None or v<0: continue
            item=' '.join(str(main.cell(row,3).value or '').split())
            if not item: continue
            out.append(dict(year=dt.year,month=dt.month,day=dt.day,kpi='庫存金額日明細',
                item=item,actual=v,target=number(main.cell(row,5).value),
                last=number(main.cell(row,4).value),unit='元',source=path.name))
    # A 類庫存來源 1/3/4：目標位於 C 欄，日期由 D 欄開始。
    for col,dt in _dates(a,4,4):
        if dt>cutoff: continue
        for row in (6,7,8):
            imp=str(a.cell(row,2).value or '').strip()
            if imp not in ('1','3','4'): continue
            v=number(a.cell(row,col).value)
            if v is None or v<0: continue
            out.append(dict(year=dt.year,month=dt.month,day=dt.day,kpi='A類庫存日明細',
                item='來源'+imp,actual=v,target=number(a.cell(row,3).value),last=None,
                unit='元',source=path.name))
    # 期貨第一區 TOYOTA 商品別+進口別；零值為有效值，不可略過。
    for col,dt in _dates(future,4,3):
        if dt>cutoff: continue
        for row in range(6,19):
            prod=str(future.cell(row,1).value or '').strip()
            imp=str(future.cell(row,2).value or '').strip()
            if not prod or not imp: continue
            v=number(future.cell(row,col).value)
            if v is None or v<0: continue
            out.append(dict(year=dt.year,month=dt.month,day=dt.day,kpi='期貨金額日明細',
                item=f'{prod}-{imp}',actual=v,target=None,last=None,unit='元',source=path.name))
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
    for r in out: unique[(r['year'],r['month'],r.get('day'),r['kpi'],r['item'])]=r
    print(f'{path.name}：正式資料截止 {cutoff}，取得 {len(unique)} 筆日報 KPI')
    return list(unique.values())

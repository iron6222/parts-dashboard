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
    """Read dates, resolving same-workbook =日期!C2 references when cached dates are invalid."""
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    from openpyxl import load_workbook
    from pathlib import Path
    import re
    result=[]
    formula_ws=None
    if getattr(ws, '_daily_source_path', None):
        formula_ws=ws._daily_formula_sheet
    for col in range(start_col,ws.max_column+1):
        dt=_date(ws.cell(header_row,col).value)
        if dt is None and formula_ws is not None:
            formula=formula_ws.cell(header_row,col).value
            match=re.fullmatch(r"=['\"]?日期['\"]?!\$?([A-Z]+)\$?(\d+)",str(formula or '').strip(),re.I)
            if match:
                ref=f'{match.group(1)}{match.group(2)}'
                dt=_date(ws._daily_date_sheet[ref].value)
        if dt is not None:result.append((col,dt))
    return result

def _metric(out,ws,header_row,start_col,data_row,kpi,unit,source,cutoff,validator=None,target_by_month=None):
    dated=[(col,dt,number(ws.cell(data_row,col).value)) for col,dt in _dates(ws,header_row,start_col) if dt<=cutoff]
    # 月內尚未出現有效非零實績時，整月的預設零值不視為已發布資料。
    # 尚未有正值前的零通常是報表預設值；正式開工後的零才保留。
    first_valid={}
    for col,dt,v in dated:
        if v is not None and v>0 and (validator is None or validator(v)):
            key=(dt.year,dt.month)
            first_valid[key]=min(first_valid.get(key,dt),dt)
    for col,dt,v in dated:
        # 供應率 0% 是已發布日期的有效資料（如休假日），不可跳過。
        if kpi!='全省供應率日推移' and dt<first_valid.get((dt.year,dt.month),date.max): continue
        if v is None: continue
        if validator and not validator(v): continue
        out.append(dict(year=dt.year,month=dt.month,kpi=kpi,
            item=f'{dt.year}/{dt.month}/{dt.day}',actual=v,
            target=(target_by_month or {}).get((dt.year,dt.month)),last=None,
            unit=unit,source=source))

def _check(ws,row,col,label):
    v=str(ws.cell(row,col).value or '').replace(' ','').replace('　','')
    if label not in v:
        raise ValueError(f'{ws.title}!{ws.cell(row,col).coordinate} 應為「{label}」，實際是「{v}」；報表格式可能已改版')

def convert(path):
    path=Path(path)
    wb=load_workbook(path,data_only=True,read_only=False)
    for name in ('當月份','附件一(A庫存金額)','附件二(期貨)'):
        if name not in wb.sheetnames: raise ValueError(f'{path.name}：缺少工作表「{name}」')
    main=wb['當月份']; a=wb['附件一(A庫存金額)']; future=wb['附件二(期貨)']
    if '日期' in wb.sheetnames:
        main._daily_source_path=path
        main._daily_date_sheet=wb['日期']
        main._daily_formula_book=load_workbook(path,read_only=False,data_only=False)
        main._daily_formula_sheet=main._daily_formula_book['當月份']
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
    # For dated daily files, the official cutoff is the day before the report date.
    # Do not publish future/preliminary values even if cells contain nonzero formulas.
    import re
    from datetime import timedelta
    match=re.search(r'需給日報-(\d{1,2})-(\d{1,2})(?:\s*\(\d+\))?$',path.stem)
    if match:
        month,day=map(int,match.groups()[:2])
        report_dates=[]
        for year in (cutoff.year-1,cutoff.year,cutoff.year+1):
            try: report_dates.append(date(year,month,day))
            except ValueError: pass
        report_date=min(report_dates,key=lambda d:abs((d-cutoff).days))
        expected=report_date-timedelta(days=1)
        if expected not in valid:
            raise ValueError(f'{path.name}：缺少報表日前一天 {expected} 的有效資料')
        cutoff=expected
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
    # 供應率目標：當月份表第 38 列的 YYYY/M 對應第 47 列 TOYOTA全商品。
    # 每月各自取目標，絕不沿用上月或硬編碼。
    supply_targets={}
    for col in range(1,9):
        label=str(main.cell(38,col).value or '').strip()
        match=__import__('re').fullmatch(r'(20\d{2})/(\d{1,2})',label)
        target=number(main.cell(47,col).value)
        if match and target is not None and 0<target<=100:
            supply_targets[(int(match.group(1)),int(match.group(2)))]=target
    _metric(out,main,38,9,47,'全省供應率日推移','%',path.name,cutoff,lambda x:0<=x<=100,target_by_month=supply_targets)
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
                    kpi=detail_kpi,item=item,actual=value,target=(number(main.cell(row,5).value) if unit=='%' else None),last=None,
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
    future_product_names={'A':'A 豐田','N':'N 用品補修件'}
    for col,dt in _dates(future,4,3):
        if dt>cutoff: continue
        for row in range(6,19):
            prod=str(future.cell(row,1).value or '').strip()
            imp=str(future.cell(row,2).value or '').strip()
            if not prod or not imp: continue
            v=number(future.cell(row,col).value)
            if v is None or v<0: continue
            out.append(dict(year=dt.year,month=dt.month,day=dt.day,kpi='期貨金額日明細',
                item=f'{future_product_names.get(prod.upper(),prod)} / 來源{imp}',actual=v,target=None,last=None,unit='元',source=path.name))
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

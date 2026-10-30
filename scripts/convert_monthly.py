"""需給月報轉換器：只擷取已結算月份；不以 0 補缺值。
放在 scripts/convert_monthly.py，原有 daily 不變。
"""
from common import *
import re

MONTHLY_KEYS = ('庫存金額', '庫存月數', '全省供應率', 'BO點數')

def _clean(v):
    return re.sub(r'\s+', '', str(v or '')).replace('　', '')

def _section(ws, title):
    for row in ws.iter_rows(min_col=1, max_col=1):
        c=row[0]
        if title in _clean(c.value): return c.row
    raise ValueError(f'{ws.title} 找不到區段：{title}')

def _find(ws, start, end, name, col):
    for i in range(start+1, end):
        if _clean(ws.cell(i,col).value)==name:
            return i
    raise ValueError(f'{ws.title} 區段 {start} 找不到 {name}')

def _layout(ws):
    avg=_section(ws,'一、庫存金額(平均)')
    end=_section(ws,'二、庫存金額(期末)')
    months=_section(ws,'三、庫存月數')
    supply=_section(ws,'四、全省供應率')
    bo=_section(ws,'六、BO點數')
    return {
      'primary':_find(ws,avg,end,'主商品',2),
      'other':_find(ws,avg,end,'其他',2),
      'months':_find(ws,months,supply,'管理全商品',1),
      'supply':_find(ws,supply,bo,'TOYOTA全商品',1),
      'bo':_find(ws,bo,ws.max_row+1,'TOYOTA全商品',1),
    }

def _value(ws,layout,month,key):
    c=month+3
    if key=='庫存金額':
        a=number(ws.cell(layout['primary'],c).value)
        b=number(ws.cell(layout['other'],c).value)
        return a+b if a is not None and b is not None and a+b>0 else None
    pos={'庫存月數':'months','全省供應率':'supply','BO點數':'bo'}[key]
    v=number(ws.cell(layout[pos],c).value)
    if v is None or v<=0:return None
    if key=='全省供應率' and v>100:return None
    return v

def convert(path):
    path=Path(path); wb=workbook(path)
    try:
        if not all(n in wb.sheetnames for n in ('當年月推移','去年月推移')):
            raise ValueError(f'{path.name}：缺少當年月推移／去年月推移')
        current=wb['當年月推移']; previous=wb['去年月推移']
        lc=_layout(current); lp=_layout(previous)
        m=re.search(r'20\d{2}[-_/](\d{1,2})',path.stem)
        y=re.search(r'20\d{2}',path.stem)
        if not y:raise ValueError(f'{path.name}：檔名缺少年份')
        year=int(y.group()); out=[]
        for yr,ws,lay in ((year-1,previous,lp),(year,current,lc)):
            for month in range(1,13):
                for key in MONTHLY_KEYS:
                    actual=_value(ws,lay,month,key)
                    if actual is None:continue
                    last=_value(previous,lp,month,key) if yr==year else None
                    out.append(dict(year=yr,month=month,warehouse='訂貨管理',category='訂貨KPI',
                        kpi=key,item='全商品',actual=actual,target=None,last=last,
                        unit={'庫存金額':'元','庫存月數':'月','全省供應率':'%','BO點數':'點'}[key],
                        source=path.name,sheet=ws.title))
        # 月份有效性依 KPI 各自的正式數值判定，不能由日報推估月結。
        if not out:raise ValueError(f'{path.name}：沒有可用的月結資料')
        print(f'{path.name}：解析 {len(out)} 筆月結 KPI')
        return out
    finally:wb.close()

# V2: 日報附件的月平均、來源別與庫存商品別月結明細。
# 注意：檔名代表報表產出月份，附件日期才是實際資料月份。
from datetime import datetime, date

def _date(v):
    if isinstance(v,(datetime,date)): return (v.year,v.month)
    m=re.search(r'(20\d{2})[/-](\d{1,2})[/-]\d{1,2}',str(v or ''))
    return (int(m.group(1)),int(m.group(2))) if m else None

def _attachment_period(ws,header_row=4):
    periods={_date(ws.cell(header_row,c).value) for c in range(3,min(ws.max_column,42)+1)}
    periods.discard(None)
    if len(periods)!=1: raise ValueError(f'{ws.title} 附件日期跨月份或缺少日期：{periods}')
    return periods.pop()

def _avg_column(ws,header_row=4):
    for c in range(3,min(ws.max_column,70)+1):
        if _clean(ws.cell(header_row,c).value).lower() in ('average','平均','月平均'):
            return c
    raise ValueError(f'{ws.title} 找不到 Average 欄')

def _detail_rows(ws,yr,mo,source):
    """只輸出 TOYOTA 主商品與其他的商品別；不混入 LEXUS。"""
    avg=_section(ws,'一、庫存金額(平均)')
    end=_section(ws,'二、庫存金額(期末)')
    out=[]
    group=None
    for r in range(avg+1,end):
        group_name=_clean(ws.cell(r,2).value)
        if group_name=='主商品':group='main';continue
        if group_name=='其他':group='other';continue
        if group_name.startswith('LEXUS'):group='lexus';continue
        if group not in ('main','other'):continue
        label=str(ws.cell(r,3).value or '').strip()
        if not re.match(r'^[A-Z0-9]\s',label):continue
        actual=number(ws.cell(r,mo+3).value)
        if actual is None or actual<=0:continue
        out.append(dict(year=yr,month=mo,warehouse='訂貨管理',category='訂貨KPI',
                        kpi='庫存金額明細',item=' '.join(label.split()),actual=actual,
                        target=None,last=None,unit='元',source=source,sheet=ws.title))
    return out

def _monthly_attachments(wb,path):
    out=[]
    ac=wb['附件一(A庫存金額)'];future=wb['附件二(期貨)']
    ap=_attachment_period(ac);fp=_attachment_period(future)
    if ap!=fp:raise ValueError(f'{path.name} A類與期貨附件日期月份不一致：{ap} vs {fp}')
    yr,mo=ap
    ac_avg=_avg_column(ac);future_avg=_avg_column(future)
    # 以來源 1/3/4 的 Average 欄作月結實績；不拿目標欄當實績。
    target_ws=wb['A類各進口別庫存金額目標'] if 'A類各進口別庫存金額目標' in wb.sheetnames else None
    targets={}
    if target_ws:
        # 年度目標表：第 2 列為月份數字；第 5~7 列分別為 A1/A3/A4。
        target_col=next((c for c in range(2,min(target_ws.max_column,16)+1)
                         if number(target_ws.cell(2,c).value)==mo),None)
        if target_col:
            for r in range(1,min(target_ws.max_row,20)+1):
                label=_clean(target_ws.cell(r,1).value).upper()
                match=re.match(r'^A([134])(?:TMC|KZ|TMAP|$)',label)
                if match:
                    v=number(target_ws.cell(r,target_col).value)
                    if v is not None and v>0:targets[match.group(1)]=v
    for r in range(6,9):
        imp=str(ac.cell(r,2).value or '').strip()
        if imp not in ('1','3','4'):continue
        v=number(ac.cell(r,ac_avg).value)
        if v is None or v<=0:continue
        out.append(dict(year=yr,month=mo,warehouse='訂貨管理',category='訂貨KPI',
                        kpi='A類庫存',item='進口別 '+imp,actual=v,target=targets.get(imp),
                        last=None,unit='元',source=path.name,sheet=ac.title))
    # 第一個「合計」以前是 TOYOTA 期貨，後方為 LEXUS，不可合併。
    for r in range(6,future.max_row+1):
        prod=str(future.cell(r,1).value or '').strip()
        if prod=='合計':break
        imp=str(future.cell(r,2).value or '').strip()
        if not prod or not imp:continue
        v=number(future.cell(r,future_avg).value)
        if v is None or v<=0:continue
        out.append(dict(year=yr,month=mo,warehouse='訂貨管理',category='訂貨KPI',
                        kpi='期貨金額',item=f'{prod}-{imp}',actual=v,target=None,
                        last=None,unit='元',source=path.name,sheet=future.title))
    return out

# 保留原先月報四項 KPI 解析，新增附件兩項與庫存第二層。
_convert_base=convert
def convert(path):
    path=Path(path)
    base=_convert_base(path)
    wb=workbook(path)
    try:
        attachments=_monthly_attachments(wb,path)
        year_match=re.search(r'20\d{2}',path.stem)
        if not year_match:raise ValueError(f'{path.name} 缺少年份')
        year=int(year_match.group())
        details=[]
        for yr,sheet_name in ((year-1,'去年月推移'),(year,'當年月推移')):
            if sheet_name not in wb.sheetnames:continue
            ws=wb[sheet_name]
            for mo in range(1,13):
                details.extend(_detail_rows(ws,yr,mo,path.name))
        # 全省供應率與 BO 商品別：從同一月推移的實際商品列擷取，不把總計混入。
        for yr,sheet_name in ((year-1,'去年月推移'),(year,'當年月推移')):
            if sheet_name not in wb.sheetnames:continue
            ws=wb[sheet_name]
            supply_start=_section(ws,'四、全省供應率')
            supply_end=_section(ws,'五、主商品高銷件全省供應率')
            bo_start=_section(ws,'六、BO點數')
            for key,start,end in (('全省供應率明細',supply_start,supply_end),('BO點數明細',bo_start,ws.max_row+1)):
                for r in range(start+1,end):
                    label=str(ws.cell(r,1).value or '').strip()
                    m=re.match(r'^([A-Z0-9])\s+(.+)',label)
                    if not m or m.group(1) not in ('A','J','O','N','H','F','U','6'):continue
                    item=m.group(1)+' '+re.sub(r'\s+','',m.group(2))
                    for mo in range(1,13):
                        v=number(ws.cell(r,mo+3).value)
                        if v is None or v<=0 or (key=='全省供應率明細' and v>100):continue
                        details.append(dict(year=yr,month=mo,warehouse='訂貨管理',category='訂貨KPI',
                          kpi=key,item=item,actual=v,target=None,last=None,
                          unit='%' if key=='全省供應率明細' else '點',source=path.name,sheet=sheet_name))
        # V6｜讀取月報「當月份」E欄各商品當月正式供應率目標。
        # 工作表 E38 實際年月可能比檔名早一個月，嚴格以儲存格為準。
        if '當月份' in wb.sheetnames:
            target_ws=wb['當月份']
            header=str(target_ws.cell(38,5).value or '')
            match=re.search(r'(20\d{2})\s*[/.-]\s*(\d{1,2})',header)
            if match:
                ty,tm=map(int,match.groups())
                for r in range(40,47):
                    label=str(target_ws.cell(r,1).value or '').strip()
                    code_match=re.match(r'^([A-Z0-9])\s+(.+)',label)
                    if not code_match or code_match.group(1) not in ('A','J','O','N','H','F','U','6'):continue
                    raw=target_ws.cell(r,5).value
                    target=number(raw)
                    if target is None or target<=0:continue
                    if target<=1:target*=100
                    if target>100:continue
                    item=code_match.group(1)+' '+re.sub(r'\s+','',code_match.group(2))
                    details.append(dict(year=ty,month=tm,warehouse='訂貨管理',category='訂貨KPI',
                        kpi='全省供應率商品目標',item=item,actual=target,target=target,last=None,
                        unit='%',source=path.name,sheet='當月份'))
        # 保留正式非零月份，跨檔去重由 build.py 執行。
        return base+attachments+details
    finally:wb.close()


# V8: 從每份歷史月報「當月份」及附件提取真正逐日數據。
def convert_daily_history(path):
    """歷史日報：主層五項及第二層商品/來源逐日明細，單次讀取工作表。"""
    from datetime import datetime, date
    path=Path(path); wb=workbook(path)
    try:
        required=('當月份','附件一(A庫存金額)','附件二(期貨)')
        if any(n not in wb.sheetnames for n in required):return []
        # read_only 模式不可重複 ws.cell() 掃描，否則 22 份會極慢。
        sheets={n:list(wb[n].iter_rows(values_only=True)) for n in required}
        out=[]
        def cell(a,r,c):
            return a[r-1][c-1] if r<=len(a) and c<=len(a[r-1]) else None
        def dates(a,r,start):
            for c in range(start,len(a[r-1])+1):
                v=cell(a,r,c)
                if isinstance(v,(datetime,date)):dt=v
                else:
                    m=re.search(r'(20\d{2})[/-](\d{1,2})[/-](\d{1,2})',str(v or ''))
                    if not m:continue
                    try:dt=date(*map(int,m.groups()))
                    except ValueError:continue
                yield c,dt
        def add(dt,kpi,item,v,unit,target=None):
            if v is None or not (0<=v<=100 if unit=='%' else v>=0):return
            out.append(dict(year=dt.year,month=dt.month,day=dt.day,
                warehouse='訂貨管理',category='訂貨KPI',kpi=kpi,item=item,
                actual=v,target=target if target is not None and target>0 else None,
                last=None,unit=unit,source=path.name))
        main=sheets['當月份'];ac=sheets['附件一(A庫存金額)'];fut=sheets['附件二(期貨)']
        inv={7:'A 豐田',8:'J 油料',9:'O 工具對策',10:'N 用品補修件',11:'C 日野',
             12:'H 輪胎',14:'E 噴射泵',15:'F 電裝',16:'U 咖啡',17:'6 國採用品零件'}
        supply={40:'A 豐田',41:'J 油料',42:'O 工具對策',43:'N 用品補修件',
                44:'6 國採用品零件',45:'H 輪胎'}
        bo={89:'A 豐田',90:'J 油料',91:'O 工具對策',92:'N 用品補修件',
            93:'6 國採用品零件',94:'H 輪胎'}
        for c,dt in dates(main,3,9):
            a=number(cell(main,6,c));b=number(cell(main,13,c))
            if a is not None and b is not None:add(dt,'庫存金額日推移',f'{dt.year}/{dt.month}/{dt.day}',a+b,'元')
            for r,name in inv.items():add(dt,'庫存金額日明細',name,number(cell(main,r,c)),'元',number(cell(main,r,5)))
        for c,dt in dates(main,38,9):
            add(dt,'全省供應率日推移',f'{dt.year}/{dt.month}/{dt.day}',number(cell(main,47,c)),'%')
            for r,name in supply.items():add(dt,'全省供應率日明細',name,number(cell(main,r,c)),'%',number(cell(main,r,5)))
        for c,dt in dates(main,87,9):
            add(dt,'BO點數日推移',f'{dt.year}/{dt.month}/{dt.day}',number(cell(main,96,c)),'點')
            for r,name in bo.items():add(dt,'BO點數日明細',name,number(cell(main,r,c)),'點')
        for c,dt in dates(ac,4,4):
            add(dt,'A類庫存日推移',f'{dt.year}/{dt.month}/{dt.day}',number(cell(ac,9,c)),'元')
            for r,name in {6:'來源1',7:'來源3',8:'來源4'}.items():
                add(dt,'A類庫存日明細',name,number(cell(ac,r,c)),'元',number(cell(ac,r,3)))
        for c,dt in dates(fut,4,3):
            add(dt,'期貨金額日推移',f'{dt.year}/{dt.month}/{dt.day}',number(cell(fut,19,c)),'元')
            for r,name in {6:'A 豐田 / 來源1',7:'A 豐田 / 來源3',8:'A 豐田 / 來源4',17:'N 用品補修件 / 來源3'}.items():
                add(dt,'期貨金額日明細',name,number(cell(fut,r,c)),'元')
        print(f'{path.name}：取得 {len(out)} 筆逐日及第二層 KPI')
        return out
    finally:wb.close()

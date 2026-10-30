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

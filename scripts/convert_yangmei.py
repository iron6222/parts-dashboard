from common import *
# 依 2026 楊配室月報「2026年」工作表既有核對欄位；G=1月、O=9月。
METRICS={
  8:('加班','進貨加班時數','hr'),15:('加班','出貨加班時數','hr'),22:('加班','包裝加班時數','hr'),
  23:('生產力','進貨生產力','i/hr'),
  26:('生產力','出貨生產力','i/hr'),
  29:('生產力','包裝生產力','i/hr'),
  58:('庫位使用率','全區','%'),
}
# 圖表分區明細，依來源表第 48-57 列標籤擷取，不假設固定分區名稱。
def convert(path):
    wb=workbook(path); out=[]
    for ws in wb:
        year=year_of_sheet(ws.title,path)
        if not year or not re.search(r'20\d{2}年',ws.title): continue
        for col in range(7,19):
            month=col-6
            if month>12: break
            for rn,(cat,kpi,unit) in METRICS.items():
                val=number(ws.cell(rn,col).value)
                if val is None: continue
                if cat=='庫位使用率' and 0<=val<=1: val*=100
                out.append(row(year,month,'楊梅園區',cat,kpi,val,85 if cat=='庫位使用率' else (36 if rn==23 else 63.6 if rn==26 else 138 if rn==29 else None),unit,'HIGH' if cat in ('庫位使用率','生產力') else 'LOW',path.name,ws.title,f'{ws.cell(rn,col).coordinate}'))
            for rn in range(48,58):
                label=str(ws.cell(rn,2).value or ws.cell(rn,3).value or ws.cell(rn,4).value or '').strip()
                val=number(ws.cell(rn,col).value)
                if not label or val is None or label in ('None','-'):continue
                if 0<=val<=1:val*=100
                out.append(row(year,month,'楊梅園區','庫位使用率',label,val,85,'%', 'HIGH',path.name,ws.title,ws.cell(rn,col).coordinate))
    return dedupe(out)

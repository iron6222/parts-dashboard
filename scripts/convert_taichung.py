from common import *
# 只接受有明確「月份/KPI/實績」標題的標準表；不推測未核對的月報欄位。
def convert(path):
    return convert_table(path,'台中副倉')

def convert_table(path,warehouse):
    wb=workbook(path); rows=[]
    for ws in wb:
        for head in ws.iter_rows(min_row=1,max_row=min(ws.max_row,35),values_only=False):
            names=[str(c.value or '').strip() for c in head]
            def idx(*opts):
                return next((i for i,s in enumerate(names) if s in opts),None)
            mi=idx('月份','月','month'); ki=idx('KPI','指標','項目'); ai=idx('實績','實際','actual','數值')
            if None in (mi,ki,ai):continue
            yi=idx('年度','年','year');ci=idx('分類','category');ui=idx('單位','unit');ti=idx('目標','target')
            for cells in ws.iter_rows(min_row=head[0].row+1,values_only=True):
                month=number(cells[mi]) if mi<len(cells) else None
                kpi=str(cells[ki] or '').strip() if ki<len(cells) else ''
                actual=number(cells[ai]) if ai<len(cells) else None
                year=int(number(cells[yi]) or 0) if yi is not None and yi<len(cells) else year_of_sheet(ws.title,path)
                if not (year and month and 1<=month<=12 and kpi and actual is not None):continue
                rows.append(row(year,month,warehouse,str(cells[ci] or '') if ci is not None else '倉儲KPI',kpi,actual,cells[ti] if ti is not None else None,str(cells[ui] or '') if ui is not None else '',source=path.name,sheet=ws.title))
            break
    if not rows:raise ValueError(f'{path.name}：找不到已確認的「月份／KPI／實績」標準表。此原始月報格式尚需建立欄位對照，禁止猜測或發布空資料。')
    return dedupe(rows)

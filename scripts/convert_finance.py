"""Convert already-normalized finance workbook to dashboard's exact inc/expMain/expSub schema."""
from common import workbook,number

SHEETS={"收入項目":"inc","費用主表":"expMain","費用子項目":"expSub"}

def text(v, default='未分類'):
    return str(v).strip() if v is not None and str(v).strip() else default

def convert(path):
    wb=workbook(path)
    output={"inc":[],"expMain":[],"expSub":[]}
    for sheet_name,kind in SHEETS.items():
        matches=[n for n in wb.sheetnames if sheet_name in n]
        if len(matches)!=1:raise ValueError(f'{path.name}: 找不到唯一工作表 {sheet_name}')
        ws=wb[matches[0]]
        iterator=ws.iter_rows(values_only=True)
        header=[text(x,'') for x in next(iterator)]
        required={'年度','月份','部門單位','預估','實績'}
        if kind=='inc':required|={'類別/廠商','項目名稱'}
        if kind=='expMain':required|={'會計科目'}
        if kind=='expSub':required|={'主會計科目','類別/廠商','項目名稱'}
        missing=required-set(header)
        if missing:raise ValueError(f'{path.name}/{sheet_name}: 缺少欄位 {sorted(missing)}')
        for rowno,vals in enumerate(iterator,2):
            if not any(v is not None for v in vals):continue
            d=dict(zip(header,vals))
            y=number(d.get('年度'));m=number(d.get('月份'))
            if y is None or m is None:raise ValueError(f'{path.name}/{sheet_name} 第 {rowno} 列缺少年月')
            if int(y)!=y or int(m)!=m or not 2000<=y<=2100 or not 1<=m<=12:
                raise ValueError(f'{path.name}/{sheet_name} 第 {rowno} 列年月錯誤')
            pb=number(d.get('預估'));actual=number(d.get('實績'))
            if pb is None or actual is None:raise ValueError(f'{path.name}/{sheet_name} 第 {rowno} 列預估/實績非數值')
            base={'year':int(y),'month':int(m),'pb':pb,'actual':actual,'dept':text(d.get('部門單位')),'note':text(d.get('異常註記'),'' )}
            if kind=='inc':base.update(vendor=text(d.get('類別/廠商')),item=text(d.get('項目名稱')))
            if kind=='expMain':base.update(account=text(d.get('會計科目')))
            if kind=='expSub':base.update(account=text(d.get('主會計科目')),vendor=text(d.get('類別/廠商')),item=text(d.get('項目名稱')))
            output[kind].append(base)
    if not all(output.values()):raise ValueError(f'{path.name}: 三張工作表不能為空')
    return output

def validate(doc):
    for kind in SHEETS.values():
        rows=doc.get(kind,[])
        if not rows:raise ValueError(f'收支分析 {kind}: 沒有資料')
        keys=set()
        for r in rows:
            key=(r['year'],r['month'],r['dept'])
            if kind=='inc':key+=r['vendor'],r['item']
            if kind=='expMain':key+=r['account'],
            if kind=='expSub':key+=r['account'],r['vendor'],r['item']
            if key in keys:raise ValueError(f'收支分析 {kind}: 重複資料鍵 {key}')
            keys.add(key)
    return {k:len(doc[k]) for k in SHEETS.values()}

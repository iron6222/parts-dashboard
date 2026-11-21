"""楊梅月報：依分類、KPI 名稱、月份表頭解析，不依賴固定列號。"""
from common import *

# (分類, 上層標籤, 指標標籤, 輸出 KPI, 單位, 方向, 目標)
SPECS = [
 ('作業量','進貨','點數','進貨點數','i','INFO',None),
 ('作業量','進貨','PCS數','進貨PCS數','PCS','INFO',None),
 ('作業量','進貨','加班時數','進貨加班時數','hr','LOW',None),
 ('作業量','出貨','點數','出貨點數','i','INFO',None),
 ('作業量','出貨','PCS數','出貨PCS數','PCS','INFO',None),
 ('作業量','出貨','加班時數','出貨加班時數','hr','LOW',None),
 ('作業量','包裝','點數','包裝點數','i','INFO',None),
 ('作業量','包裝','PCS數','包裝PCS數','PCS','INFO',None),
 ('作業量','包裝','加班時數','包裝加班時數','hr','LOW',None),
 ('生產力',None,'進貨(i/hr)','進貨生產力','i/hr','HIGH',36),
 ('生產力',None,'中小物(i/hr)','進貨中小物生產力','i/hr','HIGH',None),
 ('生產力',None,'大物(i/hr)','進貨大物生產力','i/hr','HIGH',None),
 ('生產力',None,'出貨(i/hr)','出貨生產力','i/hr','HIGH',63.6),
 ('生產力',None,'中小物(i/hr)','出貨中小物生產力','i/hr','HIGH',None),
 ('生產力',None,'大物(i/hr)','出貨大物生產力','i/hr','HIGH',None),
 ('生產力',None,'包裝(i/hr)','包裝生產力','i/hr','HIGH',138),
 ('品質','出貨品質','誤品(ppm)','出貨誤品','ppm','LOW',9),
 ('品質','出貨品質','損品(ppm)','出貨損品','ppm','LOW',8),
 ('品質','出貨品質','包裝(ppm)','包裝品質','ppm','LOW',103),
 ('品質',None,'配送品質-準時率','配送準時率','%','HIGH',98),
 ('增援運費','金額(仟元)','DLR(北區)','增援運費｜DLR(北區)','仟元','LOW',None),
 ('增援運費','金額(仟元)','SDTC','增援運費｜SDTC','仟元','LOW',None),
 ('增援運費','金額(仟元)','SDHN','增援運費｜SDHN','仟元','LOW',None),
 ('增援運費',None,'合計','增援運費｜合計','仟元','LOW',None),
 ('庫位使用率',None,'全區','全區','%','HIGH',85),
]

def norm(v):return re.sub(r'\s+','',str(v or '')).replace('（','(').replace('）',')').lower()

def month_header(v):
    m=re.fullmatch(r'(1[0-2]|[1-9])月',norm(v))
    return int(m.group(1)) if m else None

def source_rows(ws):
    """讀取左側分組標籤；僅在同一主分類內延續父標籤。"""
    section=parent=None
    for rn in range(3,ws.max_row+1):
        a,b,c=(ws.cell(rn,i).value for i in (1,2,3))
        if norm(a):section=norm(a);parent=None
        if norm(b):parent=norm(b)
        if not norm(c) and not norm(b):continue
        yield rn,section,parent,norm(c),a,b,c



def utilization_by_month(wb, path):
    """每個月獨立選來源；缺資料即省略，不跨表補值。"""
    old_name='庫位使用率'
    new_name='庫位使用率 (202608變更)'
    result=[]
    if old_name not in wb.sheetnames or new_name not in wb.sheetnames:
        raise ValueError(f'{path.name}: 找不到庫位使用率新舊工作表')
    for sheet_name, start_month, end_month, standard in ((old_name,1,7,'202608前'),(new_name,8,12,'202608後')):
        ws=wb[sheet_name]
        # 由「合計」欄標題尋找整體使用率，不使用固定 AC/AD 欄。
        totals=[c for c in range(3,ws.max_column+1)
                if norm(ws.cell(5,c).value)=='合計' and norm(ws.cell(4,c).value)=='' and norm(ws.cell(5,c+1).value)=='▲']
        if len(totals)!=1:
            raise ValueError(f'{path.name}/{sheet_name}: 整體合計欄不唯一 {totals}')
        total_col=totals[0]
        header7=next(ws.iter_rows(min_row=7,max_row=7,values_only=True))
        current=None
        for rn, values in enumerate(ws.iter_rows(min_row=8,values_only=True),start=8):
            heading=norm(values[0])
            match=re.fullmatch(r'(20\d{2})年(1[0-2]|[1-9])月',heading)
            if match:current=(int(match.group(1)),int(match.group(2)))
            if not current or current[0]!=2026 or not start_month<=current[1]<=end_month:continue
            if norm(values[1])!='使用率':continue
            year,month=current
            v=number(values[total_col-1])
            if v is None:
                print(f'警告：{sheet_name} {year}/{month} 整體使用率缺值，不顯示',flush=True)
                continue
            if 0<=v<=2.5:v*=100
            rec=row(year,month,'楊梅園區','庫位使用率','全區',v,85,'%','HIGH',path.name,sheet_name,f'{__import__("openpyxl").utils.get_column_letter(total_col)}{rn}')
            rec['statisticsStandard']=standard
            rec['statisticsStandardChangedAt']='2026-08'
            result.append(rec)
            # 第一層分類：直接讀取原始工作表小計，避免平均庫區百分比。
            subtotal_cols = ((13,'中小物區'),(25,'大物區'),(29,'三倉合計')) if standard=='202608後' else ((16,'中小物區'),(28,'大物區'))
            for subtotal_col, subtotal_label in subtotal_cols:
                sv=number(ws.cell(rn,subtotal_col).value)
                if sv is None:continue
                if 0<=sv<=2.5:sv*=100
                subtotal=row(year,month,'楊梅園區','庫位使用率',subtotal_label,sv,85,'%','HIGH',path.name,sheet_name,f'{__import__("openpyxl").utils.get_column_letter(subtotal_col)}{rn}')
                subtotal['statisticsStandard']=standard
                subtotal['statisticsStandardChangedAt']='2026-08'
                result.append(subtotal)
            # 三倉樓層以各樓層點數 / 配置數計算；合併小計不代表 2F。
            from openpyxl.utils import get_column_letter
            if standard=='202608後':
                point_row=capacity_row=None
                for rr in range(max(8,rn-6),rn):
                    caption=norm(ws.cell(rr,2).value)
                    if caption=='點數(i)':point_row=rr
                    if caption=='配置數':capacity_row=rr
                if point_row is None or capacity_row is None:
                    print(f'警告：{sheet_name} {year}/{month} 找不到三倉樓層點數或配置數',flush=True)
                else:
                    for floor,col in ((2,26),(3,27),(4,28)):
                        points=number(ws.cell(point_row,col).value)
                        capacity=number(ws.cell(capacity_row,col).value)
                        if points is None or capacity is None or capacity<=0:
                            print(f'警告：{sheet_name} {year}/{month} 三倉 {floor}F 缺少有效數據',flush=True)
                            continue
                        detail=row(year,month,'楊梅園區','庫位使用率',f'三倉 {floor}F',points/capacity*100,85,'%','HIGH',path.name,sheet_name,f'{get_column_letter(col)}{point_row}/{get_column_letter(col)}{capacity_row}')
                        detail['statisticsStandard']=standard
                        detail['statisticsStandardChangedAt']='2026-08'
                        result.append(detail)
            # 庫區使用率以「使用率」列的實際儲存值為準。
            # 合併儲存格代表一組庫區的共同使用率，不能拆成每個庫區，也不能將 Z:AB 誤稱 2F。
            from openpyxl.utils import get_column_letter
            from openpyxl.utils.cell import range_boundaries
            excluded={total_col}
            for c in range(3,total_col):
                raw=ws.cell(rn,c).value
                fv=number(raw)
                if fv is None:continue
                # 只取中小物、大物、三倉的實際庫區；排除小計與管理狀況計數欄。
                if standard=='202608後':
                    if not (3<=c<=12 or 14<=c<=24 or 26<=c<=28):continue
                else:
                    if not (3<=c<=15 or 17<=c<=26):continue
                first=last=c
                for merged in ws.merged_cells.ranges:
                    if merged.min_row<=rn<=merged.max_row and merged.min_col<=c<=merged.max_col:
                        first,last=merged.min_col,merged.max_col
                        break
                if c!=first:continue
                codes=[str(ws.cell(6,j).value or '').strip() for j in range(first,last+1)]
                codes=[x for x in codes if x]
                if standard=='202608後' and first>=26:
                    # 合併的 Z:AB 是三倉小計，已由 AC 小計讀取；不可重複建立同名明細。
                    if last>first:continue
                    label=f'三倉 {codes[0]}F'
                else:
                    group='中小物' if (first<=12 if standard=='202608後' else first<=15) else '大物'
                    label=f'{group}區（{",".join(codes)}）'
                if 0<=fv<=2.5:fv*=100
                detail=row(year,month,'楊梅園區','庫位使用率',label,fv,85,'%','HIGH',path.name,sheet_name,f'{get_column_letter(c)}{rn}')
                detail['statisticsStandard']=standard
                detail['statisticsStandardChangedAt']='2026-08'
                result.append(detail)
    return result

def merged_value(ws, rn, col):
    """Read merged-cell top-left value; distinguish merged KPI from missing data."""
    cell=ws.cell(rn,col)
    for area in ws.merged_cells.ranges:
        if area.min_row <= rn <= area.max_row and area.min_col <= col <= area.max_col:
            return ws.cell(area.min_row,area.min_col).value, str(area)
    return cell.value, None

def convert(path):
    # Normal mode is required to inspect merged ranges; common.workbook uses read_only=True.
    from openpyxl import load_workbook
    wb=load_workbook(path,data_only=True,read_only=False);out=[]
    for ws in wb:
        year=year_of_sheet(ws.title,path)
        if not year or not re.fullmatch(r'20\d{2}年',ws.title):continue
        # 月份必須由表頭解析，且不能重複。
        headers={}
        for col in range(1,ws.max_column+1):
            month=month_header(ws.cell(2,col).value)
            if month:
                if month in headers:raise ValueError(f'{path.name}/{ws.title}: {month}月表頭重複')
                headers[month]=col
        if not headers:raise ValueError(f'{path.name}/{ws.title}: 找不到月份表頭')
        records=list(source_rows(ws));matched={}
        for cat,par,metric,kpi,unit,direction,target in SPECS:
            catkey=norm(cat);parkey=norm(par);metkey=norm(metric)
            candidates=[]
            for rn,sec,p,c,a,b,rawc in records:
                if sec!=catkey:continue
                if par is not None and p!=parkey:
                    # 2022 年增援運費未標示『金額(仟元)』父分類；僅允許此已確認的歷史差異。
                    if not (year==2022 and cat=='增援運費' and p is None):continue
                # 某些 KPI 名稱位於 B 欄（例如出貨生產力、配送準時率、全區）
                if c==metkey or (cat=='增援運費' and metric=='DLR(北區)' and year<=2024 and c=='dlr') or (not c and norm(b)==metkey):
                    candidates.append(rn)
            # 進貨/出貨中小物與大物同名：用緊鄰的生產力主指標界定區段
            if cat=='生產力' and metric in ('中小物(i/hr)','大物(i/hr)'):
                anchor='進貨(i/hr)' if '進貨' in kpi else '出貨(i/hr)'
                anchors=[rn for rn,sec,p,c,a,b,rawc in records if sec==catkey and norm(b)==norm(anchor)]
                if len(anchors)!=1:raise ValueError(f'{path.name}: 生產力分類 {anchor} 不唯一')
                candidates=[rn for rn in candidates if anchors[0]<rn<=anchors[0]+2]
            if len(candidates)!=1:
                raise ValueError(f'{path.name}/{ws.title}: {cat}/{par or ""}/{metric} 找到 {len(candidates)} 筆，需檢查格式')
            rn=candidates[0]
            for month,col in headers.items():
                raw, merged=merged_value(ws,rn,col)
                val=number(raw)
                if val is None:
                    # No value, no fabricated zero. Never report a merged continuation as missing.
                    continue
                if cat=='庫位使用率':
                    # 原月報儲存格以小數表示比率；舊報表部分值 >1 為超過 100%。
                    if 0<=val<=2.5:val*=100
                out.append(row(year,month,'楊梅園區',cat,kpi,val,target,unit,direction,path.name,ws.title,ws.cell(rn,col).coordinate))
            matched[rn]=kpi
        # Verify the two-row quality label is a single combined metric, not two independent KPIs.
        quality_rows=[rn for rn,sec,par,c,a,b,rawc in records if sec==norm('品質') and par==norm('出貨品質') and c==norm('誤品(ppm)')]
        for rn in quality_rows:
            if norm(ws.cell(rn+1,3).value)!=norm('欠品(ppm)'):
                raise ValueError(f'{path.name}/{ws.title}: 誤欠品標籤不相鄰')
            for month,col in headers.items():
                first=ws.cell(rn,col); second=ws.cell(rn+1,col)
                if second.value is not None and number(second.value)!=number(first.value):
                    raise ValueError(f'{path.name}/{ws.title}: {month}月 誤品與欠品各有不同實績，需人工確認')
                _, merged=merged_value(ws,rn+1,col)
                if first.value is not None and merged is None:
                    raise ValueError(f'{path.name}/{ws.title}: {month}月 誤欠品未合併，需人工確認')
        # 分區名稱取 C 欄，避免將 B 欄「中小物區／大物區」誤當分區。
        for rn,sec,par,c,a,b,rawc in records:
            if sec!='庫位使用率' or not c:continue
            if rn in matched:continue
            label=str(rawc).strip()
            for month,col in headers.items():
                val=number(ws.cell(rn,col).value)
                if val is None:continue
                if 0<=val<=2.5:val*=100
                out.append(row(year,month,'楊梅園區','庫位使用率',label,val,85,'%','HIGH',path.name,ws.title,ws.cell(rn,col).coordinate))
    if '2026' in path.stem or '2026年' in wb.sheetnames:
        # 以新舊明細表的正式月份資料取代年度摘要的庫位使用率，避免重複或混用。
        details=utilization_by_month(wb,path)
        out=[r for r in out if not (r['year']==2026 and r['category']=='庫位使用率')]
        out.extend(details)
    return dedupe(out)

import argparse,datetime,sys
from common import *
from convert_yangmei import convert as yangmei
from convert_taichung import convert as taichung
from convert_renwu import convert as renwu
from convert_daily import convert as daily
from convert_monthly import convert as monthly, convert_daily_history
from convert_finance import convert as finance, validate as validate_finance

PARSERS={'yangmei':yangmei,'taichung':taichung,'renwu':renwu,'daily':daily,'monthly':monthly}
def validate(rows,area):
    if not rows:raise ValueError(f'{area}: 沒有有效資料')
    for r in rows:
        if not 2000<=r['year']<=2100 or not 1<=r['month']<=12:raise ValueError(f'{area}: 年月異常 {r}')
        if r.get('actual') is None:raise ValueError(f'{area}: 實績缺值 {r}')
        if r.get('unit')=='%' and not 0<=r['actual']<=100:raise ValueError(f'{area}: 百分比異常 {r}')
    keys=[canonical_key(r) for r in rows]
    if len(keys)!=len(set(keys)):raise ValueError(f'{area}: 重複 KPI 鍵')

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--area',choices=[*PARSERS,'finance','all'],default='all');args=ap.parse_args()
    areas=list(PARSERS) if args.area=='all' else ([args.area] if args.area!='finance' else [])
    reports={}; data={}
    for area in areas:
        paths=files(area)
        if not paths:
            print(f'略過 {area}：沒有新 Excel；沿用既有 JSON');continue
        rows=[]
        accepted=[]
        rejected=[]
        if area == 'daily':
            # 單份日報故障不可拖垮其他模組；但不可讓舊日報覆蓋較新的正式資料。
            # 每份報表各自解析，僅保留成功解析且符合資料日期的版本。
            import re
            from datetime import date as _date
            candidates=[]
            for path in paths:
                print(f'解析 {path.relative_to(ROOT)}', flush=True)
                try:
                    parsed=PARSERS[area](path)
                    validate(parsed, f'daily/{path.name}')
                    dates=[]
                    for r in parsed:
                        y,m,d=map(int,r['item'].split('/'))
                        dates.append(_date(y,m,d))
                    latest=max(dates)
                    # 支援舊檔名「需給日報-DD」及新檔名「需給日報-MM-DD」。
                    # 以 Excel 實際截止日為準，檔名只用來驗證報表日期是否合理。
                    match=re.fullmatch(r'(?:需給日報|(?:#U[0-9a-fA-F]{4})+)-(?:(\d{1,2})-)?(\d{1,2})(?:\s*\(\d+\))?', path.stem)
                    if match:
                        month_text, day_text = match.groups()
                        report_day=int(day_text)
                        if month_text:
                            report_month=int(month_text)
                            # Excel 通常截至報表日的前一天；允許跨月、跨年。
                            report_dates=[]
                            for year in (latest.year-1,latest.year,latest.year+1):
                                try: report_dates.append(_date(year,report_month,report_day))
                                except ValueError: pass
                            if not report_dates:
                                raise ValueError(f'檔名日期不合法：{path.name}')
                            report_date=min(report_dates,key=lambda d:abs((d-latest).days))
                            lag=(report_date-latest).days
                            if not 0<=lag<=3:
                                raise ValueError(f'檔名報表日 {report_date} 與正式截止日 {latest} 不符（差 {lag} 天）')
                        elif not 1<=report_day<=31:
                            raise ValueError(f'檔名日期不合法：{path.name}')
                        # 舊檔名不含月份，不以當前月份推斷；只用資料日期排序。
                    candidates.append((latest,path,parsed))
                except Exception as exc:
                    rejected.append({'file':path.name,'reason':str(exc)})
                    print(f'警告：略過異常日報 {path.name}：{exc}',flush=True)
            if candidates:
                # 優先採用資料截止日最新的一份完整日報，避免同月份舊資料覆蓋新資料。
                latest,path,rows=max(candidates,key=lambda x:(x[0],x[1].name))
                accepted=[path.name]
                existing=load_json(ROOT/'data'/'ordering_daily.json',{}).get('rows',[])
                if existing:
                    existing_dates=[]
                    for r in existing:
                        try:
                            y,m,d=map(int,r['item'].split('/'))
                            existing_dates.append(_date(y,m,d))
                        except (KeyError,ValueError,TypeError):pass
                    if existing_dates and max(existing_dates)>latest:
                        print(f'警告：既有 JSON 截止日 {max(existing_dates)} 比新檔 {latest} 晚，保留既有 JSON',flush=True)
                        rows=[]
                if rows:
                    rows=dedupe(rows);validate(rows,area)
                    data[area]=rows
            else:
                print('警告：沒有有效日報，沿用既有 ordering_daily.json',flush=True)
            reports[area]={'count':len(data.get(area,[])),'files':accepted,'skipped':rejected}
            continue
        for path in paths:
            print(f'解析 {path.relative_to(ROOT)}',flush=True)
            rows.extend(PARSERS[area](path))
        rows=dedupe(rows);validate(rows,area)
        data[area]=rows;reports[area]={'count':len(rows),'files':[p.name for p in paths]}
    finance_doc=None
    if args.area in ('all','finance'):
        paths=files('finance')
        if paths:
            # Only newest snapshot: never combine old complete snapshots and double-count.
            latest=max(paths,key=lambda p:(p.stat().st_mtime,p.name))
            print(f'解析收支分析 {latest.relative_to(ROOT)}')
            finance_doc=finance(latest)
            counts=validate_finance(finance_doc)
            reports['finance']={'count':sum(counts.values()),'sheets':counts,'files':[latest.name]}
        else:print('略過 finance：沒有新 Excel；沿用既有 JSON')
    # 先驗證所有更新，再寫檔，任何解析失敗則停止部署。
    for area,rows in data.items():
        name='ordering_daily.json' if area=='daily' else ('ordering_monthly.json' if area=='monthly' else f'{area}.json')
        save_json(ROOT/'data'/name,{'schemaVersion':1,'area':area,'rows':rows})
    # V8：逐日歷史資料與 10 月日報分開產生，不覆蓋現有 daily JSON。
    if args.area in ('all','monthly'):
        history=[]
        for path in files('monthly'):
            history.extend(convert_daily_history(path))
        if history:
            # 同日同 KPI 若重複出現在不同報表，以檔名較新的版本覆蓋。
            unique={}
            for r in history:unique[(r['year'],r['month'],r.get('day'),r['kpi'],r['item'])]=r
            history=list(unique.values())
            # 歷史日明細以日期作唯一鍵，不能套用月結 canonical_key。
            assert len(history)==len(unique)
            save_json(ROOT/'data'/'ordering_daily_history.json',{'schemaVersion':1,'rows':history})
            reports['monthly_daily']={'count':len(history)}
    if finance_doc is not None:save_json(ROOT/'data'/'finance.json',{'schemaVersion':1,**finance_doc})
    # 合併三倉；未更新倉別沿用既有資料，不能因單倉更新清空其他倉。
    warehouse=[]
    for area in ('yangmei','taichung','renwu'):
        warehouse.extend(data.get(area,load_json(ROOT/'data'/f'{area}.json',{}).get('rows',[])))
    if warehouse:save_json(ROOT/'data'/'warehouse.json',{'schemaVersion':1,'rows':dedupe(warehouse)})
    save_json(ROOT/'data'/'manifest.json',{'schemaVersion':1,'updatedAt':datetime.now(__import__("datetime").timezone.utc).isoformat(),'updated':reports,'warehouseRows':len(warehouse)})
    print('驗證完成',reports)
if __name__=='__main__':
    try:main()
    except Exception as e:print('驗證失敗，停止發布：',e,file=sys.stderr);sys.exit(1)

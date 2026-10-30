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
        for path in paths:
            print(f'解析 {path.relative_to(ROOT)}')
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

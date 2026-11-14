#!/usr/bin/env python3
"""Read-only pre-deployment gate for ordering_daily.json.

Checks structural completeness of the most recent published day. Does not
recalculate Excel or validate all historical monthly/warehouse KPI values.
"""
import argparse
import json
import math
import sys
from collections import Counter
from datetime import date
from pathlib import Path

SUMMARY = ('庫存金額日推移', '全省供應率日推移', 'BO點數日推移', 'A類庫存日推移', '期貨金額日推移')
DETAIL = {'庫存金額日明細': 8, '全省供應率日明細': 7, 'BO點數日明細': 7, 'A類庫存日明細': 3, '期貨金額日明細': 13}

def parse_day(row):
    if row.get('day') is not None:
        return date(int(row['year']), int(row['month']), int(row['day']))
    parts = str(row['item']).split('/')
    if len(parts) != 3: raise ValueError('date not encoded in item')
    d = date(*(int(p) for p in parts))
    if (d.year, d.month) != (int(row['year']), int(row['month'])):
        raise ValueError('item date does not match year/month')
    return d

def validate(doc, expected_cutoff=None):
    errors=[]
    rows=doc.get('rows')
    if not isinstance(rows,list) or not rows:
        return {'status':'FAIL','errors':['missing or empty rows'], 'cutoff':None}
    by_day={}
    keys=set()
    for idx,r in enumerate(rows):
        try:
            d=parse_day(r)
            kpi=r['kpi']; item=str(r['item'])
            v=r['actual']
            if not isinstance(v,(int,float)) or isinstance(v,bool) or not math.isfinite(v):
                raise ValueError('invalid numeric actual')
            if v < 0: raise ValueError('negative actual')
            if r.get('unit')=='%' and v>100: raise ValueError('percentage above 100')
            key=(d,kpi,item)
            if key in keys: errors.append(f'duplicate key {key}')
            keys.add(key)
            by_day.setdefault(d,{}).setdefault(kpi,[]).append(r)
        except (ValueError,TypeError,KeyError,OverflowError) as exc:
            errors.append(f'row {idx}: {exc}')
    if not by_day: return {'status':'FAIL','errors':errors or ['no valid dates'],'cutoff':None}
    cutoff=max(by_day)
    if expected_cutoff and cutoff!=expected_cutoff:
        errors.append(f'cutoff mismatch: actual={cutoff}, expected={expected_cutoff}')
    latest=by_day[cutoff]
    for kpi in SUMMARY:
        subset=latest.get(kpi,[])
        if len(subset)!=1:
            errors.append(f'{cutoff} {kpi}: expected 1 row, got {len(subset)}')
        elif kpi!='全省供應率日推移' and subset[0]['actual']<=0:
            errors.append(f'{cutoff} {kpi}: non-positive value')
    for kpi,count in DETAIL.items():
        subset=latest.get(kpi,[])
        if len(subset)!=count:
            errors.append(f'{cutoff} {kpi}: expected {count} rows, got {len(subset)}')
        if len({r['item'] for r in subset})!=len(subset):
            errors.append(f'{cutoff} {kpi}: duplicate item names')
    return {'status':'PASS' if not errors else 'FAIL','cutoff':str(cutoff),'rows':len(rows),
            'latest_counts':{k:len(v) for k,v in latest.items()},'errors':errors,
            'note':'庫存月數為月結指標；無當月月結資料時允許顯示 —，不納入日報完整性檢查'}

def main():
    p=argparse.ArgumentParser()
    p.add_argument('json_file',type=Path)
    p.add_argument('--expected-cutoff',type=date.fromisoformat)
    args=p.parse_args()
    try:
        result=validate(json.loads(args.json_file.read_text(encoding='utf-8')),args.expected_cutoff)
    except (OSError,ValueError,json.JSONDecodeError) as exc:
        result={'status':'FAIL','errors':[str(exc)]}
    print(json.dumps(result,ensure_ascii=False,indent=2))
    return 0 if result['status']=='PASS' else 1

if __name__=='__main__': sys.exit(main())

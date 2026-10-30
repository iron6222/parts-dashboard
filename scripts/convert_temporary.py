#!/usr/bin/env python3
"""Convert raw 追蹤彙總表 Excel files into monthly snapshot JSON. Stdlib only."""
import argparse, json, re, zipfile, xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime, timezone
N={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def read_sheet(path):
    with zipfile.ZipFile(path) as z:
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        targets={x.get('Id'):x.get('Target') for x in rel}
        sheets=wb.findall('.//m:sheets/m:sheet',N)
        sh=next((x for x in sheets if x.get('name','').startswith('追蹤彙總表')),None)
        if sh is None: raise ValueError(f'{path}: 找不到追蹤彙總表')
        rid=sh.get('{'+N['r']+'}id');target=targets[rid].lstrip('/')
        if not target.startswith('xl/'):target='xl/'+target
        ss=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            ss=[''.join(t.text or '' for t in si.findall('.//m:t',N)) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si',N)]
        root=ET.fromstring(z.read(target));out={}
        for row in root.findall('.//m:sheetData/m:row',N):
            vals={}
            for c in row.findall('m:c',N):
                v=c.find('m:v',N); inline=c.find('m:is',N)
                if v is not None:
                    s=ss[int(v.text)] if c.get('t')=='s' else v.text
                elif inline is not None:s=''.join(t.text or '' for t in inline.findall('.//m:t',N))
                else:continue
                vals[re.match(r'[A-Z]+',c.get('r')).group()]=s
            out[int(row.get('r'))]=vals
        return out

def number(v):
    try:return float(v)
    except (ValueError,TypeError):return None

def parse(path):
    rows=read_sheet(path);records=[];year=None;section=None;metric=None;category=None;valid_months=set(range(1,13))
    for r,c in rows.items():
        a=c.get('A','').replace("'",'').strip(); b=c.get('B','').replace("'",'').strip()
        h=re.search(r'(20\d{2})年\s*\n?\s*\(([^)]+)\)',a)
        if h:
            year=int(h.group(1));section=h.group(2);metric=None;category=None
            if year not in (2025,2026):year=None
            valid_months={m for m in range(1,13) if re.search(r"\d+[/\-]\d+",c.get(chr(ord("C")+m-1),""))}
            continue
        if year is None:continue
        if section in ('ALL','ALL-未結','長天期','長天期-未結'):
            if '點數' in a:metric='點數'
            elif '金額' in a:metric='金額'
            if b and b!='月/日':category=b
        elif section in ('倉缺/破損','各項目'):
            if a and a!='小計':category=a
            if b in ('點數','金額'):metric=b
        else:continue
        if b=='小計' or a=='小計':category='小計'
        if not metric or not category:continue
        for month in sorted(valid_months):
            col=chr(ord('C')+month-1);v=number(c.get(col))
            if v is None:continue
            records.append({'year':year,'month':month,'section':section,'category':category,'metric':metric,'value':v,'source_file':path.name})
    return records

def main():
    p=argparse.ArgumentParser();p.add_argument('inputs',nargs='+');p.add_argument('-o','--output',default='data/temporary.json');args=p.parse_args()
    # Most recent workbook per year wins, preserving all months in its summary.
    sources={}
    for x in args.inputs:
        path=Path(x)
        if not path.is_file():continue
        m=re.search(r'(20\d{2})(\d{2})(\d{2})',path.name)
        if not m:continue
        year=int(m.group(1))
        if year not in sources or m.group(0)>sources[year][0]:sources[year]=(m.group(0),path)
    result=[]
    for _,path in sorted(sources.values()):result+=parse(path)
    if not result:raise SystemExit('沒有可用的暫存資料；請檢查檔名日期與工作表')
    output=Path(args.output);output.parent.mkdir(parents=True,exist_ok=True)
    output.write_text(json.dumps({'schema':'temporary-snapshots-v1','generated_at':datetime.now(timezone.utc).isoformat(),'sources':[str(p.name) for _,p in sorted(sources.values())],'records':result},ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'Wrote {len(result)} records to {output}')
if __name__=='__main__':main()

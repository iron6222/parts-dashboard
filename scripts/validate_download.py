#!/usr/bin/env python3
"""Check the downloaded Excel is today's report, and JSON has yesterday's data.

Designed for scheduled and manual GitHub Actions runs in Asia/Taipei.
"""
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo
import re
import sys
from validate_daily_deploy import validate
import json

folder = Path('raw/daily')
today = datetime.now(ZoneInfo('Asia/Taipei')).date()
pattern = re.compile(r'^需給日報-(\d{1,2})-(\d{1,2})\.xlsx$')
matching = [p for p in folder.glob('*.xlsx') if (m := pattern.fullmatch(p.name)) and (int(m.group(1)),int(m.group(2))) == (today.month,today.day)]
if len(matching) != 1:
    sys.exit(f'FAIL: Expected exactly one report for {today}, found {len(matching)}; stop deployment')
result = validate(json.loads(Path('data/ordering_daily.json').read_text(encoding='utf-8')), today-timedelta(days=1))
print(json.dumps(result,ensure_ascii=False,indent=2))
if result['status'] != 'PASS':
    sys.exit('FAIL: daily KPI validation did not pass')
print('PASS: downloaded report date and KPI JSON validated')

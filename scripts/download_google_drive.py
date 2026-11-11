#!/usr/bin/env python3
"""Fetch the newest Excel daily report from a shared Google Drive folder.

Uses a short-lived Google OAuth access token supplied by GitHub Actions OIDC.
Does not log file names or file content.
"""
import argparse
import io
import json
import os
from pathlib import Path
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
import zipfile

API = "https://www.googleapis.com/drive/v3/files"
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def request_json(url, token):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}", "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as response:
        return json.load(response)


def list_excel(folder_id, token):
    result = []
    page_token = None
    escaped = folder_id.replace("\\", "\\\\").replace("'", "\\'")
    while True:
        args = {
            "q": f"'{escaped}' in parents and trashed = false",
            "fields": "nextPageToken,files(id,name,mimeType,modifiedTime,size)",
            "pageSize": 1000,
        }
        if page_token:
            args["pageToken"] = page_token
        data = request_json(API + "?" + urllib.parse.urlencode(args), token)
        result.extend(f for f in data.get("files", []) if
                      f.get("mimeType") == EXCEL_MIME and str(f.get("name", "")).lower().endswith(".xlsx"))
        page_token = data.get("nextPageToken")
        if not page_token:
            return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default="raw/daily")
    args = parser.parse_args()
    token = os.environ.get("GOOGLE_ACCESS_TOKEN", "")
    folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "")
    if not token or not folder_id:
        raise RuntimeError("Missing GOOGLE_ACCESS_TOKEN or GOOGLE_DRIVE_FOLDER_ID")

    files = list_excel(folder_id, token)
    print(f"Found {len(files)} Excel report(s) in Google Drive folder")
    if not files:
        raise RuntimeError("No .xlsx files in the specified Google Drive folder")
    newest = max(files, key=lambda f: (f.get("modifiedTime", ""), f["id"]))
    file_id = urllib.parse.quote(newest["id"], safe="")
    url = f"{API}/{file_id}?alt=media"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=90) as response:
        payload = response.read(30 * 1024 * 1024 + 1)
    if len(payload) > 30 * 1024 * 1024:
        raise RuntimeError("Excel report exceeds the 30 MB safety limit")
    if not zipfile.is_zipfile(io.BytesIO(payload)):
        raise RuntimeError("Downloaded document is not a valid .xlsx container")
    with zipfile.ZipFile(io.BytesIO(payload)) as zf:
        if "xl/workbook.xml" not in zf.namelist():
            raise RuntimeError("Downloaded document is missing the Excel workbook")

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    # Preserve the original filename when safe; do not allow paths or control characters.
    name = str(newest["name"])
    if (name in (".", "..") or Path(name).name != name or
            re.search(r"[\\/\x00-\x1f]", name)):
        raise RuntimeError("Invalid Excel filename from Google Drive")
    output = out_dir / name
    tmp = out_dir / (name + ".tmp")
    tmp.write_bytes(payload)
    tmp.replace(output)
    print(f"Downloaded one Excel workbook ({len(payload)} bytes) into {out_dir}")
    print("Drive daily report download PASSED")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError, urllib.error.URLError) as exc:
        print(f"Google Drive download failed: {type(exc).__name__}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"Google Drive download failed: {type(exc).__name__}: {exc}", file=sys.stderr)
        sys.exit(1)

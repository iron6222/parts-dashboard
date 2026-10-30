# 零件部管理 Dashboard — GitHub Actions / Pages

以 **v46.0** 為版面基準。`index.html` 原有三模組、圖表、卡片、手動資料整理功能均保留，僅在頁尾新增線上 JSON 載入器。

## 部署

1. 建立 **可公開資料** 的 GitHub repository，將本資料夾內容上傳至 repository 根目錄（包含隱藏的 `.github` 目錄）。
2. Repository → Settings → Pages → Build and deployment → **Source: GitHub Actions**。
3. Actions → Excel to Dashboard JSON → Run workflow，或將 Excel 上傳至 `raw/` 對應資料夾以觸發。
4. 等待 workflow 成功後，開啟 Pages 網址；瀏覽器重新整理時讀取 `data/manifest.json`、`warehouse.json`、`ordering_daily.json`、`finance.json`。
5. 若尚未發布 JSON 或離線雙擊 HTML，原有內建資料與手動匯入仍可使用。

## 資料更新

| 路徑 | 來源 | 轉換器 |
| --- | --- | --- |
| `raw/yangmei/*.xlsx` | 楊梅月報（`2026年` 工作表 G:O 1–9 月等） | `convert_yangmei.py` |
| `raw/taichung/*.xlsx` | 台中月報 | `convert_taichung.py` |
| `raw/renwu/*.xlsx` | 仁武月報 | `convert_renwu.py` |
| `raw/daily/*.xlsx` | 需給日報 | `convert_daily.py` |
| `raw/finance/*.xlsx` | 已整理的費用收入彙整資料庫（三張標準工作表） | `convert_finance.py` |

**重要：這是安全失敗的第一版轉換器，不是已完成的四種原始月報格式全自動對照。** 楊梅轉換器目前核對的固定列是加班（8、15、22）與全區庫位（58），分區列 48–57 依原表標籤；其餘 KPI 尚未建立原始列對照。台中／仁武只支援有「月份／KPI／實績」欄名的標準長表，**尚未支援**使用者提供的 SDTC/SDHN 原始管理月報特殊格式。需給日報只支援有「日期」及日報 KPI 欄位的標準寬表，**尚未支援**使用者原始「需給日報-30.xlsx」特殊版型與月結判讀。未辨識時 Actions 會失敗、保留上一版部署，絕不猜測資料或覆蓋成 0。

### 注意

- **請勿在公開 repository 上傳未經公司核准的資料。** GitHub Pages 和其 JSON 檔案是公開可讀取的。
- GitHub Pages 是靜態網站，無法由瀏覽器「資料整理」按鈕直接寫回 repository。更新請到 GitHub 上傳原始 Excel，或另行設計具權限的後端 API。
- 原始 HTML 內建示範/歷史資料可能在新 JSON 未覆蓋的月份繼續顯示；在轉換器完全核對前，**不可視為所有 KPI 都已在線上同步**。
- 原始程式可能存在瀏覽器 IndexedDB 的舊資料；線上載入器會在原始還原後覆寫倉儲與日報兩組資料，但訂貨月結及收支資料仍保留原邏輯。
- GitHub Pages 的每次部署會重新生成網站，並不會持久保存先前生成的 `data/`；本專案若只上傳單月 Excel，必須在 repository 保留所有需發布的歷史原始 Excel，或將歷史 JSON 另存進版控。

## 本機測試

```bash
pip install -r requirements.txt
python scripts/build.py
python -m http.server 8000
```

瀏覽 `http://localhost:8000/`。若解析失敗，請查看錯誤訊息並先補齊來源欄位對照，切勿直接跳過檢查。

## 收支分析（第五種資料來源）

將最新版 `費用收入彙整資料庫_*.xlsx` 上傳到 `raw/finance/`，Actions 轉成 `data/finance.json`，保留與 HTML 手動匯入相同的 `inc`、`expMain`、`expSub` 欄位。網站每次載入會優先使用已發布 JSON，離線開啟仍可手動上傳 Excel。**請只保留一份最新完整彙整檔**，避免選錯版本；轉換器會檢查三張工作表、年月、金額與重複鍵。

收支資料為整份快照，不是增量月報；更新時請上傳包含所有需保留年度及月份的完整彙整檔。

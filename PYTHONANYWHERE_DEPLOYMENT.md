# PythonAnywhere 部署說明

本專案會在 Web App 第一次收到請求時自動建立或升級 SQLite，並依
`seed_data/utility_points_manifest.json` 判斷是否需要更新公用點位。
相同 checksum 不會重複匯入；新版來源資料也不會覆蓋人工 EMS 設定。

## 1. 建立 virtualenv

在 PythonAnywhere Bash console 執行（請依 Web App 選用的 Python 版本調整）：

```bash
python3.12 -m venv ~/.virtualenvs/energy-system
source ~/.virtualenvs/energy-system/bin/activate
pip install -r /home/YOUR_USERNAME/energy-system/requirements.txt
```

## 2. 設定 WSGI

從 PythonAnywhere 的 **Web** 頁面建立 Manual configuration，再編輯 WSGI 檔：

```python
import os
import sys

project_path = "/home/YOUR_USERNAME/energy-system"
if project_path not in sys.path:
    sys.path.insert(0, project_path)

os.environ["ENERGY_DATABASE_PATH"] = "/home/YOUR_USERNAME/energy-data/app.db"
os.environ["ENERGY_SECRET_KEY"] = "請換成至少 32 bytes 的隨機字串"

from app import app as application
```

在 Web 頁面的 Virtualenv 欄位填入：

```text
/home/YOUR_USERNAME/.virtualenvs/energy-system
```

`energy-data` 必須保留於程式碼目錄之外：

```bash
mkdir -p /home/YOUR_USERNAME/energy-data
```

## 3. 第一次啟動與驗證

按 Web 頁面的 **Reload**，登入後開啟「EMS 點位設定」。預期顯示：

- 點位版本 `2026.09.18.1`
- UtilityDepartment 點位 388 筆
- 388 筆 EMS 設定，預設全部監控

若匯入失敗，查看 PythonAnywhere Web 頁面的 error log。

## 4. 日後部署

```bash
cd /home/YOUR_USERNAME/energy-system
git pull
source ~/.virtualenvs/energy-system/bin/activate
pip install -r requirements.txt
```

然後在 Web 頁面按 **Reload**。系統會自動比較 manifest checksum：

- 相同：略過匯入。
- 不同：更新來源欄位、新增新點位、標記已消失點位。
- 人工上下限、監控與 EMS 分類：全部保留。

## 5. 在本機產生新版點位快照

本機資料庫查詢 API 可用時執行：

```powershell
python .\scripts\export_utility_points.py --version 2026.09.18.2
```

確認 `seed_data` 的差異後提交及部署。匯出工具不會把內部 API URL 寫入部署檔。

## 6. 備份

SQLite 位於：

```text
/home/YOUR_USERNAME/energy-data/app.db
```

可使用管理首頁的「備份資料庫」下載備份。正式收集資料後，部署前也應先備份。

## 重要限制

- PythonAnywhere 無法透過 `127.0.0.1:8080` 連回開發電腦。
- PythonAnywhere 無法直接連入公司內網的 `10.x.x.x` 裝置 API。
- 此 SQLite 架構適合目前少量管理員設定；大量即時 EMS 歷史資料應改用中央資料庫。
- 正式公開前仍應完成密碼雜湊與 CSRF 防護。

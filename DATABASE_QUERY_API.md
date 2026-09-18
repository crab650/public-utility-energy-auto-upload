# 資料庫查詢 API 使用約定

本專案查詢 SQL Server 時，固定使用本機 HTTP API。除非使用者明確要求，否則不要改用 MCP、直接連線 SQL Server，或猜測資料表與欄位。

## 固定連線方式

- API Base URL：`http://127.0.0.1:8080`
- 存取範圍：僅限本機 `127.0.0.1`
- 資料庫模式：唯讀
- SQL 類型：僅允許單一 `SELECT` 或唯讀 CTE

## 標準查詢流程

每次開始查詢時，依照以下順序執行：

1. 呼叫 `GET /api/status`，確認 API 正在運行、`read_only` 為 `true`，且 `query_enabled` 為 `true`。
2. 若本次對話尚未確認可用資源，呼叫 `GET /api/resources/discover`，取得白名單內的 Table 與 View。
3. 不得猜測資料表或 View 名稱，只能使用探索結果中實際存在的資源。
4. 查詢欄位前，先透過安全的 metadata `SELECT` 或小範圍查詢確認欄位名稱與資料型別。
5. 呼叫 `POST /api/query` 執行唯讀 SQL。
6. 檢查回傳的 `success`、`row_count`、`truncated` 與 `error`。
7. 將結果整理成人類容易閱讀的答案，並在資料遭截斷時明確告知使用者。

## API 端點

### 1. 檢查服務狀態

```http
GET http://127.0.0.1:8080/api/status
```

預期回傳範例：

```json
{
  "server": "running",
  "read_only": true,
  "query_enabled": true,
  "database": "DGC"
}
```

只有在 `server` 為 `running` 且 `query_enabled` 為 `true` 時，才繼續查詢。

### 2. 探索允許查詢的資源

```http
GET http://127.0.0.1:8080/api/resources/discover
```

回傳格式：

```json
{
  "success": true,
  "tables": [
    {
      "schema": "dbo",
      "name": "TableName"
    }
  ],
  "views": [
    {
      "schema": "dbo",
      "name": "ViewName"
    }
  ]
}
```

### 3. 執行唯讀 SQL

```http
POST http://127.0.0.1:8080/api/query
Content-Type: application/json
```

請求格式：

```json
{
  "sql": "SELECT TOP (10) ColumnA, ColumnB FROM dbo.TableName WHERE ColumnA = :value",
  "parameters": {
    "value": "查詢值"
  }
}
```

回傳格式：

```json
{
  "success": true,
  "rows": [],
  "row_count": 0,
  "truncated": false
}
```

## PowerShell 呼叫範例

### 狀態檢查

```powershell
Invoke-RestMethod `
  -Uri 'http://127.0.0.1:8080/api/status' `
  -Method Get
```

### 探索資源

```powershell
Invoke-RestMethod `
  -Uri 'http://127.0.0.1:8080/api/resources/discover' `
  -Method Get
```

### 執行參數化查詢

```powershell
$body = @{
  sql = 'SELECT TOP (10) ColumnA, ColumnB FROM dbo.TableName WHERE ColumnA = :value'
  parameters = @{
    value = '查詢值'
  }
} | ConvertTo-Json -Depth 10

Invoke-RestMethod `
  -Uri 'http://127.0.0.1:8080/api/query' `
  -Method Post `
  -ContentType 'application/json' `
  -Body $body
```

## SQL 安全規則

- 只能執行單一、唯讀的 `SELECT` 或唯讀 CTE。
- 禁止 `INSERT`、`UPDATE`、`DELETE`、`MERGE`。
- 禁止 `CREATE`、`ALTER`、`DROP`、`TRUNCATE`。
- 禁止 `EXEC`、Stored Procedure 與 `SELECT ... INTO`。
- 禁止用分號串接多段 SQL。
- 動態值不得直接拼接到 SQL 字串中。
- 所有動態值都使用 `:parameter_name`，並放進 `parameters`。
- 參數名稱只能包含英文字母、數字及底線，且不能以數字開頭。
- 資源名稱一律寫成 `schema.name`，例如 `dbo.TableName`。
- 避免 `SELECT *`，只選擇回答問題所需的欄位。
- 優先使用 `WHERE` 限制日期、條件與資料範圍。
- 限制筆數時使用 SQL Server 語法 `TOP (n)`。

## AI 與使用者的溝通約定

當使用者說「查資料庫」、「查數據」、「幫我找資料」或提出資料分析問題時：

1. AI 應直接使用本文所述的 HTTP API。
2. AI 不應再次詢問要使用 MCP 還是 API。
3. AI 不應把 `AI_SYSTEM_PROMPT.md` 中的 MCP 方式當成預設連線方式。
4. 若使用者沒有指定資料表，AI 應先探索資源並依名稱判斷可能的候選資料表。
5. 若存在多個合理候選且會產生不同答案，AI 應說明候選差異，再請使用者確認。
6. 若能從資料結構安全判定查詢目標，AI 應直接完成查詢，不必要求使用者提供 SQL。
7. 回答時應提供查詢結果摘要；只有在有助於核對時才附上 SQL。

## 錯誤處理

- API 無法連線：確認 `http://127.0.0.1:8080` 服務是否啟動。
- `QUERY_DISABLED`：查詢功能尚未啟用，需由管理員開啟。
- `RESOURCE_NOT_ALLOWED`：重新探索資源，只使用白名單內的 Table 或 View。
- `QUERY_TIMEOUT` / `RESULT_TOO_LARGE`：縮小日期範圍、減少欄位、簡化 Join 或加入 `TOP`。
- `INVALID_SQL` / `INVALID_PARAMETER`：檢查 SQL 語法及 `:name` 與 `parameters` 是否一致。
- `INTERNAL_ERROR`：保留回傳的 `request_id`，供管理員查閱 Audit Log。

## 最重要的固定原則

> 本專案預設以 `http://127.0.0.1:8080` 的 HTTP API 查詢資料庫；先探索資源，再執行參數化、唯讀的 SQL。

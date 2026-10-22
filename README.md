# Enterprise Energy Reporting & Analytics System (Prototype)

A robust, web-based utility data collection and analytics portal designed for company-wide energy cost reporting, control, and monitoring. Built to support cross-border operations, the system features dynamic multi-language localization (Traditional Chinese and Vietnamese) and granular multi-dimensional permission management.

---

## 🌟 Key Features

### 1. Direct Excel Copy-Paste Integration
* **Seamless Grid Entry:** Enables users to copy ranges directly from Microsoft Excel (`Ctrl+C`) and paste them (`Ctrl+V`) into the web reporting form.
* **Auto-parsing & Validation:** The Javascript grid automatically parses cells, maps them to the appropriate fields, and flags invalid numbers (e.g., text values or malformed currencies) with visual alerts in real-time.

### 2. Multi-dimensional Authorization Grid (X-Y Permissions)
* **Granular Control:** Administrators can assign reporting scopes by mapping fillers to specific **Energy Types (Y-axis)** and **Areas/Units (X-axis)**.
* **Tailored Viewports:** Data entry grids are dynamically generated, showing fillers only the rows and columns they are authorized to fill.

### 3. Customizable Control Limits & Warnings
* **Fat-finger Protection:** System maintains historical averages for each energy category and area.
* **Flexible Recalculation:** Admins can dynamically calculate control limits using a configurable deviation threshold (±5%, ±10%, ±15%, ±20%, ±30%).
* **Pre-submission Checks:** Warns users if their entry significantly deviates from historical ranges, preventing typos before data is locked.

### 4. Dynamic Multi-Language Engine
* **Instant Language Toggle:** Supports seamless bilingual translation between **Traditional Chinese (`zh-TW`)** and **Vietnamese (`vi`)** dynamically rendered on the server side.
* **Cross-border Friendly:** Tailored for operations with manufacturing plants or administrative hubs in both regions.

### 5. System Audit Trail (Operations Log)
* **High Accountability:** Captures critical actions taken by administrators and data fillers (e.g., log in, data submission, permission modifications, limit recalculations, database backup downloads).
* **Real-time Dashboard Widget:** Displays the latest 15 operations logs on the Admin Dashboard for quick oversight.

### 6. One-click Database Backup
* **Zero-config Backups:** Admins can download a full backup copy of the SQLite `.db` database directly from the dashboard heading with a single click.

### 7. EMS Utility Point Configuration
* **Bundled Point Catalog:** Imports 388 `UtilityDepartment` points into SQLite from a versioned deployment snapshot.
* **Safe Automatic Updates:** On application startup, the manifest checksum is compared with the last successful catalog version. Source metadata is updated only when needed, while manually entered EMS settings are preserved.
* **Admin-only Configuration:** Administrators can search and filter points, enable or disable monitoring, configure upper/lower alarm limits, and classify points by energy type, measurement type, unit, location, equipment, value type, aggregation method, and flow direction.
* **Bulk Operations and Export:** Selected points can be updated in bulk, and the complete EMS configuration can be exported to Excel.
* **Offline Fill Template:** Administrators can download an Excel workbook with every UtilityDepartment point, its current EMS settings, editable columns, dropdown choices, and filling instructions. Import of the completed workbook is planned separately.

---

## 🛠️ Technology Stack
* **Backend:** Python / Flask
* **Database:** SQLite
* **Frontend:** Vanilla HTML5, CSS3, and JavaScript (engineered with a retro Windows 95/98 visual styling featuring double-bordered frames, HSL-tailored colors, and grid tables)
* **Excel Generation:** `openpyxl` (Generates formatted report exports)

---

## 📊 Database Schema

```mermaid
erDiagram
    users ||--o{ user_energy_permissions : has
    users ||--o{ user_area_permissions : has
    users ||--o{ energy_entries : submits
    users ||--o{ report_submissions : logs
    users ||--o{ audit_logs : logs
    
    energy_types ||--o{ user_energy_permissions : restricts
    energy_types ||--o{ energy_entries : categorized-by
    energy_types ||--o{ control_limits : defines

    areas ||--o{ user_area_permissions : restricts
    areas ||--o{ energy_entries : located-at
    areas ||--o{ control_limits : defines

    users {
        int id PK
        string username UNIQUE
        string password
        string display_name
        string role "admin | filler"
        int viewer "0 | 1 (guest)"
        int active "0 | 1"
        string created_at
    }

    energy_types {
        int id PK
        string name UNIQUE
        string unit
        int active "0 | 1"
        int display_order
    }

    areas {
        int id PK
        string name UNIQUE
        int active "0 | 1"
        int display_order
    }

    energy_entries {
        int id PK
        string report_month
        int user_id FK
        int energy_type_id FK
        int area_id FK
        float quantity
        float amount
        string updated_at
    }

    report_submissions {
        int id PK
        string report_month
        int user_id FK
        string status
        string submitted_at
    }

    control_limits {
        int id PK
        int energy_type_id FK
        int area_id FK
        int quantity_enabled "0 | 1"
        float quantity_average
        float quantity_min
        float quantity_max
        int amount_enabled "0 | 1"
        float amount_average
        float amount_min
        float amount_max
        string updated_at
    }

    audit_logs {
        int id PK
        int user_id FK
        string username
        string action
        string details
        string created_at
    }
```

---

## 🚀 Getting Started

### Prerequisites
* Python 3.8 or higher installed on your local machine.

### Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/crab650/public-utility-energy-auto-upload.git
   cd public-utility-energy-auto-upload/prototype
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application:**
   ```bash
   python app.py
   ```
   *Note: On first startup, the system automatically creates the database directory, seeds default energy types, areas, and creates a default administrator account.*

4. **Access the portal:**
   * Open `http://127.0.0.1:5000` in your web browser.
   * **Default Admin Credentials:**
     * **Username:** `admin`
     * **Password:** `1234`

New filler and viewer accounts receive the initial password `1234`. On first login, they must set a new password of at least eight characters before accessing reports or data entry. An administrator password reset also requires the user to change that password on the next login. Existing accounts are not forced to change their password when the database is upgraded.

### Monthly Energy Data API

Set `ENERGY_API_KEY` to a long, random secret in the web application's environment. The API returns 503 until this is configured. Call `GET /api/energy-monthly-data` with the `X-API-Key` request header. An optional `month=YYYY-MM` query parameter changes the center month; by default it uses the current month in Bangkok time. The response includes records from three months before through three months after the center month, inclusive. Months without submitted data have no records.

Each JSON array item represents one report month and area. `month` uses the first day of the month as `YYYY-MM-01T00:00:00`, and `department` is the area name. Energy quantities appear as top-level numeric properties named by each energy type's configured **API 名稱**. Energy types without an API name are omitted. Amounts are not exposed. Only submitted reports are included; quantities from multiple fillers for the same month, area, and energy type are summed.

Example request:

```http
GET /api/energy-monthly-data?month=2026-10
X-API-Key: YOUR_SECRET
```

Example response:

```json
[{"month":"2026-09-01T00:00:00","department":"Poly 51","electricity":12345.0}]
```

### Refreshing the Utility Point Catalog

On a machine that can reach the read-only local DGC query API:

```bash
python scripts/export_utility_points.py --version 2026.09.18.2
```

Commit the updated files under `seed_data/` and deploy normally. The next application reload automatically imports the new catalog when its checksum changes. New points default to monitoring enabled; missing points are retained and marked as missing; existing EMS settings are never overwritten.

For PythonAnywhere setup, persistent SQLite paths, WSGI configuration, and the deployment workflow, see [`PYTHONANYWHERE_DEPLOYMENT.md`](PYTHONANYWHERE_DEPLOYMENT.md).

---

## 💡 Recommended Future Optimizations
For moving this prototype into production, the following enhancements are recommended:
1. **Password Hashing:** Migrate from plaintext password storage to standard hashes (e.g., using `scrypt` or `bcrypt` via Flask-Bcrypt).
2. **Interactive Charting:** Integrate a front-end library like Chart.js or ECharts for more interactive dashboards, while preserving the retro style layout.
3. **Database Migration Pipeline:** Introduce Flask-Migrate (Alembic) to handle database schema upgrades and updates gracefully without requiring table rebuilds.

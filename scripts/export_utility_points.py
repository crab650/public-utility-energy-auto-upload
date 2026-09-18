"""Export the read-only DGC utility point catalog as a deployment seed.

Run this on a machine that can reach the local database query API:

    python scripts/export_utility_points.py --version 2026.09.18.1
"""

from __future__ import annotations

import argparse
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import URLError
from urllib.request import Request, urlopen


API_BASE_URL = "http://127.0.0.1:8080"
DATA_SOURCE_ID = 1
ROOT_DIR = Path(__file__).resolve().parents[1]


def api_request(path: str, payload: dict | None = None) -> dict:
    data = None
    headers = {}
    method = "GET"
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"
        method = "POST"
    request = Request(f"{API_BASE_URL}{path}", data=data, headers=headers, method=method)
    with urlopen(request, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def readonly_query(sql: str, parameters: dict | None = None) -> list[dict]:
    response = api_request("/api/query", {"sql": sql, "parameters": parameters or {}})
    if not response.get("success"):
        raise RuntimeError(response.get("error") or "Database query failed")
    if response.get("truncated"):
        raise RuntimeError("Database query result was truncated")
    rows = response.get("rows")
    if not isinstance(rows, list):
        raise RuntimeError("Database query returned invalid rows")
    if response.get("row_count") != len(rows):
        raise RuntimeError("Database query row_count does not match returned rows")
    return rows


def export_catalog(output_dir: Path, version: str) -> tuple[Path, Path, int]:
    status = api_request("/api/status")
    if status.get("server") != "running" or status.get("read_only") is not True:
        raise RuntimeError("Database query API is not running in read-only mode")
    if status.get("query_enabled") is not True or status.get("database") != "DGC":
        raise RuntimeError("DGC database query is not enabled")

    resources = api_request("/api/resources/discover")
    allowed_tables = {
        (item.get("schema"), item.get("name")) for item in resources.get("tables", [])
    }
    required = {("dbo", "BI_DeviceDataPoints"), ("dbo", "BI_DeviceSources")}
    if not required.issubset(allowed_tables):
        raise RuntimeError("Required BI device tables are not in the API allowlist")

    # Small queries verify the requested columns before reading the full catalog.
    readonly_query(
        "SELECT TOP (1) [Id], [DataSourceId], [PointName], [JsonPath], "
        "[Description], [IsActive], [PointType], [FetchIntervalMinutes] "
        "FROM dbo.BI_DeviceDataPoints WHERE DataSourceId = :data_source_id ORDER BY [Id]",
        {"data_source_id": DATA_SOURCE_ID},
    )
    source_rows = readonly_query(
        "SELECT TOP (1) [Id], [Name], [ApiUrl], [IsActive] "
        "FROM dbo.BI_DeviceSources WHERE [Id] = :data_source_id",
        {"data_source_id": DATA_SOURCE_ID},
    )
    if len(source_rows) != 1:
        raise RuntimeError(f"Data source {DATA_SOURCE_ID} was not found")
    point_rows = readonly_query(
        "SELECT TOP (1000) [Id], [DataSourceId], [PointName], [JsonPath], "
        "[Description], [IsActive], [PointType], [FetchIntervalMinutes] "
        "FROM dbo.BI_DeviceDataPoints WHERE DataSourceId = :data_source_id ORDER BY [Id]",
        {"data_source_id": DATA_SOURCE_ID},
    )
    external_ids = [row["Id"] for row in point_rows]
    if len(external_ids) != len(set(external_ids)):
        raise RuntimeError("The source contains duplicate point Id values")

    source = source_rows[0]
    catalog = {
        "schema_version": 1,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source": {
            "external_id": source["Id"],
            "name": source["Name"],
            # The internal collection URL is intentionally not bundled into a
            # deployable/public artifact. The PythonAnywhere site only needs metadata.
            "api_url": None,
            "source_is_active": bool(source["IsActive"]),
        },
        "points": [
            {
                "external_id": row["Id"],
                "data_source_id": row["DataSourceId"],
                "point_name": row["PointName"],
                "json_path": row["JsonPath"],
                "description": row["Description"],
                "source_is_active": bool(row["IsActive"]),
                "point_type": row["PointType"],
                "fetch_interval_minutes": row["FetchIntervalMinutes"],
            }
            for row in point_rows
        ],
    }
    output_dir.mkdir(parents=True, exist_ok=True)
    catalog_path = output_dir / "utility_points.json"
    catalog_bytes = (json.dumps(catalog, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    catalog_path.write_bytes(catalog_bytes)
    manifest = {
        "schema_version": 1,
        "data_source_id": DATA_SOURCE_ID,
        "version": version,
        "row_count": len(point_rows),
        "sha256": hashlib.sha256(catalog_bytes).hexdigest(),
        "generated_at": catalog["generated_at"],
    }
    manifest_path = output_dir / "utility_points_manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return catalog_path, manifest_path, len(point_rows)


def main() -> int:
    parser = argparse.ArgumentParser(description="Export UtilityDepartment points for deployment")
    parser.add_argument(
        "--version",
        default=datetime.now().strftime("%Y.%m.%d.1"),
        help="Catalog release version, for example 2026.09.18.1",
    )
    parser.add_argument("--output-dir", type=Path, default=ROOT_DIR / "seed_data")
    args = parser.parse_args()
    try:
        catalog_path, manifest_path, row_count = export_catalog(args.output_dir, args.version)
    except (RuntimeError, URLError, TimeoutError, json.JSONDecodeError) as exc:
        print(f"Export failed: {exc}")
        return 1
    print(f"Exported {row_count} points to {catalog_path}")
    print(f"Manifest written to {manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

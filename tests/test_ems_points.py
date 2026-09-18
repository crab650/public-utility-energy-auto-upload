from __future__ import annotations

import hashlib
import json
import sqlite3
import tempfile
import unittest
from contextlib import contextmanager
from pathlib import Path

import app as application


class EmsPointFeatureTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = application.DB_PATH
        self.original_seed_path = application.POINT_SEED_PATH
        self.original_manifest_path = application.POINT_MANIFEST_PATH
        application.DB_PATH = Path(self.temp_dir.name) / "app.db"
        application._INITIALIZED_DB_PATH = None
        application.app.config.update(TESTING=True, SECRET_KEY="test-secret")
        self.client = application.app.test_client()
        self.client.get("/login")

    def tearDown(self) -> None:
        application.DB_PATH = self.original_db_path
        application.POINT_SEED_PATH = self.original_seed_path
        application.POINT_MANIFEST_PATH = self.original_manifest_path
        application._INITIALIZED_DB_PATH = None
        self.temp_dir.cleanup()

    @contextmanager
    def database(self):
        connection = sqlite3.connect(application.DB_PATH)
        connection.row_factory = sqlite3.Row
        try:
            yield connection
        finally:
            connection.close()

    def login_admin(self) -> None:
        response = self.client.post(
            "/login",
            data={"username": "admin", "password": "1234"},
        )
        self.assertEqual(response.status_code, 302)

    def test_initial_catalog_import_is_complete_and_idempotent(self) -> None:
        with self.database() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM device_data_points").fetchone()[0], 388)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM ems_point_configs").fetchone()[0], 388)
            self.assertEqual(
                db.execute("SELECT COUNT(*) FROM ems_point_configs WHERE monitor_enabled = 1").fetchone()[0],
                388,
            )
            versions_before = db.execute("SELECT COUNT(*) FROM point_catalog_versions").fetchone()[0]
        application._INITIALIZED_DB_PATH = None
        self.client.get("/login")
        with self.database() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM device_data_points").fetchone()[0], 388)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM point_catalog_versions").fetchone()[0], versions_before)

    def test_admin_can_view_and_update_a_point(self) -> None:
        self.login_admin()
        page = self.client.get("/admin/ems-points")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"FQI_U704A.TOTALIZERA.PV", page.data)
        with self.database() as db:
            point_id = db.execute(
                "SELECT id FROM device_data_points WHERE data_source_id = 1 AND external_id = 2"
            ).fetchone()[0]
            category_id = db.execute(
                "SELECT id FROM ems_energy_categories WHERE name = '壓縮空氣'"
            ).fetchone()[0]
        response = self.client.post(
            f"/admin/ems-points/{point_id}",
            data={
                "monitor_enabled": "on",
                "lower_limit": "10.5",
                "upper_limit": "20.5",
                "energy_category_id": str(category_id),
                "alarm_value_mode": "interval_delta",
                "alarm_consecutive_samples": "2",
                "alarm_deadband": "0.5",
                "alarm_severity": "warning",
            },
        )
        self.assertEqual(response.status_code, 302)
        with self.database() as db:
            config = db.execute(
                "SELECT * FROM ems_point_configs WHERE point_id = ?", (point_id,)
            ).fetchone()
            self.assertEqual(config["lower_limit"], 10.5)
            self.assertEqual(config["upper_limit"], 20.5)
            self.assertEqual(config["energy_category_id"], category_id)
            self.assertEqual(config["alarm_value_mode"], "interval_delta")
            self.assertEqual(config["alarm_consecutive_samples"], 2)

    def test_invalid_limits_are_rejected(self) -> None:
        self.login_admin()
        with self.database() as db:
            point_id = db.execute(
                "SELECT id FROM device_data_points WHERE data_source_id = 1 AND external_id = 2"
            ).fetchone()[0]
        self.client.post(
            f"/admin/ems-points/{point_id}",
            data={
                "lower_limit": "100",
                "upper_limit": "10",
                "alarm_value_mode": "raw",
                "alarm_consecutive_samples": "1",
                "alarm_deadband": "0",
                "alarm_severity": "warning",
            },
        )
        with self.database() as db:
            config = db.execute(
                "SELECT lower_limit, upper_limit FROM ems_point_configs WHERE point_id = ?",
                (point_id,),
            ).fetchone()
            self.assertIsNone(config["lower_limit"])
            self.assertIsNone(config["upper_limit"])

    def test_bulk_monitor_update_and_export(self) -> None:
        self.login_admin()
        with self.database() as db:
            point_ids = [
                row[0]
                for row in db.execute(
                    "SELECT id FROM device_data_points WHERE data_source_id = 1 ORDER BY external_id LIMIT 2"
                ).fetchall()
            ]
        response = self.client.post(
            "/admin/ems-points/bulk",
            data={"point_ids": [str(value) for value in point_ids], "action": "monitor_off"},
        )
        self.assertEqual(response.status_code, 302)
        with self.database() as db:
            placeholders = ",".join("?" for _ in point_ids)
            monitored = db.execute(
                f"SELECT SUM(monitor_enabled) FROM ems_point_configs WHERE point_id IN ({placeholders})",
                point_ids,
            ).fetchone()[0]
            self.assertEqual(monitored, 0)
        export = self.client.get("/admin/ems-points/export")
        self.assertEqual(export.status_code, 200)
        self.assertEqual(
            export.mimetype,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    def test_new_catalog_metadata_does_not_overwrite_ems_config(self) -> None:
        with self.database() as db:
            point_id = db.execute(
                "SELECT id FROM device_data_points WHERE data_source_id = 1 AND external_id = 2"
            ).fetchone()[0]
            db.execute(
                "UPDATE ems_point_configs SET lower_limit = 12, upper_limit = 34 WHERE point_id = ?",
                (point_id,),
            )
            db.commit()

        catalog = json.loads(self.original_seed_path.read_text(encoding="utf-8"))
        catalog["points"][0]["description"] = "Changed by source"
        seed_path = Path(self.temp_dir.name) / "utility_points.json"
        manifest_path = Path(self.temp_dir.name) / "utility_points_manifest.json"
        seed_bytes = (json.dumps(catalog, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
        seed_path.write_bytes(seed_bytes)
        manifest_path.write_text(
            json.dumps(
                {
                    "schema_version": 1,
                    "data_source_id": 1,
                    "version": "test.2",
                    "row_count": len(catalog["points"]),
                    "sha256": hashlib.sha256(seed_bytes).hexdigest(),
                }
            ),
            encoding="utf-8",
        )
        application.POINT_SEED_PATH = seed_path
        application.POINT_MANIFEST_PATH = manifest_path
        with application.app.app_context():
            result = application.sync_bundled_device_points(application.get_db())
        self.assertEqual(result["status"], "updated")
        with self.database() as db:
            point = db.execute("SELECT description FROM device_data_points WHERE id = ?", (point_id,)).fetchone()
            config = db.execute(
                "SELECT lower_limit, upper_limit FROM ems_point_configs WHERE point_id = ?", (point_id,)
            ).fetchone()
            self.assertEqual(point["description"], "Changed by source")
            self.assertEqual(config["lower_limit"], 12)
            self.assertEqual(config["upper_limit"], 34)


if __name__ == "__main__":
    unittest.main()

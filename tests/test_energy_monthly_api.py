from __future__ import annotations

import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app as application


class EnergyMonthlyApiTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = application.DB_PATH
        application.DB_PATH = Path(self.temp_dir.name) / "app.db"
        application._INITIALIZED_DB_PATH = None
        application.app.config.update(TESTING=True, SECRET_KEY="test-secret")
        self.client = application.app.test_client()
        self.client.get("/login")

    def tearDown(self) -> None:
        application.DB_PATH = self.original_db_path
        application._INITIALIZED_DB_PATH = None
        self.temp_dir.cleanup()

    def test_seven_month_window_and_flat_energy_fields(self) -> None:
        with sqlite3.connect(application.DB_PATH) as db:
            db.execute("UPDATE energy_types SET api_name = 'electricity' WHERE id = 1")
            db.execute("INSERT INTO users (username, password, display_name, role) VALUES ('second', 'x', 'Second', 'filler')")
            second_user_id = db.execute("SELECT id FROM users WHERE username = 'second'").fetchone()[0]
            db.execute("INSERT INTO users (username, password, display_name, role) VALUES ('third', 'x', 'Third', 'filler')")
            third_user_id = db.execute("SELECT id FROM users WHERE username = 'third'").fetchone()[0]
            for month in ("2025-10", "2025-11", "2026-02", "2026-05", "2026-06"):
                db.execute("INSERT INTO report_submissions (report_month, user_id) VALUES (?, 1)", (month,))
                db.execute(
                    "INSERT INTO energy_entries (report_month, user_id, energy_type_id, area_id, quantity, amount) VALUES (?, 1, 1, 1, 10, 999)",
                    (month,),
                )
            db.execute("INSERT INTO report_submissions (report_month, user_id) VALUES ('2026-02', ?)", (second_user_id,))
            db.execute(
                "INSERT INTO energy_entries (report_month, user_id, energy_type_id, area_id, quantity) VALUES ('2026-02', ?, 1, 1, 5)",
                (second_user_id,),
            )
            db.execute(
                "INSERT INTO energy_entries (report_month, user_id, energy_type_id, area_id, quantity) VALUES ('2026-02', 1, 2, 1, 77)"
            )
            db.execute(
                "INSERT INTO energy_entries (report_month, user_id, energy_type_id, area_id, quantity) VALUES ('2026-02', ?, 1, 2, 100)",
                (third_user_id,),
            )
        db.close()

        with patch.dict(os.environ, {"ENERGY_API_KEY": "test-key"}):
            response = self.client.get("/api/energy-monthly-data?month=2026-02", headers={"X-API-Key": "test-key"})
            self.assertEqual(response.status_code, 200)
            rows = response.get_json()
            self.assertEqual([row["month"] for row in rows], [
                "2025-11-01T00:00:00", "2026-02-01T00:00:00", "2026-05-01T00:00:00",
            ])
            self.assertEqual(rows[1], {"month": "2026-02-01T00:00:00", "department": "Poly 51", "electricity": 15.0})
            with patch.object(application, "current_report_month", return_value="2026-02"):
                default_response = self.client.get("/api/energy-monthly-data", headers={"X-API-Key": "test-key"})
            self.assertEqual(default_response.get_json(), rows)

    def test_api_requires_key_and_valid_month(self) -> None:
        with patch.dict(os.environ, {"ENERGY_API_KEY": "test-key"}):
            self.assertEqual(self.client.get("/api/energy-monthly-data").status_code, 401)
            self.assertEqual(self.client.get("/api/energy-monthly-data?month=2026-13", headers={"X-API-Key": "test-key"}).status_code, 400)
        with patch.dict(os.environ, {"ENERGY_API_KEY": ""}):
            self.assertEqual(self.client.get("/api/energy-monthly-data").status_code, 503)


if __name__ == "__main__":
    unittest.main()

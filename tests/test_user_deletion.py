from __future__ import annotations

import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path

import app as application


class UserDeletionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = application.DB_PATH
        application.DB_PATH = Path(self.temp_dir.name) / "app.db"
        application._INITIALIZED_DB_PATH = None
        application.app.config.update(TESTING=True, SECRET_KEY="test-secret")
        self.admin = application.app.test_client()
        self.admin.get("/login")

    def tearDown(self) -> None:
        application.DB_PATH = self.original_db_path
        application._INITIALIZED_DB_PATH = None
        self.temp_dir.cleanup()

    def test_delete_deactivates_account_and_preserves_history(self) -> None:
        self.admin.post("/login", data={"username": "admin", "password": "1234"})
        self.admin.post("/admin/users", data={
            "username": "reporter", "password": "secret", "display_name": "Reporter",
        })
        with closing(sqlite3.connect(application.DB_PATH)) as db:
            user_id = db.execute("SELECT id FROM users WHERE username = 'reporter'").fetchone()[0]
            db.execute("INSERT INTO user_energy_permissions (user_id, energy_type_id) VALUES (?, 1)", (user_id,))
            db.execute("INSERT INTO report_submissions (report_month, user_id) VALUES ('2026-01', ?)", (user_id,))
            db.execute(
                "INSERT INTO energy_entries (report_month, user_id, energy_type_id, area_id, quantity) "
                "VALUES ('2026-01', ?, 1, 1, 42)", (user_id,),
            )
            db.commit()

        reporter = application.app.test_client()
        self.assertEqual(reporter.post("/login", data={"username": "reporter", "password": "1234"}).status_code, 302)
        response = self.admin.post(f"/admin/users/{user_id}/delete", follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertNotIn(b"reporter", response.data)
        self.assertEqual(reporter.get("/fill").location, "/login")
        self.assertNotEqual(reporter.post("/login", data={"username": "reporter", "password": "1234"}).location, "/fill")
        self.admin.post(f"/admin/users/{user_id}/password", data={"new_password": "changed"})
        self.admin.post(f"/admin/users/{user_id}/permissions", data={"energy_type_ids": ["2"]})

        with closing(sqlite3.connect(application.DB_PATH)) as db:
            self.assertEqual(db.execute("SELECT active FROM users WHERE id = ?", (user_id,)).fetchone()[0], 0)
            self.assertEqual(db.execute("SELECT password FROM users WHERE id = ?", (user_id,)).fetchone()[0], "1234")
            self.assertEqual(db.execute("SELECT energy_type_id FROM user_energy_permissions WHERE user_id = ?", (user_id,)).fetchone()[0], 1)
            self.assertEqual(db.execute("SELECT quantity FROM energy_entries WHERE user_id = ?", (user_id,)).fetchone()[0], 42)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM report_submissions WHERE user_id = ?", (user_id,)).fetchone()[0], 1)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM audit_logs WHERE action = 'delete_user'").fetchone()[0], 1)
        self.assertIn(b"Reporter", self.admin.get("/reports?report_month=2026-01").data)

    def test_only_admin_can_delete_filler_or_viewer(self) -> None:
        with closing(sqlite3.connect(application.DB_PATH)) as db:
            admin_id = db.execute("SELECT id FROM users WHERE username = 'admin'").fetchone()[0]
            db.execute("INSERT INTO users (username, password, display_name, role, viewer) VALUES ('guest', 'pw', 'Guest', 'filler', 1)")
            guest_id = db.execute("SELECT id FROM users WHERE username = 'guest'").fetchone()[0]
            db.commit()

        self.assertEqual(self.admin.post(f"/admin/users/{guest_id}/delete").location, "/login")
        self.admin.post("/login", data={"username": "guest", "password": "pw"})
        self.assertEqual(self.admin.post(f"/admin/users/{guest_id}/delete").location, "/fill")
        self.admin.post("/login", data={"username": "admin", "password": "1234"})
        self.assertEqual(self.admin.post(f"/admin/users/{admin_id}/delete").location, "/admin/users")
        self.assertEqual(self.admin.post(f"/admin/users/{guest_id}/delete").location, "/admin/users")
        with closing(sqlite3.connect(application.DB_PATH)) as db:
            self.assertEqual(db.execute("SELECT active FROM users WHERE id = ?", (admin_id,)).fetchone()[0], 1)
            self.assertEqual(db.execute("SELECT active FROM users WHERE id = ?", (guest_id,)).fetchone()[0], 0)

    def test_account_settings_open_only_after_selection(self) -> None:
        self.admin.post("/login", data={"username": "admin", "password": "1234"})
        self.admin.post("/admin/users", data={"username": "reporter", "account_type": "filler"})
        self.admin.post("/admin/users", data={"username": "observer", "account_type": "viewer"})
        with closing(sqlite3.connect(application.DB_PATH)) as db:
            reporter_id = db.execute("SELECT id FROM users WHERE username = 'reporter'").fetchone()[0]
            observer_id = db.execute("SELECT id FROM users WHERE username = 'observer'").fetchone()[0]

        list_page = self.admin.get("/admin/users").get_data(as_text=True)
        self.assertIn('name="user_id"', list_page)
        self.assertNotIn(f"/admin/users/{reporter_id}/permissions", list_page)
        self.assertNotIn(f"/admin/users/{observer_id}/delete", list_page)

        reporter_page = self.admin.get(f"/admin/users?user_id={reporter_id}").get_data(as_text=True)
        self.assertIn(f"/admin/users/{reporter_id}/permissions", reporter_page)
        self.assertNotIn(f"/admin/users/{observer_id}/delete", reporter_page)
        response = self.admin.post(
            f"/admin/users/{reporter_id}/permissions",
            data={"energy_type_ids": ["1"], "area_ids": ["1"]},
        )
        self.assertEqual(response.location, f"/admin/users?user_id={reporter_id}")
        self.assertIn('name="energy_type_ids"', self.admin.get(response.location).get_data(as_text=True))

        observer_page = self.admin.get(f"/admin/users?user_id={observer_id}").get_data(as_text=True)
        self.assertIn(f"/admin/users/{observer_id}/delete", observer_page)
        self.assertNotIn(f"/admin/users/{reporter_id}/permissions", observer_page)
        self.assertNotIn(f"/admin/users/{observer_id}/permissions", observer_page)
        self.admin.post(f"/admin/users/{observer_id}/delete")
        self.assertNotIn("observer", self.admin.get("/admin/users").get_data(as_text=True))


if __name__ == "__main__":
    unittest.main()

from __future__ import annotations

import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path

import app as application


class FirstLoginPasswordTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = application.DB_PATH
        application.DB_PATH = Path(self.temp_dir.name) / "app.db"
        application._INITIALIZED_DB_PATH = None
        application.app.config.update(TESTING=True, SECRET_KEY="test-secret")
        self.admin = application.app.test_client()
        self.admin.get("/login")
        self.admin.post("/login", data={"username": "admin", "password": "1234"})

    def tearDown(self) -> None:
        application.DB_PATH = self.original_db_path
        application._INITIALIZED_DB_PATH = None
        self.temp_dir.cleanup()

    def create_account(self, username: str, account_type: str) -> int:
        response = self.admin.post("/admin/users", data={
            "username": username, "display_name": username.title(), "account_type": account_type,
            "password": "ignored",
        })
        with closing(sqlite3.connect(application.DB_PATH)) as db:
            row = db.execute(
                "SELECT id, password, must_change_password FROM users WHERE username = ?", (username,),
            ).fetchone()
        self.assertEqual(response.location, f"/admin/users?user_id={row[0]}")
        self.assertEqual((row[1], row[2]), ("1234", 1))
        return row[0]

    def test_first_login_requires_password_change_before_access(self) -> None:
        for username, account_type, protected_path in (
            ("reporter", "filler", "/fill"),
            ("observer", "viewer", "/reports"),
        ):
            with self.subTest(account_type=account_type):
                self.create_account(username, account_type)
                user = application.app.test_client()
                self.assertEqual(
                    user.post("/login", data={"username": username, "password": "1234"}).location,
                    "/change-password",
                )
                self.assertEqual(user.get(protected_path).location, "/change-password")
                self.assertEqual(user.post(protected_path, data={"report_month": "2026-01"}).location, "/change-password")
                self.assertEqual(user.get("/change-password").status_code, 200)
                self.assertEqual(user.post("/change-password", data={
                    "current_password": "wrong", "new_password": "new-secret", "confirm_password": "new-secret",
                }).status_code, 200)
                self.assertEqual(user.post("/change-password", data={
                    "current_password": "1234", "new_password": "short", "confirm_password": "short",
                }).status_code, 200)
                self.assertEqual(user.post("/change-password", data={
                    "current_password": "1234", "new_password": "1234", "confirm_password": "1234",
                }).status_code, 200)
                self.assertEqual(user.post("/change-password", data={
                    "current_password": "1234", "new_password": "new-secret", "confirm_password": "different",
                }).status_code, 200)
                self.assertEqual(user.get(protected_path).location, "/change-password")
                self.assertEqual(user.post("/change-password", data={
                    "current_password": "1234", "new_password": "new-secret", "confirm_password": "new-secret",
                }).location, "/")
                self.assertEqual(user.get(protected_path).status_code, 200)
                with closing(sqlite3.connect(application.DB_PATH)) as db:
                    self.assertEqual(
                        db.execute("SELECT password, must_change_password FROM users WHERE username = ?", (username,)).fetchone(),
                        ("new-secret", 0),
                    )
                user.post("/logout")
                self.assertEqual(user.post("/login", data={"username": username, "password": "new-secret"}).location, "/")

    def test_admin_reset_requires_change_again(self) -> None:
        user_id = self.create_account("reporter", "filler")
        user = application.app.test_client()
        user.post("/login", data={"username": "reporter", "password": "1234"})
        user.post("/change-password", data={
            "current_password": "1234", "new_password": "first-secret", "confirm_password": "first-secret",
        })
        self.assertEqual(user.get("/fill").status_code, 200)
        self.admin.post(f"/admin/users/{user_id}/password", data={"new_password": "reset-secret"})
        self.assertEqual(user.get("/fill").location, "/change-password")
        user.post("/logout")
        self.assertEqual(user.post("/login", data={"username": "reporter", "password": "reset-secret"}).location, "/change-password")

    def test_existing_database_accounts_are_not_forced_on_migration(self) -> None:
        legacy_path = Path(self.temp_dir.name) / "legacy.db"
        with closing(sqlite3.connect(legacy_path)) as db:
            db.execute(
                "CREATE TABLE users (id INTEGER PRIMARY KEY, username TEXT UNIQUE, password TEXT, "
                "display_name TEXT, role TEXT, viewer INTEGER DEFAULT 0, active INTEGER DEFAULT 1, "
                "created_at TEXT DEFAULT CURRENT_TIMESTAMP)"
            )
            db.execute(
                "INSERT INTO users (username, password, display_name, role) "
                "VALUES ('existing', 'old-secret', 'Existing', 'filler')"
            )
            db.commit()
        application.DB_PATH = legacy_path
        application._INITIALIZED_DB_PATH = None
        client = application.app.test_client()
        client.get("/login")
        self.assertEqual(client.post("/login", data={"username": "existing", "password": "old-secret"}).location, "/")
        with closing(sqlite3.connect(legacy_path)) as db:
            self.assertEqual(db.execute("SELECT must_change_password FROM users WHERE username = 'existing'").fetchone()[0], 0)


if __name__ == "__main__":
    unittest.main()

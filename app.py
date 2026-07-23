from __future__ import annotations

import sqlite3
from datetime import date
from pathlib import Path

from flask import Flask, flash, g, redirect, render_template, request, session, url_for


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "instance" / "app.db"


app = Flask(__name__)
app.config["SECRET_KEY"] = "dev-secret-change-me"


DEFAULT_ENERGY_TYPES = [
    ("產量(T)", "T"),
    ("水(in-自來水)", "公噸"),
    ("水(out-排放水)", "公噸"),
    ("水(in-湖水)", "公噸"),
    ("氮氣(液氮)", "公噸"),
    ("電(台電、供電局)", "度"),
    ("電(尖峰)", "度"),
    ("電(離峰)", "度"),
    ("電(正常)", "度"),
    ("蒸汽(自產-北廠熱煤鍋爐)", "公噸"),
    ("蒸汽(自產-南廠蒸汽鍋爐)", "公噸"),
    ("煤碳(北廠熱煤鍋爐)", "公噸"),
    ("煤碳(南廠蒸汽鍋爐)", "公噸"),
    ("煤碳-蒸汽(南廠蒸汽鍋爐)", "公噸"),
    ("煤碳-熱媒(北廠熱煤鍋爐)", "公噸"),
    ("木屑粒(北廠熱煤鍋爐)", "公噸"),
    ("天然氣(北廠熱媒鍋爐)", "Sm3"),
    ("稻殼粒(南廠蒸汽鍋爐)", "公噸"),
]

DEFAULT_AREAS = [
    "Poly 51",
    "SSP 51",
    "SSP Rpet",
    "Poly 52",
    "FL 51",
    "POY",
    "DTY",
    "POY造粒",
]


def current_report_month() -> str:
    return date.today().strftime("%Y-%m")


def valid_report_month(value: str) -> bool:
    try:
        date.fromisoformat(f"{value}-01")
    except ValueError:
        return False
    return len(value) == 7


def previous_report_month(value: str) -> str:
    year, month = (int(part) for part in value.split("-"))
    if month == 1:
        return f"{year - 1}-12"
    return f"{year}-{month - 1:02d}"


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        DB_PATH.parent.mkdir(exist_ok=True)
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(error: BaseException | None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db() -> None:
    db = get_db()
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            display_name TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'filler')),
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS energy_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            unit TEXT NOT NULL,
            active INTEGER NOT NULL DEFAULT 1,
            display_order INTEGER NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS areas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            active INTEGER NOT NULL DEFAULT 1,
            display_order INTEGER NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS user_energy_permissions (
            user_id INTEGER NOT NULL,
            energy_type_id INTEGER NOT NULL,
            PRIMARY KEY (user_id, energy_type_id),
            FOREIGN KEY (user_id) REFERENCES users(id),
            FOREIGN KEY (energy_type_id) REFERENCES energy_types(id)
        );

        CREATE TABLE IF NOT EXISTS energy_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_month TEXT NOT NULL,
            user_id INTEGER NOT NULL,
            energy_type_id INTEGER NOT NULL,
            area_id INTEGER NOT NULL,
            quantity REAL,
            amount REAL,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(report_month, user_id, energy_type_id, area_id),
            FOREIGN KEY (user_id) REFERENCES users(id),
            FOREIGN KEY (energy_type_id) REFERENCES energy_types(id),
            FOREIGN KEY (area_id) REFERENCES areas(id)
        );

        CREATE TABLE IF NOT EXISTS report_submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_month TEXT NOT NULL,
            user_id INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'submitted',
            submitted_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(report_month, user_id),
            FOREIGN KEY (user_id) REFERENCES users(id)
        );
        """
    )
    seed_defaults(db)
    db.commit()


def seed_defaults(db: sqlite3.Connection) -> None:
    db.execute(
        """
        INSERT OR IGNORE INTO users (username, password, display_name, role)
        VALUES ('admin', '1234', '系統管理者', 'admin')
        """
    )
    for index, (name, unit) in enumerate(DEFAULT_ENERGY_TYPES, start=1):
        db.execute(
            """
            INSERT OR IGNORE INTO energy_types (name, unit, display_order)
            VALUES (?, ?, ?)
            """,
            (name, unit, index),
        )
    for index, name in enumerate(DEFAULT_AREAS, start=1):
        db.execute(
            """
            INSERT OR IGNORE INTO areas (name, display_order)
            VALUES (?, ?)
            """,
            (name, index),
        )


@app.before_request
def before_request() -> None:
    init_db()


def current_user() -> sqlite3.Row | None:
    user_id = session.get("user_id")
    if not user_id:
        return None
    return get_db().execute(
        "SELECT * FROM users WHERE id = ? AND active = 1",
        (user_id,),
    ).fetchone()


def login_required():
    if current_user() is None:
        flash("請先登入。", "warning")
        return redirect(url_for("login"))
    return None


def admin_required():
    guard = login_required()
    if guard:
        return guard
    if session.get("role") != "admin":
        flash("此功能限管理者使用。", "danger")
        return redirect(url_for("fill_report"))
    return None


def active_energy_types() -> list[sqlite3.Row]:
    return get_db().execute(
        """
        SELECT * FROM energy_types
        WHERE active = 1
        ORDER BY display_order, id
        """
    ).fetchall()


def active_areas() -> list[sqlite3.Row]:
    return get_db().execute(
        """
        SELECT * FROM areas
        WHERE active = 1
        ORDER BY display_order, id
        """
    ).fetchall()


@app.route("/", methods=["GET"])
def index():
    if not session.get("user_id"):
        return redirect(url_for("login"))
    if session.get("role") == "admin":
        return redirect(url_for("admin_dashboard"))
    return redirect(url_for("fill_report"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = get_db().execute(
            """
            SELECT * FROM users
            WHERE username = ? AND password = ? AND active = 1
            """,
            (username, password),
        ).fetchone()

        if user:
            session.clear()
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["display_name"] = user["display_name"]
            session["role"] = user["role"]
            flash("登入成功。", "success")
            return redirect(url_for("index"))

        flash("帳號或密碼錯誤。", "danger")

    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    flash("已登出。", "info")
    return redirect(url_for("login"))


@app.route("/admin", methods=["GET"])
def admin_dashboard():
    guard = admin_required()
    if guard:
        return guard

    db = get_db()
    report_month = request.args.get("report_month", "").strip() or current_report_month()
    if not valid_report_month(report_month):
        flash("分析月份格式不正確。", "danger")
        return redirect(url_for("admin_dashboard", report_month=current_report_month()))

    previous_month = previous_report_month(report_month)
    current_amount = db.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM energy_entries WHERE report_month = ?",
        (report_month,),
    ).fetchone()[0]
    previous_amount = db.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM energy_entries WHERE report_month = ?",
        (previous_month,),
    ).fetchone()[0]
    month_change = (
        ((current_amount - previous_amount) / previous_amount) * 100
        if previous_amount
        else None
    )
    active_fillers = db.execute(
        "SELECT COUNT(*) FROM users WHERE role = 'filler' AND active = 1"
    ).fetchone()[0]
    submitted_fillers = db.execute(
        """
        SELECT COUNT(DISTINCT rs.user_id)
        FROM report_submissions rs
        JOIN users u ON u.id = rs.user_id
        WHERE rs.report_month = ? AND u.role = 'filler' AND u.active = 1
        """,
        (report_month,),
    ).fetchone()[0]
    valid_entries = db.execute(
        """
        SELECT COUNT(*) FROM energy_entries
        WHERE report_month = ? AND (quantity IS NOT NULL OR amount IS NOT NULL)
        """,
        (report_month,),
    ).fetchone()[0]

    monthly_rows = db.execute(
        """
        SELECT report_month, COALESCE(SUM(amount), 0) AS total_amount
        FROM energy_entries
        GROUP BY report_month
        ORDER BY report_month DESC
        LIMIT 12
        """
    ).fetchall()[::-1]
    max_month_amount = max((row["total_amount"] for row in monthly_rows), default=0)
    monthly_trend = [
        {
            "month": row["report_month"],
            "amount": row["total_amount"],
            "percent": (row["total_amount"] / max_month_amount * 100) if max_month_amount else 0,
            "selected": row["report_month"] == report_month,
        }
        for row in monthly_rows
    ]

    energy_rows = db.execute(
        """
        SELECT et.name, COALESCE(SUM(ee.amount), 0) AS total_amount
        FROM energy_entries ee
        JOIN energy_types et ON et.id = ee.energy_type_id
        WHERE ee.report_month = ? AND ee.amount IS NOT NULL
        GROUP BY et.id, et.name
        HAVING SUM(ee.amount) != 0
        ORDER BY total_amount DESC
        """,
        (report_month,),
    ).fetchall()
    energy_total = sum(row["total_amount"] for row in energy_rows)
    energy_breakdown = [
        {
            "name": row["name"],
            "amount": row["total_amount"],
            "percent": (row["total_amount"] / energy_total * 100) if energy_total else 0,
        }
        for row in energy_rows
    ]

    area_rows = db.execute(
        """
        SELECT a.name, COALESCE(SUM(ee.amount), 0) AS total_amount
        FROM areas a
        LEFT JOIN energy_entries ee
            ON ee.area_id = a.id AND ee.report_month = ? AND ee.amount IS NOT NULL
        WHERE a.active = 1
        GROUP BY a.id, a.name
        ORDER BY total_amount DESC, a.display_order, a.id
        """,
        (report_month,),
    ).fetchall()
    max_area_amount = max((row["total_amount"] for row in area_rows), default=0)
    area_ranking = [
        {
            "name": row["name"],
            "amount": row["total_amount"],
            "percent": (row["total_amount"] / max_area_amount * 100) if max_area_amount else 0,
            "share": (row["total_amount"] / current_amount * 100) if current_amount else 0,
        }
        for row in area_rows
    ]
    stats = {
        "fillers": db.execute("SELECT COUNT(*) FROM users WHERE role = 'filler'").fetchone()[0],
        "energy_types": db.execute("SELECT COUNT(*) FROM energy_types WHERE active = 1").fetchone()[0],
        "areas": db.execute("SELECT COUNT(*) FROM areas WHERE active = 1").fetchone()[0],
        "entries": db.execute("SELECT COUNT(*) FROM energy_entries").fetchone()[0],
    }
    return render_template(
        "admin_dashboard.html",
        stats=stats,
        report_month=report_month,
        previous_month=previous_month,
        current_amount=current_amount,
        previous_amount=previous_amount,
        month_change=month_change,
        active_fillers=active_fillers,
        submitted_fillers=submitted_fillers,
        pending_fillers=max(active_fillers - submitted_fillers, 0),
        valid_entries=valid_entries,
        monthly_trend=monthly_trend,
        energy_breakdown=energy_breakdown,
        area_ranking=area_ranking,
    )


@app.route("/admin/users", methods=["GET", "POST"])
def admin_users():
    guard = admin_required()
    if guard:
        return guard

    db = get_db()
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        display_name = request.form.get("display_name", "").strip() or username
        if not username or not password:
            flash("請輸入帳號與密碼。", "danger")
        else:
            try:
                db.execute(
                    """
                    INSERT INTO users (username, password, display_name, role)
                    VALUES (?, ?, ?, 'filler')
                    """,
                    (username, password, display_name),
                )
                db.commit()
                flash("填寫員已建立。", "success")
            except sqlite3.IntegrityError:
                flash("帳號已存在。", "danger")
        return redirect(url_for("admin_users"))

    users = db.execute(
        """
        SELECT * FROM users
        WHERE role = 'filler'
        ORDER BY id DESC
        """
    ).fetchall()
    permissions = {
        row["user_id"]: set(row["energy_ids"].split(",") if row["energy_ids"] else [])
        for row in db.execute(
            """
            SELECT user_id, GROUP_CONCAT(energy_type_id) AS energy_ids
            FROM user_energy_permissions
            GROUP BY user_id
            """
        ).fetchall()
    }
    return render_template(
        "admin_users.html",
        users=users,
        energy_types=active_energy_types(),
        permissions=permissions,
    )


@app.route("/admin/users/<int:user_id>/permissions", methods=["POST"])
def update_user_permissions(user_id: int):
    guard = admin_required()
    if guard:
        return guard

    db = get_db()
    user = db.execute(
        "SELECT * FROM users WHERE id = ? AND role = 'filler'",
        (user_id,),
    ).fetchone()
    if user is None:
        flash("找不到填寫員。", "warning")
        return redirect(url_for("admin_users"))

    selected = request.form.getlist("energy_type_ids")
    db.execute("DELETE FROM user_energy_permissions WHERE user_id = ?", (user_id,))
    db.executemany(
        """
        INSERT INTO user_energy_permissions (user_id, energy_type_id)
        VALUES (?, ?)
        """,
        [(user_id, energy_id) for energy_id in selected],
    )
    db.commit()
    flash("填寫權限已更新。", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/users/<int:user_id>/password", methods=["POST"])
def reset_user_password(user_id: int):
    guard = admin_required()
    if guard:
        return guard

    new_password = request.form.get("new_password", "").strip()
    if not new_password:
        flash("請輸入新密碼。", "danger")
        return redirect(url_for("admin_users"))

    result = get_db().execute(
        """
        UPDATE users
        SET password = ?
        WHERE id = ? AND role = 'filler'
        """,
        (new_password, user_id),
    )
    get_db().commit()
    if result.rowcount:
        flash("密碼已重設。", "success")
    else:
        flash("找不到填寫員。", "warning")
    return redirect(url_for("admin_users"))


@app.route("/admin/settings", methods=["GET", "POST"])
def admin_settings():
    guard = admin_required()
    if guard:
        return guard

    db = get_db()
    action = request.form.get("action")
    if request.method == "POST" and action == "add_energy":
        name = request.form.get("name", "").strip()
        unit = request.form.get("unit", "").strip()
        if not name or not unit:
            flash("請輸入能源種類與單位。", "danger")
        else:
            try:
                next_order = db.execute("SELECT COALESCE(MAX(display_order), 0) + 1 FROM energy_types").fetchone()[0]
                db.execute(
                    "INSERT INTO energy_types (name, unit, display_order) VALUES (?, ?, ?)",
                    (name, unit, next_order),
                )
                db.commit()
                flash("能源種類已新增。", "success")
            except sqlite3.IntegrityError:
                flash("能源種類已存在。", "danger")
        return redirect(url_for("admin_settings"))

    if request.method == "POST" and action == "add_area":
        name = request.form.get("name", "").strip()
        if not name:
            flash("請輸入區域名稱。", "danger")
        else:
            try:
                next_order = db.execute("SELECT COALESCE(MAX(display_order), 0) + 1 FROM areas").fetchone()[0]
                db.execute(
                    "INSERT INTO areas (name, display_order) VALUES (?, ?)",
                    (name, next_order),
                )
                db.commit()
                flash("區域已新增。", "success")
            except sqlite3.IntegrityError:
                flash("區域已存在。", "danger")
        return redirect(url_for("admin_settings"))

    return render_template(
        "admin_settings.html",
        energy_types=active_energy_types(),
        areas=active_areas(),
    )


@app.route("/fill", methods=["GET", "POST"])
def fill_report():
    guard = login_required()
    if guard:
        return guard
    if session.get("role") == "admin":
        return redirect(url_for("admin_dashboard"))

    db = get_db()
    user_id = session["user_id"]
    report_month = (
        request.form.get("report_month", "").strip()
        if request.method == "POST"
        else request.args.get("report_month", "").strip()
    ) or current_report_month()
    if not valid_report_month(report_month):
        flash("填報月份格式不正確。", "danger")
        return redirect(url_for("fill_report", report_month=current_report_month()))

    energy_types = db.execute(
        """
        SELECT et.*
        FROM energy_types et
        JOIN user_energy_permissions p ON p.energy_type_id = et.id
        WHERE p.user_id = ? AND et.active = 1
        ORDER BY et.display_order, et.id
        """,
        (user_id,),
    ).fetchall()
    areas = active_areas()
    submission = db.execute(
        """
        SELECT * FROM report_submissions
        WHERE report_month = ? AND user_id = ?
        """,
        (report_month, user_id),
    ).fetchone()
    is_submitted = submission is not None

    if request.method == "POST":
        if not energy_types:
            flash("目前尚未被指派可填寫的能源種類，無法送出。", "warning")
            return redirect(url_for("fill_report", report_month=report_month))

        for energy_type in energy_types:
            for area in areas:
                quantity = request.form.get(f"quantity_{energy_type['id']}_{area['id']}", "").strip()
                amount = request.form.get(f"amount_{energy_type['id']}_{area['id']}", "").strip()
                db.execute(
                    """
                    INSERT INTO energy_entries (
                        report_month, user_id, energy_type_id, area_id, quantity, amount, updated_at
                    )
                    VALUES (?, ?, ?, ?, NULLIF(?, ''), NULLIF(?, ''), CURRENT_TIMESTAMP)
                    ON CONFLICT(report_month, user_id, energy_type_id, area_id)
                    DO UPDATE SET
                        quantity = excluded.quantity,
                        amount = excluded.amount,
                        updated_at = CURRENT_TIMESTAMP
                    """,
                    (report_month, user_id, energy_type["id"], area["id"], quantity, amount),
                )
        db.execute(
            """
            INSERT INTO report_submissions (report_month, user_id)
            VALUES (?, ?)
            ON CONFLICT(report_month, user_id)
            DO UPDATE SET
                submitted_at = CURRENT_TIMESTAMP,
                status = 'submitted'
            """,
            (report_month, user_id),
        )
        db.commit()
        flash(f"{report_month} 能源資料已送出。若需修正，可再次編輯後重新送出。", "success")
        return redirect(url_for("fill_report", report_month=report_month))

    entries = {
        (row["energy_type_id"], row["area_id"]): row
        for row in db.execute(
            """
            SELECT * FROM energy_entries
            WHERE report_month = ? AND user_id = ?
            """,
            (report_month, user_id),
        ).fetchall()
    }
    return render_template(
        "fill_report.html",
        report_month=report_month,
        energy_types=energy_types,
        areas=areas,
        entries=entries,
        is_submitted=is_submitted,
        submitted_at=submission["submitted_at"] if submission else None,
    )


@app.route("/admin/reports", methods=["GET"])
def admin_reports():
    guard = admin_required()
    if guard:
        return guard

    report_month = request.args.get("report_month") or current_report_month()
    rows = get_db().execute(
        """
        SELECT
            ee.report_month,
            u.display_name,
            rs.submitted_at,
            et.name AS energy_name,
            et.unit,
            a.name AS area_name,
            ee.quantity,
            ee.amount,
            ee.updated_at
        FROM energy_entries ee
        JOIN users u ON u.id = ee.user_id
        JOIN energy_types et ON et.id = ee.energy_type_id
        JOIN areas a ON a.id = ee.area_id
        LEFT JOIN report_submissions rs
            ON rs.report_month = ee.report_month
            AND rs.user_id = ee.user_id
        WHERE ee.report_month = ?
        ORDER BY u.display_name, et.display_order, a.display_order
        """,
        (report_month,),
    ).fetchall()
    return render_template("admin_reports.html", report_month=report_month, rows=rows)


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)

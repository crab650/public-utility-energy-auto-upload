from __future__ import annotations

import sqlite3
from datetime import date
from io import BytesIO
from pathlib import Path

from flask import Flask, flash, g, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "instance" / "app.db"
APP_VERSION = "0.6.7"

VI_TRANSLATIONS = {
    "能源資料填報系統": "Hệ thống khai báo dữ liệu năng lượng",
    "管理首頁": "Trang quản trị",
    "帳號權限管理": "Quản lý tài khoản và quyền",
    "能源/區域維護": "Quản lý năng lượng/khu vực",
    "填報查詢": "Tra cứu báo cáo",
    "資料查詢與匯出": "Tra cứu và xuất dữ liệu",
    "能源填報": "Khai báo năng lượng",
    "登出": "Đăng xuất",
    "登入": "Đăng nhập",
    "帳號": "Tài khoản",
    "密碼": "Mật khẩu",
    "管理者預設帳號": "Tài khoản quản trị mặc định",
    "能源資料輸入": "Nhập dữ liệu năng lượng",
    "填報月份": "Tháng khai báo",
    "切換月份": "Đổi tháng",
    "填報資料": "Dữ liệu khai báo",
    "此月份資料已於": "Dữ liệu tháng này đã được gửi lúc",
    "送出。若需修正，可編輯後重新送出。": "Nếu cần sửa, có thể chỉnh sửa và gửi lại.",
    "請填寫所選月份資料。按下送出後，資料會交給管理者查詢；之後仍可修正並重新送出。": "Vui lòng nhập dữ liệu của tháng đã chọn. Sau khi gửi, quản trị viên có thể tra cứu; bạn vẫn có thể sửa và gửi lại.",
    "Excel 複製貼上操作": "Cách sao chép và dán từ Excel",
    "可以直接從 Excel 複製多格資料，並一次貼到下方填報表格。": "Bạn có thể sao chép nhiều ô từ Excel và dán một lần vào bảng khai báo bên dưới.",
    "在 Excel 依照填報表格的順序選取資料範圍。": "Trong Excel, chọn vùng dữ liệu theo đúng thứ tự của bảng khai báo.",
    "按 Ctrl+C 複製選取的 Excel 資料。": "Nhấn Ctrl+C để sao chép dữ liệu đã chọn trong Excel.",
    "回到本頁，點擊要開始填入的第一個欄位。": "Quay lại trang này và nhấp vào ô đầu tiên muốn điền.",
    "按 Ctrl+V，資料會依照 Excel 的列與欄自動填入。": "Nhấn Ctrl+V; dữ liệu sẽ tự động điền theo hàng và cột trong Excel.",
    "確認綠色欄位與提示訊息；若有紅色錯誤欄位，請修正後再送出。": "Kiểm tra các ô màu xanh và thông báo; nếu có ô lỗi màu đỏ, hãy sửa trước khi gửi.",
    "注意：請不要複製 Excel 的標題列，只需複製數量與金額資料。": "Lưu ý: Không sao chép hàng tiêu đề trong Excel; chỉ sao chép dữ liệu số lượng và số tiền.",
    "管制值管理": "Quản lý giới hạn kiểm soát",
    "依歷史平均值設定各能源種類與單位的管制範圍。超出範圍時會警告，但不會阻止送出。": "Thiết lập phạm vi kiểm soát theo giá trị trung bình lịch sử cho từng loại năng lượng và đơn vị. Hệ thống sẽ cảnh báo khi vượt phạm vi nhưng vẫn cho phép gửi.",
    "重新依歷史平均值計算全部項目（±10%）": "Tính lại tất cả theo trung bình lịch sử (±10%)",
    "數量管制": "Kiểm soát số lượng",
    "金額管制": "Kiểm soát số tiền",
    "歷史平均": "Trung bình lịch sử",
    "下限": "Giới hạn dưới",
    "上限": "Giới hạn trên",
    "啟用": "Bật",
    "儲存管制值": "Lưu giới hạn kiểm soát",
    "尚無歷史資料": "Chưa có dữ liệu lịch sử",
    "輸入值超出管制範圍": "Giá trị nhập vượt phạm vi kiểm soát",
    "筆資料超出管制範圍，仍可送出。": "giá trị vượt phạm vi kiểm soát; vẫn có thể gửi.",
    "管制值已更新。": "Đã cập nhật giới hạn kiểm soát.",
    "管制值已依歷史平均值重新計算。": "Đã tính lại giới hạn theo trung bình lịch sử.",
    "管制值格式不正確，請確認上下限。": "Định dạng giới hạn không hợp lệ; vui lòng kiểm tra giới hạn dưới và trên.",
    "下限不可大於上限。": "Giới hạn dưới không được lớn hơn giới hạn trên.",
    "確定要送出": "Bạn có chắc muốn gửi dữ liệu tháng",
    "的填報資料嗎？": "không?",
    "送出後仍可修改並重新送出。": "Sau khi gửi, bạn vẫn có thể chỉnh sửa và gửi lại.",
    "目前有資料超出管制範圍，確定仍要送出嗎？": "Hiện có dữ liệu vượt phạm vi kiểm soát. Bạn vẫn muốn gửi?",
    "超出管制值明細：": "Chi tiết vượt giới hạn kiểm soát:",
    "輸入值：": "Giá trị nhập:",
    "管制範圍：": "Phạm vi kiểm soát:",
    "點擊可移至該欄位": "Nhấp để chuyển đến ô này",
    "目前尚未被指派可填寫的能源種類或單位，請洽管理者設定。": "Bạn chưa được cấp quyền cho loại năng lượng hoặc đơn vị. Vui lòng liên hệ quản trị viên.",
    "能源種類 / 單位": "Loại năng lượng / Đơn vị",
    "數量": "Số lượng",
    "金額(未稅)": "Số tiền (chưa thuế)",
    "重新送出填報資料": "Gửi lại dữ liệu",
    "送出填報資料": "Gửi dữ liệu",
    "請輸入有效數字": "Vui lòng nhập số hợp lệ",
    "貼上的內容不是有效數字。": "Nội dung dán không phải là số hợp lệ.",
    "不是有效數字": "không phải là số hợp lệ",
    "已填入": "Đã điền",
    "格超出表格範圍": "ô vượt ngoài phạm vi bảng",
    "格": "ô",
    "Excel 貼上成功，共填入": "Dán từ Excel thành công, tổng cộng",
    "帳號與權限管理": "Quản lý tài khoản và quyền",
    "新增填寫員或訪客": "Thêm nhân viên nhập liệu hoặc khách",
    "顯示名稱": "Tên hiển thị",
    "帳號權限": "Quyền tài khoản",
    "填寫員": "Nhân viên nhập liệu",
    "訪客（僅查看及匯出）": "Khách (chỉ xem và xuất)",
    "建立": "Tạo",
    "訪客（唯讀）": "Khách (chỉ đọc)",
    "Y 軸：可填寫能源種類": "Trục Y: Loại năng lượng được phép nhập",
    "X 軸：可填寫單位": "Trục X: Đơn vị được phép nhập",
    "儲存權限": "Lưu quyền",
    "此帳號可查看所有年份資料並匯出 Excel，不可填報或修改資料。": "Tài khoản này có thể xem dữ liệu mọi năm và xuất Excel, nhưng không thể nhập hoặc sửa dữ liệu.",
    "重設密碼": "Đặt lại mật khẩu",
    "輸入新密碼": "Nhập mật khẩu mới",
    "尚未建立填寫員。": "Chưa có tài khoản.",
    "指定月份（優先）": "Chọn tháng (ưu tiên)",
    "或選擇年份": "Hoặc chọn năm",
    "全部年份": "Tất cả các năm",
    "查詢": "Tra cứu",
    "全部資料": "Tất cả dữ liệu",
    "匯出 Excel": "Xuất Excel",
    "月份": "Tháng",
    "能源種類": "Loại năng lượng",
    "能源單位": "Đơn vị năng lượng",
    "單位": "Đơn vị",
    "區域": "Khu vực",
    "送出時間": "Thời gian gửi",
    "更新時間": "Thời gian cập nhật",
    "查詢範圍內尚無填報資料。": "Không có dữ liệu trong phạm vi tìm kiếm.",
    "新增能源種類": "Thêm loại năng lượng",
    "例如：水(in-自來水)": "Ví dụ: Nước (nước máy vào)",
    "例如：公噸": "Ví dụ: tấn",
    "新增區域": "Thêm khu vực",
    "區域名稱": "Tên khu vực",
    "例如：Poly 51": "Ví dụ: Poly 51",
    "目前能源種類": "Các loại năng lượng hiện tại",
    "排序": "Thứ tự",
    "目前區域": "Các khu vực hiện tại",
    "能源分析看板": "Bảng phân tích năng lượng",
    "掌握各月份能源費用、填報進度與主要成本分布。": "Theo dõi chi phí năng lượng, tiến độ khai báo và phân bổ chi phí chính theo tháng.",
    "分析月份": "Tháng phân tích",
    "更新": "Cập nhật",
    "能源總金額（未稅）": "Tổng chi phí năng lượng (chưa thuế)",
    "筆有效填報資料": "bản ghi hợp lệ",
    "較": "So với",
    "增減": "thay đổi",
    "上月": "Tháng trước",
    "上月沒有可比較資料": "Không có dữ liệu tháng trước để so sánh",
    "已填報人數": "Số người đã khai báo",
    "依有效填寫員帳號計算": "Tính theo tài khoản nhập liệu đang hoạt động",
    "尚未填報人數": "Số người chưa khai báo",
    "仍需追蹤填報進度": "Cần tiếp tục theo dõi tiến độ",
    "本月皆已完成送出": "Tất cả đã gửi trong tháng này",
    "每月能源費用趨勢": "Xu hướng chi phí năng lượng hàng tháng",
    "最近 12 個有資料的月份": "12 tháng gần nhất có dữ liệu",
    "目前尚無趨勢資料。": "Hiện chưa có dữ liệu xu hướng.",
    "能源費用占比": "Tỷ trọng chi phí năng lượng",
    "依未稅金額排序": "Sắp xếp theo số tiền chưa thuế",
    "此月份尚無能源金額資料。": "Tháng này chưa có dữ liệu chi phí năng lượng.",
    "各區域費用排名": "Xếp hạng chi phí theo khu vực",
    "含占當月總費用比例": "Bao gồm tỷ lệ trong tổng chi phí tháng",
    "目前尚未設定區域。": "Hiện chưa thiết lập khu vực.",
    "請先登入。": "Vui lòng đăng nhập trước.",
    "此功能限管理者使用。": "Chức năng này chỉ dành cho quản trị viên.",
    "此功能限管理者或訪客使用。": "Chức năng này chỉ dành cho quản trị viên hoặc khách.",
    "登入成功。": "Đăng nhập thành công.",
    "帳號或密碼錯誤。": "Tài khoản hoặc mật khẩu không đúng.",
    "已登出。": "Đã đăng xuất.",
    "分析月份格式不正確。": "Định dạng tháng phân tích không hợp lệ.",
    "請輸入帳號與密碼。": "Vui lòng nhập tài khoản và mật khẩu.",
    "填寫員已建立。": "Đã tạo tài khoản.",
    "帳號已存在。": "Tài khoản đã tồn tại.",
    "找不到填寫員。": "Không tìm thấy tài khoản.",
    "填寫權限已更新。": "Đã cập nhật quyền nhập liệu.",
    "請輸入新密碼。": "Vui lòng nhập mật khẩu mới.",
    "密碼已重設。": "Đã đặt lại mật khẩu.",
    "請輸入能源種類與單位。": "Vui lòng nhập loại năng lượng và đơn vị.",
    "能源種類已新增。": "Đã thêm loại năng lượng.",
    "能源種類已存在。": "Loại năng lượng đã tồn tại.",
    "請輸入區域名稱。": "Vui lòng nhập tên khu vực.",
    "區域已新增。": "Đã thêm khu vực.",
    "區域已存在。": "Khu vực đã tồn tại.",
    "填報月份格式不正確。": "Định dạng tháng khai báo không hợp lệ.",
    "能源資料已送出。若需修正，可再次編輯後重新送出。": "Dữ liệu năng lượng đã được gửi. Nếu cần sửa, hãy chỉnh sửa và gửi lại.",
    "能源填報資料": "Dữ liệu khai báo năng lượng",
    "年": "Năm",
    "備份資料庫": "Sao lưu cơ sở dữ liệu",
    "系統日誌": "Nhật ký hệ thống",
    "時間": "Thời gian",
    "人員": "Nhân viên",
    "動作": "Hành động",
    "詳細內容": "Chi tiết",
    "重新計算偏離值範圍：": "Tính lại phạm vi độ lệch:",
    "重新計算全部項目": "Tính lại tất cả các mục",
    "資料庫備份成功。": "Sao lưu cơ sở dữ liệu thành công.",
}


app = Flask(__name__)
app.config["SECRET_KEY"] = "dev-secret-change-me"


def translate(text: str) -> str:
    if session.get("language", "zh-TW") == "vi":
        return VI_TRANSLATIONS.get(text, text)
    return text


@app.context_processor
def inject_app_version() -> dict[str, str]:
    return {
        "app_version": APP_VERSION,
        "current_language": session.get("language", "zh-TW"),
    }


@app.after_request
def translate_html_response(response):
    if (
        session.get("language") == "vi"
        and response.content_type
        and response.content_type.startswith("text/html")
    ):
        html = response.get_data(as_text=True)
        for chinese, vietnamese in sorted(
            VI_TRANSLATIONS.items(), key=lambda item: len(item[0]), reverse=True
        ):
            html = html.replace(chinese, vietnamese)
        response.set_data(html)
    return response


@app.route("/language/<language>", methods=["POST"])
def set_language(language: str):
    if language in {"zh-TW", "vi"}:
        session["language"] = language
    next_url = request.form.get("next", "")
    if not next_url.startswith("/") or next_url.startswith("//"):
        next_url = url_for("index")
    return redirect(next_url)


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
    area_permissions_exist = db.execute(
        """
        SELECT 1 FROM sqlite_master
        WHERE type = 'table' AND name = 'user_area_permissions'
        """
    ).fetchone() is not None
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            display_name TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'filler')),
            viewer INTEGER NOT NULL DEFAULT 0,
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

        CREATE TABLE IF NOT EXISTS user_area_permissions (
            user_id INTEGER NOT NULL,
            area_id INTEGER NOT NULL,
            PRIMARY KEY (user_id, area_id),
            FOREIGN KEY (user_id) REFERENCES users(id),
            FOREIGN KEY (area_id) REFERENCES areas(id)
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

        CREATE TABLE IF NOT EXISTS control_limits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            energy_type_id INTEGER NOT NULL,
            area_id INTEGER NOT NULL,
            quantity_enabled INTEGER NOT NULL DEFAULT 0,
            quantity_average REAL,
            quantity_min REAL,
            quantity_max REAL,
            amount_enabled INTEGER NOT NULL DEFAULT 0,
            amount_average REAL,
            amount_min REAL,
            amount_max REAL,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(energy_type_id, area_id),
            FOREIGN KEY (energy_type_id) REFERENCES energy_types(id),
            FOREIGN KEY (area_id) REFERENCES areas(id)
        );

        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username TEXT,
            action TEXT NOT NULL,
            details TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        );
        """
    )
    if "viewer" not in {
        row["name"] for row in db.execute("PRAGMA table_info(users)").fetchall()
    }:
        db.execute("ALTER TABLE users ADD COLUMN viewer INTEGER NOT NULL DEFAULT 0")
    seed_defaults(db)
    seed_control_limits(db)
    if not area_permissions_exist:
        # Existing fillers used to have access to every area. Preserve that
        # behavior when introducing X-axis permissions.
        db.execute(
            """
            INSERT OR IGNORE INTO user_area_permissions (user_id, area_id)
            SELECT u.id, a.id
            FROM users u
            CROSS JOIN areas a
            WHERE u.role = 'filler' AND a.active = 1
            """
        )
    db.commit()


def seed_control_limits(db: sqlite3.Connection, replace: bool = False, percentage: float = 0.1) -> None:
    if replace:
        db.execute("DELETE FROM control_limits")
    db.execute(
        """
        WITH averages AS (
            SELECT
                energy_type_id,
                area_id,
                AVG(quantity) AS quantity_average,
                AVG(amount) AS amount_average
            FROM energy_entries
            GROUP BY energy_type_id, area_id
        )
        INSERT OR IGNORE INTO control_limits (
            energy_type_id, area_id,
            quantity_enabled, quantity_average, quantity_min, quantity_max,
            amount_enabled, amount_average, amount_min, amount_max
        )
        SELECT
            et.id,
            a.id,
            CASE WHEN av.quantity_average IS NULL THEN 0 ELSE 1 END,
            av.quantity_average,
            av.quantity_average - ABS(av.quantity_average) * ?,
            av.quantity_average + ABS(av.quantity_average) * ?,
            CASE WHEN av.amount_average IS NULL THEN 0 ELSE 1 END,
            av.amount_average,
            av.amount_average - ABS(av.amount_average) * ?,
            av.amount_average + ABS(av.amount_average) * ?
        FROM energy_types et
        CROSS JOIN areas a
        LEFT JOIN averages av
            ON av.energy_type_id = et.id
            AND av.area_id = a.id
        WHERE et.active = 1 AND a.active = 1
        """,
        (percentage, percentage, percentage, percentage),
    )


def log_action(action: str, details: str = None) -> None:
    try:
        db = get_db()
        user_id = session.get("user_id")
        username = session.get("username", "system")
        db.execute(
            "INSERT INTO audit_logs (user_id, username, action, details) VALUES (?, ?, ?, ?)",
            (user_id, username, action, details),
        )
        db.commit()
    except Exception:
        pass


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


def report_access_required():
    guard = login_required()
    if guard:
        return guard
    if session.get("role") not in {"admin", "viewer"}:
        flash("此功能限管理者或訪客使用。", "danger")
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
    if session.get("role") == "viewer":
        return redirect(url_for("admin_reports"))
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
            language = session.get("language", "zh-TW")
            session.clear()
            session["language"] = language
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["display_name"] = user["display_name"]
            session["role"] = "viewer" if user["viewer"] else user["role"]
            flash("登入成功。", "success")
            log_action("login", f"Successful login from role: {session['role']}")
            return redirect(url_for("index"))

        flash("帳號或密碼錯誤。", "danger")

    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    log_action("logout")
    language = session.get("language", "zh-TW")
    session.clear()
    session["language"] = language
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
        "SELECT COUNT(*) FROM users WHERE role = 'filler' AND viewer = 0 AND active = 1"
    ).fetchone()[0]
    submitted_fillers = db.execute(
        """
        SELECT COUNT(DISTINCT rs.user_id)
        FROM report_submissions rs
        JOIN users u ON u.id = rs.user_id
        WHERE rs.report_month = ? AND u.role = 'filler' AND u.viewer = 0 AND u.active = 1
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
        "fillers": db.execute("SELECT COUNT(*) FROM users WHERE role = 'filler' AND viewer = 0").fetchone()[0],
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
        audit_logs=db.execute("SELECT * FROM audit_logs ORDER BY id DESC LIMIT 15").fetchall(),
    )


@app.route("/admin/backup", methods=["GET"])
def admin_backup():
    guard = admin_required()
    if guard:
        return guard
    try:
        log_action("backup_db", "Downloaded database backup")
        return send_file(
            DB_PATH,
            as_attachment=True,
            download_name=f"energy_system_backup_{date.today().strftime('%Y%m%d')}.db",
            mimetype="application/x-sqlite3"
        )
    except Exception as e:
        flash(f"備份失敗: {str(e)}", "danger")
        return redirect(url_for("admin_dashboard"))


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
        account_type = request.form.get("account_type", "filler")
        if account_type not in {"filler", "viewer"}:
            account_type = "filler"
        if not username or not password:
            flash("請輸入帳號與密碼。", "danger")
        else:
            try:
                db.execute(
                    """
                    INSERT INTO users (username, password, display_name, role, viewer)
                    VALUES (?, ?, ?, 'filler', ?)
                    """,
                    (username, password, display_name, account_type == "viewer"),
                )
                user_id = db.execute(
                    "SELECT id FROM users WHERE username = ?",
                    (username,),
                ).fetchone()["id"]
                if account_type == "filler":
                    db.execute(
                        """
                        INSERT INTO user_area_permissions (user_id, area_id)
                        SELECT ?, id FROM areas WHERE active = 1
                        """,
                        (user_id,),
                    )
                db.commit()
                log_action("create_user", f"Created account {username} as {account_type}")
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
    energy_permissions = {
        row["user_id"]: set(row["energy_ids"].split(",") if row["energy_ids"] else [])
        for row in db.execute(
            """
            SELECT user_id, GROUP_CONCAT(energy_type_id) AS energy_ids
            FROM user_energy_permissions
            GROUP BY user_id
            """
        ).fetchall()
    }
    area_permissions = {
        row["user_id"]: set(row["area_ids"].split(",") if row["area_ids"] else [])
        for row in db.execute(
            """
            SELECT user_id, GROUP_CONCAT(area_id) AS area_ids
            FROM user_area_permissions
            GROUP BY user_id
            """
        ).fetchall()
    }
    return render_template(
        "admin_users.html",
        users=users,
        energy_types=active_energy_types(),
        areas=active_areas(),
        energy_permissions=energy_permissions,
        area_permissions=area_permissions,
    )


@app.route("/admin/users/<int:user_id>/permissions", methods=["POST"])
def update_user_permissions(user_id: int):
    guard = admin_required()
    if guard:
        return guard

    db = get_db()
    user = db.execute(
        "SELECT * FROM users WHERE id = ? AND role = 'filler' AND viewer = 0",
        (user_id,),
    ).fetchone()
    if user is None:
        flash("找不到填寫員。", "warning")
        return redirect(url_for("admin_users"))

    selected_energy_ids = request.form.getlist("energy_type_ids")
    selected_area_ids = request.form.getlist("area_ids")
    db.execute("DELETE FROM user_energy_permissions WHERE user_id = ?", (user_id,))
    db.executemany(
        """
        INSERT INTO user_energy_permissions (user_id, energy_type_id)
        VALUES (?, ?)
        """,
        [(user_id, energy_id) for energy_id in selected_energy_ids],
    )
    db.execute("DELETE FROM user_area_permissions WHERE user_id = ?", (user_id,))
    db.executemany(
        """
        INSERT INTO user_area_permissions (user_id, area_id)
        VALUES (?, ?)
        """,
        [(user_id, area_id) for area_id in selected_area_ids],
    )
    db.commit()
    log_action("update_permissions", f"Updated Y/X permissions for filler: {user['username']}")
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
        log_action("reset_password", f"Reset password for filler ID: {user_id}")
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
                log_action("add_energy_type", f"Added: {name} (Unit: {unit})")
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
                log_action("add_area", f"Added area: {name}")
                flash("區域已新增。", "success")
            except sqlite3.IntegrityError:
                flash("區域已存在。", "danger")
        return redirect(url_for("admin_settings"))

    return render_template(
        "admin_settings.html",
        energy_types=active_energy_types(),
        areas=active_areas(),
    )


@app.route("/admin/control-limits", methods=["GET", "POST"])
def admin_control_limits():
    guard = admin_required()
    if guard:
        return guard

    db = get_db()
    if request.method == "POST" and request.form.get("action") == "recalculate":
        pct_val = request.form.get("percentage", "10")
        try:
            percentage = float(pct_val) / 100.0
        except ValueError:
            percentage = 0.1
        seed_control_limits(db, replace=True, percentage=percentage)
        db.commit()
        log_action("recalculate_control_limits", f"Recalculated with deviation ±{pct_val}%")
        flash("管制值已依歷史平均值重新計算。", "success")
        return redirect(url_for("admin_control_limits"))

    if request.method == "POST":
        rows = db.execute("SELECT id FROM control_limits ORDER BY id").fetchall()
        updates = []
        try:
            for row in rows:
                control_id = row["id"]
                quantity_enabled = f"quantity_enabled_{control_id}" in request.form
                amount_enabled = f"amount_enabled_{control_id}" in request.form

                def optional_float(field: str) -> float | None:
                    value = request.form.get(f"{field}_{control_id}", "").strip()
                    return float(value) if value else None

                quantity_min = optional_float("quantity_min")
                quantity_max = optional_float("quantity_max")
                amount_min = optional_float("amount_min")
                amount_max = optional_float("amount_max")
                if quantity_enabled and (quantity_min is None or quantity_max is None):
                    raise ValueError
                if amount_enabled and (amount_min is None or amount_max is None):
                    raise ValueError
                if quantity_min is not None and quantity_max is not None and quantity_min > quantity_max:
                    flash("下限不可大於上限。", "danger")
                    return redirect(url_for("admin_control_limits"))
                if amount_min is not None and amount_max is not None and amount_min > amount_max:
                    flash("下限不可大於上限。", "danger")
                    return redirect(url_for("admin_control_limits"))
                updates.append((
                    quantity_enabled, quantity_min, quantity_max,
                    amount_enabled, amount_min, amount_max, control_id,
                ))
        except ValueError:
            flash("管制值格式不正確，請確認上下限。", "danger")
            return redirect(url_for("admin_control_limits"))

        db.executemany(
            """
            UPDATE control_limits
            SET quantity_enabled = ?, quantity_min = ?, quantity_max = ?,
                amount_enabled = ?, amount_min = ?, amount_max = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            updates,
        )
        db.commit()
        log_action("update_control_limits", "Manually updated control limit grid")
        flash("管制值已更新。", "success")
        return redirect(url_for("admin_control_limits"))

    rows = db.execute(
        """
        SELECT cl.*, et.name AS energy_name, et.unit, a.name AS area_name
        FROM control_limits cl
        JOIN energy_types et ON et.id = cl.energy_type_id
        JOIN areas a ON a.id = cl.area_id
        WHERE et.active = 1 AND a.active = 1
        ORDER BY et.display_order, et.id, a.display_order, a.id
        """
    ).fetchall()
    return render_template("admin_control_limits.html", rows=rows)


@app.route("/fill", methods=["GET", "POST"])
def fill_report():
    guard = login_required()
    if guard:
        return guard
    if session.get("role") == "admin":
        return redirect(url_for("admin_dashboard"))
    if session.get("role") == "viewer":
        return redirect(url_for("admin_reports"))

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
    areas = db.execute(
        """
        SELECT a.*
        FROM areas a
        JOIN user_area_permissions p ON p.area_id = a.id
        WHERE p.user_id = ? AND a.active = 1
        ORDER BY a.display_order, a.id
        """,
        (user_id,),
    ).fetchall()
    control_limits = {
        (row["energy_type_id"], row["area_id"]): row
        for row in db.execute(
            """
            SELECT cl.*
            FROM control_limits cl
            JOIN user_energy_permissions ep
                ON ep.energy_type_id = cl.energy_type_id AND ep.user_id = ?
            JOIN user_area_permissions ap
                ON ap.area_id = cl.area_id AND ap.user_id = ?
            """,
            (user_id, user_id),
        ).fetchall()
    }
    submission = db.execute(
        """
        SELECT * FROM report_submissions
        WHERE report_month = ? AND user_id = ?
        """,
        (report_month, user_id),
    ).fetchone()
    is_submitted = submission is not None

    if request.method == "POST":
        if not energy_types or not areas:
            flash("目前尚未被指派可填寫的能源種類或單位，無法送出。", "warning")
            return redirect(url_for("fill_report", report_month=report_month))

        control_warning_count = 0
        for energy_type in energy_types:
            for area in areas:
                quantity = request.form.get(f"quantity_{energy_type['id']}_{area['id']}", "").strip()
                amount = request.form.get(f"amount_{energy_type['id']}_{area['id']}", "").strip()
                limit = control_limits.get((energy_type["id"], area["id"]))
                for value, enabled_key, min_key, max_key in (
                    (quantity, "quantity_enabled", "quantity_min", "quantity_max"),
                    (amount, "amount_enabled", "amount_min", "amount_max"),
                ):
                    if (
                        value and limit and limit[enabled_key]
                        and limit[min_key] is not None and limit[max_key] is not None
                    ):
                        try:
                            numeric_value = float(value)
                        except ValueError:
                            continue
                        if numeric_value < limit[min_key] or numeric_value > limit[max_key]:
                            control_warning_count += 1
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
        log_action("submit_report", f"Submitted data for {report_month} (Warnings: {control_warning_count})")
        flash(f"{report_month} 能源資料已送出。若需修正，可再次編輯後重新送出。", "success")
        if control_warning_count:
            flash(f"{control_warning_count} 筆資料超出管制範圍，仍可送出。", "warning")
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
        control_limits=control_limits,
        is_submitted=is_submitted,
        submitted_at=submission["submitted_at"] if submission else None,
    )


def report_rows(report_month: str, report_year: str) -> list[sqlite3.Row]:
    filters = []
    params = []
    if report_month:
        filters.append("ee.report_month = ?")
        params.append(report_month)
    elif report_year:
        filters.append("substr(ee.report_month, 1, 4) = ?")
        params.append(report_year)
    where_sql = f"WHERE {' AND '.join(filters)}" if filters else ""
    return get_db().execute(
        f"""
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
        {where_sql}
        ORDER BY ee.report_month DESC, u.display_name, et.display_order, a.display_order
        """,
        params,
    ).fetchall()


def report_filters() -> tuple[str, str]:
    report_month = request.args.get("report_month", "").strip()
    report_year = request.args.get("report_year", "").strip()
    if report_month and not valid_report_month(report_month):
        report_month = ""
    if report_year and (len(report_year) != 4 or not report_year.isdigit()):
        report_year = ""
    return report_month, report_year


@app.route("/reports", methods=["GET"])
@app.route("/admin/reports", methods=["GET"])
def admin_reports():
    guard = report_access_required()
    if guard:
        return guard

    report_month, report_year = report_filters()
    years = [
        row[0] for row in get_db().execute(
            """
            SELECT DISTINCT substr(report_month, 1, 4) AS year
            FROM energy_entries ORDER BY year DESC
            """
        ).fetchall()
    ]
    return render_template(
        "admin_reports.html",
        report_month=report_month,
        report_year=report_year,
        years=years,
        rows=report_rows(report_month, report_year),
    )


@app.route("/reports/export", methods=["GET"])
def export_reports():
    guard = report_access_required()
    if guard:
        return guard

    report_month, report_year = report_filters()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = translate("能源填報資料")
    sheet.append([
        translate("月份"), translate("填寫員"), translate("能源種類"),
        translate("能源單位"), translate("區域"), translate("數量"),
        translate("金額(未稅)"), translate("送出時間"), translate("更新時間"),
    ])
    for row in report_rows(report_month, report_year):
        sheet.append([
            row["report_month"], row["display_name"], row["energy_name"], row["unit"],
            row["area_name"], row["quantity"], row["amount"],
            row["submitted_at"], row["updated_at"],
        ])
    sheet.freeze_panes = "A2"
    for column in sheet.columns:
        sheet.column_dimensions[column[0].column_letter].width = min(
            max(len(str(cell.value or "")) for cell in column) + 2,
            40,
        )
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    suffix = report_month or report_year or "all"
    return send_file(
        output,
        as_attachment=True,
        download_name=f"energy-report-{suffix}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)

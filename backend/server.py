from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import os
import re
import secrets
import sqlite3
import sys
import time
import uuid
from datetime import datetime
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from xlsx_utils import read_xlsx_rows, write_xlsx_bytes


PROJECT_ROOT = Path(__file__).resolve().parents[1]
BACKEND_ROOT = Path(__file__).resolve().parent
FRONTEND_ROOT = PROJECT_ROOT / "frontend"
DATA_DIR = BACKEND_ROOT / "data"
UPLOAD_DIR = BACKEND_ROOT / "uploads"
DB_PATH = Path(os.environ.get("BIOTECH_PM_DB", DATA_DIR / "biotech_pm.sqlite3"))

DATA_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

SESSIONS: dict[str, dict] = {}

PROJECT_STAGES = [
    "样本入库",
    "样本核对",
    "建库",
    "测序",
    "数据存放",
    "生信分析",
    "合同资料",
    "回款",
    "项目结束",
]

SAMPLE_FIELDS = [
    "row_no",
    "hospital_department",
    "doctor_owner",
    "sample_code",
    "patient_name",
    "sample_category",
    "sample_subcategory",
    "group_name",
    "volume",
    "test_project",
    "sequencing_mode",
    "data_amount",
    "sample_note",
    "sample_status",
    "sequencing_date",
    "source_sample_code",
    "lab_code",
    "sample_source",
    "handling_method",
    "testing_flow",
    "extraction_concentration",
    "library_input_amount",
    "sub_library_code",
    "adapter_code",
    "umsi_code",
    "library_concentration",
    "sequencing_library_amount",
    "chip_code",
    "reagent_lot",
    "operator_name",
    "total_reads",
    "host_rate",
    "q20",
    "q30",
    "report_result",
    "detected_pathogens",
    "data_storage_path",
    "storage_notes",
    "bioinfo_status",
    "bioinfo_task_content",
]

PROJECT_FIELDS = [
    "name",
    "hospital_name",
    "customer_name",
    "research_manager",
    "oa_project_code",
    "order_code",
    "sample_type",
    "sample_nature",
    "received_sample_count",
    "sent_sample_count",
    "arrival_date",
    "priority",
    "lab_flow",
    "operation_note",
    "sequencing_plan",
    "sequencing_platform",
    "sequencing_batch_no",
    "big_pool_no",
    "small_pool_no",
    "sequencing_mode",
    "data_amount_per_sample",
    "data_location_code",
    "current_progress",
    "progress_time",
    "progress_percent",
    "status",
]


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def row_to_dict(row: sqlite3.Row | None) -> dict | None:
    if row is None:
        return None
    return {key: row[key] for key in row.keys()}


def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 200_000)
    return f"{salt}${digest.hex()}"


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        salt, _digest = stored_hash.split("$", 1)
    except ValueError:
        return False
    return secrets.compare_digest(hash_password(password, salt), stored_hash)


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def execute_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            display_name TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            status TEXT NOT NULL DEFAULT 'active',
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS register_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            display_name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'pending',
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            owner_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            hospital_name TEXT DEFAULT '',
            customer_name TEXT DEFAULT '',
            research_manager TEXT DEFAULT '',
            oa_project_code TEXT DEFAULT '',
            order_code TEXT DEFAULT '',
            sample_type TEXT DEFAULT '',
            sample_nature TEXT DEFAULT '',
            received_sample_count INTEGER DEFAULT 0,
            sent_sample_count INTEGER DEFAULT 0,
            arrival_date TEXT DEFAULT '',
            priority TEXT DEFAULT '常规',
            lab_flow TEXT DEFAULT '',
            operation_note TEXT DEFAULT '',
            sequencing_plan TEXT DEFAULT '',
            sequencing_platform TEXT DEFAULT '',
            sequencing_batch_no TEXT DEFAULT '',
            big_pool_no TEXT DEFAULT '',
            small_pool_no TEXT DEFAULT '',
            sequencing_mode TEXT DEFAULT '',
            data_amount_per_sample TEXT DEFAULT '',
            data_location_code TEXT DEFAULT '',
            current_progress TEXT DEFAULT '样本入库',
            progress_time TEXT DEFAULT '',
            progress_percent INTEGER DEFAULT 10,
            status TEXT NOT NULL DEFAULT 'active',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(owner_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS project_members (
            project_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            role TEXT NOT NULL DEFAULT 'member',
            PRIMARY KEY (project_id, user_id),
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS samples (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            row_no TEXT DEFAULT '',
            hospital_department TEXT DEFAULT '',
            doctor_owner TEXT DEFAULT '',
            sample_code TEXT DEFAULT '',
            patient_name TEXT DEFAULT '',
            sample_category TEXT DEFAULT '',
            sample_subcategory TEXT DEFAULT '',
            group_name TEXT DEFAULT '',
            volume TEXT DEFAULT '',
            test_project TEXT DEFAULT '',
            sequencing_mode TEXT DEFAULT '',
            data_amount TEXT DEFAULT '',
            sample_note TEXT DEFAULT '',
            sample_status TEXT DEFAULT '待核对',
            sequencing_date TEXT DEFAULT '',
            source_sample_code TEXT DEFAULT '',
            lab_code TEXT DEFAULT '',
            sample_source TEXT DEFAULT '',
            handling_method TEXT DEFAULT '',
            testing_flow TEXT DEFAULT '',
            extraction_concentration TEXT DEFAULT '',
            library_input_amount TEXT DEFAULT '',
            sub_library_code TEXT DEFAULT '',
            adapter_code TEXT DEFAULT '',
            umsi_code TEXT DEFAULT '',
            library_concentration TEXT DEFAULT '',
            sequencing_library_amount TEXT DEFAULT '',
            chip_code TEXT DEFAULT '',
            reagent_lot TEXT DEFAULT '',
            operator_name TEXT DEFAULT '',
            total_reads TEXT DEFAULT '',
            host_rate TEXT DEFAULT '',
            q20 TEXT DEFAULT '',
            q30 TEXT DEFAULT '',
            report_result TEXT DEFAULT '',
            detected_pathogens TEXT DEFAULT '',
            data_storage_path TEXT DEFAULT '',
            storage_notes TEXT DEFAULT '',
            bioinfo_status TEXT DEFAULT '',
            bioinfo_task_content TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS uploaded_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            module TEXT NOT NULL,
            original_name TEXT NOT NULL,
            stored_name TEXT NOT NULL,
            path TEXT NOT NULL,
            uploaded_by INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(uploaded_by) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS bioinfo_tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            content TEXT DEFAULT '',
            status TEXT NOT NULL DEFAULT '待开始',
            owner TEXT DEFAULT '',
            due_date TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL UNIQUE,
            total_amount REAL DEFAULT 0,
            received_amount REAL DEFAULT 0,
            payment_note TEXT DEFAULT '',
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS operation_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            sample_id INTEGER,
            module TEXT NOT NULL,
            action TEXT NOT NULL,
            actor TEXT NOT NULL,
            detail TEXT DEFAULT '',
            created_at TEXT NOT NULL
        );
        """
    )
    conn.commit()


def seed_data(conn: sqlite3.Connection) -> None:
    user_count = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    if user_count == 0:
        created_at = now_text()
        conn.execute(
            "INSERT INTO users (username, password_hash, display_name, role, status, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            ("admin", hash_password("admin123"), "系统管理员", "admin", "active", created_at),
        )
        conn.execute(
            "INSERT INTO users (username, password_hash, display_name, role, status, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            ("demo", hash_password("demo123"), "项目成员", "user", "active", created_at),
        )
        conn.commit()

    project_count = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
    if project_count == 0:
        admin_id = conn.execute("SELECT id FROM users WHERE username = 'admin'").fetchone()["id"]
        demo_id = conn.execute("SELECT id FROM users WHERE username = 'demo'").fetchone()["id"]
        template = PROJECT_ROOT / "4.样本到样信息汇总统计表-最终版-0528.xlsx"
        if template.exists():
            rows = read_xlsx_rows(template, "样本到样信息汇总统计", max_rows=60)
            for row in rows[1:]:
                padded = row + [""] * 29
                if not any(padded[:5]):
                    continue
                name = f"{padded[3] or padded[4]} {padded[0]} {padded[1]}".strip()
                progress = padded[21] or "样本入库"
                now = now_text()
                cur = conn.execute(
                    """
                    INSERT INTO projects (
                        owner_id, name, hospital_name, customer_name, research_manager,
                        oa_project_code, order_code, sample_type, sample_nature,
                        received_sample_count, sent_sample_count, arrival_date, priority,
                        lab_flow, operation_note, sequencing_plan, sequencing_platform,
                        sequencing_batch_no, big_pool_no, small_pool_no, sequencing_mode,
                        data_amount_per_sample, data_location_code, current_progress,
                        progress_time, progress_percent, status, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        admin_id,
                        name,
                        padded[0],
                        padded[1],
                        padded[2],
                        padded[3],
                        padded[4],
                        padded[5],
                        padded[6],
                        int(float(padded[7] or 0)),
                        int(float(padded[8] or 0)),
                        padded[9],
                        padded[10] or "常规",
                        padded[11],
                        padded[12],
                        padded[13],
                        padded[14],
                        padded[15],
                        padded[16],
                        padded[17],
                        padded[18],
                        padded[19],
                        padded[20],
                        progress,
                        padded[22],
                        progress_to_percent(progress),
                        "active",
                        now,
                        now,
                    ),
                )
                project_id = cur.lastrowid
                conn.execute(
                    "INSERT OR IGNORE INTO project_members (project_id, user_id, role) VALUES (?, ?, ?)",
                    (project_id, demo_id, "member"),
                )
        conn.commit()

    sample_count = conn.execute("SELECT COUNT(*) FROM samples").fetchone()[0]
    if sample_count == 0:
        first_project = conn.execute("SELECT id FROM projects ORDER BY id LIMIT 1").fetchone()
        template = PROJECT_ROOT / "1.样本信息表.xlsx"
        if first_project and template.exists():
            rows = read_xlsx_rows(template, "返样信息单", max_rows=20)
            for row in rows[1:]:
                values = row + [""] * 13
                if not values[3] or values[0].startswith("填写说明"):
                    continue
                insert_sample(
                    conn,
                    first_project["id"],
                    {
                        "row_no": values[0],
                        "hospital_department": values[1],
                        "doctor_owner": values[2],
                        "sample_code": values[3],
                        "patient_name": values[4],
                        "sample_category": values[5],
                        "sample_subcategory": values[6],
                        "group_name": values[7],
                        "volume": values[8],
                        "test_project": values[9],
                        "sequencing_mode": values[10],
                        "data_amount": values[11],
                        "sample_note": values[12],
                        "sample_status": "待核对",
                    },
                )
            conn.commit()


def progress_to_percent(progress: str) -> int:
    text = progress or ""
    if "结束" in text or "完成" in text:
        return 100
    if "回款" in text:
        return 90
    if "合同" in text:
        return 80
    if "生信" in text:
        return 70
    if "数据" in text or "下机" in text:
        return 60
    if "测序" in text or "pooling" in text or "排机" in text:
        return 50
    if "文库" in text or "建库" in text:
        return 40
    if "提取" in text:
        return 25
    if "核对" in text:
        return 20
    return 10


def init_db() -> None:
    with connect() as conn:
        execute_schema(conn)
        seed_data(conn)


def log_operation(
    conn: sqlite3.Connection,
    user: dict,
    module: str,
    action: str,
    project_id: int | None = None,
    sample_id: int | None = None,
    detail: str = "",
) -> None:
    conn.execute(
        """
        INSERT INTO operation_logs (project_id, sample_id, module, action, actor, detail, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (project_id, sample_id, module, action, user.get("display_name") or user.get("username"), detail, now_text()),
    )


def insert_sample(conn: sqlite3.Connection, project_id: int, data: dict) -> int:
    payload = clean_payload(data, SAMPLE_FIELDS)
    payload["project_id"] = project_id
    payload["created_at"] = now_text()
    payload["updated_at"] = now_text()
    columns = list(payload.keys())
    placeholders = ", ".join("?" for _ in columns)
    cur = conn.execute(
        f"INSERT INTO samples ({', '.join(columns)}) VALUES ({placeholders})",
        [payload[column] for column in columns],
    )
    return int(cur.lastrowid)


def update_sample(conn: sqlite3.Connection, sample_id: int, data: dict) -> None:
    payload = clean_payload(data, SAMPLE_FIELDS)
    payload["updated_at"] = now_text()
    assignments = ", ".join(f"{column} = ?" for column in payload)
    conn.execute(
        f"UPDATE samples SET {assignments} WHERE id = ?",
        [payload[column] for column in payload] + [sample_id],
    )


def clean_payload(data: dict, allowed_fields: list[str]) -> dict:
    payload: dict[str, object] = {}
    for field in allowed_fields:
        if field in data:
            value = data[field]
            if value is None:
                value = ""
            payload[field] = value
    return payload


def safe_filename(filename: str) -> str:
    name = Path(filename or "upload.bin").name
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    return name or "upload.bin"


def save_upload(
    conn: sqlite3.Connection,
    user: dict,
    project_id: int,
    module: str,
    filename: str,
    content: bytes,
) -> dict:
    original = safe_filename(filename)
    suffix = Path(original).suffix
    stored_name = f"{module}_{uuid.uuid4().hex}{suffix}"
    module_dir = UPLOAD_DIR / str(project_id)
    module_dir.mkdir(parents=True, exist_ok=True)
    file_path = module_dir / stored_name
    file_path.write_bytes(content)
    created_at = now_text()
    cur = conn.execute(
        """
        INSERT INTO uploaded_files (project_id, module, original_name, stored_name, path, uploaded_by, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (project_id, module, original, stored_name, str(file_path), user["id"], created_at),
    )
    return {
        "id": cur.lastrowid,
        "project_id": project_id,
        "module": module,
        "original_name": original,
        "stored_name": stored_name,
        "url": f"/api/files/{cur.lastrowid}/download",
        "created_at": created_at,
    }


def sample_match(conn: sqlite3.Connection, project_id: int, values: dict) -> sqlite3.Row | None:
    lab_code = values.get("lab_code") or ""
    sample_code = values.get("sample_code") or values.get("source_sample_code") or ""
    sub_library_code = values.get("sub_library_code") or ""
    if lab_code:
        row = conn.execute(
            "SELECT * FROM samples WHERE project_id = ? AND lab_code = ? ORDER BY id LIMIT 1",
            (project_id, lab_code),
        ).fetchone()
        if row:
            return row
    if sample_code:
        row = conn.execute(
            "SELECT * FROM samples WHERE project_id = ? AND sample_code = ? ORDER BY id LIMIT 1",
            (project_id, sample_code),
        ).fetchone()
        if row:
            return row
    if sub_library_code:
        return conn.execute(
            "SELECT * FROM samples WHERE project_id = ? AND sub_library_code = ? ORDER BY id LIMIT 1",
            (project_id, sub_library_code),
        ).fetchone()
    return None


def recalculate_project_progress(conn: sqlite3.Connection, project_id: int) -> None:
    sample_count = conn.execute("SELECT COUNT(*) FROM samples WHERE project_id = ?", (project_id,)).fetchone()[0]
    library_count = conn.execute(
        "SELECT COUNT(*) FROM samples WHERE project_id = ? AND sub_library_code != ''",
        (project_id,),
    ).fetchone()[0]
    sequencing_count = conn.execute(
        "SELECT COUNT(*) FROM samples WHERE project_id = ? AND (total_reads != '' OR chip_code != '')",
        (project_id,),
    ).fetchone()[0]
    storage_count = conn.execute(
        "SELECT COUNT(*) FROM samples WHERE project_id = ? AND data_storage_path != ''",
        (project_id,),
    ).fetchone()[0]
    task_count = conn.execute("SELECT COUNT(*) FROM bioinfo_tasks WHERE project_id = ?", (project_id,)).fetchone()[0]
    contract_count = conn.execute(
        "SELECT COUNT(*) FROM uploaded_files WHERE project_id = ? AND module = 'contract'",
        (project_id,),
    ).fetchone()[0]
    payment = conn.execute("SELECT total_amount, received_amount FROM payments WHERE project_id = ?", (project_id,)).fetchone()
    project = conn.execute("SELECT status FROM projects WHERE id = ?", (project_id,)).fetchone()

    percent = 10
    progress = "样本入库"
    if sample_count:
        percent, progress = 20, "样本核对"
    if library_count:
        percent, progress = 40, "建库"
    if sequencing_count:
        percent, progress = 55, "测序"
    if storage_count:
        percent, progress = 65, "数据存放"
    if task_count:
        percent, progress = 75, "生信分析"
    if contract_count:
        percent, progress = 85, "合同资料"
    if payment and payment["total_amount"] and payment["received_amount"] >= payment["total_amount"]:
        percent, progress = 95, "回款"
    if project and project["status"] == "completed":
        percent, progress = 100, "项目结束"

    conn.execute(
        """
        UPDATE projects
        SET progress_percent = ?, current_progress = ?, progress_time = ?, updated_at = ?
        WHERE id = ?
        """,
        (percent, progress, datetime.now().strftime("%Y-%m-%d"), now_text(), project_id),
    )


class ApiHandler(SimpleHTTPRequestHandler):
    server_version = "BiotechPM/0.1"
    protocol_version = "HTTP/1.1"

    def handle_expect_100(self) -> bool:
        self.send_response_only(HTTPStatus.CONTINUE)
        self.end_headers()
        return True

    def end_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
        super().end_headers()

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.end_headers()

    def do_GET(self) -> None:
        self.dispatch("GET")

    def do_POST(self) -> None:
        self.dispatch("POST")

    def do_PUT(self) -> None:
        self.dispatch("PUT")

    def do_DELETE(self) -> None:
        self.dispatch("DELETE")

    def dispatch(self, method: str) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        try:
            if path.startswith("/api/"):
                self.handle_api(method, path)
            else:
                self.serve_frontend(path)
        except ApiError as exc:
            self.send_json(exc.status, {"ok": False, "error": exc.message})
        except Exception as exc:
            print(f"ERROR {method} {path}: {exc}", file=sys.stderr)
            self.send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"ok": False, "error": str(exc)})

    def handle_api(self, method: str, path: str) -> None:
        if method == "GET" and path == "/api/health":
            self.send_json(HTTPStatus.OK, {"ok": True, "time": now_text()})
            return

        if path == "/api/auth/login" and method == "POST":
            self.login()
            return
        if path == "/api/auth/register-request" and method == "POST":
            self.register_request()
            return
        if path == "/api/auth/forgot-password" and method == "POST":
            self.forgot_password()
            return

        with connect() as conn:
            user = self.require_user(conn)

            if path == "/api/me" and method == "GET":
                self.send_json(HTTPStatus.OK, {"ok": True, "user": public_user(user)})
                return

            if path == "/api/projects" and method == "GET":
                self.list_projects(conn, user)
                return
            if path == "/api/projects" and method == "POST":
                self.create_project(conn, user)
                return
            if path == "/api/projects/export-summary" and method == "GET":
                self.export_summary(conn, user)
                return

            match = re.fullmatch(r"/api/projects/(\d+)", path)
            if match and method == "GET":
                project = self.require_project(conn, user, int(match.group(1)))
                self.get_project(conn, project)
                return
            if match and method == "PUT":
                project = self.require_project(conn, user, int(match.group(1)))
                self.update_project(conn, user, project["id"])
                return

            match = re.fullmatch(r"/api/projects/(\d+)/complete", path)
            if match and method == "POST":
                project = self.require_project(conn, user, int(match.group(1)))
                self.complete_project(conn, user, project["id"])
                return

            match = re.fullmatch(r"/api/projects/(\d+)/samples", path)
            if match and method == "GET":
                project = self.require_project(conn, user, int(match.group(1)))
                self.list_samples(conn, project["id"])
                return
            if match and method == "POST":
                project = self.require_project(conn, user, int(match.group(1)))
                self.create_sample(conn, user, project["id"])
                return

            match = re.fullmatch(r"/api/projects/(\d+)/samples/import", path)
            if match and method == "POST":
                project = self.require_project(conn, user, int(match.group(1)))
                self.import_samples(conn, user, project["id"])
                return

            match = re.fullmatch(r"/api/projects/(\d+)/library/import", path)
            if match and method == "POST":
                project = self.require_project(conn, user, int(match.group(1)))
                self.import_library(conn, user, project["id"])
                return

            match = re.fullmatch(r"/api/samples/(\d+)", path)
            if match and method == "PUT":
                sample = self.require_sample(conn, user, int(match.group(1)))
                self.update_sample_endpoint(conn, user, sample)
                return
            if match and method == "DELETE":
                sample = self.require_sample(conn, user, int(match.group(1)))
                self.delete_sample(conn, user, sample)
                return

            match = re.fullmatch(r"/api/projects/(\d+)/bioinfo/tasks", path)
            if match and method == "GET":
                project = self.require_project(conn, user, int(match.group(1)))
                self.list_bioinfo_tasks(conn, project["id"])
                return
            if match and method == "POST":
                project = self.require_project(conn, user, int(match.group(1)))
                self.create_bioinfo_task(conn, user, project["id"])
                return

            match = re.fullmatch(r"/api/bioinfo/tasks/(\d+)", path)
            if match and method == "PUT":
                self.update_bioinfo_task(conn, user, int(match.group(1)))
                return

            match = re.fullmatch(r"/api/projects/(\d+)/bioinfo/export", path)
            if match and method == "GET":
                project = self.require_project(conn, user, int(match.group(1)))
                self.export_bioinfo(conn, project)
                return

            match = re.fullmatch(r"/api/projects/(\d+)/files", path)
            if match and method == "GET":
                project = self.require_project(conn, user, int(match.group(1)))
                self.list_files(conn, project["id"])
                return

            match = re.fullmatch(r"/api/projects/(\d+)/files/([A-Za-z0-9_-]+)", path)
            if match and method == "POST":
                project = self.require_project(conn, user, int(match.group(1)))
                self.upload_file(conn, user, project["id"], match.group(2))
                return

            match = re.fullmatch(r"/api/files/(\d+)/download", path)
            if match and method == "GET":
                self.download_file(conn, user, int(match.group(1)))
                return

            match = re.fullmatch(r"/api/projects/(\d+)/payment", path)
            if match and method == "GET":
                project = self.require_project(conn, user, int(match.group(1)))
                self.get_payment(conn, project["id"])
                return
            if match and method == "PUT":
                project = self.require_project(conn, user, int(match.group(1)))
                self.update_payment(conn, user, project["id"])
                return

        raise ApiError(HTTPStatus.NOT_FOUND, "接口不存在")

    def serve_frontend(self, path: str) -> None:
        if path in ("", "/"):
            file_path = FRONTEND_ROOT / "index.html"
        else:
            relative = path.lstrip("/")
            file_path = (FRONTEND_ROOT / relative).resolve()
            if FRONTEND_ROOT not in file_path.parents and file_path != FRONTEND_ROOT:
                raise ApiError(HTTPStatus.FORBIDDEN, "禁止访问")
            if not file_path.exists():
                file_path = FRONTEND_ROOT / "index.html"

        if not file_path.exists() or not file_path.is_file():
            raise ApiError(HTTPStatus.NOT_FOUND, "页面不存在")

        content = file_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mimetypes.guess_type(file_path.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def login(self) -> None:
        data = self.read_json()
        username = str(data.get("username", "")).strip()
        password = str(data.get("password", ""))
        with connect() as conn:
            row = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
            if not row or row["status"] != "active" or not verify_password(password, row["password_hash"]):
                raise ApiError(HTTPStatus.UNAUTHORIZED, "账号或密码错误")
            user = row_to_dict(row)
            token = secrets.token_urlsafe(32)
            SESSIONS[token] = user
            self.send_json(HTTPStatus.OK, {"ok": True, "token": token, "user": public_user(user)})

    def register_request(self) -> None:
        data = self.read_json()
        username = str(data.get("username", "")).strip()
        password = str(data.get("password", ""))
        display_name = str(data.get("display_name", "")).strip() or username
        if not username or len(password) < 6:
            raise ApiError(HTTPStatus.BAD_REQUEST, "账号不能为空，密码至少 6 位")
        with connect() as conn:
            existing = conn.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
            if existing:
                raise ApiError(HTTPStatus.CONFLICT, "账号已存在")
            conn.execute(
                "INSERT INTO register_requests (username, display_name, password_hash, status, created_at) VALUES (?, ?, ?, ?, ?)",
                (username, display_name, hash_password(password), "pending", now_text()),
            )
            conn.commit()
        self.send_json(HTTPStatus.OK, {"ok": True, "message": "申请已提交，等待管理员审核"})

    def forgot_password(self) -> None:
        data = self.read_json()
        username = str(data.get("username", "")).strip()
        self.send_json(HTTPStatus.OK, {"ok": True, "message": f"{username or '账号'} 的重置申请已记录"})

    def list_projects(self, conn: sqlite3.Connection, user: dict) -> None:
        if user["role"] == "admin":
            rows = conn.execute(project_list_sql(order_by="p.updated_at DESC")).fetchall()
        else:
            rows = conn.execute(
                project_list_sql(
                    where="""
                    p.owner_id = ? OR EXISTS (
                        SELECT 1 FROM project_members pm WHERE pm.project_id = p.id AND pm.user_id = ?
                    )
                    """,
                    order_by="p.updated_at DESC",
                ),
                (user["id"], user["id"]),
            ).fetchall()
        self.send_json(HTTPStatus.OK, {"ok": True, "projects": [row_to_dict(row) for row in rows]})

    def create_project(self, conn: sqlite3.Connection, user: dict) -> None:
        data = self.read_json()
        payload = clean_payload(data, PROJECT_FIELDS)
        if not payload.get("name"):
            pieces = [payload.get("oa_project_code"), payload.get("hospital_name"), payload.get("customer_name")]
            payload["name"] = " ".join(str(piece) for piece in pieces if piece).strip() or "未命名项目"
        payload.setdefault("priority", "常规")
        payload.setdefault("current_progress", "样本入库")
        payload["progress_percent"] = int(payload.get("progress_percent") or progress_to_percent(str(payload["current_progress"])))
        payload.setdefault("status", "active")
        payload["owner_id"] = user["id"]
        payload["created_at"] = now_text()
        payload["updated_at"] = now_text()
        columns = list(payload.keys())
        cur = conn.execute(
            f"INSERT INTO projects ({', '.join(columns)}) VALUES ({', '.join('?' for _ in columns)})",
            [payload[column] for column in columns],
        )
        project_id = int(cur.lastrowid)
        conn.execute(
            "INSERT OR IGNORE INTO project_members (project_id, user_id, role) VALUES (?, ?, ?)",
            (project_id, user["id"], "owner"),
        )
        conn.execute(
            "INSERT OR IGNORE INTO payments (project_id, total_amount, received_amount, payment_note, updated_at) VALUES (?, 0, 0, '', ?)",
            (project_id, now_text()),
        )
        log_operation(conn, user, "项目", "创建项目", project_id, detail=payload["name"])
        conn.commit()
        self.send_json(HTTPStatus.CREATED, {"ok": True, "project_id": project_id})

    def get_project(self, conn: sqlite3.Connection, project: sqlite3.Row) -> None:
        row = conn.execute(project_list_sql(where="p.id = ?"), (project["id"],)).fetchone()
        self.send_json(HTTPStatus.OK, {"ok": True, "project": row_to_dict(row)})

    def update_project(self, conn: sqlite3.Connection, user: dict, project_id: int) -> None:
        payload = clean_payload(self.read_json(), PROJECT_FIELDS)
        if not payload:
            raise ApiError(HTTPStatus.BAD_REQUEST, "没有可更新字段")
        payload["updated_at"] = now_text()
        assignments = ", ".join(f"{field} = ?" for field in payload)
        conn.execute(
            f"UPDATE projects SET {assignments} WHERE id = ?",
            [payload[field] for field in payload] + [project_id],
        )
        log_operation(conn, user, "项目", "更新项目", project_id)
        conn.commit()
        self.send_json(HTTPStatus.OK, {"ok": True})

    def complete_project(self, conn: sqlite3.Connection, user: dict, project_id: int) -> None:
        conn.execute(
            """
            UPDATE projects
            SET status = 'completed', current_progress = '项目结束', progress_percent = 100, progress_time = ?, updated_at = ?
            WHERE id = ?
            """,
            (datetime.now().strftime("%Y-%m-%d"), now_text(), project_id),
        )
        log_operation(conn, user, "项目结束", "项目结束", project_id)
        conn.commit()
        self.send_json(HTTPStatus.OK, {"ok": True})

    def list_samples(self, conn: sqlite3.Connection, project_id: int) -> None:
        rows = conn.execute("SELECT * FROM samples WHERE project_id = ? ORDER BY id", (project_id,)).fetchall()
        self.send_json(HTTPStatus.OK, {"ok": True, "samples": [row_to_dict(row) for row in rows]})

    def create_sample(self, conn: sqlite3.Connection, user: dict, project_id: int) -> None:
        data = self.read_json()
        sample_id = insert_sample(conn, project_id, data)
        log_operation(conn, user, "样本核对", "新增样本", project_id, sample_id)
        recalculate_project_progress(conn, project_id)
        conn.commit()
        self.send_json(HTTPStatus.CREATED, {"ok": True, "sample_id": sample_id})

    def import_samples(self, conn: sqlite3.Connection, user: dict, project_id: int) -> None:
        files, _fields = self.read_multipart()
        upload = files.get("file")
        if not upload:
            raise ApiError(HTTPStatus.BAD_REQUEST, "没有收到文件")
        save_upload(conn, user, project_id, "sample_import", upload["filename"], upload["content"])
        rows = read_xlsx_rows(upload["content"], "返样信息单", max_rows=5000)
        imported = 0
        updated = 0
        errors: list[str] = []
        for index, row in enumerate(rows[1:], start=2):
            values = row + [""] * 13
            if not any(values[:13]) or values[0].startswith("填写说明"):
                continue
            if not values[3]:
                if not any(values[2:13]) or "不够行数" in values[1]:
                    continue
                errors.append(f"第 {index} 行缺少样本管编号")
                continue
            data = {
                "row_no": values[0],
                "hospital_department": values[1],
                "doctor_owner": values[2],
                "sample_code": values[3],
                "patient_name": values[4],
                "sample_category": values[5],
                "sample_subcategory": values[6],
                "group_name": values[7],
                "volume": values[8],
                "test_project": values[9],
                "sequencing_mode": values[10],
                "data_amount": values[11],
                "sample_note": values[12],
                "sample_status": "待核对",
            }
            existing = conn.execute(
                """
                SELECT id FROM samples
                WHERE project_id = ? AND sample_code = ? AND COALESCE(group_name, '') = ?
                ORDER BY id LIMIT 1
                """,
                (project_id, values[3], values[7] or ""),
            ).fetchone()
            if existing:
                update_sample(conn, existing["id"], data)
                updated += 1
            else:
                insert_sample(conn, project_id, data)
                imported += 1
        log_operation(conn, user, "样本入库", "导入样本信息表", project_id, detail=f"新增 {imported}，更新 {updated}")
        recalculate_project_progress(conn, project_id)
        conn.commit()
        self.send_json(HTTPStatus.OK, {"ok": True, "imported": imported, "updated": updated, "errors": errors[:50]})

    def import_library(self, conn: sqlite3.Connection, user: dict, project_id: int) -> None:
        files, _fields = self.read_multipart()
        upload = files.get("file")
        if not upload:
            raise ApiError(HTTPStatus.BAD_REQUEST, "没有收到文件")
        save_upload(conn, user, project_id, "library_import", upload["filename"], upload["content"])
        rows = read_xlsx_rows(upload["content"], "上机信息表", max_rows=10000)
        imported = 0
        updated = 0
        for row in rows[1:]:
            values = row + [""] * 25
            if not any(values[:20]) or not (values[1] or values[2] or values[10]):
                continue
            data = {
                "sequencing_date": values[0],
                "source_sample_code": values[1],
                "lab_code": values[2],
                "patient_name": values[3],
                "sample_subcategory": values[4],
                "testing_flow": values[5],
                "sample_source": values[6],
                "handling_method": values[7],
                "extraction_concentration": values[8],
                "library_input_amount": values[9],
                "sub_library_code": values[10],
                "adapter_code": values[11],
                "umsi_code": values[12],
                "library_concentration": values[13],
                "sequencing_library_amount": values[14],
                "chip_code": values[15],
                "reagent_lot": values[16],
                "operator_name": values[17],
                "sample_note": values[18],
                "total_reads": values[19],
                "host_rate": values[20],
                "q20": values[21],
                "q30": values[22],
                "report_result": values[23],
                "detected_pathogens": values[24],
                "sample_status": "已建库",
            }
            matched = sample_match(conn, project_id, {**data, "sample_code": values[1]})
            if matched:
                update_sample(conn, matched["id"], data)
                updated += 1
            else:
                data["sample_code"] = values[1] or values[2] or values[10]
                insert_sample(conn, project_id, data)
                imported += 1
        log_operation(conn, user, "建库", "导入实验室上机表", project_id, detail=f"新增 {imported}，更新 {updated}")
        recalculate_project_progress(conn, project_id)
        conn.commit()
        self.send_json(HTTPStatus.OK, {"ok": True, "imported": imported, "updated": updated})

    def update_sample_endpoint(self, conn: sqlite3.Connection, user: dict, sample: sqlite3.Row) -> None:
        update_sample(conn, sample["id"], self.read_json())
        log_operation(conn, user, "样本核对", "更新样本", sample["project_id"], sample["id"])
        recalculate_project_progress(conn, sample["project_id"])
        conn.commit()
        self.send_json(HTTPStatus.OK, {"ok": True})

    def delete_sample(self, conn: sqlite3.Connection, user: dict, sample: sqlite3.Row) -> None:
        conn.execute("DELETE FROM samples WHERE id = ?", (sample["id"],))
        log_operation(conn, user, "样本核对", "删除样本", sample["project_id"], sample["id"])
        recalculate_project_progress(conn, sample["project_id"])
        conn.commit()
        self.send_json(HTTPStatus.OK, {"ok": True})

    def list_bioinfo_tasks(self, conn: sqlite3.Connection, project_id: int) -> None:
        rows = conn.execute("SELECT * FROM bioinfo_tasks WHERE project_id = ? ORDER BY id DESC", (project_id,)).fetchall()
        self.send_json(HTTPStatus.OK, {"ok": True, "tasks": [row_to_dict(row) for row in rows]})

    def create_bioinfo_task(self, conn: sqlite3.Connection, user: dict, project_id: int) -> None:
        data = self.read_json()
        title = str(data.get("title", "")).strip()
        if not title:
            raise ApiError(HTTPStatus.BAD_REQUEST, "任务名称不能为空")
        cur = conn.execute(
            """
            INSERT INTO bioinfo_tasks (project_id, title, content, status, owner, due_date, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                project_id,
                title,
                data.get("content", ""),
                data.get("status", "待开始"),
                data.get("owner", ""),
                data.get("due_date", ""),
                now_text(),
                now_text(),
            ),
        )
        log_operation(conn, user, "生信分析", "新增生信任务", project_id, detail=title)
        recalculate_project_progress(conn, project_id)
        conn.commit()
        self.send_json(HTTPStatus.CREATED, {"ok": True, "task_id": cur.lastrowid})

    def update_bioinfo_task(self, conn: sqlite3.Connection, user: dict, task_id: int) -> None:
        task = conn.execute("SELECT * FROM bioinfo_tasks WHERE id = ?", (task_id,)).fetchone()
        if not task:
            raise ApiError(HTTPStatus.NOT_FOUND, "任务不存在")
        self.require_project(conn, user, task["project_id"])
        data = self.read_json()
        allowed = ["title", "content", "status", "owner", "due_date"]
        payload = clean_payload(data, allowed)
        if not payload:
            raise ApiError(HTTPStatus.BAD_REQUEST, "没有可更新字段")
        payload["updated_at"] = now_text()
        assignments = ", ".join(f"{key} = ?" for key in payload)
        conn.execute(f"UPDATE bioinfo_tasks SET {assignments} WHERE id = ?", [payload[key] for key in payload] + [task_id])
        log_operation(conn, user, "生信分析", "更新生信任务", task["project_id"], detail=str(task_id))
        conn.commit()
        self.send_json(HTTPStatus.OK, {"ok": True})

    def export_bioinfo(self, conn: sqlite3.Connection, project: sqlite3.Row) -> None:
        samples = conn.execute("SELECT * FROM samples WHERE project_id = ? ORDER BY id", (project["id"],)).fetchall()
        rows: list[list[object]] = [
            ["广州微远基因科技有限公司"],
            [
                "序号",
                "项目名称",
                "样品编号",
                "患者姓名",
                "分组信息",
                "实验室编号",
                "子文库编号",
                "接头编号",
                "Adapter编号",
                "检测流程",
                "提取浓度（ng/ul）",
                "文库终浓度(ng/ul）",
                "备注",
            ],
        ]
        for index, sample in enumerate(samples, start=1):
            rows.append(
                [
                    index,
                    project["name"],
                    sample["source_sample_code"] or sample["sample_code"],
                    sample["patient_name"],
                    sample["group_name"],
                    sample["lab_code"],
                    sample["sub_library_code"],
                    sample["adapter_code"],
                    sample["umsi_code"],
                    sample["testing_flow"],
                    sample["extraction_concentration"],
                    sample["library_concentration"],
                    sample["bioinfo_task_content"] or sample["sample_note"],
                ]
            )
        content = write_xlsx_bytes(rows, "生信模板表")
        self.send_file_bytes(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "生信信息表.xlsx")

    def list_files(self, conn: sqlite3.Connection, project_id: int) -> None:
        rows = conn.execute("SELECT * FROM uploaded_files WHERE project_id = ? ORDER BY id DESC", (project_id,)).fetchall()
        files = []
        for row in rows:
            item = row_to_dict(row)
            item["url"] = f"/api/files/{row['id']}/download"
            files.append(item)
        self.send_json(HTTPStatus.OK, {"ok": True, "files": files})

    def upload_file(self, conn: sqlite3.Connection, user: dict, project_id: int, module: str) -> None:
        files, _fields = self.read_multipart()
        upload = files.get("file")
        if not upload:
            raise ApiError(HTTPStatus.BAD_REQUEST, "没有收到文件")
        saved = save_upload(conn, user, project_id, module, upload["filename"], upload["content"])
        log_operation(conn, user, module, "上传文件", project_id, detail=saved["original_name"])
        recalculate_project_progress(conn, project_id)
        conn.commit()
        self.send_json(HTTPStatus.CREATED, {"ok": True, "file": saved})

    def download_file(self, conn: sqlite3.Connection, user: dict, file_id: int) -> None:
        row = conn.execute("SELECT * FROM uploaded_files WHERE id = ?", (file_id,)).fetchone()
        if not row:
            raise ApiError(HTTPStatus.NOT_FOUND, "文件不存在")
        self.require_project(conn, user, row["project_id"])
        file_path = Path(row["path"])
        if not file_path.exists():
            raise ApiError(HTTPStatus.NOT_FOUND, "文件已丢失")
        self.send_file_bytes(
            file_path.read_bytes(),
            mimetypes.guess_type(row["original_name"])[0] or "application/octet-stream",
            row["original_name"],
        )

    def get_payment(self, conn: sqlite3.Connection, project_id: int) -> None:
        row = conn.execute("SELECT * FROM payments WHERE project_id = ?", (project_id,)).fetchone()
        if not row:
            conn.execute(
                "INSERT INTO payments (project_id, total_amount, received_amount, payment_note, updated_at) VALUES (?, 0, 0, '', ?)",
                (project_id, now_text()),
            )
            conn.commit()
            row = conn.execute("SELECT * FROM payments WHERE project_id = ?", (project_id,)).fetchone()
        payment = row_to_dict(row)
        payment["progress_percent"] = payment_percent(payment["total_amount"], payment["received_amount"])
        self.send_json(HTTPStatus.OK, {"ok": True, "payment": payment})

    def update_payment(self, conn: sqlite3.Connection, user: dict, project_id: int) -> None:
        data = self.read_json()
        total = parse_float(data.get("total_amount"))
        received = parse_float(data.get("received_amount"))
        note = str(data.get("payment_note", ""))
        conn.execute(
            """
            INSERT INTO payments (project_id, total_amount, received_amount, payment_note, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(project_id) DO UPDATE SET
                total_amount = excluded.total_amount,
                received_amount = excluded.received_amount,
                payment_note = excluded.payment_note,
                updated_at = excluded.updated_at
            """,
            (project_id, total, received, note, now_text()),
        )
        log_operation(conn, user, "回款", "更新回款", project_id, detail=f"{received}/{total}")
        recalculate_project_progress(conn, project_id)
        conn.commit()
        self.send_json(HTTPStatus.OK, {"ok": True, "progress_percent": payment_percent(total, received)})

    def export_summary(self, conn: sqlite3.Connection, user: dict) -> None:
        if user["role"] == "admin":
            rows = conn.execute(project_list_sql(order_by="p.id")).fetchall()
        else:
            rows = conn.execute(
                project_list_sql(
                    where="""
                    p.owner_id = ? OR EXISTS (
                        SELECT 1 FROM project_members pm WHERE pm.project_id = p.id AND pm.user_id = ?
                    )
                    """,
                    order_by="p.id",
                ),
                (user["id"], user["id"]),
            ).fetchall()
        output = [
            [
                "医院名称",
                "客户名称",
                "科研经理",
                "OA立项编号",
                "订单编号",
                "样本类型",
                "样本性质",
                "实收样本量",
                "外送测序样本量",
                "到样时间",
                "优先级",
                "具体实验室流程",
                "具体操作补充",
                "测序安排",
                "上机平台",
                "上机批次号",
                "外送大pool号",
                "外送小pool号",
                "测序模式",
                "测序数据量/例",
                "网盘数据位置编号",
                "目前进展",
                "当前进展的时间点",
                "样本数量",
                "回款总金额",
                "已回款金额",
                "回款进度",
            ]
        ]
        for row in rows:
            output.append(
                [
                    row["hospital_name"],
                    row["customer_name"],
                    row["research_manager"],
                    row["oa_project_code"],
                    row["order_code"],
                    row["sample_type"],
                    row["sample_nature"],
                    row["received_sample_count"],
                    row["sent_sample_count"],
                    row["arrival_date"],
                    row["priority"],
                    row["lab_flow"],
                    row["operation_note"],
                    row["sequencing_plan"],
                    row["sequencing_platform"],
                    row["sequencing_batch_no"],
                    row["big_pool_no"],
                    row["small_pool_no"],
                    row["sequencing_mode"],
                    row["data_amount_per_sample"],
                    row["data_location_code"],
                    row["current_progress"],
                    row["progress_time"],
                    row["sample_count"],
                    row["total_amount"],
                    row["received_amount"],
                    f"{payment_percent(row['total_amount'], row['received_amount'])}%",
                ]
            )
        content = write_xlsx_bytes(output, "样本到样信息汇总统计")
        self.send_file_bytes(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "项目信息汇总.xlsx")

    def require_user(self, conn: sqlite3.Connection) -> dict:
        header = self.headers.get("Authorization", "")
        token = header[7:] if header.startswith("Bearer ") else ""
        user = SESSIONS.get(token)
        if not user:
            raise ApiError(HTTPStatus.UNAUTHORIZED, "未登录或登录已过期")
        row = conn.execute("SELECT * FROM users WHERE id = ? AND status = 'active'", (user["id"],)).fetchone()
        if not row:
            raise ApiError(HTTPStatus.UNAUTHORIZED, "账号不可用")
        return row_to_dict(row)

    def require_project(self, conn: sqlite3.Connection, user: dict, project_id: int) -> sqlite3.Row:
        if user["role"] == "admin":
            project = conn.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()
        else:
            project = conn.execute(
                """
                SELECT p.*
                FROM projects p
                WHERE p.id = ?
                  AND (
                    p.owner_id = ?
                    OR EXISTS (SELECT 1 FROM project_members pm WHERE pm.project_id = p.id AND pm.user_id = ?)
                  )
                """,
                (project_id, user["id"], user["id"]),
            ).fetchone()
        if not project:
            raise ApiError(HTTPStatus.NOT_FOUND, "项目不存在或无权限")
        return project

    def require_sample(self, conn: sqlite3.Connection, user: dict, sample_id: int) -> sqlite3.Row:
        sample = conn.execute("SELECT * FROM samples WHERE id = ?", (sample_id,)).fetchone()
        if not sample:
            raise ApiError(HTTPStatus.NOT_FOUND, "样本不存在")
        self.require_project(conn, user, sample["project_id"])
        return sample

    def read_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or "0")
        if length == 0:
            return {}
        raw = self.rfile.read(length)
        try:
            return json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise ApiError(HTTPStatus.BAD_REQUEST, f"JSON 格式错误：{exc}") from exc

    def read_multipart(self) -> tuple[dict[str, dict], dict[str, str]]:
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            raise ApiError(HTTPStatus.BAD_REQUEST, "请求必须是 multipart/form-data")
        boundary_match = re.search(r"boundary=(?:\"([^\"]+)\"|([^;]+))", content_type)
        if not boundary_match:
            raise ApiError(HTTPStatus.BAD_REQUEST, "缺少 multipart boundary")
        boundary = (boundary_match.group(1) or boundary_match.group(2)).encode("utf-8")
        length = int(self.headers.get("Content-Length") or "0")
        body = self.rfile.read(length)
        files: dict[str, dict] = {}
        fields: dict[str, str] = {}

        delimiter = b"--" + boundary
        for raw_part in body.split(delimiter):
            raw_part = raw_part.strip(b"\r\n")
            if not raw_part or raw_part == b"--":
                continue
            if raw_part.endswith(b"--"):
                raw_part = raw_part[:-2].rstrip(b"\r\n")
            if b"\r\n\r\n" not in raw_part:
                continue
            raw_headers, content = raw_part.split(b"\r\n\r\n", 1)
            if content.endswith(b"\r\n"):
                content = content[:-2]
            headers = raw_headers.decode("utf-8", errors="replace").split("\r\n")
            disposition = next((line for line in headers if line.lower().startswith("content-disposition:")), "")
            params = parse_header_params(disposition)
            name = params.get("name")
            if not name:
                continue
            filename = params.get("filename")
            if filename:
                files[name] = {"filename": filename, "content": content}
            else:
                fields[name] = content.decode("utf-8", errors="replace")
        return files, fields

    def send_json(self, status: int | HTTPStatus, payload: dict) -> None:
        content = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def send_file_bytes(self, content: bytes, content_type: str, filename: str) -> None:
        safe_name = safe_filename(filename)
        disposition = f"attachment; filename*=UTF-8''{url_quote(safe_name)}"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", disposition)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def log_message(self, fmt: str, *args) -> None:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] {self.address_string()} {fmt % args}")


class ApiError(Exception):
    def __init__(self, status: int | HTTPStatus, message: str) -> None:
        self.status = status
        self.message = message
        super().__init__(message)


def public_user(user: dict) -> dict:
    return {
        "id": user["id"],
        "username": user["username"],
        "display_name": user["display_name"],
        "role": user["role"],
    }


def parse_float(value: object) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def payment_percent(total: object, received: object) -> int:
    total_value = parse_float(total)
    received_value = parse_float(received)
    if total_value <= 0:
        return 0
    return max(0, min(100, round(received_value / total_value * 100)))


def url_quote(value: str) -> str:
    return "".join(f"%{byte:02X}" if byte > 127 or chr(byte) in " %;," else chr(byte) for byte in value.encode("utf-8"))


def parse_header_params(header_line: str) -> dict[str, str]:
    if ":" in header_line:
        header_line = header_line.split(":", 1)[1]
    params: dict[str, str] = {}
    for part in header_line.split(";"):
        part = part.strip()
        if "=" not in part:
            continue
        key, value = part.split("=", 1)
        value = value.strip()
        if value.startswith('"') and value.endswith('"'):
            value = value[1:-1]
        params[key.strip().lower()] = value
    return params


PROJECT_LIST_SELECT = """
SELECT
    p.*,
    u.display_name AS owner_name,
    COUNT(s.id) AS sample_count,
    COALESCE(pay.total_amount, 0) AS total_amount,
    COALESCE(pay.received_amount, 0) AS received_amount
FROM projects p
JOIN users u ON u.id = p.owner_id
LEFT JOIN samples s ON s.project_id = p.id
LEFT JOIN payments pay ON pay.project_id = p.id
"""


def project_list_sql(where: str = "", order_by: str = "") -> str:
    sql = PROJECT_LIST_SELECT
    if where:
        sql += f" WHERE {where}"
    sql += " GROUP BY p.id"
    if order_by:
        sql += f" ORDER BY {order_by}"
    return sql


def main() -> None:
    parser = argparse.ArgumentParser(description="Biotech project management MVP server")
    parser.add_argument("--host", default=os.environ.get("BIOTECH_PM_HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.environ.get("BIOTECH_PM_PORT", "8000")))
    args = parser.parse_args()

    init_db()
    server = ThreadingHTTPServer((args.host, args.port), ApiHandler)
    print(f"Biotech PM server running at http://{args.host}:{args.port}")
    print("Default accounts: admin/admin123, demo/demo123")
    server.serve_forever()


if __name__ == "__main__":
    main()

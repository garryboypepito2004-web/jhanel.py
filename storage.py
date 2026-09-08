import ast
import json
import os
import shutil
import sqlite3
import tempfile
import time
from pathlib import Path

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(APP_DIR, "ailyn_house.db")
BACKUP_DIR = os.path.join(APP_DIR, "backups")
SCANNER_PHOTO_DIR = os.path.join(APP_DIR, "scanner_photos")


def _default_state_dict():
    return {
        "records": [],
        "labor_records": [],
        "payroll_expenses": [],
        "planner_tasks": [],
        "budget": 0.0,
        "budget_history": [],
        "remaining_money": 0.0,
        "view": "home",
        "current_view": "home",
        "project": {
            "name": "Ailyn House Project",
            "client": "",
            "address": "",
            "manager": "",
            "status": "Active",
            "target_date": "",
        },
        "receipt_archive": [],
        "scanner_photos": [],
        "dark_mode": False,
        "selected_role": "Labor",
    }


def _ensure_db():
    os.makedirs(APP_DIR, exist_ok=True)
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS app_state (key TEXT PRIMARY KEY, value TEXT)"
        )
        conn.commit()


def _coerce_state_value(value):
    if isinstance(value, (dict, list, int, float, bool)) or value is None:
        return value
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return {}
        try:
            return json.loads(stripped)
        except json.JSONDecodeError:
            try:
                return ast.literal_eval(stripped)
            except (ValueError, SyntaxError):
                return value
    return value


def _state_from_db():
    _ensure_db()
    with sqlite3.connect(DB_FILE) as conn:
        legacy_rows = conn.execute("SELECT key, value FROM app_state ORDER BY key").fetchall()

    if legacy_rows:
        for key, value in legacy_rows:
            if key == "state":
                parsed = _coerce_state_value(value)
                if isinstance(parsed, dict):
                    return parsed

        state = {}
        for key, value in legacy_rows:
            converted = _coerce_state_value(value)
            state[key] = converted
        return state

    return {}


def load_state():
    state = _state_from_db()
    if not state:
        default_state = _default_state_dict()
        save_state(default_state)
        return default_state
    return state


def save_state(state):
    _ensure_db()
    payload = json.dumps(state, default=str, ensure_ascii=False)
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute("INSERT OR REPLACE INTO app_state (key, value) VALUES ('state', ?)", (payload,))
        conn.execute("DELETE FROM app_state WHERE key NOT IN ('state')")
        conn.commit()


def create_backup():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = int(time.time())
    backup_name = f"ailyn_house_{timestamp}.db"
    backup_path = os.path.join(BACKUP_DIR, backup_name)
    if os.path.exists(DB_FILE):
        shutil.copy2(DB_FILE, backup_path)
    else:
        open(backup_path, "wb").close()
    return backup_path


def history_count():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    return len([name for name in os.listdir(BACKUP_DIR) if name.endswith(".db") or name.endswith(".bak")])


def restore_backup(backup_source):
    temp_path = None
    try:
        if hasattr(backup_source, "read"):
            uploaded = backup_source.read()
            temp_fd, temp_path = tempfile.mkstemp(prefix="restore_", suffix=".db")
            os.close(temp_fd)
            with open(temp_path, "wb") as fh:
                fh.write(uploaded)
            backup_source = temp_path
        if not os.path.exists(backup_source):
            raise FileNotFoundError(f"Backup not found: {backup_source}")
        shutil.copy2(backup_source, DB_FILE)
        return DB_FILE
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)


def save_scanner_photo(photo_bytes, mime_type, photo_id):
    os.makedirs(SCANNER_PHOTO_DIR, exist_ok=True)
    extension = ".jpg" if mime_type.lower().startswith("image") else ".bin"
    file_name = f"{photo_id}{extension}"
    file_path = os.path.join(SCANNER_PHOTO_DIR, file_name)
    with open(file_path, "wb") as fh:
        fh.write(photo_bytes)
    return os.path.relpath(file_path, APP_DIR)


def delete_scanner_photo(file_path):
    if not file_path:
        return
    candidate = file_path if os.path.isabs(file_path) else os.path.join(APP_DIR, file_path)
    if os.path.exists(candidate):
        os.remove(candidate)


# compatibility aliases used by the app
load_sqlite_state = load_state
save_sqlite_state = save_state

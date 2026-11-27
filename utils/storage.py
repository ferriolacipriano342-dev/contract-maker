"""로컬 SQLite DB(data/contracts.db) 기반 계약서 임시저장·누적저장 및 증빙서류 보관 모듈."""

from __future__ import annotations

import json
import re
import shutil
import sqlite3
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "contracts.db"
ATTACHMENTS_ROOT = DATA_DIR / "attachments"

ATTACHMENT_CATEGORIES = ["신분증", "등기부등본", "토지대장", "건축물대장", "기타"]

DATE_FIELD_KEYS = {
    "interim_date",
    "balance_date",
    "handover_date",
    "lease_end_date",
    "contract_date",
}


def _get_conn() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_db() -> None:
    """data/contracts.db 테이블과 첨부파일 루트 폴더를 생성한다."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ATTACHMENTS_ROOT.mkdir(parents=True, exist_ok=True)
    with _get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS contracts (
                id TEXT PRIMARY KEY,
                contract_type TEXT NOT NULL,
                contract_type_label TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'draft',
                title TEXT NOT NULL,
                address TEXT DEFAULT '',
                facility_or_store TEXT DEFAULT '',
                party1_name TEXT DEFAULT '',
                party2_name TEXT DEFAULT '',
                deposit_amount INTEGER DEFAULT 0,
                contract_date TEXT DEFAULT '',
                payload_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS attachments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                contract_id TEXT NOT NULL,
                category TEXT NOT NULL,
                original_filename TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                mime_type TEXT NOT NULL,
                file_size INTEGER NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (contract_id) REFERENCES contracts(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_contracts_updated ON contracts(updated_at DESC);
            CREATE INDEX IF NOT EXISTS idx_contracts_type ON contracts(contract_type);
            CREATE INDEX IF NOT EXISTS idx_attachments_cid ON attachments(contract_id);
            """
        )


def generate_contract_id() -> str:
    """고유 계약서 ID 생성 (예: CT-20261127-A1B2C3)."""
    day_tag = datetime.now().strftime("%Y%m%d")
    short_uid = uuid.uuid4().hex[:6].upper()
    return f"CT-{day_tag}-{short_uid}"


def _json_default(obj: Any) -> Any:
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()[:10]
    return str(obj)


def _restore_dates(form_data: dict[str, Any]) -> dict[str, Any]:
    restored = dict(form_data)
    for key in DATE_FIELD_KEYS:
        val = restored.get(key)
        if isinstance(val, str) and val.strip():
            try:
                restored[key] = date.fromisoformat(val.strip()[:10])
            except ValueError:
                pass
    return restored


def _build_contract_title(contract_type_label: str, form_data: dict[str, Any]) -> str:
    primary = (
        str(form_data.get("facility_name", "")).strip()
        or str(form_data.get("store_name", "")).strip()
        or str(form_data.get("address", "")).strip()
        or f"{contract_type_label} 미입력 계약"
    )
    if len(primary) > 34:
        primary = primary[:34] + "…"
    p1 = str(form_data.get("lessor_name", "")).strip()
    p2 = str(form_data.get("lessee_name", "")).strip()
    if p1 or p2:
        parties = f"{p1 or '미입력'}↔{p2 or '미입력'}"
        return f"{primary} ({parties})"
    return primary


def save_contract(
    contract_id: str,
    contract_type: str,
    contract_type_label: str,
    status: str,
    payload: dict[str, Any],
) -> str:
    """계약서를 임시저장('draft') 또는 최종 누적 저장('completed') 상태로 저장한다."""
    init_db()
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    form_data = payload.get("form_data", {})

    title = _build_contract_title(contract_type_label, form_data)
    address = str(form_data.get("address", "")).strip()
    facility_or_store = (
        str(form_data.get("facility_name", "")).strip()
        or str(form_data.get("store_name", "")).strip()
    )
    party1_name = str(form_data.get("lessor_name", "")).strip()
    party2_name = str(form_data.get("lessee_name", "")).strip()
    try:
        deposit_amount = int(form_data.get("deposit", 0) or 0)
    except (TypeError, ValueError):
        deposit_amount = 0
    c_date = form_data.get("contract_date", "")
    contract_date_str = c_date.isoformat()[:10] if isinstance(c_date, (date, datetime)) else str(c_date or "")

    payload_str = json.dumps(payload, ensure_ascii=False, default=_json_default)

    with _get_conn() as conn:
        existing = conn.execute("SELECT created_at FROM contracts WHERE id = ?", (contract_id,)).fetchone()
        created_at = existing["created_at"] if existing else now_iso

        conn.execute(
            """
            INSERT INTO contracts (
                id, contract_type, contract_type_label, status, title,
                address, facility_or_store, party1_name, party2_name,
                deposit_amount, contract_date, payload_json, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                contract_type = excluded.contract_type,
                contract_type_label = excluded.contract_type_label,
                status = excluded.status,
                title = excluded.title,
                address = excluded.address,
                facility_or_store = excluded.facility_or_store,
                party1_name = excluded.party1_name,
                party2_name = excluded.party2_name,
                deposit_amount = excluded.deposit_amount,
                contract_date = excluded.contract_date,
                payload_json = excluded.payload_json,
                updated_at = excluded.updated_at
            """,
            (
                contract_id,
                contract_type,
                contract_type_label,
                status,
                title,
                address,
                facility_or_store,
                party1_name,
                party2_name,
                deposit_amount,
                contract_date_str,
                payload_str,
                created_at,
                now_iso,
            ),
        )
    return contract_id


def list_contracts(
    contract_type_filter: str = "all",
    search_query: str = "",
    status_filter: str = "all",
) -> list[dict[str, Any]]:
    """저장된 계약서 목록을 최신 수정순으로 조회한다."""
    init_db()
    clauses: list[str] = []
    params: list[Any] = []

    if contract_type_filter and contract_type_filter != "all":
        clauses.append("c.contract_type = ?")
        params.append(contract_type_filter)

    if status_filter and status_filter != "all":
        clauses.append("c.status = ?")
        params.append(status_filter)

    q = search_query.strip()
    if q:
        like = f"%{q}%"
        clauses.append(
            "(c.title LIKE ? OR c.address LIKE ? OR c.facility_or_store LIKE ? "
            "OR c.party1_name LIKE ? OR c.party2_name LIKE ? OR c.id LIKE ?)"
        )
        params.extend([like, like, like, like, like, like])

    where_sql = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    sql = f"""
        SELECT c.*, COUNT(a.id) AS attachment_count
        FROM contracts c
        LEFT JOIN attachments a ON a.contract_id = c.id
        {where_sql}
        GROUP BY c.id
        ORDER BY c.updated_at DESC
    """
    with _get_conn() as conn:
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]


def load_contract(contract_id: str) -> tuple[dict[str, Any] | None, dict[str, Any] | None, list[dict[str, Any]]]:
    """계약서 메타데이터, 복원된 페이로드, 첨부 증빙서류 목록을 불러온다."""
    init_db()
    with _get_conn() as conn:
        row = conn.execute("SELECT * FROM contracts WHERE id = ?", (contract_id,)).fetchone()
        if not row:
            return None, None, []
        meta = dict(row)
        payload = json.loads(meta["payload_json"])
        if "form_data" in payload and isinstance(payload["form_data"], dict):
            payload["form_data"] = _restore_dates(payload["form_data"])
        att_rows = conn.execute(
            "SELECT * FROM attachments WHERE contract_id = ? ORDER BY created_at DESC, id DESC",
            (contract_id,),
        ).fetchall()
        return meta, payload, [dict(a) for a in att_rows]


def delete_contract(contract_id: str) -> None:
    """계약서 레코드와 연결된 첨부파일 폴더를 완전히 삭제한다."""
    init_db()
    with _get_conn() as conn:
        conn.execute("DELETE FROM contracts WHERE id = ?", (contract_id,))
    att_dir = ATTACHMENTS_ROOT / contract_id
    if att_dir.exists() and att_dir.is_dir():
        shutil.rmtree(att_dir, ignore_errors=True)


def _safe_filename(name: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", name.strip())
    return cleaned or "attachment.bin"


def ensure_contract_record_exists(
    contract_id: str,
    contract_type: str,
    contract_type_label: str,
    payload: dict[str, Any],
) -> None:
    """첨부파일 저장 전 contracts 행이 없으면 자동으로 임시저장 행을 생성한다."""
    init_db()
    with _get_conn() as conn:
        exists = conn.execute("SELECT 1 FROM contracts WHERE id = ?", (contract_id,)).fetchone()
    if not exists:
        save_contract(
            contract_id=contract_id,
            contract_type=contract_type,
            contract_type_label=contract_type_label,
            status="draft",
            payload=payload,
        )


def save_attachment(
    contract_id: str,
    category: str,
    original_filename: str,
    data_bytes: bytes,
    mime_type: str,
) -> dict[str, Any]:
    """증빙서류 파일을 data/attachments/{계약ID}/ 폴더에 저장하고 DB에 기록한다."""
    init_db()
    cat = category if category in ATTACHMENT_CATEGORIES else "기타"
    safe_name = _safe_filename(original_filename)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:19]
    stored_name = f"{cat}_{ts}_{safe_name}"

    target_dir = ATTACHMENTS_ROOT / contract_id
    target_dir.mkdir(parents=True, exist_ok=True)
    target_file = target_dir / stored_name
    target_file.write_bytes(data_bytes)

    rel_path = target_file.relative_to(BASE_DIR).as_posix()
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with _get_conn() as conn:
        cur = conn.execute(
            """
            INSERT INTO attachments (
                contract_id, category, original_filename, stored_path, mime_type, file_size, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (contract_id, cat, original_filename, rel_path, mime_type, len(data_bytes), now_iso),
        )
        att_id = cur.lastrowid

    return {
        "id": att_id,
        "contract_id": contract_id,
        "category": cat,
        "original_filename": original_filename,
        "stored_path": rel_path,
        "mime_type": mime_type,
        "file_size": len(data_bytes),
        "created_at": now_iso,
    }


def list_attachments(contract_id: str) -> list[dict[str, Any]]:
    """특정 계약서에 보관된 증빙서류 목록을 반환한다."""
    init_db()
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM attachments WHERE contract_id = ? ORDER BY created_at DESC, id DESC",
            (contract_id,),
        ).fetchall()
        return [dict(r) for r in rows]


def read_attachment_bytes(stored_path: str) -> bytes | None:
    """저장된 증빙서류 파일 바이트를 안전하게 읽어온다."""
    full_path = (BASE_DIR / stored_path).resolve()
    try:
        full_path.relative_to(ATTACHMENTS_ROOT.resolve())
    except ValueError:
        return None
    if not full_path.exists():
        return None
    return full_path.read_bytes()


def delete_attachment(attachment_id: int) -> None:
    """단일 증빙서류를 디스크와 DB에서 삭제한다."""
    init_db()
    with _get_conn() as conn:
        row = conn.execute("SELECT stored_path FROM attachments WHERE id = ?", (attachment_id,)).fetchone()
        if row:
            full_path = (BASE_DIR / row["stored_path"]).resolve()
            if full_path.exists():
                full_path.unlink(missing_ok=True)
            conn.execute("DELETE FROM attachments WHERE id = ?", (attachment_id,))

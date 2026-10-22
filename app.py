# -*- coding: utf-8 -*-
"""
Nipro Service Report -> Excel automation (Streamlit)

Run app :  streamlit run app.py
Self test: python app.py --selftest      (no AI / no Streamlit server needed)
"""
from __future__ import annotations

import base64
import difflib
import hashlib
import io
import json
import os
import re
import sqlite3
import sys
import time
import traceback
import uuid
import zipfile
from collections import defaultdict
from contextlib import contextmanager
from datetime import date, datetime, time as dtime
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import streamlit as st
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import from_excel, to_excel
from PIL import Image, ImageOps, ImageSequence

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

# ----------------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DB_PATH = DATA_DIR / "nipro_cache.db"
MASTER_PATH = DATA_DIR / "master.xlsx"

SHEET_TMS = "TMS Working Record"
SHEET_SN = "SN tracking form"
SHEET_EXP = "Expense Claim"
SHEET_ERR = "Error Code"
SHEET_SPARE = "Spare part code"
SHEET_DATA2 = "Data2 (2)"
SHEET_SVC = "ServiceType"
REQUIRED_SHEETS = [SHEET_TMS, SHEET_SN, SHEET_EXP, SHEET_ERR, SHEET_SPARE, SHEET_DATA2]

PROVIDER_CLAUDE = "Anthropic (Claude)"
PROVIDER_OPENAI = "OpenAI"
PROVIDER_DEMO = "Demo (ไม่ใช้ AI - ทดสอบระบบ)"
DEFAULT_MODELS = {PROVIDER_CLAUDE: "claude-sonnet-4-5", PROVIDER_OPENAI: "gpt-4o", PROVIDER_DEMO: "-"}

PDF_ZOOM = 2.5
MAX_SIDE = 2200
JPEG_QUALITY = 88

CONDITION_OPTIONS = ["Warranty", "PM Contract", "Fixed asset", "Work Order", "Other"]
KIND_OPTIONS = ["PM", "Service", "Installation"]
STATUS_OPTIONS = ["Issue DO", "Issue PM", "C", "W", "S", "T", "N/A"]
SERVICE_TYPE_OPTIONS = ["Warranty", "PM Contract"]
PLACEHOLDERS = {"office", "-", "n/a", "na"}

REPORT_COLS = [
    "include", "rid", "page_key", "source", "page", "report_no", "date", "customer",
    "hospital_eng", "hospital_th", "province", "model", "serials", "condition",
    "service_kind", "start_time", "end_time", "error_codes", "description",
    "status_code", "quotation_no", "remark", "confidence", "notes",
]
PART_COLS = ["pid", "rid", "part_name", "part_code", "qty", "unit", "consignment_no"]
EXP_COLS = ["date", "description", "service_type", "allowance", "petrol", "expressway", "parking", "other"]
EXP_NUM = ["allowance", "petrol", "expressway", "parking", "other"]

THAI_MONTHS = {
    "ม.ค": 1, "มกราคม": 1, "ก.พ": 2, "กุมภาพันธ์": 2, "มี.ค": 3, "มีนาคม": 3,
    "เม.ย": 4, "เมษายน": 4, "พ.ค": 5, "พฤษภาคม": 5, "มิ.ย": 6, "มิถุนายน": 6,
    "ก.ค": 7, "กรกฎาคม": 7, "ส.ค": 8, "สิงหาคม": 8, "ก.ย": 9, "กันยายน": 9,
    "ต.ค": 10, "ตุลาคม": 10, "พ.ย": 11, "พฤศจิกายน": 11, "ธ.ค": 12, "ธันวาคม": 12,
}

# ----------------------------------------------------------------------------
# Generic helpers
# ----------------------------------------------------------------------------
def is_blank(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return not v.strip()
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def norm_text(v: Any) -> str:
    if is_blank(v):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def norm_key(v: Any) -> str:
    s = norm_text(v).casefold()
    return s.replace("\u201c", '"').replace("\u201d", '"').replace("\u2018", "'").replace("\u2019", "'")


def norm_code(v: Any) -> str:
    s = norm_text(v).upper()
    for ch in ("\u2013", "\u2014", "\uff0d", "\u2212"):
        s = s.replace(ch, "-")
    return re.sub(r"\s+", "", s)


def norm_serial(v: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", norm_text(v).upper())


def to_float(v: Any, default: Optional[float] = None) -> Optional[float]:
    if is_blank(v):
        return default
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        m = re.search(r"-?\d+(?:\.\d+)?", str(v))
        return float(m.group(0)) if m else default


def is_na(v: Any) -> bool:
    return norm_key(v) in {"", "n/a", "na", "-", "none", "ไม่มี", "\u2013"}


def has_thai(s: str) -> bool:
    return bool(re.search(r"[\u0E00-\u0E7F]", s or ""))


def split_list(v: Any) -> list[str]:
    if is_blank(v):
        return []
    if isinstance(v, (list, tuple)):
        parts = [norm_text(x) for x in v]
    else:
        parts = re.split(r"[,;\n|/\s]+", str(v))
    out: list[str] = []
    for p in parts:
        p = p.strip()
        if p and p not in out:
            out.append(p)
    return out


def parse_date(v: Any) -> Optional[date]:
    if is_blank(v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 < float(v) < 80000:
            try:
                return from_excel(float(v)).date()
            except Exception:
                return None
        return None
    s = norm_text(v)
    if not s:
        return None
    y = mo = d = None
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})", s)
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if mo > 12 >= d:
                d, mo = mo, d
        else:
            m = re.search(r"(\d{1,2})\s*([\u0E00-\u0E7F\.]+)\s*(\d{2,4})", s)
            if m:
                name = m.group(2).replace(".", "")
                for k, num in THAI_MONTHS.items():
                    if name.startswith(k.replace(".", "")):
                        d, mo, y = int(m.group(1)), num, int(m.group(3))
                        break
    if y is None:
        return None
    if y < 100:
        y += 2000
    if y > 2400:
        y -= 543
    try:
        return date(y, mo, d)
    except (ValueError, TypeError):
        return None


def parse_time(v: Any) -> Optional[dtime]:
    if is_blank(v):
        return None
    if isinstance(v, datetime):
        return v.time().replace(second=0, microsecond=0)
    if isinstance(v, dtime):
        return v.replace(second=0, microsecond=0)
    s = norm_text(v)
    m = re.search(r"(\d{1,2})\s*[:.\-]\s*(\d{2})", s)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if h < 24 and mi < 60:
            return dtime(h, mi)
    return None


def fmt_time(t: Optional[dtime]) -> str:
    return t.strftime("%H:%M") if t else ""


def fmt_date(d: Optional[date]) -> str:
    return d.isoformat() if d else ""


def map_condition(text: Any) -> str:
    k = norm_key(text)
    if "warranty" in k:
        return "Warranty"
    if "contract" in k:
        return "PM Contract"
    if "fixed" in k:
        return "Fixed asset"
    if "work order" in k:
        return "Work Order"
    return "Other" if k else ""


def map_kind(text: Any) -> str:
    k = norm_key(text)
    if k.startswith("pm") or "preventive" in k:
        return "PM"
    if "install" in k:
        return "Installation"
    return "Service"


def parse_json_loose(text: str) -> dict:
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?\s*", "", t, flags=re.I)
    t = re.sub(r"\s*```$", "", t)
    try:
        return json.loads(t)
    except Exception:
        pass
    s, e = t.find("{"), t.rfind("}")
    if s != -1 and e > s:
        return json.loads(t[s:e + 1])
    raise ValueError("AI response is not valid JSON")


def _py(v: Any) -> Any:
    if is_blank(v):
        return None
    if isinstance(v, bool):
        return int(v)
    if hasattr(v, "item"):
        try:
            return v.item()
        except Exception:
            pass
    return v


# ----------------------------------------------------------------------------
# Reference data (Error Code / Spare part code / Data2 / ServiceType)
# ----------------------------------------------------------------------------
D2_FIELDS = {
    "type": "type", "province": "province", "model": "model name", "engineer": "engineer name",
    "name_th": "th - customer name (commercial)", "serial": "serial number",
    "name_en": "eng - customer name (commercial)", "status": "machine status",
    "pm_start": 'start date of "pm"', "pm_end": 'end date of "pm"',
    "pm_contract": "pm contract yes / no", "pm_count": "number of pm maintenance",
    "contract_no": "contract number",
}


def get_sheet(wb, name: str):
    for ws in wb.worksheets:
        if norm_key(ws.title) == norm_key(name):
            return ws
    return None


def _sheet_rows(ws) -> list[list[Any]]:
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _find_header(rows: list[list[Any]], labels: list[str], limit: int = 40):
    for i, row in enumerate(rows[:limit]):
        keys = [norm_key(c) for c in row]
        if all(lab in keys for lab in labels):
            return i, keys
    return None, []


def _hn(s: str) -> str:
    s = norm_key(s)
    s = re.sub(r"[^\w\u0E00-\u0E7F]+", " ", s)
    s = re.sub(r"\b(hospital|โรงพยาบาล|รพ)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _hscore(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if min(len(a), len(b)) >= 4 and (a in b or b in a):
        return 0.92
    return difflib.SequenceMatcher(None, a, b).ratio()


class Reference:
    def __init__(self) -> None:
        self.error_codes: dict[str, tuple[str, str]] = {}
        self.spare_parts: dict[str, tuple[str, str]] = {}
        self.machines: list[dict] = []
        self.serial_index: dict[str, dict] = {}
        self.service_types: list[str] = []
        self.conditions: list[str] = []
        self.missing_sheets: list[str] = []
        self._err_keys: list[str] = []
        self._part_keys: list[str] = []
        self._desc_index: list[tuple[str, str]] = []
        self._hosp_index: list[tuple[str, dict]] = []

    @classmethod
    def from_bytes(cls, data: bytes) -> "Reference":
        ref = cls()
        wb = load_workbook(io.BytesIO(data), data_only=True, keep_links=False)
        ref.missing_sheets = [n for n in REQUIRED_SHEETS if get_sheet(wb, n) is None]

        ws = get_sheet(wb, SHEET_ERR)
        if ws is not None:
            rows = _sheet_rows(ws)
            hi, keys = _find_header(rows, ["item", "error code"])
            if hi is not None:
                ci = keys.index("error code")
                di = keys.index("alarm description") if "alarm description" in keys else ci + 2
                for row in rows[hi + 1:]:
                    if ci >= len(row):
                        continue
                    code = norm_text(row[ci])
                    if is_na(code):
                        continue
                    desc = norm_text(row[di]) if di < len(row) else ""
                    key = norm_code(code)
                    ref.error_codes.setdefault(key, (code, desc))
                    base = re.sub(r"\(\d+\)$", "", key)
                    ref.error_codes.setdefault(base, (code, desc))
        ref._err_keys = list(ref.error_codes)

        ws = get_sheet(wb, SHEET_SPARE)
        if ws is not None:
            rows = _sheet_rows(ws)
            hi, keys = _find_header(rows, ["nipro order code", "product description"])
            if hi is not None:
                ci, di = keys.index("nipro order code"), keys.index("product description")
                for row in rows[hi + 1:]:
                    if ci >= len(row):
                        continue
                    code = norm_text(row[ci])
                    if is_na(code):
                        continue
                    desc = norm_text(row[di]) if di < len(row) else ""
                    key = norm_code(code)
                    if key not in ref.spare_parts:
                        ref.spare_parts[key] = (code, desc)
                        if desc:
                            ref._desc_index.append((norm_key(desc), key))
        ref._part_keys = list(ref.spare_parts)

        ws = get_sheet(wb, SHEET_DATA2)
        if ws is not None:
            rows = _sheet_rows(ws)
            hi, keys = _find_header(rows, ["serial number", "province"], limit=15)
            if hi is not None:
                idx: dict[str, int] = {}
                for field, label in D2_FIELDS.items():
                    for i, k in enumerate(keys):
                        if k == label:
                            idx[field] = i
                            break
                    else:
                        for i, k in enumerate(keys):
                            if k.startswith(label):
                                idx[field] = i
                                break
                for row in rows[hi + 1:]:
                    rec = {f: (row[i] if i < len(row) else None) for f, i in idx.items()}
                    sn = norm_serial(rec.get("serial"))
                    if not sn:
                        continue
                    rec["serial"] = norm_text(rec.get("serial"))
                    for f in ("province", "model", "name_th", "name_en", "contract_no", "type", "status", "pm_contract"):
                        rec[f] = norm_text(rec.get(f))
                    ref.machines.append(rec)
                    ref.serial_index.setdefault(sn, rec)
                    for nm in (rec["name_th"], rec["name_en"]):
                        h = _hn(nm)
                        if h:
                            ref._hosp_index.append((h, rec))

        ws = get_sheet(wb, SHEET_SVC)
        if ws is not None:
            for row in ws.iter_rows(values_only=True):
                a = norm_text(row[0]) if len(row) > 0 else ""
                b = norm_text(row[1]) if len(row) > 1 else ""
                if a and norm_key(a) not in ("service type", "type") and a not in ref.service_types:
                    ref.service_types.append(a)
                if b and norm_key(b) not in ("condition",) and b not in ref.conditions:
                    ref.conditions.append(b)
        wb.close()
        return ref

    # -- lookups -----------------------------------------------------------
    def check_error(self, code: str) -> tuple[str, str, str]:
        key = norm_code(code)
        if key in self.error_codes:
            c, d = self.error_codes[key]
            return "OK", c, d
        cm = difflib.get_close_matches(key, self._err_keys, n=1, cutoff=0.8)
        if cm:
            c, d = self.error_codes[cm[0]]
            return "SUGGEST", c, d
        return "UNKNOWN", "", ""

    def _name_suggest(self, name: str) -> Optional[tuple[str, str]]:
        nm = norm_key(name)
        if not nm:
            return None
        best, best_key = 0.0, ""
        for dn, key in self._desc_index:
            sc = 1.0 if dn == nm else difflib.SequenceMatcher(None, nm, dn).ratio()
            if sc > best:
                best, best_key = sc, key
        if best >= 0.85:
            return self.spare_parts[best_key]
        return None

    def check_part(self, code: str, name: str = "") -> tuple[str, str, str]:
        key = norm_code(code)
        if key and key in self.spare_parts:
            c, d = self.spare_parts[key]
            return "OK", c, d
        if key:
            cm = difflib.get_close_matches(key, self._part_keys, n=1, cutoff=0.85)
            if cm:
                c, d = self.spare_parts[cm[0]]
                return "SUGGEST", c, d
        hit = self._name_suggest(name)
        if hit:
            return "SUGGEST", hit[0], hit[1]
        return ("UNKNOWN", "", "") if key else ("MISSING", "", "")

    def match_serial(self, sn: str) -> Optional[dict]:
        return self.serial_index.get(norm_serial(sn))

    def suggest_serial(self, sn: str) -> Optional[dict]:
        key = norm_serial(sn)
        if not key:
            return None
        cm = difflib.get_close_matches(key, list(self.serial_index), n=1, cutoff=0.88)
        return self.serial_index[cm[0]] if cm else None

    def match_hospital(self, name: str) -> Optional[dict]:
        h = _hn(name)
        if not h:
            return None
        best, rec = 0.0, None
        for key, m in self._hosp_index:
            sc = _hscore(h, key)
            if sc > best:
                best, rec = sc, m
        return rec if best >= 0.78 else None


# ----------------------------------------------------------------------------
# SQLite persistence
# ----------------------------------------------------------------------------
@contextmanager
def db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    with db() as conn:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS reports ("
            + ", ".join(f"{c} INTEGER" if c == "include" else f"{c} TEXT" for c in REPORT_COLS)
            + ", PRIMARY KEY(rid))"
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS parts ("
            + ", ".join(f"{c} REAL" if c == "qty" else f"{c} TEXT" for c in PART_COLS)
            + ", PRIMARY KEY(pid))"
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS expenses ("
            + ", ".join(f"{c} REAL" if c in EXP_NUM else f"{c} TEXT" for c in EXP_COLS)
            + ", PRIMARY KEY(date))"
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS files (page_key TEXT PRIMARY KEY, name TEXT, page INTEGER, "
            "status TEXT, n_reports INTEGER, processed_at TEXT)"
        )


def ensure_cols(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = None
    return df[cols]


def read_table(table: str, cols: list[str]) -> pd.DataFrame:
    with db() as conn:
        df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
    return ensure_cols(df, cols)


def assign_ids(df: pd.DataFrame, col: str) -> pd.DataFrame:
    df = df.copy()
    seen: set[str] = set()
    ids = []
    for v in df[col].tolist():
        s = norm_text(v)
        if not s or s in seen:
            s = uuid.uuid4().hex[:12]
        seen.add(s)
        ids.append(s)
    df[col] = ids
    return df


def replace_table(table: str, cols: list[str], df: pd.DataFrame) -> None:
    df = ensure_cols(df, cols)
    records = [tuple(_py(v) for v in row) for row in df.itertuples(index=False, name=None)]
    with db() as conn:
        conn.execute(f"DELETE FROM {table}")
        if records:
            conn.executemany(
                f"INSERT OR REPLACE INTO {table} ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                records,
            )
    try:  # CSV mirror (cache redundancy)
        df.to_csv(DATA_DIR / f"{table}.csv", index=False, encoding="utf-8-sig")
    except Exception:
        pass


def load_reports() -> pd.DataFrame:
    df = read_table("reports", REPORT_COLS)
    df["include"] = df["include"].apply(lambda v: True if is_blank(v) else bool(int(float(v))))
    return df


def load_parts() -> pd.DataFrame:
    return read_table("parts", PART_COLS)


def load_expenses() -> pd.DataFrame:
    return read_table("expenses", EXP_COLS)


def save_reports(df: pd.DataFrame) -> None:
    df = ensure_cols(df, REPORT_COLS)
    df = df[~df.drop(columns=["include"]).isna().all(axis=1)]
    df = assign_ids(df, "rid")
    df["include"] = df["include"].apply(lambda v: 1 if is_blank(v) else int(bool(v)))
    df["source"] = df["source"].apply(lambda v: norm_text(v) or "manual")
    replace_table("reports", REPORT_COLS, df)
    with db() as conn:
        conn.execute("DELETE FROM parts WHERE rid NOT IN (SELECT rid FROM reports)")


def file_status(page_key: str) -> str:
    with db() as conn:
        row = conn.execute("SELECT status FROM files WHERE page_key=?", (page_key,)).fetchone()
    return row[0] if row else ""


def mark_file(page_key: str, name: str, page: int, status: str, n: int) -> None:
    with db() as conn:
        conn.execute(
            "INSERT OR REPLACE INTO files VALUES (?,?,?,?,?,?)",
            (page_key, name, page, status[:300], n, datetime.now().isoformat(timespec="seconds")),
        )


def delete_page_data(page_key: str) -> None:
    with db() as conn:
        conn.execute("DELETE FROM parts WHERE rid IN (SELECT rid FROM reports WHERE page_key=?)", (page_key,))
        conn.execute("DELETE FROM reports WHERE page_key=?", (page_key,))


def clear_all() -> None:
    with db() as conn:
        for t in ("reports", "parts", "expenses", "files"):
            conn.execute(f"DELETE FROM {t}")


# ----------------------------------------------------------------------------
# Page rendering (PDF/images -> JPEG)
# ----------------------------------------------------------------------------
def _to_jpeg(img: Image.Image) -> bytes:
    img = ImageOps.exif_transpose(img).convert("RGB")
    if max(img.size) > MAX_SIDE:
        img.thumbnail((MAX_SIDE, MAX_SIDE), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=JPEG_QUALITY)
    return buf.getvalue()


def count_pages(name: str, data: bytes) -> int:
    if name.lower().endswith(".pdf"):
        if fitz is None:
            raise RuntimeError("ไม่พบไลบรารี PyMuPDF (pip install PyMuPDF)")
        doc = fitz.open(stream=data, filetype="pdf")
        try:
            if doc.needs_pass:
                raise RuntimeError("PDF ถูกตั้งรหัสผ่าน")
            return doc.page_count
        finally:
            doc.close()
    with Image.open(io.BytesIO(data)) as im:
        return max(getattr(im, "n_frames", 1), 1)


def iter_pages(name: str, data: bytes):
    """Yield (page_no starting at 1, jpeg_bytes)."""
    if name.lower().endswith(".pdf"):
        if fitz is None:
            raise RuntimeError("ไม่พบไลบรารี PyMuPDF (pip install PyMuPDF)")
        doc = fitz.open(stream=data, filetype="pdf")
        try:
            for i in range(doc.page_count):
                pix = doc[i].get_pixmap(matrix=fitz.Matrix(PDF_ZOOM, PDF_ZOOM), colorspace=fitz.csRGB, alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                yield i + 1, _to_jpeg(img)
        finally:
            doc.close()
    else:
        with Image.open(io.BytesIO(data)) as im:
            for i, frame in enumerate(ImageSequence.Iterator(im)):
                yield i + 1, _to_jpeg(frame.copy())


# ----------------------------------------------------------------------------
# AI vision extraction
# ----------------------------------------------------------------------------
SYSTEM_PROMPT = (
    "You are an expert OCR/data-entry assistant for Nipro (Thailand) dialysis-machine service reports "
    "(PM reports and Service reports, Thai + English, printed and handwritten). "
    "Read the page image carefully and return ONLY valid JSON. Never invent data: use an empty string "
    "or empty list when something is not visible or unreadable."
)

USER_PROMPT = """Extract every service report visible on this page.
Return JSON exactly in this schema:
{"reports":[{
 "report_no": "Book/No. printed on the form, else ''",
 "date": "YYYY-MM-DD (convert Buddhist-era years, e.g. 2569 -> 2026; 2-digit year 26 -> 2026; day comes first in dd/mm/yy)",
 "customer_name": "hospital / center name as written",
 "machine_model": "e.g. Surdial 55, Surdial 55 Plus II, NCE-18, Phoenix One+ DS",
 "serial_numbers": ["every machine serial number listed, e.g. 21CX22513"],
 "condition": "Under warranty | Service contract | Other (the ticked box)",
 "report_type": "PM | Service | Installation",
 "error_codes": ["e.g. E97, NCUE100"],
 "service_description": "service details / engineer comments (keep Thai text)",
 "start_time": "HH:MM 24h or ''", "finish_time": "HH:MM 24h or ''",
 "parts": [{"name": "", "code": "e.g. SP-0-504-232", "qty": 1, "remark": "e.g. CF-DMS-261024"}],
 "confidence": 0.0-1.0,
 "warnings": ["anything unclear"]
}]}
Rules: only parts listed under 'Part replacement' are replaced parts. If the page is blank or is not a
service report, return {"reports": []}. One physical form = one item."""

DEMO_PAYLOAD = {"reports": [{
    "report_no": "08789", "date": "2026-09-07", "customer_name": "Rajthanee Hospital (Ayutthaya)",
    "machine_model": "Surdial 55 Plus II", "serial_numbers": ["21CX22513"], "condition": "Service contract",
    "report_type": "Service", "error_codes": ["E97"], "service_description": "เปลี่ยนท่อน้ำ เปลี่ยนอะไหล่ตามรายการ",
    "start_time": "18:00", "finish_time": "20:00",
    "parts": [{"name": "Clamp box", "code": "SP-0-504-232", "qty": 1, "remark": "CF-DMS-261024"}],
    "confidence": 0.9, "warnings": [],
}]}


class VisionClient:
    def __init__(self, provider: str, api_key: str, model: str) -> None:
        self.provider, self.api_key, self.model = provider, api_key.strip(), model.strip()

    def extract(self, jpeg: bytes) -> dict:
        if self.provider == PROVIDER_DEMO:
            return json.loads(json.dumps(DEMO_PAYLOAD))
        if not self.api_key:
            raise ValueError("ยังไม่ได้ใส่ API key ที่แถบด้านซ้าย")
        last: Optional[Exception] = None
        for attempt in range(3):
            try:
                return parse_json_loose(self._call(jpeg))
            except Exception as e:  # network / rate limit / bad JSON
                last = e
                time.sleep(2 ** attempt)
        raise RuntimeError(f"AI extraction failed: {last}")

    def _call(self, jpeg: bytes) -> str:
        b64 = base64.b64encode(jpeg).decode("ascii")
        if self.provider == PROVIDER_CLAUDE:
            import anthropic

            client = anthropic.Anthropic(api_key=self.api_key)
            msg = client.messages.create(
                model=self.model, max_tokens=4096, system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                    {"type": "text", "text": USER_PROMPT}]}],
            )
            return "".join(getattr(b, "text", "") for b in msg.content)
        if self.provider == PROVIDER_OPENAI:
            from openai import OpenAI

            client = OpenAI(api_key=self.api_key)
            resp = client.chat.completions.create(
                model=self.model, response_format={"type": "json_object"},
                messages=[{"role": "system", "content": SYSTEM_PROMPT},
                          {"role": "user", "content": [
                              {"type": "text", "text": USER_PROMPT},
                              {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}]}],
            )
            return resp.choices[0].message.content or ""
        raise ValueError(f"Unknown provider: {self.provider}")


# ----------------------------------------------------------------------------
# AI output -> DB records (with cross-check enrichment)
# ----------------------------------------------------------------------------
def build_records(raw: Any, source: str, page: int, page_key: str, ref: Optional[Reference]):
    items = raw.get("reports") if isinstance(raw, dict) else None
    if items is None:
        items = [raw] if isinstance(raw, dict) and raw else []
    reports: list[dict] = []
    parts: list[dict] = []
    for idx, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        rid = hashlib.sha1(f"{page_key}:{idx}".encode()).hexdigest()[:12]
        serials = split_list(it.get("serial_numbers") or it.get("serial_number"))
        d = parse_date(it.get("date"))
        customer = norm_text(it.get("customer_name"))
        notes = [norm_text(w) for w in (it.get("warnings") or []) if norm_text(w)] if isinstance(it.get("warnings"), list) else []
        machine = None
        if ref is not None:
            for sn in serials:
                machine = ref.match_serial(sn)
                if machine:
                    break
            if machine is None and customer:
                machine = ref.match_hospital(customer)
                if machine:
                    notes.append("จับคู่โรงพยาบาลจากชื่อ (ไม่พบ serial ใน Data2)")
        eng = norm_text(machine.get("name_en")) if machine else ""
        th = norm_text(machine.get("name_th")) if machine else ""
        if not eng and not th and customer:
            if has_thai(customer):
                th = customer
            else:
                eng = customer
        if d is None:
            notes.append("อ่านวันที่ไม่ได้/รูปแบบไม่ถูกต้อง")
        reports.append({
            "include": 1, "rid": rid, "page_key": page_key, "source": source, "page": str(page),
            "report_no": norm_text(it.get("report_no")), "date": fmt_date(d), "customer": customer,
            "hospital_eng": eng, "hospital_th": th,
            "province": norm_text(machine.get("province")) if machine else "",
            "model": norm_text(it.get("machine_model")) or (norm_text(machine.get("model")) if machine else ""),
            "serials": ", ".join(serials), "condition": map_condition(it.get("condition")),
            "service_kind": map_kind(it.get("report_type")),
            "start_time": fmt_time(parse_time(it.get("start_time"))),
            "end_time": fmt_time(parse_time(it.get("finish_time") or it.get("end_time"))),
            "error_codes": ", ".join(split_list(it.get("error_codes"))),
            "description": norm_text(it.get("service_description")), "status_code": "",
            "quotation_no": "", "remark": "", "confidence": norm_text(it.get("confidence")),
            "notes": "; ".join(notes),
        })
        plist = it.get("parts") if isinstance(it.get("parts"), list) else []
        for pi, p in enumerate(plist):
            if not isinstance(p, dict):
                continue
            name, code = norm_text(p.get("name")), norm_text(p.get("code")).upper()
            if not name and not code:
                continue
            if ref is not None and code and norm_code(code) in ref.spare_parts:
                code = ref.spare_parts[norm_code(code)][0]
            remark = norm_text(p.get("remark"))
            m = re.search(r"CF-?DMS-?\d+", remark, flags=re.I)
            parts.append({
                "pid": hashlib.sha1(f"{rid}:{pi}".encode()).hexdigest()[:12], "rid": rid,
                "part_name": name, "part_code": code, "qty": to_float(p.get("qty"), 1.0),
                "unit": norm_text(p.get("unit")) or "pcs",
                "consignment_no": m.group(0).upper() if m else remark,
            })
    return reports, parts


def store_records(reports: list[dict], parts: list[dict]) -> None:
    with db() as conn:
        for r in reports:
            conn.execute(
                f"INSERT OR REPLACE INTO reports ({','.join(REPORT_COLS)}) VALUES ({','.join('?' * len(REPORT_COLS))})",
                tuple(_py(r.get(c)) if c != "include" else int(r.get(c, 1)) for c in REPORT_COLS),
            )
        for p in parts:
            conn.execute(
                f"INSERT OR REPLACE INTO parts ({','.join(PART_COLS)}) VALUES ({','.join('?' * len(PART_COLS))})",
                tuple(_py(p.get(c)) for c in PART_COLS),
            )


def process_upload(files: list[tuple[str, bytes]], client: VisionClient, ref: Optional[Reference],
                   reprocess: bool, on_progress=None) -> dict:
    stats = {"pages": 0, "reports": 0, "skipped": 0, "errors": []}
    plan = []
    for name, data in files:
        try:
            plan.append((name, data, count_pages(name, data)))
        except Exception as e:
            stats["errors"].append(f"{name}: เปิดไฟล์ไม่ได้ ({e})")
    total = max(sum(p[2] for p in plan), 1)
    done = 0
    for name, data, _n in plan:
        fhash = hashlib.sha1(data).hexdigest()[:16]
        try:
            for page_no, jpeg in iter_pages(name, data):
                done += 1
                page_key = f"{fhash}:{page_no}"
                if on_progress:
                    on_progress(done / total, f"{name} - หน้า {page_no}")
                if not reprocess and file_status(page_key) == "ok":
                    stats["skipped"] += 1
                    continue
                try:
                    raw = client.extract(jpeg)
                    recs, prts = build_records(raw, name, page_no, page_key, ref)
                    delete_page_data(page_key)
                    store_records(recs, prts)
                    mark_file(page_key, name, page_no, "ok", len(recs))
                    stats["pages"] += 1
                    stats["reports"] += len(recs)
                except Exception as e:
                    mark_file(page_key, name, page_no, f"error: {e}", 0)
                    stats["errors"].append(f"{name} หน้า {page_no}: {e}")
        except Exception as e:
            stats["errors"].append(f"{name}: {e}")
    return stats


# ----------------------------------------------------------------------------
# Cross-check validation
# ----------------------------------------------------------------------------
def _included(reports: pd.DataFrame) -> list[dict]:
    if reports is None or reports.empty:
        return []
    rep = reports.copy()
    rep["include"] = rep["include"].apply(lambda v: True if is_blank(v) else bool(v))
    return rep[rep["include"]].to_dict("records")


def validate_all(reports: pd.DataFrame, parts: pd.DataFrame, ref: Reference) -> pd.DataFrame:
    issues: list[dict] = []
    pmap: dict[str, list[dict]] = defaultdict(list)
    if parts is not None and not parts.empty:
        for p in parts.to_dict("records"):
            pmap[norm_text(p.get("rid"))].append(p)

    for r in _included(reports):
        rid = norm_text(r.get("rid"))
        base = {"rid": rid, "date": norm_text(r.get("date")),
                "hospital": norm_text(r.get("hospital_eng")) or norm_text(r.get("customer"))}

        def add(check, value, status, suggestion="", detail=""):
            issues.append({**base, "check": check, "value": value, "status": status,
                           "suggestion": suggestion, "detail": detail})

        if parse_date(r.get("date")) is None:
            add("Date", norm_text(r.get("date")), "ERROR", "", "วันที่ว่างหรือรูปแบบไม่ถูกต้อง")
        serials = split_list(r.get("serials"))
        if not serials:
            add("Serial", "", "MISSING", "", "ไม่มี Serial number")
        for sn in serials:
            if ref.match_serial(sn):
                add("Serial", sn, "OK")
            else:
                sug = ref.suggest_serial(sn)
                if sug:
                    add("Serial", sn, "SUGGEST", sug["serial"], "ไม่พบใน Data2 - ใกล้เคียงกับรายการนี้")
                else:
                    add("Serial", sn, "UNKNOWN", "", "ไม่พบใน Data2")
        for code in split_list(r.get("error_codes")):
            if is_na(code):
                continue
            st_, canon, desc = ref.check_error(code)
            add("Error code", code, st_, canon if st_ != "OK" else "", desc if st_ == "OK" else "ไม่พบใน Error Code sheet")
        for p in pmap.get(rid, []):
            code, name = norm_text(p.get("part_code")), norm_text(p.get("part_name"))
            if is_na(code) and not name:
                continue
            st_, canon, desc = ref.check_part(code, name)
            label = code or name
            add("Spare part", label, st_, canon if st_ != "OK" else "",
                desc if st_ == "OK" else ("ไม่มีรหัสอะไหล่" if st_ == "MISSING" else "ไม่พบใน Spare part code sheet"))
            q = to_float(p.get("qty"))
            if q is None or q <= 0:
                add("Spare part qty", label, "WARN", "", "จำนวนว่าง/น้อยกว่าหรือเท่ากับ 0")
    cols = ["rid", "date", "hospital", "check", "value", "status", "suggestion", "detail"]
    return pd.DataFrame(issues, columns=cols)


def apply_suggestions(reports: pd.DataFrame, parts: pd.DataFrame, ref: Reference):
    n = 0
    reports, parts = reports.copy(), parts.copy()
    for i, r in reports.iterrows():
        codes = split_list(r.get("error_codes"))
        new = []
        for c in codes:
            if is_na(c):
                new.append(c)
                continue
            st_, canon, _ = ref.check_error(c)
            if st_ in ("OK", "SUGGEST") and canon and canon != c:
                new.append(canon)
                n += 1
            else:
                new.append(c)
        if new != codes:
            reports.at[i, "error_codes"] = ", ".join(new)
    for i, p in parts.iterrows():
        code, name = norm_text(p.get("part_code")), norm_text(p.get("part_name"))
        st_, canon, desc = ref.check_part(code, name)
        if st_ in ("OK", "SUGGEST") and canon and canon != code:
            parts.at[i, "part_code"] = canon
            if not name:
                parts.at[i, "part_name"] = desc
            n += 1
    return reports, parts, n


# ----------------------------------------------------------------------------
# Business logic: build rows for each target sheet
# ----------------------------------------------------------------------------
def display_name(r: dict, lang: str) -> str:
    eng, th, cu = norm_text(r.get("hospital_eng")), norm_text(r.get("hospital_th")), norm_text(r.get("customer"))
    return (eng or cu or th) if lang == "ENG" else (th or cu or eng)


# ค่าที่ใช้ในชีต SN tracking form (ตามถ้อยคำของแถวเดิมและตารางสรุป)
SN_COND = {"Warranty": "Warranty", "PM Contract": "Contract", "Fixed asset": "Fixed asset", "Work Order": "Work Order"}
SN_KIND = {"PM": "PM", "Service": "Service", "Installation": "Installation"}


def build_sn_rows(reports: pd.DataFrame, parts: pd.DataFrame, ref: Optional[Reference], opts: dict) -> list[dict]:
    recs = _included(reports)
    recs.sort(key=lambda r: (parse_date(r.get("date")) or date.max, norm_text(r.get("source")), norm_text(r.get("page"))))
    pmap: dict[str, list[dict]] = defaultdict(list)
    if parts is not None and not parts.empty:
        for p in parts.to_dict("records"):
            if not is_blank(p.get("part_code")) or not is_blank(p.get("part_name")):
                pmap[norm_text(p.get("rid"))].append(p)
    rows: list[dict] = []
    for r in recs:
        d = parse_date(r.get("date"))
        serials = split_list(r.get("serials")) or [""]
        errors = [e for e in split_list(r.get("error_codes")) if not is_na(e)]
        plist = pmap.get(norm_text(r.get("rid")), [])
        kind = norm_text(r.get("service_kind")) or "Service"
        stype = SN_KIND.get(kind, "Service")
        cond_out = SN_COND.get(norm_text(r.get("condition")), "")
        details = norm_text(r.get("description")) or ("PM" if kind == "PM" else "")
        quotation = norm_text(r.get("quotation_no")) or "-"
        for si, sn in enumerate(serials):
            machine = ref.match_serial(sn) if (ref and sn) else None
            n = max(len(errors), len(plist), 1) if si == 0 else 1
            for i in range(n):
                if si == 0 and i < len(errors):
                    ecode, edesc = errors[i], ""
                    if ref:
                        s_, canon, desc = ref.check_error(ecode)
                        if s_ == "OK":
                            ecode, edesc = canon, desc
                else:
                    ecode, edesc = ("PM", "-") if kind == "PM" else ("N/A", "-")
                pcode, pdesc, qty, unit, cons = "N/A", "-", "-", "", "-"
                if si == 0 and i < len(plist):
                    p = plist[i]
                    pcode, pname = norm_text(p.get("part_code")), norm_text(p.get("part_name"))
                    pdesc = pname or "-"
                    if ref and pcode:
                        s_, canon, desc = ref.check_part(pcode, pname)
                        if s_ == "OK":
                            pcode, pdesc = canon, desc
                    pcode = pcode or "N/A"
                    q = to_float(p.get("qty"), 1.0)
                    qty = int(q) if q is not None and float(q).is_integer() else q
                    unit = norm_text(p.get("unit"))
                    cons = norm_text(p.get("consignment_no")) or "-"
                contract = start_c = end_c = freq = ""
                if machine:
                    contract = machine.get("contract_no") or ""
                    start_c, end_c = parse_date(machine.get("pm_start")), parse_date(machine.get("pm_end"))
                    cnt = to_float(machine.get("pm_count"))
                    if cnt and cnt > 0:
                        freq = int(round(12 / cnt))
                rows.append({
                    "person": opts.get("staff", ""), "date": d,
                    "hospital": norm_text(r.get("hospital_eng")) or norm_text(r.get("customer")),
                    "model": norm_text(r.get("model")) or (machine.get("model") if machine else ""),
                    "serial": sn, "condition": cond_out, "service_type": stype,
                    "error_code": ecode, "error_desc": edesc, "part_code": pcode, "part_desc": pdesc,
                    "qty": qty, "unit": unit, "consignment": cons, "quotation": quotation,
                    "details": details, "status": norm_text(r.get("status_code")),
                    "contract_no": contract, "start_contract": start_c, "end_contract": end_c, "freq": freq,
                    "customer": (machine.get("name_th") if machine and machine.get("name_th") else norm_text(r.get("customer"))),
                    "remark": norm_text(r.get("remark")),
                })
    return rows


def aggregate_days(reports: pd.DataFrame, lang: str) -> dict[date, dict]:
    days: dict[date, dict] = {}
    for r in _included(reports):
        d = parse_date(r.get("date"))
        if d is None:
            continue
        a = days.setdefault(d, {"pm": [], "svc": [], "inst": [], "pm_q": 0, "svc_q": 0, "inst_q": 0,
                                "province": "", "start": None, "finish": None})
        kind = norm_text(r.get("service_kind"))
        b = "pm" if kind == "PM" else "inst" if kind == "Installation" else "svc"
        name = display_name(r, lang)
        if name and name not in a[b]:
            a[b].append(name)
        a[b + "_q"] += max(len(split_list(r.get("serials"))), 1)
        if not a["province"]:
            a["province"] = norm_text(r.get("province"))
        s, e = parse_time(r.get("start_time")), parse_time(r.get("end_time"))
        if s and (a["start"] is None or s < a["start"]):
            a["start"] = s
        if e and (a["finish"] is None or e > a["finish"]):
            a["finish"] = e
    return days


def build_expense_df(reports: pd.DataFrame, saved: pd.DataFrame, lang: str) -> pd.DataFrame:
    gen: dict[str, dict] = {}
    for r in _included(reports):
        d = parse_date(r.get("date"))
        if d is None:
            continue
        key = d.isoformat()
        g = gen.setdefault(key, {"date": key, "description": [], "service_type": ""})
        nm = display_name(r, lang)
        if nm and nm not in g["description"]:
            g["description"].append(nm)
        cond = norm_text(r.get("condition"))
        if not g["service_type"] and cond in ("Warranty", "PM Contract"):
            g["service_type"] = cond
    rows = {k: {"date": k, "description": " / ".join(v["description"]), "service_type": v["service_type"],
                **{c: None for c in EXP_NUM}} for k, v in gen.items()}
    if saved is not None and not saved.empty:
        for s in saved.to_dict("records"):
            key = norm_text(s.get("date"))
            if not key:
                continue
            row = rows.setdefault(key, {"date": key, "description": "", "service_type": "", **{c: None for c in EXP_NUM}})
            for c in ("description", "service_type"):
                if norm_text(s.get(c)):
                    row[c] = norm_text(s.get(c))
            for c in EXP_NUM:
                if not is_blank(s.get(c)):
                    row[c] = to_float(s.get(c))
    df = pd.DataFrame(list(rows.values()), columns=EXP_COLS)
    return df.sort_values("date").reset_index(drop=True) if not df.empty else df


# ----------------------------------------------------------------------------
# Excel writing: openpyxl is used ONLY to read/locate. The file itself is patched at XML level,
# so external links, shapes, styles and every untouched sheet stay byte-identical.
# ----------------------------------------------------------------------------
class Recorder:
    def __init__(self) -> None:
        self.reset()

    def reset(self) -> None:
        self.cells: dict = defaultdict(list)     # sheet -> [(row, col, value)]
        self.formulas: dict = defaultdict(list)  # sheet -> [(row, col, formula_without_equals)]
        self.styles: dict = defaultdict(list)    # sheet -> [(src_row, dst_row, ncols)]

    def edits(self) -> dict:
        names = set(self.cells) | set(self.formulas) | set(self.styles)
        return {n: {"cells": self.cells.get(n, []), "formulas": self.formulas.get(n, []),
                    "styles": self.styles.get(n, [])} for n in names}


REC = Recorder()


def row_values(ws, r: int) -> tuple:
    if r < 1 or r > ws.max_row:
        return ()
    return next(ws.iter_rows(min_row=r, max_row=r, values_only=True), ())


def find_row(ws, col: int, pred, start: int = 1, end: int = 60) -> Optional[int]:
    for r in range(start, min(end, ws.max_row) + 1):
        if pred(norm_key(ws.cell(r, col).value)):
            return r
    return None


def locate_cols(ws, header_rows: tuple, spec: dict, log: list, sheet: str) -> dict:
    cache = {hr: [norm_key(v) for v in row_values(ws, hr)] for hr in header_rows}
    found = {}
    for key, (labels, fallback) in spec.items():
        col = None
        for mode in ("exact", "prefix"):
            for lab in labels:
                for hr in header_rows:
                    for ci, k in enumerate(cache[hr], start=1):
                        if (mode == "exact" and k == lab) or (mode == "prefix" and k and k.startswith(lab)):
                            col = ci
                            break
                    if col:
                        break
                if col:
                    break
            if col:
                break
        if col is None:
            col = fallback
            log.append({"sheet": sheet, "item": key, "result": f"warning: ไม่พบหัวคอลัมน์ '{labels[0]}' ใช้คอลัมน์ {fallback} แทน"})
        found[key] = col
    return found


def set_cell(ws, row: int, col: int, value: Any, overwrite: bool = False,
             allow_placeholder: bool = False) -> str:
    """Plan a write (kept in memory + recorded). Never touches formulas."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return "empty"
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        return "merged-cell"
    cur = cell.value
    if isinstance(cur, str) and cur.startswith("="):
        return "kept-formula"
    if not is_blank(cur) and not overwrite:
        if not (allow_placeholder and norm_key(cur) in PLACEHOLDERS):
            return "occupied"
    cell.value = value
    REC.cells[ws.title].append((row, col, value))
    return "written"


def _cached(wsv, r: int, c: int) -> Any:
    if wsv is None:
        return None
    try:
        return wsv.cell(r, c).value
    except Exception:
        return None


def _sheet_month(ws, wsv) -> Optional[date]:
    mr = find_row(ws, 1, lambda t: t.startswith("month"), end=15)
    if mr is None:
        return None
    for src in (wsv, ws):
        if src is None:
            continue
        d = parse_date(src.cell(mr, 2).value)
        if d:
            return d.replace(day=1)
    return None


def detect_tms(ws, wsv, base_hint: Optional[date], log: list) -> dict:
    hdr = find_row(ws, 1, lambda t: t == "date", end=40)
    if hdr is None:
        log.append({"sheet": SHEET_TMS, "item": "header", "result": "warning: ไม่พบหัวตาราง 'Date' ใช้แถว 5"})
        hdr = 5
    first = hdr + 2
    for r in range(hdr, min(hdr + 4, ws.max_row) + 1):
        if "(units)" in norm_key(ws.cell(r, 5).value):
            first = r + 1
            break
    labels: dict[str, int] = {}
    for ci, v in enumerate(row_values(ws, hdr), start=1):
        k = norm_key(v)
        if k and k not in labels:
            labels[k] = ci

    def get(k: str, fb: int) -> int:
        if k in labels:
            return labels[k]
        log.append({"sheet": SHEET_TMS, "item": k, "result": f"warning: ไม่พบหัวคอลัมน์ '{k}' ใช้คอลัมน์ {fb}"})
        return fb

    pm, svc, inst = get("pm", 16), get("service", 18), get("installation", 20)
    wt, route, rem = get("working time", 22), get("route (province)", 7), get("remark", 25)
    base = parse_date(_cached(wsv, first, 1)) or _sheet_month(ws, wsv) or base_hint
    return {"first": first, "base": base, "cols": {
        "province": pm - 1, "dep": route, "arr": route + 1, "pm_h": pm, "pm_q": pm + 1,
        "svc_h": svc, "svc_q": svc + 1, "inst_h": inst, "inst_q": inst + 1,
        "start": wt, "finish": wt + 1, "remark": rem}}


def fill_tms(ws, wsv, days: dict, opts: dict, log: list) -> int:
    hint = opts.get("month_override") or (min(days).replace(day=1) if days else None)
    lay = detect_tms(ws, wsv, hint, log)
    base = opts.get("month_override") or lay["base"]
    if base is None:
        log.append({"sheet": SHEET_TMS, "item": "month", "result": "error: หาเดือนของชีตไม่ได้ (ระบุที่ Month override)"})
        return 0
    lay["base"] = base
    c, ow, written = lay["cols"], opts.get("overwrite", False), 0
    for d, a in sorted(days.items()):
        off = (d - base).days
        if not 0 <= off <= 30:
            log.append({"sheet": SHEET_TMS, "item": d.isoformat(), "result": f"skipped: วันที่อยู่นอกเดือน {base:%Y-%m}"})
            continue
        r = lay["first"] + off
        cached = parse_date(_cached(wsv, r, 1))
        if cached and cached != d:
            log.append({"sheet": SHEET_TMS, "item": d.isoformat(), "result": f"skipped: แถว {r} เป็นวันที่ {cached} ไม่ตรง"})
            continue
        res = []
        for key, val in (("pm_h", " / ".join(a["pm"])), ("pm_q", a["pm_q"] or None),
                         ("svc_h", " / ".join(a["svc"])), ("svc_q", a["svc_q"] or None),
                         ("inst_h", " / ".join(a["inst"])), ("inst_q", a["inst_q"] or None),
                         ("province", a["province"]), ("start", a["start"]), ("finish", a["finish"])):
            res.append((key, set_cell(ws, r, c[key], val, ow)))
        prov = norm_text(a["province"])
        if opts.get("fill_route", True) and prov and norm_key(prov) not in ("bkk", "bangkok"):
            res.append(("dep", set_cell(ws, r, c["dep"], "BKK", ow)))
            res.append(("arr", set_cell(ws, r, c["arr"], prov, ow)))
        for key, rs in res:
            if rs == "written":
                written += 1
            elif rs not in ("empty",):
                log.append({"sheet": SHEET_TMS, "item": f"{d.isoformat()} / {key} (row {r})", "result": rs})
    return written


SN_SPEC = {
    "person": (["person in charge (eng)"], 1), "date": (["date"], 2), "hospital": (["hospital name (eng)"], 3),
    "model": (["model"], 4), "serial": (["serial number"], 5), "condition": (["condition"], 6),
    "service_type": (["service type"], 7), "error_code": (["error code"], 8),
    "error_desc": (["error description"], 9), "part_code": (["spare part code"], 10),
    "part_desc": (["spare part description"], 11), "qty": (["qty"], 12), "unit": (["unit"], 13),
    "consignment": (["consignment no."], 14), "quotation": (["quotation no."], 15),
    "details": (["service details (th)"], 16), "status": (["status of service type"], 17),
    "contract_no": (["contract no."], 18), "start_contract": (["start contract"], 19),
    "end_contract": (["end contract"], 20), "freq": (["frequency of pm"], 21),
    "customer": (["customer name"], 22), "remark": (["remark"], 23),
}


def fill_sn(ws, rows: list[dict], opts: dict, log: list) -> int:
    hdr = find_row(ws, 1, lambda t: t.startswith("person in charge"), end=40)
    if hdr is None:
        raise ValueError("ไม่พบหัวตาราง 'Person in charge (ENG)' ในชีต SN tracking form")
    cols = locate_cols(ws, (hdr, hdr + 1), SN_SPEC, log, SHEET_SN)
    first = hdr + 2 if "(only name)" in norm_key(ws.cell(hdr + 1, 1).value) else hdr + 1
    ncols = max(cols.values())
    limit = min(ws.max_row, first + 3000)

    def used(r: int) -> bool:
        for k in ("date", "hospital", "serial", "model"):
            v = ws.cell(r, cols[k]).value
            if not is_blank(v) and not (isinstance(v, str) and v.startswith("=")):
                return True
        return False

    # 1) หาตารางสรุป (Warranty / Contract / Work Order / Fixed asset) ที่อยู่ใต้ข้อมูล
    stop = None
    for r in range(first, limit + 1):
        if any(norm_key(ws.cell(r, cols[k]).value) == "warranty" for k in ("model", "hospital", "date")):
            stop = r
            break
    end = (stop - 1) if stop else limit
    last_used = first - 1
    for r in range(first, end + 1):
        if used(r):
            last_used = r
    tmpl = last_used if last_used >= first else first

    # 2) แถวว่างที่เขียนได้ = ต่อจากแถวข้อมูลสุดท้าย จนถึงก่อนตารางสรุป
    if stop:
        free = list(range(last_used + 1, stop))
    else:
        free = list(range(last_used + 1, last_used + 1 + len(rows)))
    if len(rows) > len(free):
        raise ValueError(
            f"ชีต SN tracking form มีแถวว่างก่อนตารางสรุป (แถว {stop}) เหลือ {len(free)} แถว แต่ต้องใช้ {len(rows)} แถว "
            f"กรุณาเปิดไฟล์ Master ใน Excel แล้วแทรกแถวว่างเพิ่มอย่างน้อย {len(rows) - len(free)} แถว "
            f"โดยแทรกไว้ 'กลาง' ช่วงข้อมูล (เช่น คลิกขวาที่แถวสุดท้ายของข้อมูล > Insert) เพื่อให้สูตรสรุปขยายตาม "
            f"แล้วอัปโหลด Master ใหม่"
        )

    # 3) คอลัมน์ที่แถวต้นแบบเป็นสูตร -> คัดลอกสูตรไปแถวใหม่ ไม่เขียนค่าทับ
    tf: dict[int, str] = {}
    for c in range(1, ncols + 1):
        v = ws.cell(tmpl, c).value
        if isinstance(v, str) and v.startswith("="):
            tf[c] = v

    written = 0
    for rec, r in zip(rows, free):
        REC.styles[ws.title].append((tmpl, r, ncols))
        for c, f in tf.items():
            try:
                nf = Translator(f, origin=f"{get_column_letter(c)}{tmpl}").translate_formula(f"{get_column_letter(c)}{r}")
            except Exception as e:
                log.append({"sheet": SHEET_SN, "item": f"row {r} / col {c}", "result": f"formula copy failed: {e}"})
                continue
            REC.formulas[ws.title].append((r, c, nf[1:]))
            written += 1
        for key, val in rec.items():
            col = cols.get(key)
            if col is None or col in tf:
                continue
            res = set_cell(ws, r, col, val, True)
            if res == "written":
                written += 1
            elif res not in ("empty", "occupied"):
                log.append({"sheet": SHEET_SN, "item": f"row {r} / {key}", "result": res})
    letters = ", ".join(get_column_letter(c) for c in sorted(tf)) or "-"
    log.append({"sheet": SHEET_SN, "item": "layout",
                "result": f"แถวต้นแบบ {tmpl}; เขียนแถว {free[0] if free else '-'}-{free[len(rows) - 1] if rows else '-'}; "
                          f"ตารางสรุปเริ่มแถว {stop or '-'}; คัดลอกสูตรคอลัมน์ {letters}"})
    return written


EXP_SPEC = {
    "description": (["description"], 2), "allowance": (["allowance"], 7), "petrol": (["petrol"], 10),
    "expressway": (["express way"], 11), "parking": (["parking"], 12), "other": (["other"], 13),
    "service_type": (["service type"], 15),
}


def fill_expense(ws, wsv, exp_df: pd.DataFrame, base_hint: Optional[date], opts: dict, log: list) -> int:
    hdr = None
    for r in range(1, min(40, ws.max_row) + 1):
        if norm_key(ws.cell(r, 1).value) == "date" and norm_key(ws.cell(r, 2).value).startswith("description"):
            hdr = r
            break
    if hdr is None:
        raise ValueError("ไม่พบหัวตาราง 'Date / Description' ในชีต Expense Claim")
    cols = locate_cols(ws, (hdr,), EXP_SPEC, log, SHEET_EXP)
    first = hdr + 1
    v = ws.cell(first, 1).value
    if isinstance(v, str) and v.strip() and not v.startswith("=") and parse_date(v) is None:
        first = hdr + 2
    base = opts.get("month_override") or parse_date(_cached(wsv, first, 1)) or base_hint
    if base is None:
        log.append({"sheet": SHEET_EXP, "item": "month", "result": "error: หาเดือนของชีตไม่ได้"})
        return 0
    ow, written = opts.get("overwrite", False), 0
    for rec in exp_df.to_dict("records"):
        d = parse_date(rec.get("date"))
        if d is None:
            continue
        off = (d - base).days
        if not 0 <= off <= 30:
            log.append({"sheet": SHEET_EXP, "item": d.isoformat(), "result": "skipped: นอกเดือน"})
            continue
        r = first + off
        cached = parse_date(_cached(wsv, r, 1))
        if cached and cached != d:
            log.append({"sheet": SHEET_EXP, "item": d.isoformat(), "result": f"skipped: แถว {r} เป็นวันที่ {cached}"})
            continue
        stype = norm_text(rec.get("service_type"))
        items = [("description", norm_text(rec.get("description")), True),
                 ("service_type", stype, False)]
        for k in EXP_NUM:
            val = to_float(rec.get(k))
            items.append((k, val if val else None, False))
        for key, val, ph in items:
            res = set_cell(ws, r, cols[key], val, ow, allow_placeholder=ph)
            if res == "written":
                written += 1
            elif res not in ("empty",):
                log.append({"sheet": SHEET_EXP, "item": f"{d.isoformat()} / {key} (row {r})", "result": res})
    return written


# ---- XML-level patcher ------------------------------------------------------
_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_XML_NS = "http://www.w3.org/XML/1998/namespace"


def _q(tag: str) -> str:
    return f"{{{_NS}}}{tag}"


def _sheet_paths(zf: zipfile.ZipFile) -> dict[str, str]:
    wbx = etree.fromstring(zf.read("xl/workbook.xml"))
    rels = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_target = {r.get("Id"): (r.get("Target") or "") for r in rels}
    out: dict[str, str] = {}
    sheets = wbx.find(_q("sheets"))
    for s in (sheets if sheets is not None else []):
        t = rid_target.get(s.get(f"{{{_RNS}}}id"), "").lstrip("/")
        out[norm_key(s.get("name"))] = t if t.startswith("xl/") else "xl/" + t
    return out


def _xml_value(v: Any) -> tuple[str, str]:
    if isinstance(v, bool):
        return "b", "1" if v else "0"
    if isinstance(v, (datetime, date, dtime)):
        return "n", repr(float(to_excel(v)))
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (v != v or v in (float("inf"), float("-inf"))):
            return "s", ""
        return "n", repr(v)
    return "s", str(v)


def _patch_sheet(xml: bytes, cells: list, formulas: list, styles: list) -> bytes:
    root = etree.fromstring(xml)
    sd = root.find(_q("sheetData"))
    if sd is None:
        raise ValueError("ไม่พบ sheetData ในชีต")
    rows = {int(r.get("r")): r for r in sd.findall(_q("row")) if r.get("r")}

    def get_row(n: int):
        row = rows.get(n)
        if row is not None:
            row.attrib.pop("spans", None)
            return row
        row = etree.Element(_q("row"))
        row.set("r", str(n))
        later = [k for k in rows if k > n]
        if later:
            rows[min(later)].addprevious(row)
        else:
            sd.append(row)
        rows[n] = row
        return row

    def peek(n: int, c: int):
        row = rows.get(n)
        if row is None:
            return None
        ref = f"{get_column_letter(c)}{n}"
        for ce in row.findall(_q("c")):
            if ce.get("r") == ref:
                return ce
        return None

    def get_cell(n: int, c: int):
        found = peek(n, c)
        if found is not None:
            return found
        row = get_row(n)
        new = etree.Element(_q("c"))
        new.set("r", f"{get_column_letter(c)}{n}")
        for ce in row.findall(_q("c")):
            m = re.match(r"[A-Z]+", ce.get("r") or "")
            if m and column_index_from_string(m.group(0)) > c:
                ce.addprevious(new)
                return new
        row.append(new)
        return new

    def reset(ce) -> None:
        for ch in list(ce):
            ce.remove(ch)
        ce.attrib.pop("t", None)

    for src_r, dst_r, ncols in styles:          # 1) รูปแบบ (เส้นขอบ/รูปแบบวันที่) จากแถวต้นแบบ
        for c in range(1, ncols + 1):
            src = peek(src_r, c)
            dst = get_cell(dst_r, c)
            if src is not None and src.get("s") is not None:
                dst.set("s", src.get("s"))
    for r, c, f in formulas:                    # 2) สูตรที่คัดลอกจากแถวต้นแบบ
        ce = get_cell(r, c)
        reset(ce)
        etree.SubElement(ce, _q("f")).text = f
    for r, c, v in cells:                       # 3) ค่า (ไม่เขียนทับเซลล์ที่มีสูตร)
        ce = get_cell(r, c)
        if ce.find(_q("f")) is not None:
            continue
        reset(ce)
        kind, text = _xml_value(v)
        if kind == "n":
            etree.SubElement(ce, _q("v")).text = text
        elif kind == "b":
            ce.set("t", "b")
            etree.SubElement(ce, _q("v")).text = text
        else:
            ce.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(ce, _q("is")), _q("t"))
            t.text = text
            if text != text.strip():
                t.set(f"{{{_XML_NS}}}space", "preserve")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _force_recalc(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    calc = root.find(_q("calcPr"))
    if calc is None:
        calc = etree.Element(_q("calcPr"))
        anchor = None
        for tag in ("sheets", "functionGroups", "externalReferences", "definedNames"):
            e = root.find(_q(tag))
            if e is not None:
                anchor = e
        if anchor is not None:
            anchor.addnext(calc)
        else:
            root.append(calc)
    calc.set("fullCalcOnLoad", "1")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_xlsx(master: bytes, edits: dict, force_recalc: bool = True) -> bytes:
    zin = zipfile.ZipFile(io.BytesIO(master))
    try:
        names = set(zin.namelist())
        paths = _sheet_paths(zin)
        new_parts: dict[str, bytes] = {}
        for sheet, e in edits.items():
            path = paths.get(norm_key(sheet))
            if not path or path not in names:
                raise ValueError(f"ไม่พบไฟล์ XML ของชีต {sheet}")
            new_parts[path] = _patch_sheet(zin.read(path), e["cells"], e["formulas"], e["styles"])
        if force_recalc:
            new_parts["xl/workbook.xml"] = _force_recalc(zin.read("xl/workbook.xml"))
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_parts.get(item.filename)
                zout.writestr(item, data if data is not None else zin.read(item.filename))
        return out.getvalue()
    finally:
        zin.close()


def fill_workbook(master: bytes, reports: pd.DataFrame, parts: pd.DataFrame, expenses: pd.DataFrame,
                  ref: Optional[Reference], opts: dict) -> tuple[bytes, list[dict]]:
    log: list[dict] = []
    REC.reset()
    wb = load_workbook(io.BytesIO(master), keep_links=False)   # อ่านเพื่อหาตำแหน่งเท่านั้น ไม่ Save
    try:
        wbv = load_workbook(io.BytesIO(master), data_only=True, keep_links=False)
    except Exception:
        wbv = None
    sheets = {n: get_sheet(wb, n) for n in (SHEET_TMS, SHEET_SN, SHEET_EXP)}
    missing = [n for n, w in sheets.items() if w is None]
    if missing:
        raise ValueError("ไม่พบชีต: " + ", ".join(missing))
    vals = {n: (get_sheet(wbv, n) if wbv else None) for n in sheets}

    days = aggregate_days(reports, opts.get("lang", "ENG"))
    n_tms = fill_tms(sheets[SHEET_TMS], vals[SHEET_TMS], days, opts, log)
    sn_rows = build_sn_rows(reports, parts, ref, opts)
    n_sn = fill_sn(sheets[SHEET_SN], sn_rows, opts, log)
    hint = opts.get("month_override") or (min(days).replace(day=1) if days else None)
    n_exp = fill_expense(sheets[SHEET_EXP], vals[SHEET_EXP], expenses, hint, opts, log)
    out = patch_xlsx(master, REC.edits(), opts.get("force_recalc", True))
    log.insert(0, {"sheet": "SUMMARY", "item": "cells written",
                   "result": f"TMS={n_tms}, SN={n_sn} (rows={len(sn_rows)}), Expense={n_exp}"})
    return out, log


# ----------------------------------------------------------------------------
# Self test (python app.py --selftest)
# ----------------------------------------------------------------------------
def _demo_master() -> bytes:
    from openpyxl.styles import Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_TMS
    ws["A1"], ws["A2"], ws["B2"] = "Monthly Report : TMS Business Trip", "Month :", datetime(2026, 9, 1)
    for c, t in {1: "Date", 3: "Province", 4: "Hospital Name", 5: "Qty", 7: "Route (Province)", 9: "Mileage",
                 15: "Province", 16: "PM", 17: "Qty", 18: "Service", 19: "Qty", 20: "Installation", 21: "Qty",
                 22: "Working Time", 25: "Remark"}.items():
        ws.cell(5, c, t)
    ws["E6"] = "(units)"
    ws["A7"] = "=B2"
    for r in range(8, 38):
        ws.cell(r, 1, f"=A{r - 1}+1")
    for r in range(7, 38):
        ws.cell(r, 24, f"=W{r}-V{r}")

    sn = wb.create_sheet(SHEET_SN)
    for c, t in enumerate(["Person in charge (ENG)", "Date", "Hospital Name (ENG)", "Model", "Serial number",
                           "Condition", "Service type", "Urgent call", "Spare parts"], 1):
        sn.cell(5, c, t)
    for c, t in {16: "Service Details (TH)", 17: "Status of Service Type", 18: "Contract No.", 19: "Start Contract",
                 20: "End contract", 21: "Frequency of PM", 22: "Customer name", 23: "Remark"}.items():
        sn.cell(5, c, t)
    for c, t in {1: "(only name)", 6: "(Drop Down List)", 7: "(Drop Down List)", 8: "Error code",
                 9: "Error Description", 10: "Spare part code", 11: "Spare part Description", 12: "QTY",
                 13: "Unit", 14: "Consignment No.", 15: "Quotation No.", 21: "(every xx mths)"}.items():
        sn.cell(6, c, t)
    # แถวข้อมูลเดิม (ต้นแบบ) + มีรูปแบบวันที่/เส้นขอบ + คอลัมน์ F เป็นสูตร
    sn["A7"], sn["B7"], sn["C7"] = "Mr. Suvanai P.", datetime(2026, 9, 1), "Nipro Excellent Renal Care"
    sn["D7"], sn["E7"] = "NCU-18", "17DG0362"
    sn["F7"] = '=IF(E7="","",E7)'
    sn["G7"], sn["H7"] = "PM", "PM"
    sn["B7"].number_format = "dd-mmm-yy"
    sn["B7"].border = Border(bottom=Side(style="thin"))
    # ตารางสรุปที่อยู่ใต้ข้อมูล (เหลือแถวว่าง 8-10 เพียง 3 แถว)
    sn["D11"], sn["E11"] = "Warranty", 20
    sn["D12"], sn["E12"] = "Contract", 49

    ex = wb.create_sheet(SHEET_EXP)
    for c, t in {1: "Date", 2: "Description", 7: "Allowance", 10: "Petrol", 11: "Express way", 12: "Parking",
                 13: "Other", 15: "Service Type", 16: "Total Amount"}.items():
        ex.cell(7, c, t)
    ex["A8"] = "=DATE(2026,9,1)"
    for r in range(9, 39):
        ex.cell(r, 1, f"=A{r - 1}+1")
    for r in range(8, 39):
        ex.cell(r, 2, "Office")
        ex.cell(r, 16, f"=SUM(J{r}:M{r})")

    er = wb.create_sheet(SHEET_ERR)
    er.append(["Trouble Shooting List"])
    er.append(["Item", "Error Code", "Error Code (1)", "Alarm Description", "Escalation mode", "ICU mode"])
    er.append([1, "E97", "E97", "SHORT CHEMICAL SOLUTION ALARM", "HIGH", "LOW"])
    er.append([2, "E114(1)", "E114(1)", "TEMPARATURE CANNOT REACH", "HIGH", "-"])
    er.append([3, "N/A", "N/A", "-", "", ""])

    sp = wb.create_sheet(SHEET_SPARE)
    sp.append(["Item", "NIPRO Order Code", "Product description", "Remark"])
    sp.append([1, "SP-0-504-232", "Clamp box", ""])
    sp.append([2, "SP-0-201-332", "Regulator", ""])
    sp.append([3, "N/A", "-", ""])

    d2 = wb.create_sheet(SHEET_DATA2)
    d2.append(["Z", "Type", "Province", "Model name", "Engineer name", "TH - Customer name (Commercial)",
               "Serial number", None, None, "CRM", "Machine installation location",
               "ENG - Customer name (Commercial)", "Machine status", None, "Year of Installation", "Install date",
               'Start date of "PM"', 'End date of "PM"', "End PM", "Remaining days", "Quotation and Contract issued date",
               "PM contract\nYes / No", "Number of PM maintenance", None, None, None, "Contract number"])
    d2.append(["", "HD", "Ayutthaya", "Surdial 55 Plus II", "Mr. Suvanai", "โรงพยาบาลราชธานี (อยุธยา)", "21CX22513",
               None, None, "", "", "Rajthanee Hospital (Ayutthaya)", "Active", None, 2022, datetime(2022, 1, 1),
               datetime(2025, 10, 1), datetime(2026, 9, 30), None, None, None, "Yes", 4, None, None, None, "CT-001"])

    sv = wb.create_sheet(SHEET_SVC)
    for a, b in [("PM", "Warranty"), ("PM (after 18.00)", "PM Contract"), ("PM (SAT-SUN)", None),
                 ("Urgent call", None), ("Urgent call (after 18.00)", None), ("Urgent call (SAT-SUN)", None),
                 ("Installation", None)]:
        sv.append([a, b])
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def run_selftest() -> None:
    global DATA_DIR, DB_PATH
    import tempfile

    tmp = Path(tempfile.mkdtemp())
    DATA_DIR, DB_PATH = tmp, tmp / "test.db"
    init_db()

    assert parse_date("28/9/26") == date(2026, 9, 28)
    assert parse_date("9 กันยายน 2569") == date(2026, 9, 9)
    assert parse_date(46266) == date(2026, 9, 1)
    assert parse_date("2026-09-07") == date(2026, 9, 7)
    assert parse_time("18.00") == dtime(18, 0)
    assert split_list("E97, E05 / E06") == ["E97", "E05", "E06"]
    assert parse_json_loose('```json\n{"a": 1}\n```') == {"a": 1}

    master = _demo_master()
    ref = Reference.from_bytes(master)
    assert not ref.missing_sheets, ref.missing_sheets
    assert ref.check_error("e97")[0] == "OK"
    assert ref.check_error("E114")[0] == "OK"
    assert ref.check_part("SP-0-504-232")[0] == "OK"
    assert ref.check_part("SP-0-504-233")[0] in ("SUGGEST", "UNKNOWN")
    assert ref.match_serial("21cx22513") is not None

    recs, prts = build_records(DEMO_PAYLOAD, "demo.pdf", 1, "k:1", ref)
    store_records(recs, prts)
    mark_file("k:1", "demo.pdf", 1, "ok", len(recs))
    assert file_status("k:1") == "ok"
    reports, parts = load_reports(), load_parts()
    assert len(reports) == 1 and len(parts) == 1
    assert reports.loc[0, "hospital_eng"] == "Rajthanee Hospital (Ayutthaya)"
    assert reports.loc[0, "condition"] == "PM Contract"

    val = validate_all(reports, parts, ref)
    assert (val["status"] == "OK").all(), val
    exp = build_expense_df(reports, load_expenses(), "ENG")
    opts = {"staff": "Mr. Suvanai P.", "overwrite": False, "lang": "ENG", "fill_route": True,
            "month_override": None, "force_recalc": True}
    out, log = fill_workbook(master, reports, parts, exp, ref, opts)

    wb = load_workbook(io.BytesIO(out))
    sn, tms, ex = wb[SHEET_SN], wb[SHEET_TMS], wb[SHEET_EXP]
    # SN: เขียนต่อจากแถว 7 (แถวใหม่ = 8) ไม่ล้ำไปทับตารางสรุปที่แถว 11
    assert sn["A7"].value == "Mr. Suvanai P." and sn["E7"].value == "17DG0362"
    assert sn["A8"].value == "Mr. Suvanai P.", sn["A8"].value
    assert sn["B8"].value == datetime(2026, 9, 7), sn["B8"].value
    assert sn["B8"].number_format == "dd-mmm-yy", sn["B8"].number_format
    assert sn["B8"].border.bottom.style == "thin"
    assert sn["E8"].value == "21CX22513"
    assert sn["F8"].value == '=IF(E8="","",E8)', sn["F8"].value   # คัดลอกสูตรจากแถวต้นแบบ
    assert sn["G8"].value == "Service", sn["G8"].value
    assert sn["H8"].value == "E97" and str(sn["I8"].value).startswith("SHORT CHEMICAL")
    assert sn["J8"].value == "SP-0-504-232" and sn["K8"].value == "Clamp box"
    assert sn["R8"].value == "CT-001" and sn["U8"].value == 3
    assert sn["D11"].value == "Warranty" and sn["E11"].value == 20      # ตารางสรุปไม่ถูกทับ
    assert sn["D12"].value == "Contract" and sn["E12"].value == 49
    # TMS / Expense
    assert tms["R13"].value == "Rajthanee Hospital (Ayutthaya)", tms["R13"].value
    assert tms["S13"].value == 1 and tms["O13"].value == "Ayutthaya"
    assert abs(float(tms["V13"].value) - 0.75) < 1e-9, tms["V13"].value
    assert str(tms["A8"].value).startswith("="), "formula must be preserved"
    assert str(tms["X13"].value).startswith("="), "formula must be preserved"
    assert ex["B14"].value == "Rajthanee Hospital (Ayutthaya)", ex["B14"].value
    assert str(ex["P14"].value).startswith("=")
    # ส่วนที่ไม่เกี่ยวข้องต้องเหมือนต้นฉบับทุกไบต์
    zi, zo = zipfile.ZipFile(io.BytesIO(master)), zipfile.ZipFile(io.BytesIO(out))
    assert zi.namelist() == zo.namelist()
    untouched = _sheet_paths(zi)[norm_key(SHEET_ERR)]
    assert zi.read(untouched) == zo.read(untouched)
    assert b"fullCalcOnLoad" in zo.read("xl/workbook.xml")
    # แถวว่างไม่พอ ต้องหยุดพร้อมข้อความ ไม่เขียนทับตารางสรุป
    big = pd.concat([reports] * 4, ignore_index=True)
    try:
        fill_workbook(master, big, parts, exp, ref, opts)
    except ValueError:
        pass
    else:
        raise AssertionError("overflow was not detected")
    print("SELFTEST PASSED")
    for item in log:
        print(item)


# ----------------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------------
def _stretch(fn, *args, **kwargs):
    """Full-width widget that works on both old and new Streamlit versions."""
    try:
        return fn(*args, width="stretch", **kwargs)
    except Exception:
        return fn(*args, use_container_width=True, **kwargs)


def _only_options(df: pd.DataFrame, col: str, options: list[str]) -> pd.DataFrame:
    df = df.copy()
    df[col] = df[col].apply(lambda v: norm_text(v) if norm_text(v) in options else None)
    return df


def _as_text_df(rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows).astype(str)
    return df.replace({"None": "", "nan": "", "NaT": ""})


def _secret(name: str) -> str:
    try:
        return str(st.secrets.get(name, "")) or os.environ.get(name, "")
    except Exception:
        return os.environ.get(name, "")


def get_reference() -> Optional[Reference]:
    if not MASTER_PATH.exists():
        return None
    stamp = MASTER_PATH.stat().st_mtime_ns
    if st.session_state.get("ref_stamp") != stamp:
        try:
            st.session_state["ref"] = Reference.from_bytes(MASTER_PATH.read_bytes())
            st.session_state["ref_err"] = ""
        except Exception as e:
            st.session_state["ref"] = None
            st.session_state["ref_err"] = f"อ่านไฟล์ Master ไม่ได้: {e}"
        st.session_state["ref_stamp"] = stamp
    return st.session_state.get("ref")


def sidebar() -> dict:
    sb = st.sidebar
    sb.header("1) Master Excel")
    up = sb.file_uploader("อัปโหลด Master (.xlsx)", type=["xlsx"], key="master_up")
    if up is not None:
        data = up.getvalue()
        if not MASTER_PATH.exists() or hashlib.sha1(data).hexdigest() != hashlib.sha1(MASTER_PATH.read_bytes()).hexdigest():
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            MASTER_PATH.write_bytes(data)
    ref = get_reference()
    if st.session_state.get("ref_err"):
        sb.error(st.session_state["ref_err"])
    elif ref is None:
        sb.warning("ยังไม่มีไฟล์ Master")
    else:
        sb.success(f"Master พร้อมใช้ | Error codes {len(ref.error_codes)} | Spare parts {len(ref.spare_parts)} | เครื่อง {len(ref.machines)}")
        if ref.missing_sheets:
            sb.error("ชีตที่ขาด: " + ", ".join(ref.missing_sheets))

    sb.header("2) AI Vision")
    provider = sb.selectbox("Provider", [PROVIDER_CLAUDE, PROVIDER_OPENAI, PROVIDER_DEMO])
    env_name = "ANTHROPIC_API_KEY" if provider == PROVIDER_CLAUDE else "OPENAI_API_KEY"
    key = sb.text_input("API key", value=_secret(env_name), type="password", disabled=provider == PROVIDER_DEMO)
    model = sb.text_input("Model", value=DEFAULT_MODELS[provider], key=f"model_{provider}")

    sb.header("3) ตัวเลือก Export")
    staff = sb.text_input("ชื่อผู้รับผิดชอบ (SN tracking)", value="Mr. Suvanai P.")
    lang = sb.radio("ภาษาชื่อ รพ. ใน TMS/Expense", ["ENG", "TH"], horizontal=True)
    overwrite = sb.checkbox("เขียนทับค่าที่มีอยู่แล้ว (ไม่แตะสูตร)", value=False)
    fill_route = sb.checkbox("เติมเส้นทาง BKK -> จังหวัด ใน TMS", value=True)
    force_recalc = sb.checkbox("ให้ Excel คำนวณสูตรใหม่ทั้งไฟล์ตอนเปิด", value=True,
                               help="ถ้าเปิดไฟล์แล้วเจอ Circular reference ให้ลองปิดตัวเลือกนี้แล้ว Export ใหม่")
    mo = sb.text_input("Month override (YYYY-MM) เว้นว่าง = อัตโนมัติ")
    month_override = None
    if mo.strip():
        d = parse_date(mo.strip() + "-01")
        if d is None:
            sb.error("รูปแบบเดือนไม่ถูกต้อง ใช้ YYYY-MM")
        month_override = d
    return {"ref": ref, "provider": provider, "key": key, "model": model, "staff": staff, "lang": lang,
            "overwrite": overwrite, "fill_route": fill_route, "month_override": month_override,
            "force_recalc": force_recalc}


def tab_upload(cfg: dict) -> None:
    st.subheader("อัปโหลดรายงาน (PDF หลายหน้า / รูปภาพ)")
    files = st.file_uploader("เลือกไฟล์ได้หลายไฟล์", type=["pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp"],
                             accept_multiple_files=True)
    reprocess = st.checkbox("ประมวลผลซ้ำหน้าที่เคยทำแล้ว (จะแทนที่ข้อมูลเดิมของหน้านั้น)")
    if cfg["ref"] is None:
        st.info("แนะนำให้อัปโหลด Master Excel ก่อน เพื่อให้ระบบตรวจสอบรหัสและเติมข้อมูลสัญญาให้")
    if st.button("เริ่มประมวลผล", type="primary", disabled=not files):
        client = VisionClient(cfg["provider"], cfg["key"], cfg["model"])
        bar, msg = st.progress(0.0), st.empty()

        def cb(p: float, text: str) -> None:
            bar.progress(min(max(p, 0.0), 1.0))
            msg.write(f"กำลังประมวลผล: {text}")

        stats = process_upload([(f.name, f.getvalue()) for f in files], client, cfg["ref"], reprocess, cb)
        bar.progress(1.0)
        msg.empty()
        st.success(f"เสร็จสิ้น: ประมวลผล {stats['pages']} หน้า, พบรายงาน {stats['reports']} ฉบับ, ข้าม {stats['skipped']} หน้า")
        for e in stats["errors"]:
            st.error(e)
        st.session_state["ver"] = st.session_state.get("ver", 0) + 1
    with st.expander("ประวัติการประมวลผล (cache)"):
        with db() as conn:
            hist = pd.read_sql_query("SELECT name, page, status, n_reports, processed_at FROM files ORDER BY processed_at DESC", conn)
        _stretch(st.dataframe, hist, hide_index=True)
        ok = st.checkbox("ยืนยันล้างข้อมูลทั้งหมด")
        if st.button("ล้างข้อมูลที่เก็บไว้", disabled=not ok):
            clear_all()
            st.session_state["ver"] = st.session_state.get("ver", 0) + 1
            st.rerun()


def tab_review(cfg: dict) -> None:
    ver = st.session_state.get("ver", 0)
    reports, parts = load_reports(), load_parts()
    st.subheader("1) รายงาน (Reports)")
    st.caption("แก้ไขได้ทุกช่อง เพิ่ม/ลบแถวได้ | วันที่รูปแบบ YYYY-MM-DD | เวลา HH:MM | Error code/Serial คั่นด้วย ,")
    rview = _only_options(reports, "condition", CONDITION_OPTIONS)
    rview = _only_options(rview, "service_kind", KIND_OPTIONS)
    rview = _only_options(rview, "status_code", STATUS_OPTIONS)
    edited = _stretch(
        st.data_editor, rview, key=f"rep_{ver}", num_rows="dynamic", hide_index=True,
        column_config={
            "include": st.column_config.CheckboxColumn("ใช้", default=True),
            "rid": None, "page_key": None,
            "condition": st.column_config.SelectboxColumn("condition", options=CONDITION_OPTIONS),
            "service_kind": st.column_config.SelectboxColumn("service_kind", options=KIND_OPTIONS),
            "status_code": st.column_config.SelectboxColumn("status_code", options=STATUS_OPTIONS),
        },
        disabled=["source", "page"],
    )
    if st.button("บันทึก Reports", type="primary"):
        save_reports(edited)
        st.session_state["ver"] = ver + 1
        st.rerun()

    st.subheader("2) อะไหล่ที่เปลี่ยน (Parts)")
    rep_now = load_reports()
    labels = {}
    for r in rep_now.to_dict("records"):
        labels[r["rid"]] = (f"{norm_text(r['rid'])[:8]} | {norm_text(r['date']) or '?'} | "
                            f"{(norm_text(r['hospital_eng']) or norm_text(r['customer']))[:30]} | {norm_text(r['serials'])[:30]}")
    if not labels:
        st.info("ยังไม่มีรายงาน จึงยังเพิ่มอะไหล่ไม่ได้ (บันทึก Reports ก่อน)")
    else:
        inv = {v: k for k, v in labels.items()}
        pdisp = parts.copy()
        pdisp.insert(0, "report", pdisp["rid"].map(labels))
        pdisp = pdisp.drop(columns=["rid"])
        pedit = _stretch(
            st.data_editor, pdisp, key=f"parts_{ver}", num_rows="dynamic", hide_index=True,
            column_config={
                "pid": None,
                "report": st.column_config.SelectboxColumn("report", options=list(labels.values()), required=True),
                "qty": st.column_config.NumberColumn("qty", min_value=0, step=1),
            },
        )
        if st.button("บันทึก Parts"):
            out = pedit.copy()
            out["rid"] = out["report"].map(inv)
            bad = int(out["rid"].isna().sum())
            out = out[out["rid"].notna()].drop(columns=["report"])
            out = assign_ids(ensure_cols(out, PART_COLS), "pid")
            replace_table("parts", PART_COLS, out)
            if bad:
                st.warning(f"ตัดทิ้ง {bad} แถวที่ไม่ได้เลือก report")
            st.session_state["ver"] = ver + 1
            st.rerun()

    st.subheader("3) ค่าใช้จ่ายรายวัน (Expense Claim)")
    st.caption("กรอก Petrol / Express way / Parking / Other / Allowance เอง (รายงานไม่มีข้อมูลนี้)")
    exp = build_expense_df(rep_now, load_expenses(), cfg["lang"])
    exp = _only_options(exp, "service_type", SERVICE_TYPE_OPTIONS)
    eedit = _stretch(
        st.data_editor, exp, key=f"exp_{ver}", num_rows="dynamic", hide_index=True,
        column_config={
            "service_type": st.column_config.SelectboxColumn("service_type", options=SERVICE_TYPE_OPTIONS),
            **{c: st.column_config.NumberColumn(c, min_value=0, format="%.2f") for c in EXP_NUM},
        },
    )
    c1, c2 = st.columns(2)
    if c1.button("บันทึก Expense"):
        e = ensure_cols(eedit, EXP_COLS)
        e = e[e["date"].apply(lambda v: parse_date(v) is not None)].copy()
        e["date"] = e["date"].apply(lambda v: parse_date(v).isoformat())
        e = e.drop_duplicates("date", keep="last")
        replace_table("expenses", EXP_COLS, e)
        st.session_state["ver"] = ver + 1
        st.rerun()
    if c2.button("รีเซ็ตตาราง Expense"):
        replace_table("expenses", EXP_COLS, pd.DataFrame(columns=EXP_COLS))
        st.session_state["ver"] = ver + 1
        st.rerun()


def tab_validate(cfg: dict) -> None:
    st.subheader("ตรวจสอบรหัสกับ Master (Error Code / Spare part code / Data2)")
    ref = cfg["ref"]
    if ref is None:
        st.warning("กรุณาอัปโหลด Master Excel ก่อน")
        return
    reports, parts = load_reports(), load_parts()
    val = validate_all(reports, parts, ref)
    if val.empty:
        st.info("ยังไม่มีข้อมูลให้ตรวจสอบ")
        return
    counts = val["status"].value_counts()
    cols = st.columns(5)
    for col, s in zip(cols, ["OK", "SUGGEST", "UNKNOWN", "MISSING", "ERROR"]):
        col.metric(s, int(counts.get(s, 0)))
    only_issue = st.checkbox("แสดงเฉพาะรายการที่มีปัญหา", value=True)
    view = val[val["status"] != "OK"] if only_issue else val
    _stretch(st.dataframe, view, hide_index=True)
    if st.button("ปรับรหัสตามคำแนะนำอัตโนมัติ (Error code / Spare part)"):
        r2, p2, n = apply_suggestions(reports, parts, ref)
        save_reports(r2)
        replace_table("parts", PART_COLS, p2)
        st.session_state["ver"] = st.session_state.get("ver", 0) + 1
        st.success(f"ปรับแล้ว {n} รายการ (Serial ไม่ถูกแก้อัตโนมัติ ต้องตรวจเอง)")
        st.rerun()


def tab_export(cfg: dict) -> None:
    st.subheader("ตัวอย่างข้อมูลที่จะเขียนลง Excel และดาวน์โหลด")
    reports, parts = load_reports(), load_parts()
    ref = cfg["ref"]
    opts = {k: cfg[k] for k in ("staff", "lang", "overwrite", "fill_route", "month_override", "force_recalc")}
    sn_rows = build_sn_rows(reports, parts, ref, opts)
    if not sn_rows:
        st.info("ยังไม่มีข้อมูลที่เลือกใช้ (include)")
        return
    st.markdown("**SN tracking form**")
    _stretch(st.dataframe, _as_text_df(sn_rows), hide_index=True)
    days = aggregate_days(reports, cfg["lang"])
    st.markdown("**TMS Working Record (รายวัน)**")
    _stretch(st.dataframe, _as_text_df([
        {"date": d, "PM": " / ".join(a["pm"]), "PM qty": a["pm_q"], "Service": " / ".join(a["svc"]),
         "Service qty": a["svc_q"], "Installation": " / ".join(a["inst"]), "Inst qty": a["inst_q"],
         "Province": a["province"], "Start": fmt_time(a["start"]), "Finish": fmt_time(a["finish"])}
        for d, a in sorted(days.items())]) if days else pd.DataFrame(), hide_index=True)
    st.caption("ไฟล์ที่ได้แก้เฉพาะ 3 ชีตโดยตรงที่ระดับ XML ส่วนอื่น (External link, รูปร่าง, ชีตอื่น) ไม่ถูกแตะ "
               "ควรเก็บไฟล์ Master ต้นฉบับไว้เสมอ")
    if not MASTER_PATH.exists():
        st.warning("กรุณาอัปโหลด Master Excel ก่อน")
        return
    if st.button("สร้างไฟล์ Excel", type="primary"):
        try:
            exp = build_expense_df(reports, load_expenses(), cfg["lang"])
            data, log = fill_workbook(MASTER_PATH.read_bytes(), reports, parts, exp, ref, opts)
            st.session_state["export"] = (data, log)
        except Exception as e:
            st.session_state.pop("export", None)
            st.error(f"สร้างไฟล์ไม่สำเร็จ: {e}")
            with st.expander("รายละเอียดข้อผิดพลาด"):
                st.code(traceback.format_exc())
    if "export" in st.session_state:
        data, log = st.session_state["export"]
        st.success("สร้างไฟล์สำเร็จ")
        st.download_button("ดาวน์โหลด Excel", data=data, file_name=f"TMS_filled_{datetime.now():%Y%m%d_%H%M}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.markdown("**Log การเขียนไฟล์**")
        _stretch(st.dataframe, pd.DataFrame(log).astype(str), hide_index=True)


def main() -> None:
    st.set_page_config(page_title="Nipro Service Report Automation", page_icon="🛠️", layout="wide")
    try:
        init_db()
    except Exception as e:
        st.error(f"สร้างฐานข้อมูล cache ไม่สำเร็จ: {e}")
        st.stop()
    st.title("Nipro Service Report -> Excel Automation")
    if fitz is None:
        st.error("ไม่พบ PyMuPDF: ติดตั้งด้วย pip install PyMuPDF (ยังอ่านไฟล์รูปภาพได้)")
    cfg = sidebar()
    t1, t2, t3, t4 = st.tabs(["1) อัปโหลด & ประมวลผล", "2) ตรวจทาน / แก้ไข", "3) ตรวจสอบรหัส", "4) Export Excel"])
    with t1:
        tab_upload(cfg)
    with t2:
        tab_review(cfg)
    with t3:
        tab_validate(cfg)
    with t4:
        tab_export(cfg)


if __name__ == "__main__" and "--selftest" in sys.argv:
    run_selftest()
elif __name__ == "__main__":
    main()
else:  # streamlit may import the script as a module in some versions
    main()

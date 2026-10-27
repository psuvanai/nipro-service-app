# -*- coding: utf-8 -*-
"""
Nipro Service Report -> Excel automation (Streamlit)
Flow: scans -> AI reads -> SN tracking rows (editable) -> TMS + Expense derived from SN rows -> patched .xlsx

Run app :  streamlit run app.py
Self test: python app.py --selftest
"""
from __future__ import annotations

import base64
import difflib
import hashlib
import io
import json
import os
import posixpath
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
import requests
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
DB_PATH = DATA_DIR / "nipro_cache_v3.db"
MASTER_PATH = DATA_DIR / "master.xlsx"

SHEET_TMS = "TMS Working Record"
SHEET_SN = "SN tracking form"
SHEET_EXP = "Expense Claim"
SHEET_ERR = "Error Code"
SHEET_SPARE = "Spare part code"
SHEET_DATA2 = "Data2 (2)"
REQUIRED_SHEETS = [SHEET_TMS, SHEET_SN, SHEET_EXP, SHEET_ERR, SHEET_SPARE, SHEET_DATA2]

PROVIDER_GEMINI = "Google Gemini (ฟรี)"
PROVIDER_CLAUDE = "Anthropic (Claude)"
PROVIDER_OPENAI = "OpenAI"
PROVIDER_DEMO = "Demo (ไม่ใช้ AI - ทดสอบระบบ)"
DEFAULT_MODELS = {PROVIDER_GEMINI: "gemini-2.5-flash", PROVIDER_CLAUDE: "claude-sonnet-4-5",
                  PROVIDER_OPENAI: "gpt-4o", PROVIDER_DEMO: "-"}

PDF_ZOOM = 2.0
MAX_SIDE = 2000
JPEG_QUALITY = 88

KIND_OPTIONS = ["PM", "Service", "Installation"]
STATUS_OPTIONS = ["Issue DO", "Issue PM", "Claim (Warranty)", "Claim (Issue DO)"]
PLACEHOLDERS = {"office", "-", "n/a", "na", "sat", "sun"}

RED_CF = "กรุณากรอกข้อมูลช่อง CF-DMS"
RED_SPARE = "กรุณากรอกข้อมูลช่อง Spare part code"
RED_STATUS = "กรุณาลงข้อมูล Status"
RED_CHECK_STATUS = "ตรวจสอบ Status of Service Type อีกครั้ง"
RED_FILL = "โปรดลงข้อมูล"

SN_COLS = [
    "include", "cont", "date", "person", "hospital", "model", "serial", "condition", "service_type",
    "error_code", "error_desc", "part_code", "part_desc", "qty", "unit", "consignment", "quotation",
    "details", "status", "contract_no", "start_contract", "end_contract", "freq", "customer",
    "remark", "remark_red", "province", "start_time", "end_time", "source", "page", "seq", "rid", "page_key",
]
SN_INT = {"include", "cont", "seq"}
EXP_NUM = ["allowance", "petrol", "expressway", "parking", "other"]
EXP_COLS = ["date"] + EXP_NUM

DIST_KM = {
    "bangkok": 30, "bkk": 30, "nonthaburi": 30, "pathum thani": 30, "pathumthani": 30,
    "กรุงเทพมหานคร": 30, "กรุงเทพ": 30, "นนทบุรี": 30, "ปทุมธานี": 30,
    "ayutthaya": 60, "phra nakhon si ayutthaya": 60, "พระนครศรีอยุธยา": 60, "อยุธยา": 60,
}
UNIT_MAP = {"pcs": "PC", "pc": "PC", "piece": "PC", "set": "SET", "unit": "UNIT", "m": "M", "meter": "M",
            "bag": "BAG", "gallon": "GALLON", "box": "BOX"}
COND_TO_EXP = {"warranty": "Warranty", "contract": "PM Contract", "pm contract": "PM Contract"}

THAI_MONTHS = {
    "ม.ค": 1, "มกราคม": 1, "ก.พ": 2, "กุมภาพันธ์": 2, "มี.ค": 3, "มีนาคม": 3,
    "เม.ย": 4, "เมษายน": 4, "พ.ค": 5, "พฤษภาคม": 5, "มิ.ย": 6, "มิถุนายน": 6,
    "ก.ค": 7, "กรกฎาคม": 7, "ส.ค": 8, "สิงหาคม": 8, "ก.ย": 9, "กันยายน": 9,
    "ต.ค": 10, "ตุลาคม": 10, "พ.ย": 11, "พฤศจิกายน": 11, "ธ.ค": 12, "ธันวาคม": 12,
}
CONF_MAP = str.maketrans({"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "S": "5", "B": "8", "Z": "2", "G": "6"})


class Rich(list):
    """Rich text for a cell: list of (text, is_red)."""


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


def _truthy(v: Any) -> bool:
    if is_blank(v):
        return False
    if isinstance(v, str):
        return v.strip().lower() in ("1", "true", "yes", "y")
    return bool(v)


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
    m = re.search(r"(\d{1,2})\s*[:.\-]\s*(\d{2})", norm_text(v))
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if h < 24 and mi < 60:
            return dtime(h, mi)
    return None


def fmt_time(t: Optional[dtime]) -> str:
    return t.strftime("%H:%M") if t else ""


def fmt_date(d: Optional[date]) -> str:
    return d.isoformat() if d else ""


def map_kind(text: Any) -> str:
    k = norm_key(text)
    if k.startswith("pm") or "preventive" in k:
        return "PM"
    if "install" in k:
        return "Installation"
    return "Service"


def norm_unit(u: Any) -> str:
    k = norm_key(u)
    if not k:
        return "PC"
    return UNIT_MAP.get(k, norm_text(u).upper())


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


def uniq_join(items: list[str], sep: str = " | ") -> str:
    return sep.join(dict.fromkeys(norm_text(x) for x in items if norm_text(x)))


# ----------------------------------------------------------------------------
# Reference data (Error Code / Spare part code / Data2)
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
                    ref.error_codes.setdefault(re.sub(r"\(\d+\)$", "", key), (code, desc))
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
                    if not code or norm_key(code) in ("n/a", "na", "-"):
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
        wb.close()
        return ref

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
        return self.spare_parts[best_key] if best >= 0.85 else None

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

    def machines_like(self, name: str) -> list[dict]:
        rec = self.match_hospital(name) if name else None
        if not rec:
            return []
        key = norm_key(rec.get("name_en")) or norm_key(rec.get("name_th"))
        return [m for m in self.machines if (norm_key(m.get("name_en")) or norm_key(m.get("name_th"))) == key]

    def resolve_serial(self, sn: str, hospital: str = "") -> tuple[str, str, Optional[dict]]:
        """Return (serial, status, machine). status: OK / FIXED / UNSURE / MISSING."""
        key = norm_serial(sn)
        if not key:
            return "", "MISSING", None
        m = self.serial_index.get(key)
        if m:
            return m["serial"], "OK", m
        pool = self.machines_like(hospital)
        restricted = bool(pool)
        if not pool:
            pool = self.machines

        def uniq(cands: list[dict]) -> list[dict]:
            return list({norm_serial(c["serial"]): c for c in cands}.values())

        hits = [c for c in pool if len(norm_serial(c["serial"])) >= 6 and key.startswith(norm_serial(c["serial"]))]
        if hits:
            best = max(len(norm_serial(c["serial"])) for c in hits)
            top = uniq([c for c in hits if len(norm_serial(c["serial"])) == best])
            if len(top) == 1:
                return top[0]["serial"], "FIXED", top[0]
        if len(key) >= 6:
            top = uniq([c for c in pool if norm_serial(c["serial"]).startswith(key)])
            if len(top) == 1:
                return top[0]["serial"], "FIXED", top[0]
        fk = key.translate(CONF_MAP)
        top = uniq([c for c in pool if norm_serial(c["serial"]).translate(CONF_MAP) == fk])
        if len(top) == 1:
            return top[0]["serial"], "FIXED", top[0]
        scored = sorted(((difflib.SequenceMatcher(None, key, norm_serial(c["serial"])).ratio(), c)
                         for c in uniq(pool)), key=lambda x: -x[0])
        cutoff = 0.75 if restricted else 0.9
        if scored and scored[0][0] >= cutoff and (len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.05):
            return scored[0][1]["serial"], "FIXED", scored[0][1]
        return norm_text(sn), "UNSURE", None


def special_code(ref: Optional[Reference], ch: str) -> tuple[str, str]:
    if ref is not None:
        hit = ref.spare_parts.get(norm_code(ch))
        if hit:
            return hit[0], (hit[1] or "-")
    return ch, "-"


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
        conn.execute("CREATE TABLE IF NOT EXISTS sn_rows ("
                     + ", ".join(f"{c} INTEGER" if c in SN_INT else f"{c} TEXT" for c in SN_COLS)
                     + ", PRIMARY KEY(rid))")
        conn.execute("CREATE TABLE IF NOT EXISTS expenses ("
                     + ", ".join(f"{c} REAL" if c in EXP_NUM else f"{c} TEXT" for c in EXP_COLS)
                     + ", PRIMARY KEY(date))")
        conn.execute("CREATE TABLE IF NOT EXISTS files (page_key TEXT PRIMARY KEY, name TEXT, page INTEGER, "
                     "status TEXT, n_reports INTEGER, processed_at TEXT)")


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
                f"INSERT OR REPLACE INTO {table} ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})", records)
    try:  # CSV mirror
        df.to_csv(DATA_DIR / f"{table}.csv", index=False, encoding="utf-8-sig")
    except Exception:
        pass


def clean_sn(df: pd.DataFrame) -> pd.DataFrame:
    df = ensure_cols(df, SN_COLS)
    if df.empty:
        return df
    meta = ["include", "cont", "rid", "page_key", "source", "page", "seq"]
    mask = df.drop(columns=meta).apply(lambda r: all(is_blank(v) for v in r), axis=1)
    df = df[~mask].copy()
    df = assign_ids(df, "rid")
    df["include"] = df["include"].apply(lambda v: 1 if is_blank(v) else int(bool(v)))
    df["cont"] = df["cont"].apply(lambda v: 1 if _truthy(v) else 0)
    df["source"] = df["source"].apply(lambda v: norm_text(v) or "manual")
    df["seq"] = df["seq"].apply(lambda v: int(to_float(v, 9999) or 0))
    return df


def save_sn(df: pd.DataFrame) -> None:
    replace_table("sn_rows", SN_COLS, clean_sn(df))


def load_sn() -> pd.DataFrame:
    df = read_table("sn_rows", SN_COLS)
    df["include"] = df["include"].apply(lambda v: True if is_blank(v) else bool(int(float(v))))
    df["cont"] = df["cont"].apply(lambda v: False if is_blank(v) else bool(int(float(v))))
    if df.empty:
        return df
    df["_d"] = df["date"].apply(lambda v: norm_text(v) or "9999")
    df["_s"] = df["source"].apply(norm_text)
    df["_p"] = df["page"].apply(lambda v: to_float(v, 0.0) or 0.0)
    df["_q"] = df["seq"].apply(lambda v: to_float(v, 9999.0) or 0.0)
    df = df.sort_values(["_d", "_s", "_p", "_q"], kind="stable").drop(columns=["_d", "_s", "_p", "_q"])
    return df.reset_index(drop=True)


def load_expenses() -> pd.DataFrame:
    return read_table("expenses", EXP_COLS)


def store_sn_rows(rows: list[dict]) -> None:
    with db() as conn:
        for r in rows:
            vals = []
            for c in SN_COLS:
                vals.append(int(r.get(c, 1 if c == "include" else 0)) if c in ("include", "cont") else _py(r.get(c)))
            conn.execute(
                f"INSERT OR REPLACE INTO sn_rows ({','.join(SN_COLS)}) VALUES ({','.join('?' * len(SN_COLS))})",
                tuple(vals))


def file_status(page_key: str) -> str:
    with db() as conn:
        row = conn.execute("SELECT status FROM files WHERE page_key=?", (page_key,)).fetchone()
    return row[0] if row else ""


def mark_file(page_key: str, name: str, page: int, status: str, n: int) -> None:
    with db() as conn:
        conn.execute("INSERT OR REPLACE INTO files VALUES (?,?,?,?,?,?)",
                     (page_key, name, page, status[:300], n, datetime.now().isoformat(timespec="seconds")))


def delete_page_data(page_key: str) -> None:
    with db() as conn:
        conn.execute("DELETE FROM sn_rows WHERE page_key=?", (page_key,))


def clear_all() -> None:
    with db() as conn:
        for t in ("sn_rows", "expenses", "files"):
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
            raise RuntimeError("ไม่พบไลบรารี PyMuPDF")
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
    if name.lower().endswith(".pdf"):
        if fitz is None:
            raise RuntimeError("ไม่พบไลบรารี PyMuPDF")
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
 "date": "YYYY-MM-DD. Day comes first in dd/mm/yy. Buddhist-era years: 2569 -> 2026; 2-digit year 26 -> 2026. Use the date written on the form.",
 "customer_name": "hospital / center name as written",
 "machine_model": "model as written",
 "serial_numbers": ["serial numbers written in the machine / serial-number field of the form, e.g. 21CX22513"],
 "comment_serials": ["serial numbers that appear ONLY inside the Engineer comments text"],
 "report_type": "PM | Service | Installation",
 "errors": [{"code": "e.g. E49", "outcome": "repair | quote | replaced", "detail": "short note of what was done for THIS error (keep Thai/English as written)"}],
 "service_description": "text of the Service description column (keep Thai exactly as written)",
 "engineer_comment": "text of the 'Engineer comments' box (bottom right). Keep Thai exactly as written",
 "quotation_no": "e.g. Q25-DMS-724 or ''",
 "start_time": "HH:MM 24h or ''", "finish_time": "HH:MM 24h or ''",
 "parts": [{"name": "", "code": "e.g. SP-0-504-232", "qty": 1, "unit": "PC | SET | UNIT | M | BAG | GALLON | BOX", "consignment_no": "e.g. CF-DMS-261024 or ''", "where": "table | description"}],
 "confidence": 0.0-1.0,
 "warnings": ["anything unclear"]
}]}
Rules:
- errors: one item per error code on the form. outcome = "quote" when the engineer comment says เสนอราคา for it,
  "replaced" when a part was replaced and listed under 'Part replacement', otherwise "repair".
- parts: items listed under 'Part replacement' have where="table". Only if that table is empty and the
  description says a part was replaced (with a part code), add it with where="description".
- If the page is blank or not a service report, return {"reports": []}. One physical form = one item."""

PAYLOAD_SERVICE = {"reports": [{
    "report_no": "08789", "date": "2026-09-07", "customer_name": "Rajthanee Hospital (Ayutthaya)",
    "machine_model": "Surdial 55 Plus II", "serial_numbers": ["21CX22513"], "comment_serials": ["21CX22517"],
    "report_type": "Service", "errors": [{"code": "E97", "outcome": "replaced", "detail": "เปลี่ยนท่อน้ำ"}],
    "service_description": "เปลี่ยนท่อน้ำ", "engineer_comment": "เปลี่ยนแล้ว",
    "quotation_no": "", "start_time": "18:00", "finish_time": "20:00",
    "parts": [{"name": "Clamp box", "code": "SP-0-504-232", "qty": 1, "unit": "", "consignment_no": "CF-DMS-261010",
               "where": "table"}],
    "confidence": 0.9, "warnings": [],
}]}
PAYLOAD_QUOTE = {"reports": [{
    "report_no": "08759", "date": "2026-09-03", "customer_name": "Krungthai Renal Care Dialysis Clinic",
    "machine_model": "Surdial 55", "serial_numbers": ["17CX6407"], "comment_serials": [],
    "report_type": "Service",
    "errors": [{"code": "E97", "outcome": "quote", "detail": ""}, {"code": "E114", "outcome": "quote", "detail": ""}],
    "service_description": "", "engineer_comment": "เสนอราคา SW Power supply SP-0-505-009",
    "quotation_no": "", "start_time": "09:00", "finish_time": "10:00", "parts": [],
    "confidence": 0.9, "warnings": [],
}]}
PAYLOAD_PM = {"reports": [{
    "report_no": "08800", "date": "2026-09-30", "customer_name": "Nipro Excellent Renal Care",
    "machine_model": "Surdial 55", "serial_numbers": ["19CX14513", "19CX145147", "19CX14515"], "comment_serials": [],
    "report_type": "PM", "errors": [], "service_description": "PM", "engineer_comment": "เปลี่ยน Endotoxin filter",
    "quotation_no": "", "start_time": "13:00", "finish_time": "16:00", "parts": [],
    "confidence": 0.9, "warnings": [],
}]}
PAYLOAD_INST = {"reports": [{
    "report_no": "08700", "date": "2026-09-03", "customer_name": "สถาบันทรวงอก",
    "machine_model": "PY-YD-340", "serial_numbers": ["2403-055"], "comment_serials": [],
    "report_type": "Installation", "errors": [], "service_description": "ติดตั้งเครื่อง",
    "engineer_comment": "ติดตั้งเครื่อง", "quotation_no": "", "start_time": "13:00", "finish_time": "15:00",
    "parts": [], "confidence": 0.9, "warnings": [],
}]}


class VisionClient:
    def __init__(self, provider: str, api_key: str, model: str) -> None:
        self.provider, self.api_key, self.model = provider, api_key.strip(), model.strip()
        self.pause = 6.5 if provider == PROVIDER_GEMINI else 0.0   # stay inside the free-tier rate limit

    def extract(self, jpeg: bytes) -> dict:
        if self.provider == PROVIDER_DEMO:
            return json.loads(json.dumps(PAYLOAD_SERVICE))
        if not self.api_key:
            raise ValueError("ยังไม่ได้ใส่ API key ที่แถบด้านซ้าย")
        last: Optional[Exception] = None
        for wait in (4, 15, 40, 0):
            try:
                return parse_json_loose(self._call(jpeg))
            except Exception as e:
                last = e
                if wait:
                    time.sleep(wait)
        raise RuntimeError(f"AI extraction failed: {last}")

    @staticmethod
    def _post(url: str, headers: dict, payload: dict) -> dict:
        r = requests.post(url, headers=headers, json=payload, timeout=180)
        if r.status_code != 200:
            raise RuntimeError(f"HTTP {r.status_code}: {r.text[:300]}")
        return r.json()

    def _call(self, jpeg: bytes) -> str:
        b64 = base64.b64encode(jpeg).decode("ascii")
        if self.provider == PROVIDER_GEMINI:
            gen: dict[str, Any] = {"responseMimeType": "application/json", "temperature": 0}
            if "2.5" in self.model and "pro" not in self.model:
                gen["thinkingConfig"] = {"thinkingBudget": 0}
            data = self._post(
                f"https://generativelanguage.googleapis.com/v1beta/models/{self.model}:generateContent",
                {"x-goog-api-key": self.api_key, "Content-Type": "application/json"},
                {"systemInstruction": {"parts": [{"text": SYSTEM_PROMPT}]},
                 "contents": [{"role": "user", "parts": [{"text": USER_PROMPT},
                                                          {"inlineData": {"mimeType": "image/jpeg", "data": b64}}]}],
                 "generationConfig": gen})
            cands = data.get("candidates") or []
            if not cands:
                raise RuntimeError(f"Gemini ไม่ตอบ: {data.get('promptFeedback')}")
            parts = (cands[0].get("content") or {}).get("parts") or []
            return "".join(p.get("text", "") for p in parts)
        if self.provider == PROVIDER_CLAUDE:
            data = self._post(
                "https://api.anthropic.com/v1/messages",
                {"x-api-key": self.api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
                {"model": self.model, "max_tokens": 4096, "system": SYSTEM_PROMPT,
                 "messages": [{"role": "user", "content": [
                     {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                     {"type": "text", "text": USER_PROMPT}]}]})
            return "".join(b.get("text", "") for b in data.get("content", []))
        if self.provider == PROVIDER_OPENAI:
            data = self._post(
                "https://api.openai.com/v1/chat/completions",
                {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"},
                {"model": self.model, "response_format": {"type": "json_object"},
                 "messages": [{"role": "system", "content": SYSTEM_PROMPT},
                              {"role": "user", "content": [
                                  {"type": "text", "text": USER_PROMPT},
                                  {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}]}]})
            return data["choices"][0]["message"]["content"] or ""
        raise ValueError(f"Unknown provider: {self.provider}")


# ----------------------------------------------------------------------------
# AI output -> SN tracking rows (the master table)
# ----------------------------------------------------------------------------
def _outcome(v: Any) -> str:
    k = norm_key(v)
    if "เสนอราคา" in k or k.startswith("quot"):
        return "quote"
    if k.startswith("repl") or k.startswith("change"):
        return "replaced"
    if k.startswith("repa") or k.startswith("fix") or k.startswith("no part"):
        return "repair"
    return ""


def _parse_errors(it: dict) -> list[dict]:
    out: list[dict] = []
    raw = it.get("errors")
    if isinstance(raw, list):
        for e in raw:
            if isinstance(e, dict):
                code, oc, det = norm_text(e.get("code")), _outcome(e.get("outcome")), norm_text(e.get("detail"))
            else:
                code, oc, det = norm_text(e), "", ""
            if code and not is_na(code):
                out.append({"code": code, "outcome": oc, "detail": det})
    if not out:
        for c in split_list(it.get("error_codes")):
            if not is_na(c):
                out.append({"code": c, "outcome": "", "detail": ""})
    return out


def _rows_for_report(it: dict, idx: int, source: str, page: int, page_key: str,
                     ref: Optional[Reference], staff: str, seq0: int) -> list[dict]:
    d = parse_date(it.get("date"))
    customer = norm_text(it.get("customer_name"))
    kind = map_kind(it.get("report_type"))
    comment = norm_text(it.get("engineer_comment"))
    svc_desc = norm_text(it.get("service_description"))
    quote_flag = "เสนอราคา" in (comment + " " + svc_desc)
    conf = to_float(it.get("confidence"), 1.0)
    conf = 1.0 if conf is None else conf
    ai_warn = [norm_text(w) for w in it.get("warnings") if norm_text(w)] if isinstance(it.get("warnings"), list) else []
    serials = split_list(it.get("serial_numbers") or it.get("serial_number"))
    csn = split_list(it.get("comment_serials"))
    if not serials and csn:
        serials, csn = csn[:1], csn[1:]
    serials = serials or [""]
    errs = _parse_errors(it)
    plist = [p for p in (it.get("parts") if isinstance(it.get("parts"), list) else [])
             if isinstance(p, dict) and (norm_text(p.get("name")) or norm_text(p.get("code")))]
    blob = " ".join([comment, svc_desc] + [norm_text(p.get("consignment_no")) + " " + norm_text(p.get("remark"))
                                            for p in plist])
    quotation = norm_text(it.get("quotation_no"))
    if not quotation:
        mq = re.search(r"\bQ\d{2}-[A-Z]{2,5}-\d+\b", blob)
        quotation = mq.group(0) if mq else ""
    start_t = fmt_time(parse_time(it.get("start_time")))
    end_t = fmt_time(parse_time(it.get("finish_time") or it.get("end_time")))

    red0: list[str] = []
    blk0: list[str] = []
    if d is None:
        red0.append("กรุณาตรวจสอบวันที่")
    if csn:
        blk0.append("Serial อื่นในใบงาน: " + ", ".join(csn))
    if ai_warn:
        blk0.append("AI: " + "; ".join(ai_warn)[:150])
    if conf < 0.6:
        red0.append("AI ไม่มั่นใจ กรุณาตรวจสอบทั้งแถว")
    if kind == "PM" and (plist or errs):
        red0.append("ใบงาน PM มีรายการ Error/อะไหล่ กรุณาตรวจสอบ")
    if kind != "PM" and errs and plist and (len(errs) > 1 or len(plist) > 1):
        red0.append("กรุณาตรวจสอบการจับคู่ Error กับอะไหล่")

    rows: list[dict] = []
    seq = seq0
    for si, sn_raw in enumerate(serials):
        first_serial = si == 0
        if ref is not None and sn_raw:
            sn, sstat, m = ref.resolve_serial(sn_raw, customer)
        elif sn_raw:
            sn, sstat, m = sn_raw, "OK", None
        else:
            sn, sstat, m = "", "MISSING", None
        red_s: list[str] = []
        blk_s: list[str] = []
        if sstat == "UNSURE":
            red_s.append("กรุณาตรวจสอบ Serial (ไม่พบใน Data2)")
        elif sstat == "FIXED":
            blk_s.append(f"Serial ในใบงาน {sn_raw} ปรับตาม Data2")
        elif sstat == "MISSING":
            red_s.append("กรุณากรอก Serial")

        n = 1 if kind == "PM" else (max(len(errs), len(plist), 1) if first_serial else 1)
        for i in range(n):
            red: list[str] = []
            blk: list[str] = []
            if i == 0:
                red += red_s
                blk += blk_s
                if first_serial:
                    red += red0
                    blk += blk0
            usable = first_serial and kind != "PM"
            err = errs[i] if (usable and i < len(errs)) else None
            part = plist[i] if (usable and i < len(plist)) else None
            row_detail = (err["detail"] if err else "") or svc_desc or comment

            ecode, edesc = ("PM", "-") if kind == "PM" else ("N/A", "-")
            if err:
                ecode, edesc = err["code"], ""
                if ref is not None:
                    s_, canon, desc = ref.check_error(ecode)
                    if s_ == "OK":
                        ecode, edesc = canon, desc
                    elif s_ == "SUGGEST":
                        red.append(f"Error {ecode} ไม่พบในชีต Error Code (ใกล้เคียง {canon})")
                    else:
                        red.append(f"Error {ecode} ไม่พบในชีต Error Code")

            pcode, pdesc, qty, unit, cons = "N/A", "-", "-", "", "-"
            details = row_detail
            status = ""
            if kind == "PM":
                details = "PM"
            elif part is not None:
                pname, pcode = norm_text(part.get("name")), norm_text(part.get("code")).upper()
                pdesc = pname or "-"
                raw_c = norm_text(part.get("consignment_no")) + " " + norm_text(part.get("remark"))
                mm = re.search(r"CF-?DMS-?\d+", raw_c, flags=re.I)
                if mm:
                    cons = re.sub(r"(?i)CF-?DMS-?", "CF-DMS-", mm.group(0))
                if ref is not None and pcode:
                    s_, canon, desc = ref.check_part(pcode, pname)
                    if s_ == "OK":
                        pcode, pdesc = canon, desc
                    elif s_ == "SUGGEST":
                        red.append(f"รหัสอะไหล่ {pcode} ไม่พบในชีต Spare part code (ใกล้เคียง {canon})")
                    else:
                        red.append(f"รหัสอะไหล่ {pcode} ไม่พบในชีต Spare part code")
                elif not pcode:
                    s_, canon, desc = ref.check_part("", pname) if ref is not None else ("MISSING", "", "")
                    if s_ == "SUGGEST":
                        pcode, pdesc = canon, desc
                        red.append(f"รหัสอะไหล่เดาจากชื่อ ({canon}) กรุณาตรวจสอบ")
                    else:
                        red.append(RED_SPARE)
                if norm_key(part.get("where")) == "description":
                    red.append("รหัสอะไหล่อ่านจากข้อความ Service description กรุณาตรวจสอบ")
                q = to_float(part.get("qty"), 1.0) or 1.0
                qty = int(q) if float(q).is_integer() else q
                unit = norm_unit(part.get("unit"))
                if quote_flag and not (err and err["outcome"]):
                    blk.append("เสนอราคา")
                    red.append("ใบงานมีคำว่าเสนอราคาและมีอะไหล่ใน Part replacement กรุณาตรวจสอบ")
                if mm:
                    status = "Issue DO"
                    red.append(RED_CHECK_STATUS)
                else:
                    red.append(RED_CF)
            elif kind == "Installation":
                pcode, pdesc = "-", "-"
            elif not first_serial:
                red.append("ใบงานมีหลาย Serial กรุณาตรวจสอบ Error/อะไหล่ของเครื่องนี้")
            else:
                outcome = err["outcome"] if err else ""
                if conf < 0.6:
                    pcode, pdesc = "", "-"
                    red.append(RED_SPARE)
                elif outcome == "quote" or (not outcome and quote_flag):
                    pcode, pdesc = special_code(ref, "w")
                    details = "เสนอราคา"
                    blk.append("เสนอราคา")
                else:
                    pcode, pdesc = special_code(ref, "s")
            if not status:
                red.append(RED_STATUS)

            seq += 1
            rows.append({
                "include": 1, "cont": 1 if (first_serial and i > 0) else 0,
                "date": fmt_date(d), "person": staff,
                "hospital": norm_text(m.get("name_en")) if m and m.get("name_en") else customer,
                "model": norm_text(m.get("model")) if m and m.get("model") else norm_text(it.get("machine_model")),
                "serial": sn, "condition": norm_text(m.get("status")) if m else "", "service_type": kind,
                "error_code": ecode, "error_desc": edesc, "part_code": pcode, "part_desc": pdesc,
                "qty": qty, "unit": unit, "consignment": cons, "quotation": quotation or "-",
                "details": details, "status": status,
                "contract_no": norm_text(m.get("contract_no")) if m else "",
                "start_contract": fmt_date(parse_date(m.get("pm_start"))) if m else "",
                "end_contract": fmt_date(parse_date(m.get("pm_end"))) if m else "",
                "freq": (int(round(12 / to_float(m.get("pm_count"))))
                         if m and (to_float(m.get("pm_count")) or 0) > 0 else ""),
                "customer": (norm_text(m.get("name_th")) if m and m.get("name_th") else customer),
                "remark": uniq_join(blk), "remark_red": uniq_join(red),
                "province": norm_text(m.get("province")) if m else "",
                "start_time": start_t, "end_time": end_t,
                "source": source, "page": str(page), "seq": seq,
                "rid": hashlib.sha1(f"{page_key}:{idx}:{si}:{i}".encode()).hexdigest()[:12],
                "page_key": page_key,
            })
    return rows


def build_sn_rows_from_ai(raw: Any, source: str, page: int, page_key: str,
                          ref: Optional[Reference], staff: str) -> list[dict]:
    items = raw.get("reports") if isinstance(raw, dict) else None
    if items is None:
        items = [raw] if isinstance(raw, dict) and raw else []
    rows: list[dict] = []
    for idx, it in enumerate(items):
        if isinstance(it, dict):
            rows += _rows_for_report(it, idx, source, page, page_key, ref, staff, len(rows))
    return rows


def process_upload(files: list[tuple[str, bytes]], client: VisionClient, ref: Optional[Reference],
                   staff: str, reprocess: bool, on_progress=None) -> dict:
    stats = {"pages": 0, "rows": 0, "skipped": 0, "errors": []}
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
                    rows = build_sn_rows_from_ai(raw, name, page_no, page_key, ref, staff)
                    delete_page_data(page_key)
                    store_sn_rows(rows)
                    mark_file(page_key, name, page_no, "ok", len(rows))
                    stats["pages"] += 1
                    stats["rows"] += len(rows)
                except Exception as e:
                    mark_file(page_key, name, page_no, f"error: {e}", 0)
                    stats["errors"].append(f"{name} หน้า {page_no}: {e}")
                if client.pause:
                    time.sleep(client.pause)
        except Exception as e:
            stats["errors"].append(f"{name}: {e}")
    return stats


# ----------------------------------------------------------------------------
# Review helpers
# ----------------------------------------------------------------------------
def _included(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        return []
    d = df.copy()
    d["include"] = d["include"].apply(lambda v: True if is_blank(v) else bool(v))
    return d[d["include"]].to_dict("records")


def refill_from_data2(df: pd.DataFrame, ref: Reference) -> tuple[pd.DataFrame, int]:
    df, n = df.copy(), 0
    for i, r in df.iterrows():
        m = ref.match_serial(r.get("serial"))
        if not m:
            continue
        df.at[i, "hospital"] = m.get("name_en") or r.get("hospital")
        df.at[i, "model"] = m.get("model") or r.get("model")
        df.at[i, "condition"] = m.get("status")
        df.at[i, "province"] = m.get("province")
        df.at[i, "contract_no"] = m.get("contract_no")
        df.at[i, "start_contract"] = fmt_date(parse_date(m.get("pm_start")))
        df.at[i, "end_contract"] = fmt_date(parse_date(m.get("pm_end")))
        cnt = to_float(m.get("pm_count"))
        df.at[i, "freq"] = int(round(12 / cnt)) if cnt and cnt > 0 else ""
        df.at[i, "customer"] = m.get("name_th") or r.get("customer")
        n += 1
    return df, n


def validate_rows(df: pd.DataFrame, ref: Reference) -> pd.DataFrame:
    issues: list[dict] = []
    for r in _included(df):
        base = {"date": norm_text(r.get("date")), "serial": norm_text(r.get("serial")),
                "hospital": norm_text(r.get("hospital"))}

        def add(check, value, status, detail=""):
            issues.append({**base, "check": check, "value": value, "status": status, "detail": detail})

        if parse_date(r.get("date")) is None:
            add("Date", norm_text(r.get("date")), "ERROR", "วันที่ว่างหรือรูปแบบไม่ถูกต้อง")
        sn = norm_text(r.get("serial"))
        if not sn:
            add("Serial", "", "MISSING", "ไม่มี Serial")
        elif ref.match_serial(sn):
            add("Serial", sn, "OK")
        else:
            add("Serial", sn, "UNKNOWN", "ไม่พบใน Data2")
        ec = norm_text(r.get("error_code"))
        if ec and norm_key(ec) not in ("pm", "n/a"):
            s_, canon, _ = ref.check_error(ec)
            add("Error code", ec, s_, f"ใกล้เคียง {canon}" if s_ == "SUGGEST" else "")
        pc = norm_text(r.get("part_code"))
        if pc and norm_key(pc) not in ("n/a", "-"):
            s_, canon, _ = ref.check_part(pc)
            add("Spare part", pc, s_, f"ใกล้เคียง {canon}" if s_ == "SUGGEST" else "")
        elif not pc:
            add("Spare part", "", "MISSING", "ช่องรหัสอะไหล่ว่าง")
        if norm_text(r.get("remark_red")):
            add("Red remark", norm_text(r.get("remark_red"))[:120], "TODO", "ต้องตรวจ/กรอกเอง")
    return pd.DataFrame(issues, columns=["date", "serial", "hospital", "check", "value", "status", "detail"])


# ----------------------------------------------------------------------------
# Trips (TMS / Expense are derived from the SN rows)
# ----------------------------------------------------------------------------
def distance_km(provs: list[str]) -> tuple[Optional[int], list[str]]:
    total, unknown = 0, []
    for p in provs:
        k = norm_key(p)
        if k in DIST_KM:
            total += DIST_KM[k]
        else:
            unknown.append(p)
    return (total or None), unknown


def build_trips(df: pd.DataFrame, lang: str = "ENG") -> dict[date, list[dict]]:
    """date -> list of trips (one per hospital per day, in SN order)."""
    days: dict[date, list[dict]] = defaultdict(list)
    index: dict[tuple, dict] = {}
    for r in _included(df):
        d = parse_date(r.get("date"))
        if d is None:
            continue
        hosp = norm_text(r.get("hospital"))
        name = hosp if lang == "ENG" else (norm_text(r.get("customer")) or hosp)
        key = (d, norm_key(name) or norm_key(r.get("serial")))
        t = index.get(key)
        if t is None:
            t = {"name": name, "prov": "", "kinds": {k: set() for k in KIND_OPTIONS}, "cond": [],
                 "start": None, "finish": None}
            index[key] = t
            days[d].append(t)
        kind = norm_text(r.get("service_type"))
        kind = kind if kind in KIND_OPTIONS else "Service"
        t["kinds"][kind].add(norm_serial(r.get("serial")) or f"row:{norm_text(r.get('rid'))}")
        p = norm_text(r.get("province"))
        if p and not t["prov"]:
            t["prov"] = p
        c = norm_text(r.get("condition"))
        if c and c not in t["cond"]:
            t["cond"].append(c)
        s, e = parse_time(r.get("start_time")), parse_time(r.get("end_time"))
        if s and (t["start"] is None or s < t["start"]):
            t["start"] = s
        if e and (t["finish"] is None or e > t["finish"]):
            t["finish"] = e
    return dict(days)


def build_expense_df(df: pd.DataFrame, saved: pd.DataFrame, lang: str) -> pd.DataFrame:
    cols = ["date", "hospitals"] + EXP_NUM
    rows: dict[str, dict] = {}
    for d, ts in build_trips(df, lang).items():
        rows[d.isoformat()] = {"date": d.isoformat(), "hospitals": " / ".join(t["name"] for t in ts),
                               **{c: None for c in EXP_NUM}}
    if saved is not None and not saved.empty:
        for s in saved.to_dict("records"):
            d = parse_date(s.get("date"))
            if d is None:
                continue
            row = rows.setdefault(d.isoformat(), {"date": d.isoformat(), "hospitals": "", **{c: None for c in EXP_NUM}})
            for c in EXP_NUM:
                if not is_blank(s.get(c)):
                    row[c] = to_float(s.get(c))
    out = pd.DataFrame(list(rows.values()), columns=cols)
    return out.sort_values("date").reset_index(drop=True) if not out.empty else out


# ----------------------------------------------------------------------------
# Excel writing. openpyxl only READS/locates; the file is patched at XML level
# so external links, shapes, styles and untouched sheets stay as they are.
# ----------------------------------------------------------------------------
class Recorder:
    def __init__(self) -> None:
        self.reset()

    def reset(self) -> None:
        self.cells: dict = defaultdict(list)   # sheet title -> [(final_row, col, value)]
        self.ins: dict = defaultdict(dict)     # sheet title -> {after_row: n_rows}


REC = Recorder()
SHEET_MODES = {norm_key(SHEET_SN): "all"}      # formula-copy mode for inserted rows (default "noA")


class RowMap:
    """Maps final row numbers (after row insertions) back to original rows."""

    def __init__(self, ins: dict) -> None:
        self.ins = sorted((a, k) for a, k in ins.items() if k > 0)

    def final(self, r: int) -> int:
        return r + sum(k for a, k in self.ins if a < r)

    def extra(self, after: int, j: int) -> int:
        return self.final(after) + 1 + j

    def origin(self, f: int) -> tuple[int, bool]:
        shift = 0
        for a, k in self.ins:
            fa = a + shift
            if f <= fa:
                break
            if f <= fa + k:
                return a, True
            shift += k
        return f - shift, False


def _copyable(f: str, col: int, mode: str) -> bool:
    if mode == "all":
        return True
    if col <= 2:
        return False
    return not re.search(r"(?<![A-Za-z0-9_.!])\$?A\$?\d+(?![\w(])", f)


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


def write_cell(ws, rmap: RowMap, f_row: int, col: int, value: Any, overwrite: bool = False,
               placeholders: bool = False) -> str:
    """Plan a write at FINAL row f_row. Never touches formulas."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return "empty"
    orig, inserted = rmap.origin(f_row)
    cell = ws.cell(orig, col)
    if isinstance(cell, MergedCell):
        return "merged-cell"
    cur = cell.value
    mode = SHEET_MODES.get(norm_key(ws.title), "noA")
    if isinstance(cur, str) and cur.startswith("="):
        if not inserted or _copyable(cur, col, mode):
            return "kept-formula"
        cur = None
    elif inserted:
        cur = None
    if not is_blank(cur) and not overwrite:
        if not (placeholders and norm_key(cur) in PLACEHOLDERS):
            return "occupied"
    REC.cells[ws.title].append((f_row, col, value))
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


def make_remark(black: Any, red: Any) -> Optional[Rich]:
    segs: list[tuple[str, bool]] = []
    if norm_text(black):
        segs.append((norm_text(black), False))
    if norm_text(red):
        if segs:
            segs.append((" | ", False))
        segs.append((norm_text(red), True))
    return Rich(segs) if segs else None


# ---- SN tracking form -------------------------------------------------------
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
CONT_SKIP = {"date", "hospital", "model", "serial", "condition", "service_type", "contract_no",
             "start_contract", "end_contract", "freq", "customer"}


def sn_rows_for_excel(df: pd.DataFrame, staff: str) -> list[dict]:
    out = []
    for r in _included(df):
        cont = _truthy(r.get("cont"))
        rec: dict[str, Any] = {}
        for k in SN_SPEC:
            v = r.get(k)
            if cont and k in CONT_SKIP:
                rec[k] = None
            elif k in ("date", "start_contract", "end_contract"):
                rec[k] = parse_date(v)
            elif k == "qty":
                t = norm_text(v)
                f = to_float(t) if re.fullmatch(r"-?\d+(\.\d+)?", t) else None
                rec[k] = (int(f) if float(f).is_integer() else f) if f is not None else t
            elif k == "freq":
                f = to_float(v)
                rec[k] = int(f) if f is not None else None
            elif k == "remark":
                rec[k] = make_remark(r.get("remark"), r.get("remark_red"))
            elif k == "person":
                rec[k] = norm_text(v) or staff
            else:
                rec[k] = norm_text(v)
        out.append(rec)
    return out


def fill_sn(ws, rows: list[dict], opts: dict, log: list) -> int:
    hdr = find_row(ws, 1, lambda t: t.startswith("person in charge"), end=40)
    if hdr is None:
        raise ValueError("ไม่พบหัวตาราง 'Person in charge (ENG)' ในชีต SN tracking form")
    cols = locate_cols(ws, (hdr, hdr + 1), SN_SPEC, log, SHEET_SN)
    first = hdr + 2 if "(only name)" in norm_key(ws.cell(hdr + 1, 1).value) else hdr + 1
    limit = min(ws.max_row, first + 5000)

    # summary table (Warranty / Contract ...) located under the data
    stop = None
    for r in range(first, limit):
        for c in range(1, 11):
            if norm_key(ws.cell(r, c).value) == "warranty" and norm_key(ws.cell(r + 1, c).value) == "contract":
                stop = r
                break
        if stop:
            break

    def const(r: int, c: int) -> bool:
        v = ws.cell(r, c).value
        return not is_blank(v) and not (isinstance(v, str) and v.startswith("="))

    if stop is None:      # fallback: first non-empty constant row below the data
        used = [r for r in range(first, limit + 1)
                if any(const(r, cols[k]) for k in ("date", "serial", "hospital", "model"))]
        last0 = used[-1] if used else first - 1
        for r in range(last0 + 1, limit + 1):
            if any(const(r, c) for c in range(1, 24)):
                stop = r
                break
    end = (stop - 1) if stop else limit
    last_used = first - 1
    for r in range(first, end + 1):
        if any(const(r, cols[k]) for k in ("date", "serial", "hospital", "model")):
            last_used = r
    first_free = last_used + 1
    n_rows = len(rows)
    if stop:
        free = max(stop - first_free, 0)
        if n_rows > free:
            k = n_rows - free
            REC.ins[ws.title][stop - 1] = k          # copy the row above, extend ranges that end there
            log.append({"sheet": SHEET_SN, "item": "insert rows",
                        "result": f"แทรกแถวเพิ่ม {k} แถว (คัดลอกสูตร/รูปแบบจากแถว {stop - 1}) ก่อนตารางสรุปที่แถว {stop}"})
    rmap = RowMap(REC.ins.get(ws.title, {}))
    written = 0
    warned = False
    for i, rec in enumerate(rows):
        r = first_free + i
        for key, val in rec.items():
            col = cols.get(key)
            if col is None:
                continue
            if key == "date" and val is not None and not warned:
                o, _ = rmap.origin(r)
                if ws.cell(o, col).number_format == "General":
                    log.append({"sheet": SHEET_SN, "item": "Date format",
                                "result": "warning: คอลัมน์ Date เป็น General ให้ตั้งเป็น dd-mmm-yy ใน Excel"})
                    warned = True
            res = write_cell(ws, rmap, r, col, val, True)
            if res == "written":
                written += 1
            elif res not in ("empty", "kept-formula"):
                log.append({"sheet": SHEET_SN, "item": f"row {r} / {key}", "result": res})
    log.append({"sheet": SHEET_SN, "item": "layout",
                "result": f"เขียนแถว {first_free}-{first_free + max(n_rows, 1) - 1}; ตารางสรุปเดิมเริ่มแถว {stop or '-'}"})
    return written


# ---- TMS Working Record -----------------------------------------------------
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
    top = [norm_key(v) for v in row_values(ws, hdr)]
    sub = [norm_key(v) for v in row_values(ws, hdr + 1)]

    def nth(lst: list, label: str, n: int, fb: int) -> int:
        idx = [i + 1 for i, k in enumerate(lst) if k == label]
        if len(idx) > n:
            return idx[n]
        log.append({"sheet": SHEET_TMS, "item": label, "result": f"warning: ไม่พบหัวคอลัมน์ '{label}' ใช้คอลัมน์ {fb}"})
        return fb

    pm, svc, inst = nth(top, "pm", 0, 16), nth(top, "service", 0, 18), nth(top, "installation", 0, 20)
    route, wt = nth(top, "route (province)", 0, 7), nth(top, "working time", 0, 22)
    cols = {
        "plan_prov": nth(top, "province", 0, 3), "plan_name": nth(top, "hospital name", 0, 4),
        "plan_q": nth(top, "qty", 0, 5), "dep": route, "arr": route + 1,
        "mileage": nth(top, "mileage", 0, 9),
        "tr_from": nth(sub, "from", 0, 10), "tr_to": nth(sub, "to", 0, 11), "tr_hrs": nth(sub, "(hrs)", 0, 12),
        "province": nth(top, "province", 1, 15),
        "svc_h": svc, "svc_q": svc + 1, "inst_h": inst, "inst_q": inst + 1,
        "start": nth(sub, "start", 0, wt), "finish": nth(sub, "finish", 0, wt + 1),
        "remark": nth(top, "remark", 0, 25),
    }
    base = parse_date(_cached(wsv, first, 1)) or _sheet_month(ws, wsv) or base_hint
    return {"first": first, "base": base, "cols": cols}


def tms_items(t: dict, opts: dict) -> tuple[list, Optional[str]]:
    pm, sv, ins = (len(t["kinds"][k]) for k in KIND_OPTIONS)
    prov = t["prov"]
    items: list = []
    note = None
    if pm:   # PM Plan side only; PM Actual is left blank for the user
        items += [("plan_prov", prov or None), ("plan_name", t["name"]), ("plan_q", pm)]
    if opts.get("fill_route", True):
        km, unknown = distance_km([prov]) if prov else (None, [])
        items += [("dep", "BKK"), ("arr", prov or None), ("mileage", km)]
        if not prov:
            note = "กรุณากรอกจังหวัดและระยะทาง"
        elif unknown:
            note = "กรุณากรอกระยะทาง: " + prov
    if sv or ins:
        items.append(("province", prov or None))
    if sv:
        items += [("svc_h", t["name"]), ("svc_q", sv)]
    if ins:
        items += [("inst_h", t["name"]), ("inst_q", ins)]
    s, e = t["start"], t["finish"]
    if s:
        items.append(("start", s))
        if s.hour >= 1:
            items += [("tr_from", dtime(s.hour - 1, s.minute)), ("tr_to", s), ("tr_hrs", dtime(1, 0))]
    if e:
        items.append(("finish", e))
    if note:
        items.append(("remark", Rich([(note, True)])))
    return items, note


def fill_tms(ws, wsv, days: dict, opts: dict, log: list) -> int:
    hint = opts.get("month_override") or (min(days).replace(day=1) if days else None)
    lay = detect_tms(ws, wsv, hint, log)
    base = opts.get("month_override") or lay["base"]
    if base is None:
        log.append({"sheet": SHEET_TMS, "item": "month", "result": "error: หาเดือนของชีตไม่ได้ (ระบุที่ Month override)"})
        return 0
    c, ow = lay["cols"], opts.get("overwrite", False)
    plans = []
    for d, trips in sorted(days.items()):
        off = (d - base).days
        if not 0 <= off <= 30:
            log.append({"sheet": SHEET_TMS, "item": d.isoformat(), "result": f"skipped: วันที่อยู่นอกเดือน {base:%Y-%m}"})
            continue
        r = lay["first"] + off
        cached = parse_date(_cached(wsv, r, 1))
        if cached and cached != d:
            log.append({"sheet": SHEET_TMS, "item": d.isoformat(), "result": f"skipped: แถว {r} เป็นวันที่ {cached} ไม่ตรง"})
            continue
        plans.append((d, r, trips))
    for d, r, trips in plans:
        if len(trips) > 1:
            REC.ins[ws.title][r] = len(trips) - 1
    rmap = RowMap(REC.ins.get(ws.title, {}))
    written = 0
    for d, r, trips in plans:
        for j, t in enumerate(trips):
            f_row = rmap.final(r) if j == 0 else rmap.extra(r, j - 1)
            items, _note = tms_items(t, opts)
            for key, val in items:
                rs = write_cell(ws, rmap, f_row, c[key], val, ow)
                if rs == "written":
                    written += 1
                elif rs not in ("empty", "kept-formula"):
                    log.append({"sheet": SHEET_TMS, "item": f"{d.isoformat()} / {key} (row {f_row})", "result": rs})
    return written


# ---- Expense Claim ----------------------------------------------------------
EXP_SPEC = {
    "description": (["description"], 2), "allowance": (["allowance"], 7), "petrol": (["petrol"], 10),
    "expressway": (["express way"], 11), "parking": (["parking"], 12), "other": (["other"], 13),
    "service_type": (["service type"], 15),
}


def fill_expense(ws, wsv, days: dict, amounts: pd.DataFrame, base_hint: Optional[date], opts: dict, log: list) -> int:
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
    ow = opts.get("overwrite", False)
    amt: dict[date, dict] = {}
    if amounts is not None and not amounts.empty:
        for s in amounts.to_dict("records"):
            d = parse_date(s.get("date"))
            if d:
                amt[d] = {k: to_float(s.get(k)) for k in EXP_NUM}
    all_dates = sorted(set(days) | {d for d, v in amt.items() if any(v.values())})
    plans = []
    for d in all_dates:
        off = (d - base).days
        if not 0 <= off <= 30:
            log.append({"sheet": SHEET_EXP, "item": d.isoformat(), "result": "skipped: นอกเดือน"})
            continue
        r = first + off
        cached = parse_date(_cached(wsv, r, 1))
        if cached and cached != d:
            log.append({"sheet": SHEET_EXP, "item": d.isoformat(), "result": f"skipped: แถว {r} เป็นวันที่ {cached}"})
            continue
        plans.append((d, r))
    for d, r in plans:
        if len(days.get(d, [])) > 1:
            REC.ins[ws.title][r] = len(days[d]) - 1
    rmap = RowMap(REC.ins.get(ws.title, {}))
    written = 0

    def put(f_row: int, key: str, val: Any, ph: bool, dd: date) -> None:
        nonlocal written
        rs = write_cell(ws, rmap, f_row, cols[key], val, ow, ph)
        if rs == "written":
            written += 1
        elif rs not in ("empty", "kept-formula"):
            log.append({"sheet": SHEET_EXP, "item": f"{dd.isoformat()} / {key} (row {f_row})", "result": rs})

    for d, r in plans:
        trips = days.get(d, [])
        for j, t in enumerate(trips):
            f_row = rmap.final(r) if j == 0 else rmap.extra(r, j - 1)
            put(f_row, "description", t["name"], True, d)
            put(f_row, "service_type", Rich([(RED_FILL, True)]), True, d)
        a = amt.get(d)
        if a:
            f_row = rmap.final(r)
            for k in EXP_NUM:
                if a.get(k):
                    put(f_row, k, a[k], False, d)
    return written


# ---- reference shifting (used when rows are inserted) -----------------------
_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_XML_NS = "http://www.w3.org/XML/1998/namespace"
_REF_RE = re.compile(
    r"(?<![\w.])(?P<pre>(?:'[^']+'|[A-Za-z_][\w.]*)!)?"
    r"(?P<c1>\$?[A-Z]{1,3})(?P<r1>\$?)(?P<n1>\d+)"
    r"(?::(?P<c2>\$?[A-Z]{1,3})(?P<r2>\$?)(?P<n2>\d+))?"
    r"(?![\w(!])"
)


def _q(tag: str) -> str:
    return f"{{{_NS}}}{tag}"


def shift_formula(f: str, host_is_target: bool, names: set, at: int, n: int, ext: bool = True) -> str:
    """Shift row numbers >= at by n for references that point to the target sheet.
    ext=True also extends the END of a range that stops at row at-1 (like inserting inside a range)."""
    def sh(num: str, lim: int) -> str:
        v = int(num)
        return str(v + n) if v >= lim else num

    def repl(m: re.Match) -> str:
        pre = m.group("pre")
        if pre:
            applies = norm_key(pre[:-1].strip("'").replace("''", "'")) in names
        else:
            applies = host_is_target
        if not applies:
            return m.group(0)
        out = (pre or "") + m.group("c1") + m.group("r1") + sh(m.group("n1"), at)
        if m.group("c2"):
            out += ":" + m.group("c2") + m.group("r2") + sh(m.group("n2"), at - 1 if ext else at)
        return out

    parts = re.split(r'("(?:[^"]|"")*")', f)
    for i in range(0, len(parts), 2):
        parts[i] = _REF_RE.sub(repl, parts[i])
    return "".join(parts)


def _expand_shared(root, ws) -> None:
    for c in root.iter(_q("c")):
        f = c.find(_q("f"))
        if f is not None and f.get("t") == "shared":
            val = ws[c.get("r")].value if ws is not None else None
            if isinstance(val, str) and val.startswith("="):
                for a in ("t", "ref", "si"):
                    f.attrib.pop(a, None)
                f.text = val[1:]


def _shift_formulas(root, host_is_target: bool, names: set, at: int, n: int) -> None:
    for f in root.iter(_q("f")):
        if f.text:
            f.text = shift_formula(f.text, host_is_target, names, at, n)
        if host_is_target and f.get("ref"):
            f.set("ref", shift_formula(f.get("ref"), True, names, at, n))


def _shift_structs(root, at: int, n: int) -> None:
    for tag, attr, ext in (("mergeCell", "ref", False), ("conditionalFormatting", "sqref", True),
                           ("dataValidation", "sqref", True), ("autoFilter", "ref", True),
                           ("dimension", "ref", True), ("hyperlink", "ref", True),
                           ("ignoredError", "sqref", True), ("protectedRange", "sqref", True)):
        for el in root.iter(_q(tag)):
            v = el.get(attr)
            if v:
                el.set(attr, shift_formula(v, True, set(), at, n, ext))
    for tag in ("formula", "formula1", "formula2"):
        for el in root.iter(_q(tag)):
            if el.text:
                el.text = shift_formula(el.text, True, set(), at, n)


def _insert_rows(root, at: int, n: int, src_row: int, mode: str) -> None:
    sd = root.find(_q("sheetData"))
    src_el = None
    for row in sd.findall(_q("row")):
        rn = int(row.get("r"))
        if rn == src_row:
            src_el = row
        if rn >= at:
            row.set("r", str(rn + n))
            row.attrib.pop("spans", None)
            for c in row.findall(_q("c")):
                m = re.match(r"([A-Z]+)\d+$", c.get("r"))
                c.set("r", f"{m.group(1)}{rn + n}")
    anchor = next((row for row in sd.findall(_q("row")) if int(row.get("r")) >= at + n), None)
    for k in range(n):
        rn = at + k
        new = etree.Element(_q("row"))
        if src_el is not None:
            for a, v in src_el.attrib.items():
                if a not in ("r", "spans"):
                    new.set(a, v)
            for c in src_el.findall(_q("c")):
                col = re.match(r"([A-Z]+)\d+$", c.get("r")).group(1)
                nc = etree.SubElement(new, _q("c"))
                nc.set("r", f"{col}{rn}")
                if c.get("s") is not None:
                    nc.set("s", c.get("s"))
                f = c.find(_q("f"))
                if f is not None and f.text and f.get("t") in (None, "normal"):
                    if _copyable(f.text, column_index_from_string(col), mode):
                        try:
                            nf = Translator("=" + f.text, origin=f"{col}{src_row}").translate_formula(f"{col}{rn}")
                            etree.SubElement(nc, _q("f")).text = nf[1:]
                        except Exception:
                            pass
        new.set("r", str(rn))
        if anchor is not None:
            anchor.addprevious(new)
        else:
            sd.append(new)
    mc = root.find(_q("mergeCells"))        # replicate single-row merges (e.g. B:F description)
    if mc is not None:
        for m in list(mc.findall(_q("mergeCell"))):
            mm = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", m.get("ref") or "")
            if mm and int(mm.group(2)) == src_row == int(mm.group(4)):
                for k in range(n):
                    nm = etree.SubElement(mc, _q("mergeCell"))
                    nm.set("ref", f"{mm.group(1)}{at + k}:{mm.group(3)}{at + k}")
        mc.set("count", str(len(mc.findall(_q("mergeCell")))))


def _shift_drawing(data: bytes, at: int, n: int) -> bytes:
    root = etree.fromstring(data)
    for tag in ("from", "to"):
        for el in root.iter(f"{{{_XDR}}}{tag}"):
            row = el.find(f"{{{_XDR}}}row")
            if row is not None and row.text and int(row.text) >= at - 1:
                row.text = str(int(row.text) + n)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _shift_table(data: bytes, at: int, n: int) -> bytes:
    root = etree.fromstring(data)
    if root.get("ref"):
        root.set("ref", shift_formula(root.get("ref"), True, set(), at, n))
    for el in root.iter(_q("autoFilter")):
        if el.get("ref"):
            el.set("ref", shift_formula(el.get("ref"), True, set(), at, n))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


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


def _xml_num(x: float) -> str:
    return str(int(x)) if float(x).is_integer() else repr(float(x))


def _patch_root(root, cells: list) -> None:
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

    def get_cell(n: int, c: int):
        row = get_row(n)
        ref = f"{get_column_letter(c)}{n}"
        for ce in row.findall(_q("c")):
            if ce.get("r") == ref:
                return ce
        new = etree.Element(_q("c"))
        new.set("r", ref)
        for ce in row.findall(_q("c")):
            m = re.match(r"[A-Z]+", ce.get("r") or "")
            if m and column_index_from_string(m.group(0)) > c:
                ce.addprevious(new)
                return new
        row.append(new)
        return new

    for r, c, v in cells:
        ce = get_cell(r, c)
        if ce.find(_q("f")) is not None:
            continue                                  # never overwrite a formula
        for ch in list(ce):
            ce.remove(ch)
        ce.attrib.pop("t", None)
        if isinstance(v, Rich):
            ce.set("t", "inlineStr")
            is_ = etree.SubElement(ce, _q("is"))
            for text, red in v:
                run = etree.SubElement(is_, _q("r"))
                rpr = etree.SubElement(run, _q("rPr"))
                etree.SubElement(rpr, _q("rFont")).set("val", "Arial")
                if red:
                    etree.SubElement(rpr, _q("color")).set("rgb", "FFFF0000")
                etree.SubElement(rpr, _q("sz")).set("val", "10")
                t = etree.SubElement(run, _q("t"))
                t.text = text
                t.set(f"{{{_XML_NS}}}space", "preserve")
        elif isinstance(v, bool):
            ce.set("t", "b")
            etree.SubElement(ce, _q("v")).text = "1" if v else "0"
        elif isinstance(v, (datetime, date, dtime)):
            etree.SubElement(ce, _q("v")).text = _xml_num(to_excel(v))
        elif isinstance(v, (int, float)):
            etree.SubElement(ce, _q("v")).text = _xml_num(v)
        else:
            ce.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(ce, _q("is")), _q("t"))
            t.text = str(v)
            t.set(f"{{{_XML_NS}}}space", "preserve")


def _tostring(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_xlsx(master: bytes, cells_by_sheet: dict, ins_by_sheet: dict, force_recalc: bool, wb) -> bytes:
    zin = zipfile.ZipFile(io.BytesIO(master))
    try:
        names = set(zin.namelist())
        paths = _sheet_paths(zin)
        key_of_path = {v: k for k, v in paths.items()}
        out_parts: dict[str, bytes] = {}
        roots: dict[str, Any] = {}
        expanded: set[str] = set()
        wbroot = etree.fromstring(zin.read("xl/workbook.xml"))

        def root_of(path: str):
            if path not in roots:
                roots[path] = etree.fromstring(zin.read(path))
            return roots[path]

        def expand(path: str) -> None:
            if path in expanded:
                return
            expanded.add(path)
            key = key_of_path.get(path)
            ws_o = next((w for w in wb.worksheets if norm_key(w.title) == key), None)
            _expand_shared(root_of(path), ws_o)

        for sheet, ins in ins_by_sheet.items():
            skey = norm_key(sheet)
            spath = paths.get(skey)
            if not spath or spath not in names:
                raise ValueError(f"ไม่พบไฟล์ XML ของชีต {sheet}")
            mode = SHEET_MODES.get(skey, "noA")
            for after in sorted(ins, reverse=True):          # bottom-up keeps earlier row numbers valid
                k = ins[after]
                if k <= 0:
                    continue
                at = after + 1
                tgt = {skey}
                for key, path in paths.items():
                    if path not in names:
                        continue
                    if key != skey and sheet.encode("utf-8") not in zin.read(path):
                        continue
                    expand(path)
                    r = root_of(path)
                    _shift_formulas(r, key == skey, tgt, at, k)
                    if key == skey:
                        _shift_structs(r, at, k)
                        _insert_rows(r, at, k, after, mode)
                for dn in wbroot.iter(_q("definedName")):
                    if dn.text:
                        dn.text = shift_formula(dn.text, False, tgt, at, k)
                rel_path = posixpath.join(posixpath.dirname(spath), "_rels", posixpath.basename(spath) + ".rels")
                if rel_path in names:
                    for rel in etree.fromstring(zin.read(rel_path)):
                        typ = rel.get("Type") or ""
                        tp = rel.get("Target") or ""
                        tp = tp.lstrip("/") if tp.startswith("/") else posixpath.normpath(
                            posixpath.join(posixpath.dirname(spath), tp))
                        if tp not in names:
                            continue
                        if typ.endswith("/drawing"):
                            out_parts[tp] = _shift_drawing(out_parts.get(tp) or zin.read(tp), at, k)
                        elif typ.endswith("/table"):
                            out_parts[tp] = _shift_table(out_parts.get(tp) or zin.read(tp), at, k)

        for sheet, cells in cells_by_sheet.items():
            path = paths.get(norm_key(sheet))
            if not path or path not in names:
                raise ValueError(f"ไม่พบไฟล์ XML ของชีต {sheet}")
            _patch_root(root_of(path), cells)

        for path, r in roots.items():
            out_parts[path] = _tostring(r)

        if force_recalc:
            calc = wbroot.find(_q("calcPr"))
            if calc is None:
                calc = etree.Element(_q("calcPr"))
                anchor = None
                for tag in ("sheets", "functionGroups", "externalReferences", "definedNames"):
                    e = wbroot.find(_q(tag))
                    if e is not None:
                        anchor = e
                if anchor is not None:
                    anchor.addnext(calc)
                else:
                    wbroot.append(calc)
            calc.set("fullCalcOnLoad", "1")
        out_parts["xl/workbook.xml"] = _tostring(wbroot)

        drop: set[str] = set()
        if "xl/calcChain.xml" in names:                        # stale after edits -> Excel rebuilds it
            drop.add("xl/calcChain.xml")
            ct = etree.fromstring(zin.read("[Content_Types].xml"))
            for el in list(ct):
                if el.get("PartName") == "/xl/calcChain.xml":
                    ct.remove(el)
            out_parts["[Content_Types].xml"] = _tostring(ct)
            rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
            for el in list(rels):
                if (el.get("Type") or "").endswith("/calcChain"):
                    rels.remove(el)
            out_parts["xl/_rels/workbook.xml.rels"] = _tostring(rels)

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in drop:
                    continue
                data = out_parts.get(item.filename)
                zout.writestr(item, data if data is not None else zin.read(item.filename))
        return out.getvalue()
    finally:
        zin.close()


def fill_workbook(master: bytes, sn_df: pd.DataFrame, amounts: pd.DataFrame, opts: dict) -> tuple[bytes, list[dict]]:
    log: list[dict] = []
    REC.reset()
    wb = load_workbook(io.BytesIO(master), keep_links=False)       # read only, never saved
    try:
        wbv = load_workbook(io.BytesIO(master), data_only=True, keep_links=False)
    except Exception:
        wbv = None
    sheets = {n: get_sheet(wb, n) for n in (SHEET_TMS, SHEET_SN, SHEET_EXP)}
    missing = [n for n, w in sheets.items() if w is None]
    if missing:
        raise ValueError("ไม่พบชีต: " + ", ".join(missing))
    vals = {n: (get_sheet(wbv, n) if wbv else None) for n in sheets}
    lang = opts.get("lang", "ENG")

    # 1) SN tracking form first; 2) TMS and 3) Expense are derived from the same SN rows
    sn_rows = sn_rows_for_excel(sn_df, opts.get("staff", ""))
    n_sn = fill_sn(sheets[SHEET_SN], sn_rows, opts, log)
    days = build_trips(sn_df, lang)
    n_tms = fill_tms(sheets[SHEET_TMS], vals[SHEET_TMS], days, opts, log)
    hint = opts.get("month_override") or (min(days).replace(day=1) if days else None)
    n_exp = fill_expense(sheets[SHEET_EXP], vals[SHEET_EXP], days, amounts, hint, opts, log)

    out = patch_xlsx(master, dict(REC.cells), {k: dict(v) for k, v in REC.ins.items()},
                     opts.get("force_recalc", True), wb)
    ins_txt = ", ".join(f"{k}: +{sum(v.values())}" for k, v in REC.ins.items() if v) or "-"
    log.insert(0, {"sheet": "SUMMARY", "item": "cells written",
                   "result": f"SN={n_sn} (rows={len(sn_rows)}), TMS={n_tms}, Expense={n_exp}; แทรกแถว {ins_txt}"})
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
                 10: "Travelling Time", 15: "Province", 16: "PM", 17: "Qty", 18: "Service", 19: "Qty",
                 20: "Installation", 21: "Qty", 22: "Working Time", 25: "Remark"}.items():
        ws.cell(5, c, t)
    for c, t in {5: "(units)", 7: "Departure", 8: "Arrvial", 9: "(kms)", 10: "From", 11: "To", 12: "(Hrs)",
                 22: "Start", 23: "Finish", 24: "(Hrs)"}.items():
        ws.cell(6, c, t)
    ws["A7"] = "=B2"
    for r in range(8, 38):
        ws.cell(r, 1, f"=A{r - 1}+1")
    for r in range(7, 38):
        ws.cell(r, 12, f"=K{r}-J{r}")
        ws.cell(r, 24, f"=W{r}-V{r}")
    ws["I38"] = "=SUM(I7:I37)"
    ws["Z1"] = "=COUNTA('SN tracking form'!$E$7:$E$10)"

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
    for r in range(7, 11):   # 4 free rows; C, D, F are formula columns
        sn.cell(r, 3, f"=IFERROR(INDEX('Data2 (2)'!$L:$L,MATCH($E{r},'Data2 (2)'!$G:$G,0)),\"\")")
        sn.cell(r, 4, f"=IFERROR(INDEX('Data2 (2)'!$D:$D,MATCH($E{r},'Data2 (2)'!$G:$G,0)),\"\")")
        sn.cell(r, 6, f"=IFERROR(INDEX('Data2 (2)'!$M:$M,MATCH($E{r},'Data2 (2)'!$G:$G,0)),\"\")")
        sn.cell(r, 2).number_format = "dd-mmm-yy"
        sn.cell(r, 2).border = Border(bottom=Side(style="thin"))
    sn["D11"], sn["D12"] = "Warranty", "Contract"
    sn["E11"], sn["E12"] = '=COUNTIF($F$7:$F$10,D11)', '=COUNTIF($F$7:$F$10,D12)'

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

    sp = wb.create_sheet(SHEET_SPARE)
    sp.append(["Item", "NIPRO Order Code", "Product description", "Remark"])
    sp.append([1, "SP-0-504-232", "Clamp box", ""])
    sp.append([2, "w", "รอเสนอราคาเปลี่ยนอะไหล่", ""])
    sp.append([3, "s", "ซ่อมแบบไม่เปลี่ยนอะไหล่", ""])
    sp.append([4, "N/A", "-", ""])

    def d2row(prov, model, th, sn_, en, status, ctr, cnt):
        return ["", "HD", prov, model, "Mr. Suvanai", th, sn_, None, None, "", "", en, status, None, 2022,
                datetime(2022, 1, 1), datetime(2025, 10, 1), datetime(2026, 9, 30), None, None, None, "Yes", cnt,
                None, None, None, ctr]

    d2 = wb.create_sheet(SHEET_DATA2)
    d2.append(["Z", "Type", "Province", "Model name", "Engineer name", "TH - Customer name (Commercial)",
               "Serial number", None, None, "CRM", "Machine installation location",
               "ENG - Customer name (Commercial)", "Machine status", None, "Year of Installation", "Install date",
               'Start date of "PM"', 'End date of "PM"', "End PM", "Remaining days", "Quotation and Contract issued date",
               "PM contract\nYes / No", "Number of PM maintenance", None, None, None, "Contract number"])
    d2.append(d2row("Ayutthaya", "Surdial 55 Plus II", "โรงพยาบาลราชธานี (อยุธยา)", "21CX22513",
                    "Rajthanee Hospital (Ayutthaya)", "Contract", "CT-001", 4))
    d2.append(d2row("Bangkok", "Surdial 55", "คลินิกไตเทียมกรุงไทย", "17CX6407",
                    "Krungthai Renal Care Dialysis Clinic", "Non-Active", "", None))
    d2.append(d2row("Nonthaburi", "PY-YD-340", "สถาบันทรวงอก", "2403-055", "สถาบันทรวงอก", "Warranty", "", None))
    for s_ in ("19CX14513", "19CX14514", "19CX14515"):
        d2.append(d2row("Bangkok", "Surdial 55 PLUS type II C", "เนอร์ซ", s_, "Nipro Excellent Renal Care",
                        "Fixed asset", "", None))
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def run_selftest() -> None:
    global DATA_DIR, DB_PATH
    import tempfile

    tmp = Path(tempfile.mkdtemp())
    DATA_DIR, DB_PATH = tmp, tmp / "test.db"
    init_db()

    # helpers
    assert parse_date("28/9/26") == date(2026, 9, 28)
    assert parse_date("9 กันยายน 2569") == date(2026, 9, 9)
    assert parse_date(46266) == date(2026, 9, 1)
    assert parse_time("18.00") == dtime(18, 0)
    assert shift_formula("COUNTIF($F$7:$F$10,D11)", True, set(), 11, 3) == "COUNTIF($F$7:$F$13,D14)"
    assert shift_formula("'SN tracking form'!E90", False, {"sn tracking form"}, 86, 3) == "'SN tracking form'!E93"
    assert shift_formula("'Data2 (2)'!$G$2:$G$1995", False, {"sn tracking form"}, 86, 3) == "'Data2 (2)'!$G$2:$G$1995"
    assert shift_formula("LOG10(A200)", True, set(), 86, 3) == "LOG10(A203)"
    assert shift_formula("SUM(I7:I37)", True, set(), 10, 1) == "SUM(I7:I38)"
    assert shift_formula("A9+1", True, set(), 10, 1) == "A9+1"
    rm = RowMap({9: 1})
    assert rm.final(13) == 14 and rm.extra(9, 0) == 10 and rm.origin(10) == (9, True) and rm.origin(14) == (13, False)

    master = _demo_master()
    ref = Reference.from_bytes(master)
    assert not ref.missing_sheets, ref.missing_sheets
    assert ref.check_part("w")[0] == "OK" and ref.check_error("E114")[0] == "OK"
    assert ref.resolve_serial("19CX145147", "Nipro Excellent Renal Care")[:2] == ("19CX14514", "FIXED")
    assert ref.resolve_serial("21cx22513")[1] == "OK"

    staff = "Mr. Suvanai P."
    for key, payload in (("a", PAYLOAD_SERVICE), ("b", PAYLOAD_QUOTE), ("c", PAYLOAD_PM), ("d", PAYLOAD_INST)):
        store_sn_rows(build_sn_rows_from_ai(payload, f"{key}.pdf", 1, f"{key}:1", ref, staff))
    sn_df = load_sn()
    assert len(sn_df) == 7, len(sn_df)
    k1, k2, inst, rj, n1, n2, n3 = sn_df.to_dict("records")
    assert k1["serial"] == "17CX6407" and k1["part_code"] == "w" and k1["details"] == "เสนอราคา"
    assert "เสนอราคา" in k1["remark"] and k1["condition"] == "Non-Active" and not k1["cont"]
    assert RED_STATUS in k1["remark_red"]
    assert k2["cont"] and k2["error_code"] == "E114" and k2["part_code"] == "w"
    assert inst["part_code"] == "-" and inst["condition"] == "Warranty" and inst["service_type"] == "Installation"
    assert rj["consignment"] == "CF-DMS-261010" and rj["status"] == "Issue DO" and rj["unit"] == "PC"
    assert RED_CHECK_STATUS in rj["remark_red"] and "21CX22517" in rj["remark"]
    assert rj["error_desc"].startswith("SHORT CHEMICAL")
    assert n2["serial"] == "19CX14514" and "ปรับตาม Data2" in n2["remark"]
    assert n1["error_code"] == "PM" and n1["part_code"] == "N/A" and n1["details"] == "PM"

    opts = {"staff": staff, "overwrite": False, "lang": "ENG", "fill_route": True,
            "month_override": None, "force_recalc": True}
    out, log = fill_workbook(master, sn_df, load_expenses(), opts)

    zo = zipfile.ZipFile(io.BytesIO(out))
    sn_xml = zo.read(_sheet_paths(zo)[norm_key(SHEET_SN)])
    assert b"FFFF0000" in sn_xml, "red remark missing"
    wb = load_workbook(io.BytesIO(out))
    sn, tms, ex = wb[SHEET_SN], wb[SHEET_TMS], wb[SHEET_EXP]

    # SN: 7 rows into 4 free rows -> 3 inserted, summary moved 11 -> 14 and ranges extended
    assert sn["D14"].value == "Warranty" and sn["D15"].value == "Contract", (sn["D14"].value, sn["D15"].value)
    assert sn["E14"].value == "=COUNTIF($F$7:$F$13,D14)", sn["E14"].value
    assert tms["Z1"].value == "=COUNTA('SN tracking form'!$E$7:$E$13)", tms["Z1"].value
    assert str(sn["C7"].value).startswith("=IFERROR(INDEX") and "$E7" in sn["C7"].value
    assert "$E11" in sn["C11"].value and "$E13" in sn["C13"].value
    assert sn["B13"].number_format == "dd-mmm-yy" and sn["B13"].border.bottom.style == "thin"
    # SN values
    assert sn["A7"].value == staff and sn["B7"].value == datetime(2026, 9, 3)
    assert sn["E7"].value == "17CX6407" and sn["G7"].value == "Service" and sn["H7"].value == "E97"
    assert sn["J7"].value == "w" and sn["Q7"].value is None
    assert "กรุณาลงข้อมูล Status" in str(sn["W7"].value)
    assert sn["A8"].value == staff and sn["B8"].value is None and sn["E8"].value is None and sn["G8"].value is None
    assert sn["H8"].value == "E114" and sn["J8"].value == "w"
    assert sn["E9"].value == "2403-055" and sn["G9"].value == "Installation" and sn["J9"].value == "-"
    assert sn["E10"].value == "21CX22513" and sn["Q10"].value == "Issue DO" and sn["N10"].value == "CF-DMS-261010"
    assert sn["L10"].value == 1 and sn["M10"].value == "PC" and sn["R10"].value == "CT-001" and sn["U10"].value == 3
    assert sn["E12"].value == "19CX14514" and sn["H11"].value == "PM" and sn["J11"].value == "N/A"
    # TMS built from SN rows; 3 Sep has two hospitals -> one inserted row (no date)
    assert tms["R9"].value == "Krungthai Renal Care Dialysis Clinic" and tms["S9"].value == 1
    assert tms["G9"].value == "BKK" and tms["H9"].value == "Bangkok" and tms["I9"].value == 30
    assert abs(float(tms["J9"].value) - 8 / 24) < 1e-9 and abs(float(tms["K9"].value) - 9 / 24) < 1e-9
    assert abs(float(tms["V9"].value) - 9 / 24) < 1e-9 and abs(float(tms["W9"].value) - 10 / 24) < 1e-9
    assert str(tms["L9"].value).startswith("="), "formula must be preserved"
    assert tms["A10"].value is None and tms["B10"].value is None
    assert tms["T10"].value == "สถาบันทรวงอก" and tms["U10"].value == 1 and tms["H10"].value == "Nonthaburi"
    assert tms["I10"].value == 30 and tms["O10"].value == "Nonthaburi"
    assert tms["L10"].value == "=K10-J10" and tms["X10"].value == "=W10-V10"
    assert tms["A11"].value == "=A9+1", tms["A11"].value
    assert tms["R14"].value == "Rajthanee Hospital (Ayutthaya)" and tms["S14"].value == 1
    assert tms["H14"].value == "Ayutthaya" and tms["I14"].value == 60 and abs(float(tms["V14"].value) - 0.75) < 1e-9
    assert tms["C37"].value == "Bangkok" and tms["D37"].value == "Nipro Excellent Renal Care" and tms["E37"].value == 3
    assert tms["P37"].value is None and tms["Q37"].value is None and tms["O37"].value is None
    assert tms["I37"].value == 30
    assert tms["I39"].value == "=SUM(I7:I38)", tms["I39"].value
    # Expense built from SN rows (Service Type is always the red "fill in" text)
    assert ex["B10"].value == "Krungthai Renal Care Dialysis Clinic" and ex["O10"].value == RED_FILL
    assert ex["A11"].value is None and ex["B11"].value == "สถาบันทรวงอก" and ex["O11"].value == RED_FILL
    assert ex["A12"].value == "=A10+1", ex["A12"].value
    assert ex["B15"].value == "Rajthanee Hospital (Ayutthaya)" and ex["O15"].value == RED_FILL
    assert ex["P11"].value == "=SUM(J11:M11)" and ex["P15"].value == "=SUM(J15:M15)"
    assert ex["B38"].value == "Nipro Excellent Renal Care" and ex["O38"].value == RED_FILL
    assert ex["B9"].value == "Office"
    # untouched sheet identical
    zi = zipfile.ZipFile(io.BytesIO(master))
    p = _sheet_paths(zi)[norm_key(SHEET_ERR)]
    assert zi.read(p) == zo.read(p)
    print("SELFTEST PASSED")
    for item in log:
        print(item)


# ----------------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------------
def _stretch(fn, *args, **kwargs):
    try:
        return fn(*args, width="stretch", **kwargs)
    except Exception:
        return fn(*args, use_container_width=True, **kwargs)


def _only_options(df: pd.DataFrame, col: str, options: list[str]) -> pd.DataFrame:
    df = df.copy()
    df[col] = df[col].apply(lambda v: norm_text(v) if norm_text(v) in options else None)
    return df


def _as_text_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.astype(str).replace({"None": "", "nan": "", "NaT": ""})


def _secret(*names: str) -> str:
    for name in names:
        try:
            v = str(st.secrets.get(name, ""))
        except Exception:
            v = ""
        v = v or os.environ.get(name, "")
        if v:
            return v
    return ""


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

    sb.header("2) AI อ่านใบงาน")
    provider = sb.selectbox("Provider", [PROVIDER_GEMINI, PROVIDER_CLAUDE, PROVIDER_OPENAI, PROVIDER_DEMO])
    names = (("GEMINI_API_KEY", "GOOGLE_API_KEY") if provider == PROVIDER_GEMINI
             else ("ANTHROPIC_API_KEY",) if provider == PROVIDER_CLAUDE else ("OPENAI_API_KEY",))
    key = sb.text_input("API key", value=_secret(*names), type="password", disabled=provider == PROVIDER_DEMO)
    model = sb.text_input("Model", value=DEFAULT_MODELS[provider], key=f"model_{provider}")

    sb.header("3) ตัวเลือก Export")
    staff = sb.text_input("ชื่อผู้รับผิดชอบ (SN tracking)", value="Mr. Suvanai P.")
    lang = sb.radio("ภาษาชื่อ รพ. ใน TMS/Expense", ["ENG", "TH"], horizontal=True)
    overwrite = sb.checkbox("เขียนทับค่าที่มีอยู่แล้วใน TMS/Expense (ไม่แตะสูตร)", value=False)
    fill_route = sb.checkbox("เติมเส้นทาง BKK -> จังหวัด และระยะทางใน TMS", value=True)
    force_recalc = sb.checkbox("ให้ Excel คำนวณสูตรใหม่ทั้งไฟล์ตอนเปิด", value=True)
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
    st.subheader("อัปโหลดใบงาน (PDF หลายหน้า / รูปภาพ)")
    files = st.file_uploader("เลือกไฟล์ได้หลายไฟล์", type=["pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp"],
                             accept_multiple_files=True)
    reprocess = st.checkbox("ประมวลผลซ้ำหน้าที่เคยทำแล้ว (แทนที่แถวเดิมของหน้านั้น)")
    if cfg["ref"] is None:
        st.info("อัปโหลด Master Excel ก่อน เพื่อให้ระบบปรับ Serial และเติมข้อมูลจาก Data2 ได้")
    if st.button("เริ่มประมวลผล", type="primary", disabled=not files):
        client = VisionClient(cfg["provider"], cfg["key"], cfg["model"])
        bar, msg = st.progress(0.0), st.empty()

        def cb(p: float, text: str) -> None:
            bar.progress(min(max(p, 0.0), 1.0))
            msg.write(f"กำลังประมวลผล: {text}")

        stats = process_upload([(f.name, f.getvalue()) for f in files], client, cfg["ref"], cfg["staff"], reprocess, cb)
        bar.progress(1.0)
        msg.empty()
        st.success(f"เสร็จสิ้น: ประมวลผล {stats['pages']} หน้า, ได้ {stats['rows']} แถว, ข้าม {stats['skipped']} หน้า")
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
    df = load_sn()
    st.subheader("SN tracking form (ตารางหลัก)")
    n_red = int(sum(1 for v in df["remark_red"] if norm_text(v))) if not df.empty else 0
    st.caption(f"ทั้งหมด {len(df)} แถว | แถวที่มีหมายเหตุสีแดงให้ตรวจ: {n_red} | วันที่ใช้รูปแบบ YYYY-MM-DD "
               "(ใน Excel จะแสดงตามรูปแบบของช่อง เช่น 07-Sep-26) | remark_red = ข้อความที่จะแสดงเป็นตัวแดง | "
               "cont = แถวต่อเนื่องของ Serial เดียวกัน (ไม่ลงวันที่/Serial ใน Excel)")
    view = _only_options(df, "status", STATUS_OPTIONS)
    edited = _stretch(
        st.data_editor, view, key=f"sn_{ver}", num_rows="dynamic", hide_index=True,
        column_config={
            "include": st.column_config.CheckboxColumn("ใช้", default=True),
            "cont": st.column_config.CheckboxColumn("cont", default=False),
            "rid": None, "page_key": None,
            "status": st.column_config.SelectboxColumn("status", options=STATUS_OPTIONS),
        },
    )
    c1, c2 = st.columns(2)
    if c1.button("บันทึกการแก้ไข", type="primary"):
        save_sn(edited)
        st.session_state["ver"] = ver + 1
        st.rerun()
    if c2.button("เติมข้อมูลจาก Data2 ใหม่ตาม Serial", disabled=cfg["ref"] is None):
        new, n = refill_from_data2(edited, cfg["ref"])
        save_sn(new)
        st.session_state["ver"] = ver + 1
        st.success(f"เติมใหม่ {n} แถว")
        st.rerun()

    st.subheader("ค่าใช้จ่ายรายวัน (Expense Claim)")
    st.caption("วันที่และชื่อ รพ. สร้างจากตาราง SN ด้านบน ส่วน Petrol / Express way / Parking / Other / Allowance กรอกเอง "
               "(ช่อง Service Type ใน Excel จะเป็นข้อความแดง \"โปรดลงข้อมูล\" ให้เลือกเอง)")
    exp = build_expense_df(load_sn(), load_expenses(), cfg["lang"])
    eedit = _stretch(
        st.data_editor, exp, key=f"exp_{ver}", num_rows="dynamic", hide_index=True,
        disabled=["hospitals"],
        column_config={c: st.column_config.NumberColumn(c, min_value=0, format="%.2f") for c in EXP_NUM},
    )
    if st.button("บันทึก Expense"):
        e = ensure_cols(eedit, EXP_COLS)
        e = e[e["date"].apply(lambda v: parse_date(v) is not None)].copy()
        e["date"] = e["date"].apply(lambda v: parse_date(v).isoformat())
        replace_table("expenses", EXP_COLS, e.drop_duplicates("date", keep="last"))
        st.session_state["ver"] = ver + 1
        st.rerun()


def tab_validate(cfg: dict) -> None:
    st.subheader("ตรวจสอบรหัสกับ Master (Data2 / Error Code / Spare part code)")
    ref = cfg["ref"]
    if ref is None:
        st.warning("กรุณาอัปโหลด Master Excel ก่อน")
        return
    val = validate_rows(load_sn(), ref)
    if val.empty:
        st.info("ยังไม่มีข้อมูลให้ตรวจสอบ")
        return
    counts = val["status"].value_counts()
    cols = st.columns(6)
    for col, s in zip(cols, ["OK", "SUGGEST", "UNKNOWN", "MISSING", "ERROR", "TODO"]):
        col.metric(s, int(counts.get(s, 0)))
    only_issue = st.checkbox("แสดงเฉพาะรายการที่มีปัญหา", value=True)
    view = val[val["status"] != "OK"] if only_issue else val
    _stretch(st.dataframe, view, hide_index=True)


def tab_export(cfg: dict) -> None:
    st.subheader("ตัวอย่างและ Export ไฟล์ Excel")
    sn_df = load_sn()
    if sn_df.empty:
        st.info("ยังไม่มีข้อมูล")
        return
    opts = {k: cfg[k] for k in ("staff", "lang", "overwrite", "fill_route", "month_override", "force_recalc")}
    st.markdown("**TMS Working Record (สร้างจาก SN tracking form: 1 แถวต่อ 1 รพ. ต่อวัน)**")
    rows = []
    for d, trips in sorted(build_trips(sn_df, cfg["lang"]).items()):
        for j, t in enumerate(trips):
            km, _u = distance_km([t["prov"]]) if t["prov"] else (None, [])
            rows.append({"date": d if j == 0 else "", "hospital": t["name"], "PM": len(t["kinds"]["PM"]),
                         "Service": len(t["kinds"]["Service"]), "Installation": len(t["kinds"]["Installation"]),
                         "Province": t["prov"], "km": km, "Start": fmt_time(t["start"]), "Finish": fmt_time(t["finish"])})
    if rows:
        _stretch(st.dataframe, _as_text_df(pd.DataFrame(rows)), hide_index=True)
    st.caption("ไฟล์ที่ได้แก้เฉพาะ 3 ชีตที่ระดับ XML ส่วนอื่นไม่ถูกแตะ ควรเก็บไฟล์ Master ต้นฉบับไว้เสมอ")
    if not MASTER_PATH.exists():
        st.warning("กรุณาอัปโหลด Master Excel ก่อน")
        return
    if st.button("สร้างไฟล์ Excel", type="primary"):
        try:
            data, log = fill_workbook(MASTER_PATH.read_bytes(), sn_df, load_expenses(), opts)
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
    t1, t2, t3, t4 = st.tabs(["1) อัปโหลด & ประมวลผล", "2) ตรวจทาน SN tracking", "3) ตรวจสอบรหัส", "4) Export Excel"])
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
else:  # streamlit may import the script as a module
    main()

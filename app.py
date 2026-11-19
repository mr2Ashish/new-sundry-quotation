import io
import json
import os
import re
import tempfile
from datetime import datetime
from difflib import SequenceMatcher

import openpyxl
import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Optional dependencies (graceful fallback)
try:
    from PIL import Image as PILImage
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import fpdf
    from fpdf import FPDF
    HAS_FPDF = True
    IS_FPDF2 = int(str(getattr(fpdf, "__version__", "1")).split(".")[0]) >= 2
except ImportError:
    HAS_FPDF = False
    IS_FPDF2 = False

try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False


# ============================================================
# PAGE CONFIG, CONSTANTS & CSS
# ============================================================
st.set_page_config(
    page_title="SundryQuote Pro — Enterprise BOM Quotation System",
    page_icon="📊",
    layout="wide",
)

DB_FILE = "mastersheet_db.csv"
INITIAL_SEED_FILE = "final_table_checked.csv"
ALIAS_FILE = "alias_memory.csv"
HISTORY_FILE = "quotation_history.json"
CLOUD_CFG_FILE = "gsheet_config.json"
LOGO_FILE = "company_logo.png"

MASTER_COLS = ["main_material", "sub_material", "purchase_price", "discount"]
ALIAS_COLS = ["bom_strict_key", "bom_sample_name", "master_strict_key", "master_sub_material", "created_at"]

DEFAULT_TERMS = """1. Prices: Ex-Works, exclusive of GST and P&F unless specified above.
2. Delivery: 2-3 weeks from receipt of technically & commercially clear PO.
3. Payment Terms: 100% against proforma invoice prior to dispatch / 30 days credit as agreed.
4. Validity: This quotation is valid for 30 days from the date of issue.
5. Make: Standard reputed industrial make unless specifically stated."""

st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(90deg, #1F4E78 0%, #2B6CB0 100%);
        padding: 1.15rem 1.5rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 1.1rem;
    }
    .main-header h1 { margin: 0; font-size: 1.75rem; color: white; }
    .main-header p { margin: 0.25rem 0 0 0; opacity: 0.92; font-size: 0.93rem; }

    /* Wrap long sub-material names inside Streamlit selectbox menus */
    div[data-baseweb="popover"] ul[role="listbox"] li {
        white-space: normal !important;
        word-wrap: break-word !important;
        height: auto !important;
        min-height: 38px !important;
        line-height: 1.35 !important;
        padding-top: 8px !important;
        padding-bottom: 8px !important;
        border-bottom: 1px solid #EDF2F7 !important;
    }
    div[data-baseweb="select"] > div {
        min-height: 42px !important;
        height: auto !important;
    }
    div[data-baseweb="select"] span {
        white-space: normal !important;
        word-break: break-word !important;
    }
    .bom-name-box {
        background-color: #F8FAFC;
        border-left: 4px solid #2D3748;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 0.92rem;
        font-weight: 600;
        color: #1A202C;
        margin-bottom: 6px;
        word-break: break-word;
    }
    .master-name-box {
        background-color: #EBF8FF;
        border-left: 4px solid #2B6CB0;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 0.91rem;
        color: #1A202C;
        margin-top: 6px;
        margin-bottom: 6px;
        word-break: break-word;
    }
    .warn-badge {
        background-color: #FFF5F5;
        border: 1px solid #FEB2B2;
        color: #C53030;
        padding: 4px 10px;
        border-radius: 5px;
        font-size: 0.84rem;
        font-weight: 600;
        display: inline-block;
        margin-top: 4px;
    }
    .safe-badge {
        background-color: #F0FFF4;
        border: 1px solid #9AE6B4;
        color: #22543D;
        padding: 4px 10px;
        border-radius: 5px;
        font-size: 0.84rem;
        font-weight: 600;
        display: inline-block;
        margin-top: 4px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# COMPANY LOGO UTILITIES (UPLOAD, NORMALIZE, SAVE & LOAD)
# ============================================================
def normalize_logo_png(raw_bytes: bytes, max_w: int = 380, max_h: int = 120) -> bytes:
    if not HAS_PIL:
        return raw_bytes
    im = PILImage.open(io.BytesIO(raw_bytes))
    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGBA")
    im.thumbnail((max_w, max_h), PILImage.LANCZOS)
    out = io.BytesIO()
    im.save(out, format="PNG")
    return out.getvalue()


def save_logo_bytes(raw_bytes: bytes) -> bytes:
    norm = normalize_logo_png(raw_bytes)
    try:
        with open(LOGO_FILE, "wb") as f:
            f.write(norm)
    except Exception:
        pass
    st.session_state["company_logo_bytes"] = norm
    return norm


def load_logo_bytes():
    if st.session_state.get("company_logo_bytes"):
        return st.session_state["company_logo_bytes"]
    if os.path.exists(LOGO_FILE):
        try:
            with open(LOGO_FILE, "rb") as f:
                data = f.read()
            st.session_state["company_logo_bytes"] = data
            return data
        except Exception:
            return None
    return None


def clear_logo():
    st.session_state["company_logo_bytes"] = None
    if os.path.exists(LOGO_FILE):
        try:
            os.remove(LOGO_FILE)
        except Exception:
            pass


# ============================================================
# STRICT KEY & MATCHING UTILITIES
# ============================================================
def strict_alnum_key(text: str) -> str:
    """Ignores spaces, commas, punctuation, symbols, and case. Keeps all A-Z and 0-9 strictly."""
    return re.sub(r"[^A-Za-z0-9]", "", str(text).upper())


def extract_tokens(text: str) -> set:
    return set(re.findall(r"[A-Z0-9]+", str(text).upper()))


def extract_numbers(text: str) -> set:
    return set(re.findall(r"\d+(?:\.\d+)?", str(text)))


def get_top_close_matches(bom_sub: str, bom_main: str, master_df: pd.DataFrame, top_n: int = 5):
    bom_key = strict_alnum_key(bom_sub)
    bom_tokens = extract_tokens(bom_sub)
    bom_nums = extract_numbers(bom_sub)
    bom_main_tokens = extract_tokens(bom_main)

    scored = []
    for idx, row in master_df.iterrows():
        m_sub = row["sub_material"]
        m_main = row["main_material"]
        m_key = strict_alnum_key(m_sub)
        m_tokens = extract_tokens(m_sub)
        m_nums = extract_numbers(m_sub)
        m_main_tokens = extract_tokens(m_main)

        seq_score = SequenceMatcher(None, bom_key, m_key).ratio()
        tok_union = len(bom_tokens | m_tokens)
        tok_score = (len(bom_tokens & m_tokens) / tok_union) if tok_union else 0.0

        num_mismatch = False
        if bom_nums and m_nums:
            num_score = len(bom_nums & m_nums) / max(len(bom_nums), len(m_nums))
            if bom_nums != m_nums:
                num_mismatch = True
        elif not bom_nums and not m_nums:
            num_score = 1.0
        else:
            num_score = 0.0
            num_mismatch = True

        main_bonus = 0.06 if (bom_main_tokens & m_main_tokens) else 0.0
        mismatch_penalty = 0.15 if num_mismatch else 0.0
        raw_score = (0.45 * seq_score) + (0.25 * tok_score) + (0.30 * num_score) + main_bonus - mismatch_penalty
        final_score = max(0.0, min(0.99, raw_score))

        scored.append(
            {
                "master_idx": idx,
                "main_material": m_main,
                "sub_material": m_sub,
                "purchase_price": float(row["purchase_price"]),
                "discount": float(row["discount"]),
                "score": round(final_score * 100, 1),
                "num_mismatch": num_mismatch,
                "bom_only_nums": sorted(bom_nums - m_nums),
                "master_only_nums": sorted(m_nums - bom_nums),
            }
        )

    scored.sort(key=lambda x: x["score"], reverse=True)
    return scored[:top_n]


# ============================================================
# LOCAL + OPTIONAL GOOGLE SHEETS CLOUD STORAGE & TOML GENERATOR
# ============================================================
def generate_streamlit_toml(gsheet_url: str, sa_dict: dict) -> str:
    """Converts Google Sheet URL + Service Account JSON into valid copy-paste Streamlit Secrets TOML."""
    lines = [f"gsheet_url = {json.dumps(str(gsheet_url).strip())}", "", "[gcp_service_account]"]
    for k, v in sa_dict.items():
        lines.append(f"{k} = {json.dumps(v)}")
    return "\n".join(lines)


def load_cloud_config() -> dict:
    try:
        if "gcp_service_account" in st.secrets and "gsheet_url" in st.secrets:
            return {
                "enabled": True,
                "auto_sync": True,
                "gsheet_url": st.secrets["gsheet_url"],
                "service_account": dict(st.secrets["gcp_service_account"]),
                "source": "streamlit_secrets",
            }
    except Exception:
        pass

    if os.path.exists(CLOUD_CFG_FILE):
        try:
            with open(CLOUD_CFG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                cfg["source"] = "local_config"
                return cfg
        except Exception:
            pass
    return {"enabled": False, "auto_sync": False, "gsheet_url": "", "service_account": None, "source": "none"}


def save_cloud_config(cfg: dict) -> None:
    to_save = {
        "enabled": bool(cfg.get("enabled", False)),
        "auto_sync": bool(cfg.get("auto_sync", False)),
        "gsheet_url": str(cfg.get("gsheet_url", "")).strip(),
        "service_account": cfg.get("service_account"),
    }
    with open(CLOUD_CFG_FILE, "w", encoding="utf-8") as f:
        json.dump(to_save, f, indent=2)


def get_gsheet_spreadsheet(cfg: dict):
    if not HAS_GSPREAD:
        raise RuntimeError("Missing libraries. Run: pip install gspread google-auth")
    if not cfg.get("service_account") or not cfg.get("gsheet_url"):
        raise RuntimeError("Google Sheet URL or Service Account JSON is not configured.")
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(cfg["service_account"], scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_url(cfg["gsheet_url"])


def get_or_create_worksheet(sh, title: str, headers: list):
    try:
        ws = sh.worksheet(title)
    except Exception:
        ws = sh.add_worksheet(title=title, rows=500, cols=max(10, len(headers)))
        ws.update("A1", [headers])
    return ws


def push_df_to_gsheet(df: pd.DataFrame, tab_name: str, cfg: dict):
    sh = get_gsheet_spreadsheet(cfg)
    ws = get_or_create_worksheet(sh, tab_name, list(df.columns))
    ws.clear()
    rows = [list(df.columns)] + df.fillna("").astype(str).values.tolist()
    ws.update("A1", rows)


def pull_df_from_gsheet(tab_name: str, expected_cols: list, cfg: dict) -> pd.DataFrame:
    sh = get_gsheet_spreadsheet(cfg)
    ws = get_or_create_worksheet(sh, tab_name, expected_cols)
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=expected_cols)
    return pd.DataFrame(records)


def push_history_to_gsheet(history: list, cfg: dict):
    rows = []
    for h in history:
        rows.append(
            {
                "id": h.get("id", ""),
                "annexure_no": h.get("annexure_no", ""),
                "revision": h.get("revision", 0),
                "created_at": h.get("created_at", ""),
                "customer": h.get("customer", ""),
                "project": h.get("project", ""),
                "grand_total": h.get("grand_total", 0.0),
                "payload_json": json.dumps(h, ensure_ascii=False),
            }
        )
    cols = ["id", "annexure_no", "revision", "created_at", "customer", "project", "grand_total", "payload_json"]
    df_h = pd.DataFrame(rows, columns=cols)
    push_df_to_gsheet(df_h, "Quotation_History", cfg)


def pull_history_from_gsheet(cfg: dict) -> list:
    cols = ["id", "annexure_no", "revision", "created_at", "customer", "project", "grand_total", "payload_json"]
    df_h = pull_df_from_gsheet("Quotation_History", cols, cfg)
    history = []
    for _, r in df_h.iterrows():
        raw = str(r.get("payload_json", "")).strip()
        if raw:
            try:
                history.append(json.loads(raw))
            except Exception:
                continue
    return history


def bootstrap_from_cloud_once():
    """On Streamlit Cloud cold start, auto-pull live data from Google Sheets once per session."""
    if st.session_state.get("_cloud_bootstrapped"):
        return
    cfg = load_cloud_config()
    if cfg.get("source") == "streamlit_secrets" and HAS_GSPREAD:
        try:
            m_cloud = pull_df_from_gsheet("Mastersheet", MASTER_COLS, cfg)
            if not m_cloud.empty:
                save_mastersheet(m_cloud, sync_cloud=False)
            a_cloud = pull_df_from_gsheet("Alias_Memory", ALIAS_COLS, cfg)
            if not a_cloud.empty:
                save_aliases(a_cloud, sync_cloud=False)
            h_cloud = pull_history_from_gsheet(cfg)
            if h_cloud:
                save_history(h_cloud, sync_cloud=False)
        except Exception:
            pass
    st.session_state["_cloud_bootstrapped"] = True


# ============================================================
# MASTERSHEET, ALIAS MEMORY & QUOTATION HISTORY PERSISTENCE
# ============================================================
def clean_master_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    rename_map = {}
    for col in df.columns:
        c = str(col).strip().lower().replace(" ", "_")
        if "main" in c:
            rename_map[col] = "main_material"
        elif "sub" in c or "desc" in c or "item" in c:
            rename_map[col] = "sub_material"
        elif "purchase" in c or "price" in c or "rate" in c:
            rename_map[col] = "purchase_price"
        elif "disc" in c:
            rename_map[col] = "discount"
    df = df.rename(columns=rename_map)

    for col in MASTER_COLS:
        if col not in df.columns:
            df[col] = "" if col in ("main_material", "sub_material") else 0.0

    df = df[MASTER_COLS]
    df["main_material"] = df["main_material"].fillna("").astype(str).str.strip()
    df["sub_material"] = df["sub_material"].fillna("").astype(str).str.strip()
    df["purchase_price"] = pd.to_numeric(df["purchase_price"], errors="coerce").fillna(0.0)
    df["discount"] = pd.to_numeric(df["discount"], errors="coerce").fillna(0.0)
    return df[df["sub_material"] != ""].reset_index(drop=True)


def load_mastersheet() -> pd.DataFrame:
    if os.path.exists(DB_FILE):
        return clean_master_df(pd.read_csv(DB_FILE, encoding="utf-8-sig"))
    if os.path.exists(INITIAL_SEED_FILE):
        df = clean_master_df(pd.read_csv(INITIAL_SEED_FILE, encoding="utf-8-sig"))
        save_mastersheet(df, sync_cloud=False)
        return df
    starter = pd.DataFrame(
        [
            {"main_material": "CABLE TIE", "sub_material": "CABLE TIE 100MM (100 TIES /PACK)", "purchase_price": 52.0, "discount": 50.0},
            {"main_material": "CABLE TIE", "sub_material": "CABLE TIE 200MM (100 TIES /PACK)", "purchase_price": 153.0, "discount": 50.0},
            {"main_material": "BLIND - SIGNAL", "sub_material": 'BLIND FOR CABLE GLAND 1/2" NPT', "purchase_price": 133.0, "discount": 0.0},
        ]
    )
    save_mastersheet(starter, sync_cloud=False)
    return starter


def save_mastersheet(df: pd.DataFrame, sync_cloud: bool = True) -> None:
    cleaned = clean_master_df(df)
    cleaned.to_csv(DB_FILE, index=False, encoding="utf-8-sig")
    if sync_cloud:
        cfg = load_cloud_config()
        if cfg.get("enabled") and cfg.get("auto_sync") and HAS_GSPREAD:
            try:
                push_df_to_gsheet(cleaned, "Mastersheet", cfg)
            except Exception as e:
                st.toast(f"Cloud sync warning (Mastersheet): {e}", icon="⚠️")


def load_aliases() -> pd.DataFrame:
    if os.path.exists(ALIAS_FILE):
        try:
            df = pd.read_csv(ALIAS_FILE, encoding="utf-8-sig")
            for c in ALIAS_COLS:
                if c not in df.columns:
                    df[c] = ""
            df = df[ALIAS_COLS].fillna("").astype(str)
            return df[df["bom_strict_key"] != ""].drop_duplicates(subset=["bom_strict_key"], keep="last").reset_index(drop=True)
        except Exception:
            pass
    return pd.DataFrame(columns=ALIAS_COLS)


def save_aliases(df: pd.DataFrame, sync_cloud: bool = True) -> None:
    out = df.copy()
    for c in ALIAS_COLS:
        if c not in out.columns:
            out[c] = ""
    out = out[ALIAS_COLS].fillna("").astype(str)
    out = out[out["bom_strict_key"] != ""].drop_duplicates(subset=["bom_strict_key"], keep="last").reset_index(drop=True)
    out.to_csv(ALIAS_FILE, index=False, encoding="utf-8-sig")
    if sync_cloud:
        cfg = load_cloud_config()
        if cfg.get("enabled") and cfg.get("auto_sync") and HAS_GSPREAD:
            try:
                push_df_to_gsheet(out, "Alias_Memory", cfg)
            except Exception as e:
                st.toast(f"Cloud sync warning (Alias): {e}", icon="⚠️")


def learn_aliases(new_pairs: list) -> int:
    if not new_pairs:
        return 0
    existing = load_aliases()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    add_rows = []
    for p in new_pairs:
        bk = strict_alnum_key(p["bom_sub"])
        mk = strict_alnum_key(p["master_sub"])
        if bk and mk and bk != mk:
            add_rows.append(
                {
                    "bom_strict_key": bk,
                    "bom_sample_name": p["bom_sub"],
                    "master_strict_key": mk,
                    "master_sub_material": p["master_sub"],
                    "created_at": now_str,
                }
            )
    if not add_rows:
        return 0
    combined = pd.concat([existing, pd.DataFrame(add_rows)], ignore_index=True)
    save_aliases(combined)
    return len(add_rows)


def load_history() -> list:
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except Exception:
            return []
    return []


def save_history(history: list, sync_cloud: bool = True) -> None:
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, indent=2, ensure_ascii=False)
    if sync_cloud:
        cfg = load_cloud_config()
        if cfg.get("enabled") and cfg.get("auto_sync") and HAS_GSPREAD:
            try:
                push_history_to_gsheet(history, cfg)
            except Exception as e:
                st.toast(f"Cloud sync warning (History): {e}", icon="⚠️")


def get_next_revision(annexure_no: str, history: list) -> int:
    revs = [int(h.get("revision", 0)) for h in history if str(h.get("annexure_no", "")).strip() == str(annexure_no).strip()]
    return (max(revs) + 1) if revs else 0


# ============================================================
# MULTI-FORMAT BOM PARSER + HEADER METADATA EXTRACTOR
# ============================================================
def extract_qty(val) -> float:
    if pd.isna(val):
        return 0.0
    s = str(val).strip()
    if not s or s.upper() in ("NA", "N/A", "-", "--", "NIL", "NONE", "NULL", "#ERROR!", "#REF!", "#VALUE!"):
        return 0.0
    m = re.search(r"(\d+(?:\.\d+)?)", s.replace(",", ""))
    return float(m.group(1)) if m else 0.0


def parse_bom_dataframe(df_raw: pd.DataFrame, source_label: str = "", row_numbers=None):
    empty_meta = {"project": "", "customer": "", "location": "", "job_no": ""}
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=["main_material", "sub_material", "qty", "source"]), empty_meta

    df_raw = df_raw.reset_index(drop=True)
    if row_numbers is None:
        row_numbers = list(range(1, len(df_raw) + 1))

    header_idx = None
    col_map = {}

    for idx, row in df_raw.iterrows():
        vals = [str(v).strip().lower() for v in row.tolist()]
        desc_col = next(
            (i for i, v in enumerate(vals) if any(k in v for k in ["description", "sub_material", "sub material", "item name", "particular", "material name"])),
            None,
        )
        qty_col = next(
            (i for i, v in enumerate(vals) if any(k in v for k in ["qty", "quantity", "nos", "count"])),
            None,
        )
        if desc_col is not None and qty_col is not None:
            header_idx = idx
            sr_col = next((i for i, v in enumerate(vals) if any(k in v for k in ["sr", "sl", "s.no", "no."])), 0)
            main_col = next(
                (i for i, v in enumerate(vals) if any(k in v for k in ["main_material", "main material", "category", "group", "head"])),
                None,
            )
            col_map = {"sr": sr_col, "desc": desc_col, "qty": qty_col, "main": main_col}
            break

    if header_idx is None:
        return pd.DataFrame(columns=["main_material", "sub_material", "qty", "source"]), empty_meta

    meta = dict(empty_meta)
    for idx in range(header_idx):
        cells = [str(v).strip() for v in df_raw.iloc[idx].tolist() if pd.notna(v) and str(v).strip() != ""]
        row_text = " | ".join(cells)
        for c in cells:
            m_proj = re.search(r"PROJECT\s*[:\-]\s*(.+)", c, re.IGNORECASE)
            if m_proj and not meta["project"]:
                meta["project"] = m_proj.group(1).strip()
            m_cust = re.search(r"CUSTOMER\s*[:\-]\s*(.+)", c, re.IGNORECASE)
            if m_cust and not meta["customer"]:
                meta["customer"] = m_cust.group(1).strip()
            m_loc = re.search(r"LOCATION\s*[:\-]\s*(.+)", c, re.IGNORECASE)
            if m_loc and not meta["location"]:
                meta["location"] = m_loc.group(1).strip()
        m_job = re.search(r"(?:JOB\s*NO|ANNEXURE|QUOTATION\s*NO)[^\d]*(\d{4,})", row_text, re.IGNORECASE)
        if m_job and not meta["job_no"]:
            meta["job_no"] = m_job.group(1)

    items = []
    current_main = "GENERAL"

    for idx in range(header_idx + 1, len(df_raw)):
        row = df_raw.iloc[idx]
        orig_row_no = row_numbers[idx]
        sr_val = "" if col_map["sr"] >= len(row) or pd.isna(row.iloc[col_map["sr"]]) else str(row.iloc[col_map["sr"]]).strip()
        desc_val = "" if col_map["desc"] >= len(row) or pd.isna(row.iloc[col_map["desc"]]) else str(row.iloc[col_map["desc"]]).strip()
        qty_raw = row.iloc[col_map["qty"]] if col_map["qty"] < len(row) else None

        if desc_val.upper() in ("TOTAL", "GRAND TOTAL") or sr_val.upper() in ("TOTAL", "GRAND TOTAL"):
            continue

        row_source = f"{source_label} (Row {orig_row_no})"

        if col_map["main"] is not None:
            main_val = "" if pd.isna(row.iloc[col_map["main"]]) else str(row.iloc[col_map["main"]]).strip()
            if main_val:
                current_main = main_val
            if desc_val:
                items.append({"main_material": current_main, "sub_material": desc_val, "qty": extract_qty(qty_raw), "source": row_source})
        else:
            if sr_val and not sr_val.replace(".", "").isdigit() and not desc_val:
                current_main = sr_val
                continue
            if not sr_val and desc_val and (pd.isna(qty_raw) or str(qty_raw).strip() == ""):
                current_main = desc_val
                continue
            if desc_val:
                items.append({"main_material": current_main, "sub_material": desc_val, "qty": extract_qty(qty_raw), "source": row_source})

    return pd.DataFrame(items), meta


def read_excel_visible_only(raw_bytes: bytes, fname: str, ignore_hidden: bool = True):
    results = {}
    hidden_rows_skipped = 0
    hidden_sheets_skipped = 0

    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=False)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ignore_hidden and ws.sheet_state != "visible":
            hidden_sheets_skipped += 1
            continue

        visible_cols = []
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            col_dim = ws.column_dimensions.get(col_letter)
            if ignore_hidden and col_dim is not None and col_dim.hidden:
                continue
            visible_cols.append(col_idx)

        rows_data = []
        orig_row_nums = []
        for row_idx in range(1, ws.max_row + 1):
            row_dim = ws.row_dimensions.get(row_idx)
            if ignore_hidden and row_dim is not None and row_dim.hidden:
                hidden_rows_skipped += 1
                continue
            row_vals = [ws.cell(row=row_idx, column=c).value for c in visible_cols]
            if any(v is not None and str(v).strip() != "" for v in row_vals):
                rows_data.append(row_vals)
                orig_row_nums.append(row_idx)

        if rows_data:
            df_raw = pd.DataFrame(rows_data)
            parsed, meta = parse_bom_dataframe(df_raw, source_label=f"{fname} [{sheet_name}]", row_numbers=orig_row_nums)
            if not parsed.empty:
                results[f"{fname} → {sheet_name}"] = (parsed, meta)
    wb.close()
    return results, hidden_rows_skipped, hidden_sheets_skipped


def read_any_bom_file(uploaded_file, ignore_hidden: bool = True):
    fname = uploaded_file.name
    ext = os.path.splitext(fname)[1].lower()
    raw_bytes = uploaded_file.getvalue()
    results = {}
    hidden_rows_skipped = 0
    hidden_sheets_skipped = 0

    if ext in (".xlsx", ".xlsm"):
        results, hidden_rows_skipped, hidden_sheets_skipped = read_excel_visible_only(raw_bytes, fname, ignore_hidden=ignore_hidden)
    elif ext == ".xls":
        xls = pd.ExcelFile(io.BytesIO(raw_bytes), engine="xlrd")
        for sheet in xls.sheet_names:
            try:
                df_raw = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet, header=None, engine="xlrd")
                parsed, meta = parse_bom_dataframe(df_raw, source_label=f"{fname} [{sheet}]")
                if not parsed.empty:
                    results[f"{fname} → {sheet}"] = (parsed, meta)
            except Exception:
                continue
    elif ext in (".csv", ".tsv", ".txt"):
        df_raw = None
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                sep = "\t" if ext == ".tsv" else None
                df_raw = pd.read_csv(io.BytesIO(raw_bytes), header=None, encoding=enc, sep=sep, engine="python", on_bad_lines="skip")
                break
            except Exception:
                continue
        if df_raw is not None:
            parsed, meta = parse_bom_dataframe(df_raw, source_label=fname)
            if not parsed.empty:
                results[fname] = (parsed, meta)
    elif ext == ".pdf":
        all_rows = []
        with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
            for page in pdf.pages:
                for tbl in (page.extract_tables() or []):
                    for r in tbl:
                        if r and any(cell and str(cell).strip() for cell in r):
                            all_rows.append([str(c).replace("\n", " ").strip() if c else "" for c in r])
        if all_rows:
            max_cols = max(len(r) for r in all_rows)
            df_raw = pd.DataFrame([r + [""] * (max_cols - len(r)) for r in all_rows])
            parsed, meta = parse_bom_dataframe(df_raw, source_label=fname)
            if not parsed.empty:
                results[fname] = (parsed, meta)
    elif ext == ".docx":
        doc = Document(io.BytesIO(raw_bytes))
        all_rows = []
        for tbl in doc.tables:
            for row in tbl.rows:
                vals = [c.text.replace("\n", " ").strip() for c in row.cells]
                if any(vals):
                    all_rows.append(vals)
        if all_rows:
            max_cols = max(len(r) for r in all_rows)
            df_raw = pd.DataFrame([r + [""] * (max_cols - len(r)) for r in all_rows])
            parsed, meta = parse_bom_dataframe(df_raw, source_label=fname)
            if not parsed.empty:
                results[fname] = (parsed, meta)

    m = re.search(r"(\d{6,})", fname)
    fname_job = m.group(1) if m else ""
    for k, (df_p, meta) in results.items():
        if not meta.get("job_no") and fname_job:
            meta["job_no"] = fname_job
            results[k] = (df_p, meta)

    return results, hidden_rows_skipped, hidden_sheets_skipped


# ============================================================
# MARGIN & COMMERCIAL CALCULATION ENGINE
# ============================================================
def compute_commercials(items_df: pd.DataFrame, profit_pct: float, pf_pct: float, gst_pct: float):
    df = items_df.copy()
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce").fillna(0.0)
    df["purchase_price"] = pd.to_numeric(df["purchase_price"], errors="coerce").fillna(0.0)
    df["discount"] = pd.to_numeric(df["discount"], errors="coerce").fillna(0.0)

    df["Amnt"] = df["purchase_price"] * (1.0 + profit_pct / 100.0)
    df["Net Unit Rate"] = df["Amnt"] * (1.0 - df["discount"] / 100.0)
    df["Total Amnt"] = df["qty"] * df["Net Unit Rate"]
    df["Total Purchase"] = df["qty"] * df["purchase_price"]
    df["Net Profit"] = df["Total Amnt"] - df["Total Purchase"]
    df["Margin %"] = df.apply(
        lambda r: ((r["Net Unit Rate"] - r["purchase_price"]) / r["purchase_price"] * 100.0) if r["purchase_price"] > 0 else 0.0,
        axis=1,
    )
    df["Loss Alert"] = df.apply(
        lambda r: "🔴 LOSS" if (r["purchase_price"] > 0 and r["Net Unit Rate"] < r["purchase_price"] - 1e-6) else ("⚠️ ₹0 Rate" if r["purchase_price"] == 0 else "✅ OK"),
        axis=1,
    )

    basic_total = float(df["Total Amnt"].sum())
    total_purchase = float(df["Total Purchase"].sum())
    pf_amount = basic_total * (pf_pct / 100.0)
    taxable_total = basic_total + pf_amount
    gst_amount = taxable_total * (gst_pct / 100.0)
    grand_total_incl_gst = taxable_total + gst_amount
    net_profit_total = basic_total - total_purchase
    loss_rows = df[df["Loss Alert"] == "🔴 LOSS"]

    summary = {
        "basic_total": basic_total,
        "total_purchase": total_purchase,
        "pf_pct": pf_pct,
        "pf_amount": pf_amount,
        "taxable_total": taxable_total,
        "gst_pct": gst_pct,
        "gst_amount": gst_amount,
        "grand_total_incl_gst": grand_total_incl_gst,
        "net_profit_total": net_profit_total,
        "overall_margin_pct": (net_profit_total / total_purchase * 100.0) if total_purchase > 0 else 0.0,
        "loss_item_count": len(loss_rows),
        "loss_total_amount": float(loss_rows["Net Profit"].sum()) if not loss_rows.empty else 0.0,
        "zero_rate_count": int((df["purchase_price"] == 0).sum()),
    }
    return df, summary


# ============================================================
# REVISION COMPARISON ENGINE (REV A vs REV B)
# ============================================================
def compare_revisions(rev_a: dict, rev_b: dict):
    df_a, sum_a = compute_commercials(
        pd.DataFrame(rev_a.get("items", [])),
        float(rev_a.get("profit_pct", 12.0)),
        float(rev_a.get("pf_pct", 0.0)),
        float(rev_a.get("gst_pct", 18.0)),
    )
    df_b, sum_b = compute_commercials(
        pd.DataFrame(rev_b.get("items", [])),
        float(rev_b.get("profit_pct", 12.0)),
        float(rev_b.get("pf_pct", 0.0)),
        float(rev_b.get("gst_pct", 18.0)),
    )

    df_a["_k"] = df_a["sub_material"].apply(strict_alnum_key)
    df_b["_k"] = df_b["sub_material"].apply(strict_alnum_key)

    map_a = df_a.drop_duplicates("_k").set_index("_k").to_dict(orient="index")
    map_b = df_b.drop_duplicates("_k").set_index("_k").to_dict(orient="index")

    ordered_keys = list(dict.fromkeys(df_a["_k"].tolist() + df_b["_k"].tolist()))
    rows = []
    counts = {"Added": 0, "Removed": 0, "Modified": 0, "Unchanged": 0}

    for k in ordered_keys:
        a = map_a.get(k)
        b = map_b.get(k)
        if a and not b:
            counts["Removed"] += 1
            rows.append(
                {
                    "Status": "Removed",
                    "Main Material": a["main_material"],
                    "Sub Material": a["sub_material"],
                    "Old QTY": float(a["qty"]),
                    "New QTY": 0.0,
                    "Old Net Rate": round(float(a["Net Unit Rate"]), 2),
                    "New Net Rate": 0.0,
                    "Old Total": round(float(a["Total Amnt"]), 2),
                    "New Total": 0.0,
                    "Delta (INR)": round(-float(a["Total Amnt"]), 2),
                    "Change Notes": "Item removed in revised quote",
                }
            )
        elif b and not a:
            counts["Added"] += 1
            rows.append(
                {
                    "Status": "Added",
                    "Main Material": b["main_material"],
                    "Sub Material": b["sub_material"],
                    "Old QTY": 0.0,
                    "New QTY": float(b["qty"]),
                    "Old Net Rate": 0.0,
                    "New Net Rate": round(float(b["Net Unit Rate"]), 2),
                    "Old Total": 0.0,
                    "New Total": round(float(b["Total Amnt"]), 2),
                    "Delta (INR)": round(float(b["Total Amnt"]), 2),
                    "Change Notes": "New item added",
                }
            )
        else:
            qty_diff = abs(float(b["qty"]) - float(a["qty"])) > 1e-6
            rate_diff = abs(float(b["Net Unit Rate"]) - float(a["Net Unit Rate"])) > 1e-4
            disc_diff = abs(float(b["discount"]) - float(a["discount"])) > 1e-4
            notes = []
            if qty_diff:
                notes.append(f"QTY {a['qty']:g}->{b['qty']:g}")
            if rate_diff:
                notes.append(f"Rate Rs.{a['Net Unit Rate']:,.2f}->Rs.{b['Net Unit Rate']:,.2f}")
            if disc_diff:
                notes.append(f"Disc {a['discount']:g}%->{b['discount']:g}%")

            status = "Modified" if notes else "Unchanged"
            counts[status] += 1
            delta = float(b["Total Amnt"]) - float(a["Total Amnt"])
            rows.append(
                {
                    "Status": status,
                    "Main Material": b["main_material"],
                    "Sub Material": b["sub_material"],
                    "Old QTY": float(a["qty"]),
                    "New QTY": float(b["qty"]),
                    "Old Net Rate": round(float(a["Net Unit Rate"]), 2),
                    "New Net Rate": round(float(b["Net Unit Rate"]), 2),
                    "Old Total": round(float(a["Total Amnt"]), 2),
                    "New Total": round(float(b["Total Amnt"]), 2),
                    "Delta (INR)": round(delta, 2),
                    "Change Notes": "; ".join(notes) if notes else "No change",
                }
            )

    diff_df = pd.DataFrame(rows)
    basic_delta = sum_b["basic_total"] - sum_a["basic_total"]
    grand_delta = sum_b["grand_total_incl_gst"] - sum_a["grand_total_incl_gst"]
    pct_delta = (grand_delta / sum_a["grand_total_incl_gst"] * 100.0) if sum_a["grand_total_incl_gst"] > 0 else 0.0

    comp_meta = {
        "label_a": f"Annexure {rev_a['annexure_no']} (Rev {rev_a['revision']})",
        "label_b": f"Annexure {rev_b['annexure_no']} (Rev {rev_b['revision']})",
        "date_a": rev_a.get("created_at", ""),
        "date_b": rev_b.get("created_at", ""),
        "customer": rev_b.get("customer") or rev_a.get("customer", ""),
        "project": rev_b.get("project") or rev_a.get("project", ""),
        "company_name": rev_b.get("company_name") or rev_a.get("company_name", "Quotation Comparison"),
        "sum_a": sum_a,
        "sum_b": sum_b,
        "basic_delta": basic_delta,
        "grand_delta": grand_delta,
        "pct_delta": pct_delta,
        "counts": counts,
    }
    return diff_df, comp_meta


# ============================================================
# EXPORT 1: INTERNAL COSTING EXCEL (EXACT quotation 6500389 FORMAT)
# ============================================================
def build_internal_quotation_excel(
    final_df: pd.DataFrame,
    meta: dict,
    summary: dict,
    profit_pct: float,
    default_make: str,
    audit_df: pd.DataFrame = None,
) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_title = Font(name="Calibri", size=13, bold=True, color="1F4E78")
    font_cat = Font(name="Calibri", size=11, bold=True, color="000000")
    font_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10, color="000000")
    font_loss = Font(name="Calibri", size=10, bold=True, color="9C0006")
    font_total = Font(name="Calibri", size=11, bold=True, color="000000")

    fill_cat = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_profit = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_total = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_loss = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    thin = Side(style="thin", color="BFBFBF")
    med = Side(style="medium", color="1F4E78")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_total = Border(top=med, bottom=med, left=thin, right=thin)

    annexure_label = f"Annexure {meta['annexure_no']} (Rev {meta['revision']})"
    ws.cell(row=1, column=2, value=annexure_label).font = font_title
    ws.row_dimensions[1].height = 24

    groups = list(final_df.groupby("main_material", sort=False)) if not final_df.empty else []
    current_row = 2
    first_data_row = None

    for g_idx, (cat_name, grp) in enumerate(groups):
        if g_idx == 0:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            c_cat = ws.cell(row=current_row, column=1, value=cat_name)
            c_cat.font = font_cat
            c_cat.fill = fill_cat
            c_cat.alignment = Alignment(horizontal="left", vertical="center")
            ws.cell(row=current_row, column=1).border = border_cell
            ws.cell(row=current_row, column=2).border = border_cell

            headers = {3: "Make", 4: "QTY", 5: "Amnt", 6: "Discount", 7: "Total Amnt", 9: "Purchase", 11: "Profit", 13: "Margin Check"}
            for col_i, h_text in headers.items():
                c = ws.cell(row=current_row, column=col_i, value=h_text)
                c.font = font_hdr
                c.fill = fill_hdr
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = border_cell

            c_prof = ws.cell(row=current_row, column=12, value=float(profit_pct))
            c_prof.font = Font(name="Calibri", size=11, bold=True, color="9C6500")
            c_prof.fill = fill_profit
            c_prof.alignment = Alignment(horizontal="center", vertical="center")
            c_prof.border = border_cell
            ws.row_dimensions[current_row].height = 24
            current_row += 1
        else:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            c_cat = ws.cell(row=current_row, column=1, value=cat_name)
            c_cat.font = font_cat
            c_cat.alignment = Alignment(horizontal="left", vertical="center")
            for col_i in range(1, 10):
                cell = ws.cell(row=current_row, column=col_i)
                cell.fill = fill_cat
                cell.border = border_cell
            ws.row_dimensions[current_row].height = 21
            current_row += 1

        for sr_no, (_, item) in enumerate(grp.iterrows(), start=1):
            if first_data_row is None:
                first_data_row = current_row
            r = current_row
            is_loss = item.get("Loss Alert") == "🔴 LOSS"

            ws.cell(row=r, column=1, value=sr_no).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=r, column=2, value=item["sub_material"]).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.cell(row=r, column=3, value=default_make).alignment = Alignment(horizontal="center", vertical="center")

            c4 = ws.cell(row=r, column=4, value=float(item["qty"]))
            c4.number_format = "#,##0.##"
            c4.alignment = Alignment(horizontal="right", vertical="center")

            c5 = ws.cell(row=r, column=5, value=f"=I{r}*(1+$L$2/100)")
            c5.number_format = "#,##0.00"
            c5.alignment = Alignment(horizontal="right", vertical="center")

            c6 = ws.cell(row=r, column=6, value=float(item["discount"]))
            c6.number_format = "0.##"
            c6.alignment = Alignment(horizontal="right", vertical="center")

            c7 = ws.cell(row=r, column=7, value=f"=D{r}*E{r}*(1-F{r}/100)")
            c7.number_format = "#,##0.00"
            c7.alignment = Alignment(horizontal="right", vertical="center")

            c9 = ws.cell(row=r, column=9, value=float(item["purchase_price"]))
            c9.number_format = "#,##0.00"
            c9.alignment = Alignment(horizontal="right", vertical="center")

            c13 = ws.cell(row=r, column=13, value=f'=IF(I{r}=0,"ZERO RATE",IF((E{r}*(1-F{r}/100))<I{r},"LOSS","OK"))')
            c13.alignment = Alignment(horizontal="center", vertical="center")

            for col_i in (1, 2, 3, 4, 5, 6, 7, 9, 13):
                cell = ws.cell(row=r, column=col_i)
                cell.font = font_loss if is_loss else font_data
                cell.border = border_cell
                if is_loss:
                    cell.fill = fill_loss

            current_row += 1

    last_data_row = current_row - 1
    total_row = current_row
    if first_data_row is not None:
        ws.cell(row=total_row, column=5, value="TOTAL").font = font_total
        ws.cell(row=total_row, column=5).alignment = Alignment(horizontal="right", vertical="center")
        c_tot = ws.cell(row=total_row, column=7, value=f"=SUM(G{first_data_row}:G{last_data_row})")
        c_tot.font = font_total
        c_tot.number_format = "#,##0.00"
        c_tot.alignment = Alignment(horizontal="right", vertical="center")
        for col_i in range(1, 10):
            ws.cell(row=total_row, column=col_i).fill = fill_total
            ws.cell(row=total_row, column=col_i).border = border_total
        ws.row_dimensions[total_row].height = 24

    widths = {1: 10, 2: 68, 3: 12, 4: 10, 5: 13, 6: 11, 7: 16, 8: 3, 9: 13, 10: 3, 11: 10, 12: 8, 13: 14}
    for col_i, w in widths.items():
        ws.column_dimensions[get_column_letter(col_i)].width = w
    ws.freeze_panes = "A3"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws_sum = wb.create_sheet("Commercial_Summary")
    ws_sum.cell(row=1, column=1, value=f"{meta.get('company_name', 'QUOTATION')} — INTERNAL COMMERCIAL SUMMARY").font = font_title
    meta_rows = [
        ("Quotation / Annexure No", f"{meta['annexure_no']} (Rev {meta['revision']})"),
        ("Date", meta["quote_date"]),
        ("Customer", meta.get("customer", "")),
        ("Project", meta.get("project", "")),
        ("Location", meta.get("location", "")),
    ]
    r_s = 3
    for k, v in meta_rows:
        ws_sum.cell(row=r_s, column=1, value=k).font = font_cat
        ws_sum.cell(row=r_s, column=2, value=v).font = font_data
        r_s += 1

    r_s += 1
    ws_sum.cell(row=r_s, column=1, value="Commercial Component").font = font_hdr
    ws_sum.cell(row=r_s, column=1).fill = fill_hdr
    ws_sum.cell(row=r_s, column=2, value="Amount (INR)").font = font_hdr
    ws_sum.cell(row=r_s, column=2).fill = fill_hdr
    r_s += 1

    basic_ref = f"=Sheet1!G{total_row}" if first_data_row is not None else 0
    r_basic = r_s
    ws_sum.cell(row=r_basic, column=1, value="Basic Material Total (Ex-Works)").font = font_data
    ws_sum.cell(row=r_basic, column=2, value=basic_ref).number_format = "#,##0.00"

    r_pf = r_basic + 1
    ws_sum.cell(row=r_pf, column=1, value=f"Packing & Forwarding ({summary['pf_pct']:g}%)").font = font_data
    ws_sum.cell(row=r_pf, column=2, value=f"=B{r_basic}*{summary['pf_pct']/100.0}").number_format = "#,##0.00"

    r_tax = r_pf + 1
    ws_sum.cell(row=r_tax, column=1, value="Taxable Value").font = font_cat
    ws_sum.cell(row=r_tax, column=2, value=f"=B{r_basic}+B{r_pf}").number_format = "#,##0.00"

    r_gst = r_tax + 1
    ws_sum.cell(row=r_gst, column=1, value=f"GST ({summary['gst_pct']:g}%)").font = font_data
    ws_sum.cell(row=r_gst, column=2, value=f"=B{r_tax}*{summary['gst_pct']/100.0}").number_format = "#,##0.00"

    r_gt = r_gst + 1
    ws_sum.cell(row=r_gt, column=1, value="GRAND TOTAL (Inclusive of GST)").font = font_total
    c_gt = ws_sum.cell(row=r_gt, column=2, value=f"=B{r_tax}+B{r_gst}")
    c_gt.font = font_total
    c_gt.number_format = "#,##0.00"
    ws_sum.cell(row=r_gt, column=1).fill = fill_total
    ws_sum.cell(row=r_gt, column=2).fill = fill_total

    r_int = r_gt + 2
    ws_sum.cell(row=r_int, column=1, value="Internal Purchase Cost").font = font_cat
    ws_sum.cell(row=r_int, column=2, value=summary["total_purchase"]).number_format = "#,##0.00"
    ws_sum.cell(row=r_int + 1, column=1, value="Estimated Net Profit / (Loss)").font = font_cat
    ws_sum.cell(row=r_int + 1, column=2, value=f"=B{r_basic}-B{r_int}").number_format = "#,##0.00"
    ws_sum.cell(row=r_int + 2, column=1, value="Loss-Making Line Items Flagged").font = font_loss if summary["loss_item_count"] > 0 else font_cat
    ws_sum.cell(row=r_int + 2, column=2, value=summary["loss_item_count"])

    ws_sum.column_dimensions["A"].width = 38
    ws_sum.column_dimensions["B"].width = 48

    if audit_df is not None and not audit_df.empty:
        ws2 = wb.create_sheet("Match_Audit")
        for c_idx, col_name in enumerate(audit_df.columns, start=1):
            c = ws2.cell(row=1, column=c_idx, value=col_name)
            c.font = font_hdr
            c.fill = fill_hdr
            c.alignment = Alignment(horizontal="center")
        for r_idx, (_, row) in enumerate(audit_df.iterrows(), start=2):
            for c_idx, val in enumerate(row.tolist(), start=1):
                cell = ws2.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_data
                cell.border = border_cell
        for col_i in range(1, len(audit_df.columns) + 1):
            ws2.column_dimensions[get_column_letter(col_i)].width = 34
        ws2.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# EXPORT 2: CUSTOMER-FACING EXCEL (WITH LOGO & NO COSTING)
# ============================================================
def build_customer_quotation_excel(
    final_df: pd.DataFrame,
    meta: dict,
    summary: dict,
    default_make: str,
    logo_bytes: bytes = None,
) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws_cov = wb.active
    ws_cov.title = "Quotation_Cover"

    font_company = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    font_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_cat = Font(name="Calibri", size=11, bold=True, color="000000")
    font_data = Font(name="Calibri", size=10, color="000000")
    font_total = Font(name="Calibri", size=11, bold=True, color="000000")

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_cat = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_total = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_light = PatternFill(start_color="F2F6FC", end_color="F2F6FC", fill_type="solid")

    thin = Side(style="thin", color="BFBFBF")
    med = Side(style="medium", color="1F4E78")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_total = Border(top=med, bottom=med, left=thin, right=thin)

    logo_added = False
    if logo_bytes and HAS_PIL:
        try:
            pil_im = PILImage.open(io.BytesIO(logo_bytes))
            aspect = pil_im.width / float(pil_im.height or 1)
            xl_img = XLImage(io.BytesIO(logo_bytes))
            xl_img.height = 54
            xl_img.width = min(210, int(54 * aspect))
            ws_cov.add_image(xl_img, "A1")
            ws_cov.row_dimensions[1].height = 48
            ws_cov.merge_cells("B1:D1")
            c_banner = ws_cov.cell(row=1, column=2, value=f"{meta.get('company_name', 'SUNDRY ITEMS QUOTATION').upper()} — COMMERCIAL OFFER")
            c_banner.font = font_company
            c_banner.fill = fill_hdr
            c_banner.alignment = Alignment(horizontal="center", vertical="center")
            logo_added = True
        except Exception:
            logo_added = False

    if not logo_added:
        ws_cov.merge_cells("A1:D1")
        c_banner = ws_cov.cell(row=1, column=1, value=f"{meta.get('company_name', 'SUNDRY ITEMS QUOTATION').upper()} — COMMERCIAL OFFER")
        c_banner.font = font_company
        c_banner.fill = fill_hdr
        c_banner.alignment = Alignment(horizontal="center", vertical="center")
        ws_cov.row_dimensions[1].height = 36

    details = [
        ("Quotation / Ref No:", f"QTN-{meta['annexure_no']}-Rev{meta['revision']}", "Date:", meta["quote_date"]),
        ("Customer Name:", meta.get("customer", "-"), "Location:", meta.get("location", "-")),
        ("Project Name:", meta.get("project", "-"), "Subject:", "Supply of Sundry Electrical Materials"),
    ]
    for idx, (k1, v1, k2, v2) in enumerate(details, start=3):
        ws_cov.cell(row=idx, column=1, value=k1).font = font_cat
        ws_cov.cell(row=idx, column=2, value=v1).font = font_data
        ws_cov.cell(row=idx, column=3, value=k2).font = font_cat
        ws_cov.cell(row=idx, column=4, value=v2).font = font_data
        for c_i in range(1, 5):
            ws_cov.cell(row=idx, column=c_i).border = border_cell
            if c_i in (1, 3):
                ws_cov.cell(row=idx, column=c_i).fill = fill_light

    ws_ps = wb.create_sheet("Price_Schedule")
    ws_ps.merge_cells("A1:G1")
    c_ps_t = ws_ps.cell(row=1, column=1, value=f"PRICE SCHEDULE — ANNEXURE {meta['annexure_no']} (Rev {meta['revision']})")
    c_ps_t.font = font_sub
    c_ps_t.alignment = Alignment(horizontal="left", vertical="center")
    ws_ps.row_dimensions[1].height = 24

    ps_headers = ["Sr.No", "Material Description", "Make", "QTY", "Unit Rate (INR)", "Discount (%)", "Total Amount (INR)"]
    for c_i, h in enumerate(ps_headers, start=1):
        c = ws_ps.cell(row=2, column=c_i, value=h)
        c.font = font_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell
    ws_ps.row_dimensions[2].height = 24

    current_row = 3
    first_data_row = None
    for cat_name, grp in final_df.groupby("main_material", sort=False):
        ws_ps.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        c_cat = ws_ps.cell(row=current_row, column=1, value=cat_name)
        c_cat.font = font_cat
        c_cat.alignment = Alignment(horizontal="left", vertical="center")
        for c_i in range(1, 8):
            ws_ps.cell(row=current_row, column=c_i).fill = fill_cat
            ws_ps.cell(row=current_row, column=c_i).border = border_cell
        ws_ps.row_dimensions[current_row].height = 20
        current_row += 1

        for sr_no, (_, item) in enumerate(grp.iterrows(), start=1):
            if first_data_row is None:
                first_data_row = current_row
            r = current_row
            ws_ps.cell(row=r, column=1, value=sr_no).alignment = Alignment(horizontal="center", vertical="center")
            ws_ps.cell(row=r, column=2, value=item["sub_material"]).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws_ps.cell(row=r, column=3, value=default_make).alignment = Alignment(horizontal="center", vertical="center")

            c4 = ws_ps.cell(row=r, column=4, value=float(item["qty"]))
            c4.number_format = "#,##0.##"
            c4.alignment = Alignment(horizontal="right", vertical="center")

            c5 = ws_ps.cell(row=r, column=5, value=round(float(item["Amnt"]), 2))
            c5.number_format = "#,##0.00"
            c5.alignment = Alignment(horizontal="right", vertical="center")

            c6 = ws_ps.cell(row=r, column=6, value=float(item["discount"]))
            c6.number_format = "0.##"
            c6.alignment = Alignment(horizontal="right", vertical="center")

            c7 = ws_ps.cell(row=r, column=7, value=f"=D{r}*E{r}*(1-F{r}/100)")
            c7.number_format = "#,##0.00"
            c7.alignment = Alignment(horizontal="right", vertical="center")

            for c_i in range(1, 8):
                ws_ps.cell(row=r, column=c_i).font = font_data
                ws_ps.cell(row=r, column=c_i).border = border_cell
            current_row += 1

    last_data_row = current_row - 1
    ps_total_row = current_row
    ws_ps.cell(row=ps_total_row, column=6, value="SUBTOTAL").font = font_total
    ws_ps.cell(row=ps_total_row, column=6).alignment = Alignment(horizontal="right")
    c_sub = ws_ps.cell(row=ps_total_row, column=7, value=f"=SUM(G{first_data_row}:G{last_data_row})" if first_data_row else 0)
    c_sub.font = font_total
    c_sub.number_format = "#,##0.00"
    for c_i in range(1, 8):
        ws_ps.cell(row=ps_total_row, column=c_i).fill = fill_total
        ws_ps.cell(row=ps_total_row, column=c_i).border = border_total

    for col_i, w in {1: 8, 2: 64, 3: 12, 4: 10, 5: 15, 6: 13, 7: 18}.items():
        ws_ps.column_dimensions[get_column_letter(col_i)].width = w
    ws_ps.freeze_panes = "A3"
    ws_ps.page_setup.orientation = ws_ps.ORIENTATION_LANDSCAPE
    ws_ps.page_setup.fitToWidth = 1
    ws_ps.page_setup.fitToHeight = 0
    ws_ps.sheet_properties.pageSetUpPr.fitToPage = True

    r_c = 8
    ws_cov.merge_cells(start_row=r_c, start_column=1, end_row=r_c, end_column=3)
    ws_cov.cell(row=r_c, column=1, value="COMMERCIAL SUMMARY").font = font_hdr
    ws_cov.cell(row=r_c, column=4, value="AMOUNT (INR)").font = font_hdr
    for c_i in range(1, 5):
        ws_cov.cell(row=r_c, column=c_i).fill = fill_hdr
        ws_cov.cell(row=r_c, column=c_i).alignment = Alignment(horizontal="center")
        ws_cov.cell(row=r_c, column=c_i).border = border_cell

    summary_lines = [
        ("Basic Price of Sundry Items (as per Price_Schedule)", f"=Price_Schedule!G{ps_total_row}", False),
        (f"Packing & Forwarding Charges ({summary['pf_pct']:g}%)", f"=D9*{summary['pf_pct']/100.0}", False),
        ("Sub-Total (Taxable Value)", "=D9+D10", True),
        (f"GST @ {summary['gst_pct']:g}%", f"=D11*{summary['gst_pct']/100.0}", False),
        ("GRAND TOTAL (Inclusive of GST)", "=D11+D12", True),
    ]
    for offset, (lbl, formula, bold_flag) in enumerate(summary_lines, start=1):
        rr = r_c + offset
        ws_cov.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=3)
        ws_cov.cell(row=rr, column=1, value=lbl).font = font_total if bold_flag else font_data
        c_val = ws_cov.cell(row=rr, column=4, value=formula)
        c_val.font = font_total if bold_flag else font_data
        c_val.number_format = "#,##0.00"
        c_val.alignment = Alignment(horizontal="right")
        for c_i in range(1, 5):
            ws_cov.cell(row=rr, column=c_i).border = border_total if offset == 5 else border_cell
            if offset == 5:
                ws_cov.cell(row=rr, column=c_i).fill = fill_total

    r_t = 16
    ws_cov.merge_cells(start_row=r_t, start_column=1, end_row=r_t, end_column=4)
    ws_cov.cell(row=r_t, column=1, value="TERMS & CONDITIONS").font = font_sub
    for line in str(meta.get("terms", DEFAULT_TERMS)).splitlines():
        if line.strip():
            r_t += 1
            ws_cov.merge_cells(start_row=r_t, start_column=1, end_row=r_t, end_column=4)
            ws_cov.cell(row=r_t, column=1, value=line.strip()).font = font_data

    for col_i, w in {1: 28, 2: 38, 3: 18, 4: 24}.items():
        ws_cov.column_dimensions[get_column_letter(col_i)].width = w
    ws_cov.page_setup.fitToWidth = 1
    ws_cov.page_setup.fitToHeight = 1
    ws_cov.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# EXPORT 3: CUSTOMER-FACING PDF (WITH COMPANY LOGO)
# ============================================================
def safe_pdf_text(text: str) -> str:
    return str(text).replace("₹", "Rs.").encode("latin-1", "replace").decode("latin-1")


def _pdf_add_logo_header(pdf: FPDF, title_text: str, logo_bytes: bytes = None):
    if logo_bytes:
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
                tf.write(logo_bytes)
                tmp_path = tf.name
            pdf.image(tmp_path, x=10, y=10, h=14)
            pdf.set_xy(54, 10)
            pdf.set_fill_color(31, 78, 120)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(0, 14, safe_pdf_text(title_text), ln=1, align="C", fill=True)
            pdf.ln(3)
            return
        except Exception:
            pass
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

    pdf.set_fill_color(31, 78, 120)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 12, safe_pdf_text(title_text), ln=1, align="C", fill=True)
    pdf.ln(3)


def build_customer_pdf(final_df: pd.DataFrame, meta: dict, summary: dict, default_make: str, logo_bytes: bytes = None) -> io.BytesIO:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    _pdf_add_logo_header(pdf, f"{meta.get('company_name', 'COMMERCIAL QUOTATION')} - QUOTATION", logo_bytes=logo_bytes)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(32, 6, "Quotation Ref:", border=1)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(68, 6, safe_pdf_text(f"QTN-{meta['annexure_no']}-Rev{meta['revision']}"), border=1)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(25, 6, "Date:", border=1)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(0, 6, safe_pdf_text(meta["quote_date"]), border=1, ln=1)

    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(32, 6, "Customer:", border=1)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(68, 6, safe_pdf_text(meta.get("customer", "-")[:40]), border=1)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(25, 6, "Location:", border=1)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(0, 6, safe_pdf_text(meta.get("location", "-")[:30]), border=1, ln=1)

    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(32, 6, "Project:", border=1)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(0, 6, safe_pdf_text(meta.get("project", "-")[:95]), border=1, ln=1)
    pdf.ln(4)

    widths = [10, 86, 18, 14, 20, 14, 28]
    headers = ["Sr", "Material Description", "Make", "QTY", "Rate (Rs)", "Disc%", "Total (Rs)"]
    pdf.set_fill_color(31, 78, 120)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    for w, h in zip(widths, headers):
        pdf.cell(w, 7, h, border=1, align="C", fill=True)
    pdf.ln()

    pdf.set_text_color(0, 0, 0)
    for cat_name, grp in final_df.groupby("main_material", sort=False):
        pdf.set_fill_color(217, 225, 242)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(sum(widths), 6, safe_pdf_text(cat_name), border=1, ln=1, fill=True)

        pdf.set_font("Helvetica", "", 7.5)
        for sr_no, (_, item) in enumerate(grp.iterrows(), start=1):
            desc = safe_pdf_text(item["sub_material"])
            if len(desc) > 72:
                desc = desc[:69] + "..."
            pdf.cell(widths[0], 6, str(sr_no), border=1, align="C")
            pdf.cell(widths[1], 6, desc, border=1, align="L")
            pdf.cell(widths[2], 6, safe_pdf_text(default_make[:11]), border=1, align="C")
            pdf.cell(widths[3], 6, f"{item['qty']:g}", border=1, align="R")
            pdf.cell(widths[4], 6, f"{item['Amnt']:,.2f}", border=1, align="R")
            pdf.cell(widths[5], 6, f"{item['discount']:g}", border=1, align="R")
            pdf.cell(widths[6], 6, f"{item['Total Amnt']:,.2f}", border=1, align="R")
            pdf.ln()

    pdf.ln(3)
    sum_rows = [
        ("Basic Material Total", f"Rs. {summary['basic_total']:,.2f}", False),
        (f"Packing & Forwarding ({summary['pf_pct']:g}%)", f"Rs. {summary['pf_amount']:,.2f}", False),
        ("Taxable Value", f"Rs. {summary['taxable_total']:,.2f}", True),
        (f"GST ({summary['gst_pct']:g}%)", f"Rs. {summary['gst_amount']:,.2f}", False),
        ("GRAND TOTAL (Inclusive of GST)", f"Rs. {summary['grand_total_incl_gst']:,.2f}", True),
    ]
    for idx_s, (lbl, val, is_bold) in enumerate(sum_rows):
        pdf.set_font("Helvetica", "B" if is_bold else "", 9)
        if idx_s == 4:
            pdf.set_fill_color(226, 239, 218)
        pdf.cell(130, 6.5, safe_pdf_text(lbl), border=1, align="R", fill=(idx_s == 4))
        pdf.cell(60, 6.5, safe_pdf_text(val), border=1, align="R", fill=(idx_s == 4), ln=1)

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 6, "Terms & Conditions:", ln=1)
    pdf.set_font("Helvetica", "", 8)
    for line in str(meta.get("terms", DEFAULT_TERMS)).splitlines():
        if line.strip():
            pdf.multi_cell(0, 4.5, safe_pdf_text(line.strip()))

    raw_pdf = pdf.output() if IS_FPDF2 else pdf.output(dest="S")
    pdf_bytes = raw_pdf.encode("latin-1") if isinstance(raw_pdf, str) else bytes(raw_pdf)
    buf = io.BytesIO(pdf_bytes)
    buf.seek(0)
    return buf


# ============================================================
# EXPORT 4 & 5: REVISION COMPARISON EXCEL & PDF (REV A vs REV B)
# ============================================================
def build_comparison_excel(diff_df: pd.DataFrame, comp_meta: dict) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Revision_Comparison"

    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True, color="000000")
    font_data = Font(name="Calibri", size=10, color="000000")

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_add = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_rem = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_mod = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin = Side(style="thin", color="BFBFBF")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:K1")
    c_t = ws.cell(row=1, column=1, value=f"QUOTATION REVISION COMPARISON: {comp_meta['label_a']}  vs  {comp_meta['label_b']}")
    c_t.font = font_title
    c_t.fill = fill_hdr
    c_t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    sum_rows = [
        ("Customer", comp_meta["customer"], "Baseline Basic Total (INR)", comp_meta["sum_a"]["basic_total"]),
        ("Project", comp_meta["project"], "Revised Basic Total (INR)", comp_meta["sum_b"]["basic_total"]),
        ("Baseline Date", comp_meta["date_a"], "Baseline Grand Total Incl. GST", comp_meta["sum_a"]["grand_total_incl_gst"]),
        ("Revised Date", comp_meta["date_b"], "Revised Grand Total Incl. GST", comp_meta["sum_b"]["grand_total_incl_gst"]),
        (
            "Changes Summary",
            f"Added: {comp_meta['counts']['Added']} | Removed: {comp_meta['counts']['Removed']} | Modified: {comp_meta['counts']['Modified']}",
            "Net Variance (INR / %)",
            f"Rs. {comp_meta['grand_delta']:+,.2f} ({comp_meta['pct_delta']:+.1f}%)",
        ),
    ]
    for r_i, (k1, v1, k2, v2) in enumerate(sum_rows, start=3):
        ws.cell(row=r_i, column=1, value=k1).font = font_bold
        ws.merge_cells(start_row=r_i, start_column=2, end_row=r_i, end_column=6)
        ws.cell(row=r_i, column=2, value=v1).font = font_data
        ws.merge_cells(start_row=r_i, start_column=7, end_row=r_i, end_column=9)
        ws.cell(row=r_i, column=7, value=k2).font = font_bold
        ws.merge_cells(start_row=r_i, start_column=10, end_row=r_i, end_column=11)
        c_v2 = ws.cell(row=r_i, column=10, value=v2)
        c_v2.font = font_bold
        if isinstance(v2, (int, float)):
            c_v2.number_format = "#,##0.00"
            c_v2.alignment = Alignment(horizontal="right")

    hdr_row = 9
    cols = list(diff_df.columns)
    for c_i, col_name in enumerate(cols, start=1):
        c = ws.cell(row=hdr_row, column=c_i, value=col_name)
        c.font = font_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell
    ws.row_dimensions[hdr_row].height = 24

    for idx, (_, row) in enumerate(diff_df.iterrows(), start=hdr_row + 1):
        st_val = row["Status"]
        row_fill = fill_add if st_val == "Added" else (fill_rem if st_val == "Removed" else (fill_mod if st_val == "Modified" else None))
        for c_i, col_name in enumerate(cols, start=1):
            val = row[col_name]
            cell = ws.cell(row=idx, column=c_i, value=val)
            cell.font = font_bold if col_name in ("Status", "Delta (INR)") else font_data
            cell.border = border_cell
            if row_fill is not None:
                cell.fill = row_fill
            if col_name in ("Old Net Rate", "New Net Rate", "Old Total", "New Total", "Delta (INR)"):
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in ("Old QTY", "New QTY"):
                cell.number_format = "#,##0.##"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name == "Status":
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    widths = {1: 12, 2: 26, 3: 46, 4: 10, 5: 10, 6: 13, 7: 13, 8: 14, 9: 14, 10: 14, 11: 32}
    for c_i, w in widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w
    ws.freeze_panes = "A10"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_comparison_pdf(diff_df: pd.DataFrame, comp_meta: dict, logo_bytes: bytes = None) -> io.BytesIO:
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    _pdf_add_logo_header(
        pdf,
        f"REVISION COMPARISON: {comp_meta['label_a']} vs {comp_meta['label_b']}",
        logo_bytes=logo_bytes,
    )

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(135, 6, safe_pdf_text(f"Customer: {comp_meta['customer']} | Project: {comp_meta['project'][:50]}"), border=1)
    pdf.cell(
        0,
        6,
        safe_pdf_text(
            f"Grand Total Delta: Rs. {comp_meta['grand_delta']:+,.2f} ({comp_meta['pct_delta']:+.1f}%) | "
            f"+{comp_meta['counts']['Added']} Add / -{comp_meta['counts']['Removed']} Rem / {comp_meta['counts']['Modified']} Mod"
        ),
        border=1,
        ln=1,
    )
    pdf.ln(3)

    widths = [18, 92, 16, 16, 24, 24, 24, 24, 39]
    headers = ["Status", "Sub Material", "Old Q", "New Q", "Old Rate", "New Rate", "Old Total", "New Total", "Delta / Notes"]
    pdf.set_fill_color(31, 78, 120)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 7.5)
    for w, h in zip(widths, headers):
        pdf.cell(w, 7, h, border=1, align="C", fill=True)
    pdf.ln()

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 7)
    for _, r in diff_df.iterrows():
        st_val = r["Status"]
        if st_val == "Added":
            pdf.set_fill_color(226, 239, 218)
            fill = True
        elif st_val == "Removed":
            pdf.set_fill_color(255, 199, 206)
            fill = True
        elif st_val == "Modified":
            pdf.set_fill_color(255, 242, 204)
            fill = True
        else:
            fill = False

        desc = safe_pdf_text(r["Sub Material"])
        if len(desc) > 60:
            desc = desc[:57] + "..."
        note = safe_pdf_text(f"{r['Delta (INR)']:+,.0f} | {r['Change Notes']}")
        if len(note) > 28:
            note = note[:25] + "..."

        pdf.cell(widths[0], 5.5, st_val, border=1, align="C", fill=fill)
        pdf.cell(widths[1], 5.5, desc, border=1, align="L", fill=fill)
        pdf.cell(widths[2], 5.5, f"{r['Old QTY']:g}", border=1, align="R", fill=fill)
        pdf.cell(widths[3], 5.5, f"{r['New QTY']:g}", border=1, align="R", fill=fill)
        pdf.cell(widths[4], 5.5, f"{r['Old Net Rate']:,.2f}", border=1, align="R", fill=fill)
        pdf.cell(widths[5], 5.5, f"{r['New Net Rate']:,.2f}", border=1, align="R", fill=fill)
        pdf.cell(widths[6], 5.5, f"{r['Old Total']:,.2f}", border=1, align="R", fill=fill)
        pdf.cell(widths[7], 5.5, f"{r['New Total']:,.2f}", border=1, align="R", fill=fill)
        pdf.cell(widths[8], 5.5, note, border=1, align="L", fill=fill)
        pdf.ln()

    raw_pdf = pdf.output() if IS_FPDF2 else pdf.output(dest="S")
    pdf_bytes = raw_pdf.encode("latin-1") if isinstance(raw_pdf, str) else bytes(raw_pdf)
    buf = io.BytesIO(pdf_bytes)
    buf.seek(0)
    return buf


# ============================================================
# MAIN BOOTSTRAP & HEADER
# ============================================================
bootstrap_from_cloud_once()

master_df = load_mastersheet()
alias_df = load_aliases()
history_list = load_history()
cloud_cfg = load_cloud_config()
logo_bytes = load_logo_bytes()

cloud_badge = "☁️ Google Sheets Auto-Sync ON" if (cloud_cfg.get("enabled") and cloud_cfg.get("auto_sync")) else "💻 Local Offline Mode"

st.markdown(
    f"""
    <div class="main-header">
        <h1>📊 SundryQuote Pro — Enterprise BOM Quotation System</h1>
        <p>Alias Memory • Margin Loss Shield • Branded Logo Excel & PDF + GST • Rev-vs-Rev Comparison • {cloud_badge}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

tab_quote, tab_master, tab_alias, tab_history, tab_cloud = st.tabs(
    [
        "🚀 Generate Quotation",
        f"🗄️ Mastersheet ({len(master_df)})",
        f"🧠 Alias Memory ({len(alias_df)})",
        f"📜 History & Compare ({len(history_list)})",
        "☁️ Cloud Sync & Secrets",
    ]
)


# ============================================================
# TAB 1: QUOTATION GENERATOR (BOM UPLOAD OR LOADED REVISION)
# ============================================================
with tab_quote:
    loaded_rev = st.session_state.get("loaded_history_quote")
    if loaded_rev:
        st.info(
            f"📂 **Revision Mode Active:** Loaded `Annexure {loaded_rev['annexure_no']} (Rev {loaded_rev['revision']})` "
            f"for **{loaded_rev.get('customer', 'N/A')}**. Edit rates/commercials below and save as the next revision."
        )
        if st.button("❌ Exit Revision Mode & Upload New BOM"):
            del st.session_state["loaded_history_quote"]
            st.rerun()

    if master_df.empty and not loaded_rev:
        st.warning("Your Mastersheet is empty. Add or import items in the **Mastersheet** tab first.")
    else:
        all_included = []
        excluded_rows = []
        auto_meta = {"project": "", "customer": "", "location": "", "job_no": ""}

        if loaded_rev:
            auto_meta = {
                "project": loaded_rev.get("project", ""),
                "customer": loaded_rev.get("customer", ""),
                "location": loaded_rev.get("location", ""),
                "job_no": loaded_rev.get("annexure_no", ""),
            }
            all_included = loaded_rev.get("items", [])
            combine_dupes = False
        else:
            col_left, col_right = st.columns([2, 1.15])
            with col_left:
                uploaded_boms = st.file_uploader(
                    "Upload one or more BOM files (Excel .xlsx/.xlsm/.xls, CSV, PDF, or Word .docx)",
                    type=["xlsx", "xlsm", "xls", "csv", "tsv", "txt", "pdf", "docx"],
                    accept_multiple_files=True,
                )
            with col_right:
                ignore_hidden = st.checkbox("Ignore hidden Excel rows & sheets", value=True)
                use_alias_memory = st.checkbox("Use Auto-Learning Alias Memory", value=True)
                c_o1, c_o2 = st.columns(2)
                with c_o1:
                    combine_dupes = st.checkbox("Combine duplicate items", value=True)
                with c_o2:
                    skip_zero_qty = st.checkbox("Skip 0 / NA qty", value=True)

            if uploaded_boms:
                all_sources = {}
                total_hidden_rows = 0
                total_hidden_sheets = 0

                for f in uploaded_boms:
                    parsed_map, h_rows, h_sheets = read_any_bom_file(f, ignore_hidden=ignore_hidden)
                    total_hidden_rows += h_rows
                    total_hidden_sheets += h_sheets
                    for label, (df_p, meta_p) in parsed_map.items():
                        all_sources[label] = (df_p, meta_p)
                        for mk in ("job_no", "customer", "project", "location"):
                            if not auto_meta[mk] and meta_p.get(mk):
                                auto_meta[mk] = meta_p[mk]

                if ignore_hidden and (total_hidden_rows > 0 or total_hidden_sheets > 0):
                    st.success(f"🧹 Automatically skipped **{total_hidden_rows} hidden row(s)** and **{total_hidden_sheets} hidden sheet(s)**.")

                if not all_sources:
                    st.error("Could not detect a BOM table with Description and QTY columns in the uploaded file(s).")
                else:
                    selected_sources = st.multiselect(
                        "Detected BOM Sheets / Files:",
                        options=list(all_sources.keys()),
                        default=list(all_sources.keys()),
                    )

                    if selected_sources:
                        bom_raw_combined = pd.concat([all_sources[s][0] for s in selected_sources], ignore_index=True)
                        zero_qty_count = int((bom_raw_combined["qty"] <= 0).sum())
                        bom_combined = bom_raw_combined[bom_raw_combined["qty"] > 0].reset_index(drop=True) if skip_zero_qty else bom_raw_combined.copy()
                        if skip_zero_qty and zero_qty_count > 0:
                            st.caption(f"ℹ️ Filtered out **{zero_qty_count}** row(s) with `0`, blank, or `NA` quantity.")

                        if combine_dupes and not bom_combined.empty:
                            bom_combined["_strict_key"] = bom_combined["sub_material"].apply(strict_alnum_key)
                            bom_combined = (
                                bom_combined.groupby("_strict_key", sort=False, as_index=False)
                                .agg(
                                    main_material=("main_material", "first"),
                                    sub_material=("sub_material", "first"),
                                    qty=("qty", "sum"),
                                    source=("source", lambda x: " | ".join(dict.fromkeys(x))),
                                )
                                .drop(columns=["_strict_key"])
                            )

                        master_indexed = master_df.copy()
                        master_indexed["_strict_key"] = master_indexed["sub_material"].apply(strict_alnum_key)
                        master_lookup = (
                            master_indexed.drop_duplicates(subset=["_strict_key"], keep="first")
                            .set_index("_strict_key")
                            .to_dict(orient="index")
                        )
                        alias_lookup = (
                            dict(zip(alias_df["bom_strict_key"], alias_df["master_strict_key"]))
                            if (use_alias_memory and not alias_df.empty)
                            else {}
                        )

                        exact_rows, alias_rows, unmatched_rows = [], [], []

                        for idx, row in bom_combined.iterrows():
                            k = strict_alnum_key(row["sub_material"])
                            if k in master_lookup:
                                m = master_lookup[k]
                                exact_rows.append(
                                    {
                                        "bom_main": row["main_material"],
                                        "bom_sub": row["sub_material"],
                                        "main_material": m["main_material"],
                                        "sub_material": m["sub_material"],
                                        "qty": float(row["qty"]),
                                        "purchase_price": float(m["purchase_price"]),
                                        "discount": float(m["discount"]),
                                        "match_status": "Exact Auto-Match",
                                        "source": row["source"],
                                    }
                                )
                            elif k in alias_lookup and alias_lookup[k] in master_lookup:
                                m = master_lookup[alias_lookup[k]]
                                alias_rows.append(
                                    {
                                        "bom_main": row["main_material"],
                                        "bom_sub": row["sub_material"],
                                        "main_material": m["main_material"],
                                        "sub_material": m["sub_material"],
                                        "qty": float(row["qty"]),
                                        "purchase_price": float(m["purchase_price"]),
                                        "discount": float(m["discount"]),
                                        "match_status": "Learned Alias Auto-Match",
                                        "source": row["source"],
                                    }
                                )
                            else:
                                suggestions = get_top_close_matches(row["sub_material"], row["main_material"], master_df, top_n=5)
                                unmatched_rows.append(
                                    {
                                        "uid": f"{idx}_{k[:20]}",
                                        "bom_main": row["main_material"],
                                        "bom_sub": row["sub_material"],
                                        "qty": float(row["qty"]),
                                        "source": row["source"],
                                        "suggestions": suggestions,
                                    }
                                )

                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Active BOM Items", len(bom_combined))
                        m2.metric("✅ Exact Auto-Matched", len(exact_rows))
                        m3.metric("🧠 Alias Auto-Matched", len(alias_rows))
                        m4.metric("⚠️ Needs Review", len(unmatched_rows))

                        approved_close_rows = []
                        if unmatched_rows:
                            st.subheader("🔍 Review Unmatched BOM Items (Close-Match Approval)")
                            b1, b2, _ = st.columns([2.1, 1.1, 1.8])
                            with b1:
                                if st.button("✅ Accept Safe High-Confidence (≥ 80% & No Size Mismatch)"):
                                    for u in unmatched_rows:
                                        if u["suggestions"] and u["suggestions"][0]["score"] >= 80.0 and not u["suggestions"][0]["num_mismatch"]:
                                            st.session_state[f"yn_{u['uid']}"] = "Yes (Include)"
                            with b2:
                                if st.button("❌ Reject All"):
                                    for u in unmatched_rows:
                                        st.session_state[f"yn_{u['uid']}"] = "No (Exclude)"

                            all_master_records = master_df.to_dict(orient="records")

                            for i, u in enumerate(unmatched_rows, start=1):
                                top_sugg = u["suggestions"]
                                with st.container(border=True):
                                    c_top_l, c_top_r = st.columns([3.7, 1.3])
                                    with c_top_l:
                                        st.markdown(f'<div class="bom-name-box">{i}. BOM Item: {u["bom_sub"]}</div>', unsafe_allow_html=True)
                                        st.caption(f"Category: **{u['bom_main']}**  •  QTY: **{u['qty']:g}**  •  Source: `{u['source']}`")
                                    with c_top_r:
                                        decision = st.radio(
                                            "Decision",
                                            options=["No (Exclude)", "Yes (Include)"],
                                            key=f"yn_{u['uid']}",
                                            label_visibility="collapsed",
                                        )

                                    search_all = st.checkbox("Search all Mastersheet items instead of Top 5", key=f"all_{u['uid']}")

                                    if search_all:
                                        full_labels = [
                                            f"{r['sub_material']}  [Group: {r['main_material']} | ₹{r['purchase_price']:,.2f} | Disc {r['discount']:g}%]"
                                            for r in all_master_records
                                        ]
                                        chosen_idx = st.selectbox(
                                            "Select from full Mastersheet:",
                                            options=range(len(full_labels)),
                                            format_func=lambda x, lbls=full_labels: lbls[x],
                                            key=f"sel_all_{u['uid']}",
                                        )
                                        rec = all_master_records[chosen_idx]
                                        bom_nums = extract_numbers(u["bom_sub"])
                                        m_nums = extract_numbers(rec["sub_material"])
                                        chosen_match = {
                                            "main_material": rec["main_material"],
                                            "sub_material": rec["sub_material"],
                                            "purchase_price": float(rec["purchase_price"]),
                                            "discount": float(rec["discount"]),
                                            "score": 100.0,
                                            "num_mismatch": bom_nums != m_nums,
                                            "bom_only_nums": sorted(bom_nums - m_nums),
                                            "master_only_nums": sorted(m_nums - bom_nums),
                                        }
                                    else:
                                        option_labels = [
                                            f"[{s['score']}%{' ⚠️ SIZE DIFF' if s['num_mismatch'] else ''}] {s['sub_material']}  (₹{s['purchase_price']:,.2f} | Disc {s['discount']:g}%)"
                                            for s in top_sugg
                                        ]
                                        selected_opt = st.selectbox(
                                            "Suggested Close Match:",
                                            options=range(len(option_labels)),
                                            format_func=lambda x, lbls=option_labels: lbls[x],
                                            key=f"sel_{u['uid']}",
                                        )
                                        chosen_match = top_sugg[selected_opt]

                                    st.markdown(
                                        f'<div class="master-name-box">'
                                        f'<b>Selected Mastersheet Item ({chosen_match["score"]}% match):</b><br>'
                                        f'{chosen_match["sub_material"]}<br>'
                                        f'<span style="font-size:0.85rem;color:#4A5568;">'
                                        f'Group: <b>{chosen_match["main_material"]}</b> &nbsp;|&nbsp; '
                                        f'Purchase: <b>₹{chosen_match["purchase_price"]:,.2f}</b> &nbsp;|&nbsp; '
                                        f'Discount: <b>{chosen_match["discount"]:g}%</b></span>'
                                        f'</div>',
                                        unsafe_allow_html=True,
                                    )

                                    if chosen_match["num_mismatch"]:
                                        b_diff = ", ".join(chosen_match["bom_only_nums"]) or "None"
                                        m_diff = ", ".join(chosen_match["master_only_nums"]) or "None"
                                        st.markdown(
                                            f'<div class="warn-badge">⚠️ Size / Number Mismatch — BOM has [{b_diff}] vs Mastersheet has [{m_diff}]</div>',
                                            unsafe_allow_html=True,
                                        )
                                    else:
                                        st.markdown('<div class="safe-badge">✅ All numbers & sizes match</div>', unsafe_allow_html=True)

                                    with st.expander("➕ Not in Mastersheet? Add this exact BOM item to Mastersheet"):
                                        ac1, ac2, ac3, ac4 = st.columns([2, 1, 1, 1.2])
                                        with ac1:
                                            add_main = st.text_input("Main Material", value=u["bom_main"], key=f"am_{u['uid']}")
                                        with ac2:
                                            add_price = st.number_input("Purchase ₹", min_value=0.0, value=0.0, step=1.0, key=f"ap_{u['uid']}")
                                        with ac3:
                                            add_disc = st.number_input("Discount %", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key=f"ad_{u['uid']}")
                                        with ac4:
                                            st.write("")
                                            if st.button("Save to Master", key=f"ab_{u['uid']}"):
                                                new_row = pd.DataFrame(
                                                    [{"main_material": add_main.strip() or "GENERAL", "sub_material": u["bom_sub"], "purchase_price": float(add_price), "discount": float(add_disc)}]
                                                )
                                                save_mastersheet(pd.concat([master_df, new_row], ignore_index=True))
                                                st.success("Added! Auto-matching now...")
                                                st.rerun()

                                    if decision == "Yes (Include)":
                                        approved_close_rows.append(
                                            {
                                                "bom_main": u["bom_main"],
                                                "bom_sub": u["bom_sub"],
                                                "main_material": chosen_match["main_material"],
                                                "sub_material": chosen_match["sub_material"],
                                                "qty": u["qty"],
                                                "purchase_price": chosen_match["purchase_price"],
                                                "discount": chosen_match["discount"],
                                                "match_status": f"User Approved Close Match ({chosen_match['score']}%)",
                                                "source": u["source"],
                                            }
                                        )
                                    else:
                                        excluded_rows.append(
                                            {
                                                "BOM Main Material": u["bom_main"],
                                                "BOM Sub Material": u["bom_sub"],
                                                "QTY": u["qty"],
                                                "Source": u["source"],
                                                "Matched Mastersheet Item": top_sugg[0]["sub_material"] if top_sugg else "",
                                                "Status": "Excluded by User",
                                            }
                                        )

                            if approved_close_rows:
                                if st.button(f"🧠 Remember {len(approved_close_rows)} Approved Match(es) Forever in Alias Memory", type="secondary"):
                                    count_saved = learn_aliases([{"bom_sub": r["bom_sub"], "master_sub": r["sub_material"]} for r in approved_close_rows])
                                    st.success(f"✅ Saved {count_saved} alias rule(s)! Future BOMs will auto-match these automatically.")
                                    st.rerun()

                        all_included = exact_rows + alias_rows + approved_close_rows

        # ============================================================
        # COMMERCIAL SETTINGS, LOGO UPLOAD, MARGIN ALERTS & EXPORTS
        # ============================================================
        if all_included:
            st.divider()
            st.subheader("🏢 Commercial Cover Details, Company Logo & Pricing")

            # Company Logo Uploader Bar
            with st.expander("🖼️ Company Logo Branding (Embedded on Customer Excel Cover Sheet & PDF)", expanded=(logo_bytes is None)):
                lc1, lc2, lc3 = st.columns([2.2, 1.5, 1.0])
                with lc1:
                    up_logo = st.file_uploader("Upload Company Logo (PNG / JPG)", type=["png", "jpg", "jpeg"], key="logo_uploader")
                    if up_logo is not None:
                        logo_bytes = save_logo_bytes(up_logo.getvalue())
                        st.success("✅ Logo saved! It will appear on your Customer Excel & PDF offers.")
                with lc2:
                    if logo_bytes:
                        st.image(logo_bytes, caption="Active Logo Preview", width=180)
                    else:
                        st.caption("No logo uploaded yet. Exports will use a clean text banner.")
                with lc3:
                    if logo_bytes and st.button("🗑️ Remove Logo"):
                        clear_logo()
                        logo_bytes = None
                        st.rerun()

            c_m1, c_m2, c_m3, c_m4 = st.columns([1.4, 1.2, 0.8, 1.0])
            with c_m1:
                company_name = st.text_input("Your Company Name", value=(loaded_rev.get("company_name") if loaded_rev else "Your Company Name Pvt. Ltd."))
            with c_m2:
                annexure_no = st.text_input("Annexure / Job Number", value=(auto_meta["job_no"] or "6500389"))
            with c_m3:
                suggested_rev = get_next_revision(annexure_no, history_list)
                revision_no = st.number_input("Revision No", min_value=0, max_value=99, value=int(suggested_rev), step=1)
            with c_m4:
                quote_date = st.text_input("Quotation Date", value=datetime.now().strftime("%d-%m-%Y"))

            c_p1, c_p2, c_p3 = st.columns([1.5, 2.0, 1.0])
            with c_p1:
                customer_name = st.text_input("Customer Name", value=auto_meta["customer"])
            with c_p2:
                project_name = st.text_input("Project Name", value=auto_meta["project"])
            with c_p3:
                location_name = st.text_input("Location", value=auto_meta["location"] or "India")

            c_r1, c_r2, c_r3, c_r4 = st.columns(4)
            with c_r1:
                profit_pct = st.number_input("Profit % (Cell L2)", min_value=0.0, max_value=500.0, value=float(loaded_rev.get("profit_pct", 12.0) if loaded_rev else 12.0), step=0.5)
            with c_r2:
                pf_pct = st.number_input("Packing & Forwarding %", min_value=0.0, max_value=100.0, value=float(loaded_rev.get("pf_pct", 0.0) if loaded_rev else 0.0), step=0.5)
            with c_r3:
                gst_pct = st.number_input("GST %", min_value=0.0, max_value=50.0, value=float(loaded_rev.get("gst_pct", 18.0) if loaded_rev else 18.0), step=1.0)
            with c_r4:
                default_make = st.text_input("Default Make", value=str(loaded_rev.get("default_make", "Reputed") if loaded_rev else "Reputed"))

            with st.expander("📝 Edit Terms & Conditions (Printed on Customer Cover Sheet & PDF)"):
                terms_text = st.text_area("Terms & Conditions", value=str(loaded_rev.get("terms", DEFAULT_TERMS) if loaded_rev else DEFAULT_TERMS), height=130)

            base_df = pd.DataFrame(all_included)
            if combine_dupes and not loaded_rev:
                base_df = (
                    base_df.groupby(["main_material", "sub_material"], sort=False, as_index=False)
                    .agg(
                        qty=("qty", "sum"),
                        purchase_price=("purchase_price", "first"),
                        discount=("discount", "first"),
                        match_status=("match_status", lambda x: " + ".join(dict.fromkeys(x))),
                    )
                )

            st.subheader("📑 Editable Quotation & Live Profit / Loss Analysis")
            st.caption("Edit **QTY**, **Purchase ₹**, or **Discount %** below. Margin alerts recalculate instantly.")

            edited_items = st.data_editor(
                base_df[["main_material", "sub_material", "qty", "purchase_price", "discount"]],
                use_container_width=True,
                hide_index=True,
                disabled=["main_material", "sub_material"],
                column_config={
                    "main_material": st.column_config.TextColumn("Main Material"),
                    "sub_material": st.column_config.TextColumn("Sub Material", width="large"),
                    "qty": st.column_config.NumberColumn("QTY", min_value=0.0, format="%.2f"),
                    "purchase_price": st.column_config.NumberColumn("Purchase ₹", min_value=0.0, format="%.2f"),
                    "discount": st.column_config.NumberColumn("Discount %", min_value=0.0, max_value=100.0, format="%.2f"),
                },
                key="quote_rate_editor",
            )

            calc_df, summary = compute_commercials(edited_items, profit_pct=profit_pct, pf_pct=pf_pct, gst_pct=gst_pct)

            if summary["loss_item_count"] > 0:
                st.error(
                    f"🚨 **NEGATIVE MARGIN ALERT:** **{summary['loss_item_count']} item(s)** are selling **BELOW Purchase Cost** "
                    f"(combined loss on those items: **₹{summary['loss_total_amount']:,.2f}**). "
                    f"Check the red rows below or adjust Discount / Profit %."
                )

            if summary["zero_rate_count"] > 0:
                st.warning(f"⚠️ **{summary['zero_rate_count']} item(s)** currently have a **₹0.00 Purchase Price**. Edit them directly in the table above.")

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Basic Material Total", f"₹ {summary['basic_total']:,.2f}")
            k2.metric(f"Grand Total (Incl. {gst_pct:g}% GST)", f"₹ {summary['grand_total_incl_gst']:,.2f}")
            k3.metric("Total Purchase Cost", f"₹ {summary['total_purchase']:,.2f}")
            k4.metric("Net Profit / (Loss)", f"₹ {summary['net_profit_total']:,.2f}", delta=f"{summary['overall_margin_pct']:.1f}% margin")
            k5.metric("Loss-Making Items", summary["loss_item_count"])

            with st.expander("🔎 View Full Profit/Loss Margin Breakdown per Item", expanded=(summary["loss_item_count"] > 0)):
                view_cols = ["Loss Alert", "main_material", "sub_material", "qty", "purchase_price", "Amnt", "discount", "Net Unit Rate", "Total Amnt", "Net Profit", "Margin %"]
                st.dataframe(calc_df[view_cols], use_container_width=True, hide_index=True)

            act1, act2 = st.columns(2)
            with act1:
                if st.button("💾 Sync Edited Purchase Prices & Discounts Back to Mastersheet", use_container_width=True):
                    updated_master = master_df.copy()
                    updated_master["_k"] = updated_master["sub_material"].apply(strict_alnum_key)
                    for _, er in edited_items.iterrows():
                        ek = strict_alnum_key(er["sub_material"])
                        mask = updated_master["_k"] == ek
                        if mask.any():
                            updated_master.loc[mask, "purchase_price"] = float(er["purchase_price"])
                            updated_master.loc[mask, "discount"] = float(er["discount"])
                    save_mastersheet(updated_master.drop(columns=["_k"]))
                    st.success("✅ Mastersheet updated with edited rates!")
                    st.rerun()

            quote_meta = {
                "company_name": company_name.strip(),
                "annexure_no": annexure_no.strip(),
                "revision": int(revision_no),
                "quote_date": quote_date.strip(),
                "customer": customer_name.strip(),
                "project": project_name.strip(),
                "location": location_name.strip(),
                "terms": terms_text,
            }

            with act2:
                if st.button(f"📌 Save to Quotation History as Rev {int(revision_no)}", type="primary", use_container_width=True):
                    hist = load_history()
                    snapshot = {
                        "id": f"{quote_meta['annexure_no']}_Rev{quote_meta['revision']}_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "company_name": quote_meta["company_name"],
                        "annexure_no": quote_meta["annexure_no"],
                        "revision": quote_meta["revision"],
                        "quote_date": quote_meta["quote_date"],
                        "customer": quote_meta["customer"],
                        "project": quote_meta["project"],
                        "location": quote_meta["location"],
                        "profit_pct": float(profit_pct),
                        "pf_pct": float(pf_pct),
                        "gst_pct": float(gst_pct),
                        "default_make": default_make,
                        "terms": terms_text,
                        "basic_total": summary["basic_total"],
                        "grand_total": summary["grand_total_incl_gst"],
                        "net_profit": summary["net_profit_total"],
                        "item_count": len(edited_items),
                        "items": edited_items.to_dict(orient="records"),
                    }
                    hist.insert(0, snapshot)
                    save_history(hist)
                    st.success(f"✅ Saved `Annexure {quote_meta['annexure_no']} (Rev {quote_meta['revision']})` to Quotation History!")
                    st.rerun()

            audit_records = [
                {
                    "BOM Main Material": r.get("bom_main", r.get("main_material", "")),
                    "BOM Sub Material": r.get("bom_sub", r.get("sub_material", "")),
                    "QTY": r.get("qty", 0),
                    "Source": r.get("source", "History Revision"),
                    "Matched Mastersheet Item": r.get("sub_material", ""),
                    "Status": r.get("match_status", "Loaded from History"),
                }
                for r in all_included
            ] + excluded_rows
            audit_df = pd.DataFrame(audit_records)

            st.markdown("### 📥 Download Quotations")
            internal_xl = build_internal_quotation_excel(calc_df, quote_meta, summary, profit_pct, default_make, audit_df)
            customer_xl = build_customer_quotation_excel(calc_df, quote_meta, summary, default_make, logo_bytes=logo_bytes)

            d1, d2, d3 = st.columns(3)
            with d1:
                st.download_button(
                    label=f"⬇️ Internal Costing Excel (Rev {int(revision_no)})",
                    data=internal_xl,
                    file_name=f"quotation_{quote_meta['annexure_no']}_Rev{int(revision_no)}_INTERNAL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            with d2:
                st.download_button(
                    label="⬇️ Customer Excel (Branded + GST)",
                    data=customer_xl,
                    file_name=f"quotation_{quote_meta['annexure_no']}_Rev{int(revision_no)}_CUSTOMER.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with d3:
                if HAS_FPDF:
                    customer_pdf = build_customer_pdf(calc_df, quote_meta, summary, default_make, logo_bytes=logo_bytes)
                    st.download_button(
                        label="⬇️ Customer PDF Offer (Branded + GST)",
                        data=customer_pdf,
                        file_name=f"quotation_{quote_meta['annexure_no']}_Rev{int(revision_no)}_CUSTOMER.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                else:
                    st.info("Run `pip install fpdf2` to enable Customer PDF export.")


# ============================================================
# TAB 2: MASTERSHEET DATABASE
# ============================================================
with tab_master:
    st.subheader("🗄️ Mastersheet Rate & Discount Database")
    st.caption(f"Total items: **{len(master_df)}** across **{master_df['main_material'].nunique()}** Main Material groups.")

    m_tab_edit, m_tab_add, m_tab_io = st.tabs(["✏️ View, Edit & Delete", "➕ Add New Item", "📤 Import / Export Backup"])

    with m_tab_edit:
        search_q = st.text_input("🔎 Quick Search (clear search to edit/delete rows):", placeholder="Type e.g. M20, LUGS, CABLE TIE...")
        if search_q.strip():
            mask = (
                master_df["main_material"].str.contains(search_q, case=False, na=False)
                | master_df["sub_material"].str.contains(search_q, case=False, na=False)
            )
            st.info(f"Showing {mask.sum()} matching item(s). Clear the search box to edit or delete rows.")
            st.dataframe(master_df[mask], use_container_width=True, hide_index=True)
        else:
            edited_df = st.data_editor(
                master_df,
                num_rows="dynamic",
                use_container_width=True,
                key="master_editor",
                column_config={
                    "main_material": st.column_config.TextColumn("Main Material", required=True),
                    "sub_material": st.column_config.TextColumn("Sub Material", required=True, width="large"),
                    "purchase_price": st.column_config.NumberColumn("Purchase Price", min_value=0.0, format="%.2f"),
                    "discount": st.column_config.NumberColumn("Discount (%)", min_value=0.0, max_value=100.0, format="%.2f"),
                },
            )
            if st.button("💾 Save All Changes to Mastersheet", type="primary"):
                save_mastersheet(edited_df)
                st.success("✅ Mastersheet saved successfully!")
                st.rerun()

    with m_tab_add:
        existing_mains = sorted([m for m in master_df["main_material"].unique().tolist() if m])
        with st.form("add_item_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                chosen_main = st.selectbox("Select Existing Main Material", options=["-- Create New Main Material --"] + existing_mains)
            with c2:
                new_main_text = st.text_input("Or Type New Main Material Name")
            new_sub = st.text_input("Sub Material Description *")
            c3, c4 = st.columns(2)
            with c3:
                new_price = st.number_input("Purchase Price", min_value=0.0, value=0.0, step=1.0)
            with c4:
                new_disc = st.number_input("Discount (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
            if st.form_submit_button("➕ Add to Mastersheet", type="primary"):
                final_main = new_main_text.strip() if new_main_text.strip() else ("" if chosen_main.startswith("--") else chosen_main)
                if not final_main or not new_sub.strip():
                    st.error("Please provide both a **Main Material** and a **Sub Material**.")
                else:
                    new_row = pd.DataFrame([{"main_material": final_main, "sub_material": new_sub.strip(), "purchase_price": float(new_price), "discount": float(new_disc)}])
                    save_mastersheet(pd.concat([master_df, new_row], ignore_index=True))
                    st.success(f"✅ Added `{new_sub.strip()}` under `{final_main}`!")
                    st.rerun()

    with m_tab_io:
        up_master = st.file_uploader("Upload Mastersheet File (CSV / Excel)", type=["csv", "xlsx", "xls"], key="master_upload")
        import_mode = st.radio("Import Mode:", ["Merge (Add new & update existing)", "Replace Entire Mastersheet"], horizontal=True)
        if up_master and st.button("📥 Import Mastersheet Now", type="primary"):
            try:
                in_df = pd.read_csv(up_master, encoding="utf-8-sig") if up_master.name.lower().endswith(".csv") else pd.read_excel(up_master)
                in_df = clean_master_df(in_df)
                if import_mode.startswith("Replace"):
                    save_mastersheet(in_df)
                    st.success(f"✅ Replaced Mastersheet with {len(in_df)} items!")
                else:
                    combined = pd.concat([master_df, in_df], ignore_index=True)
                    combined["_k"] = combined["sub_material"].apply(strict_alnum_key)
                    combined = combined.drop_duplicates(subset=["_k"], keep="last").drop(columns=["_k"])
                    save_mastersheet(combined)
                    st.success(f"✅ Merged! Mastersheet now has {len(combined)} items.")
                st.rerun()
            except Exception as e:
                st.error(f"Import failed: {e}")

        st.divider()
        e1, e2 = st.columns(2)
        with e1:
            st.download_button("⬇️ Export CSV Backup", data=master_df.to_csv(index=False).encode("utf-8-sig"), file_name="mastersheet_backup.csv", mime="text/csv", use_container_width=True)
        with e2:
            xl_buf = io.BytesIO()
            with pd.ExcelWriter(xl_buf, engine="openpyxl") as writer:
                master_df.to_excel(writer, index=False, sheet_name="Mastersheet")
            xl_buf.seek(0)
            st.download_button("⬇️ Export Excel Backup", data=xl_buf, file_name="mastersheet_backup.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)


# ============================================================
# TAB 3: AUTO-LEARNING ALIAS MEMORY
# ============================================================
with tab_alias:
    st.subheader("🧠 Auto-Learning Alias Memory")
    st.caption("Learned mappings auto-match future BOMs using your live Mastersheet rates.")

    if alias_df.empty:
        st.info("No learned aliases saved yet. Approve close matches in **Generate Quotation** and click **Remember Approved Matches** to start building memory.")
    else:
        edited_alias = st.data_editor(
            alias_df[["bom_sample_name", "master_sub_material", "created_at"]],
            num_rows="dynamic",
            use_container_width=True,
            disabled=["created_at"],
            column_config={
                "bom_sample_name": st.column_config.TextColumn("BOM Item Name", width="large"),
                "master_sub_material": st.column_config.TextColumn("Mapped Mastersheet Item", width="large"),
                "created_at": st.column_config.TextColumn("Learned On"),
            },
            key="alias_editor",
        )
        a1, a2 = st.columns([1.2, 1.2])
        with a1:
            if st.button("💾 Save Alias Edits / Deletions", type="primary"):
                rebuilt = edited_alias.copy()
                rebuilt["bom_strict_key"] = rebuilt["bom_sample_name"].apply(strict_alnum_key)
                rebuilt["master_strict_key"] = rebuilt["master_sub_material"].apply(strict_alnum_key)
                save_aliases(rebuilt[ALIAS_COLS])
                st.success("✅ Alias Memory updated!")
                st.rerun()
        with a2:
            st.download_button("⬇️ Export Alias Memory CSV", data=alias_df.to_csv(index=False).encode("utf-8-sig"), file_name="alias_memory.csv", mime="text/csv")


# ============================================================
# TAB 4: QUOTATION HISTORY & REV-VS-REV COMPARISON VIEW
# ============================================================
with tab_history:
    st.subheader("📜 Quotation History, Revisions & Side-by-Side Comparison")
    if not history_list:
        st.info("No quotations saved yet. Generate a quotation in Tab 1 and click **Save to Quotation History**.")
    else:
        sub_hist, sub_comp = st.tabs(["📂 Saved Revisions", "⚖️ Compare Two Revisions (Rev A vs Rev B)"])

        # --- SUB-TAB A: SAVED REVISIONS ---
        with sub_hist:
            h_search = st.text_input("🔎 Filter by Annexure/Job No, Customer, or Project:")
            filtered_hist = [
                h for h in history_list
                if not h_search.strip()
                or h_search.lower() in str(h.get("annexure_no", "")).lower()
                or h_search.lower() in str(h.get("customer", "")).lower()
                or h_search.lower() in str(h.get("project", "")).lower()
            ]

            summary_table = pd.DataFrame(
                [
                    {
                        "Annexure / Job": h.get("annexure_no", ""),
                        "Revision": f"Rev {h.get('revision', 0)}",
                        "Saved On": h.get("created_at", ""),
                        "Customer": h.get("customer", ""),
                        "Project": h.get("project", ""),
                        "Items": h.get("item_count", len(h.get("items", []))),
                        "Basic Total (₹)": round(float(h.get("basic_total", 0.0)), 2),
                        "Grand Total Incl. GST (₹)": round(float(h.get("grand_total", 0.0)), 2),
                    }
                    for h in filtered_hist
                ]
            )
            st.dataframe(summary_table, use_container_width=True, hide_index=True)

            if filtered_hist:
                h_labels = [
                    f"Annexure {h['annexure_no']} (Rev {h['revision']}) — {h.get('customer', 'N/A')} | ₹{float(h.get('grand_total', 0)):,.2f} [{h.get('created_at', '')}]"
                    for h in filtered_hist
                ]
                sel_h_idx = st.selectbox("Select Saved Quotation:", options=range(len(h_labels)), format_func=lambda i: h_labels[i])
                chosen_h = filtered_hist[sel_h_idx]

                h_items_df = pd.DataFrame(chosen_h.get("items", []))
                h_calc, h_sum = compute_commercials(h_items_df, float(chosen_h.get("profit_pct", 12)), float(chosen_h.get("pf_pct", 0)), float(chosen_h.get("gst_pct", 18)))
                h_meta = {
                    "company_name": chosen_h.get("company_name", "Quotation"),
                    "annexure_no": chosen_h.get("annexure_no", ""),
                    "revision": chosen_h.get("revision", 0),
                    "quote_date": chosen_h.get("quote_date", ""),
                    "customer": chosen_h.get("customer", ""),
                    "project": chosen_h.get("project", ""),
                    "location": chosen_h.get("location", ""),
                    "terms": chosen_h.get("terms", DEFAULT_TERMS),
                }

                hc1, hc2, hc3, hc4, hc5 = st.columns(5)
                with hc1:
                    if st.button("📂 Load to Revise", type="primary", use_container_width=True):
                        st.session_state["loaded_history_quote"] = chosen_h
                        st.success("Loaded! Switch to **🚀 Generate Quotation** to edit & save next Rev.")
                with hc2:
                    st.download_button(
                        "⬇️ Internal Excel",
                        data=build_internal_quotation_excel(h_calc, h_meta, h_sum, float(chosen_h.get("profit_pct", 12)), chosen_h.get("default_make", "Reputed")),
                        file_name=f"quotation_{h_meta['annexure_no']}_Rev{h_meta['revision']}_INTERNAL.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with hc3:
                    st.download_button(
                        "⬇️ Customer Excel",
                        data=build_customer_quotation_excel(h_calc, h_meta, h_sum, chosen_h.get("default_make", "Reputed"), logo_bytes=logo_bytes),
                        file_name=f"quotation_{h_meta['annexure_no']}_Rev{h_meta['revision']}_CUSTOMER.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with hc4:
                    if HAS_FPDF:
                        st.download_button(
                            "⬇️ Customer PDF",
                            data=build_customer_pdf(h_calc, h_meta, h_sum, chosen_h.get("default_make", "Reputed"), logo_bytes=logo_bytes),
                            file_name=f"quotation_{h_meta['annexure_no']}_Rev{h_meta['revision']}_CUSTOMER.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )
                with hc5:
                    if st.button("🗑️ Delete", use_container_width=True):
                        new_hist = [h for h in history_list if h.get("id") != chosen_h.get("id")]
                        save_history(new_hist)
                        st.success("Deleted snapshot.")
                        st.rerun()

        # --- SUB-TAB B: SIDE-BY-SIDE REVISION COMPARISON ---
        with sub_comp:
            if len(history_list) < 2:
                st.info("Save at least **2 quotations / revisions** (e.g., `Rev 0` and `Rev 1`) in History to compare them side-by-side.")
            else:
                comp_labels = [
                    f"Annexure {h['annexure_no']} (Rev {h['revision']}) — {h.get('customer', 'N/A')} | ₹{float(h.get('grand_total', 0)):,.2f} [{h.get('created_at', '')}]"
                    for h in history_list
                ]
                cc1, cc2 = st.columns(2)
                with cc1:
                    idx_a = st.selectbox("Baseline Revision (Old / Rev A):", options=range(len(comp_labels)), index=min(1, len(comp_labels) - 1), format_func=lambda i: comp_labels[i], key="comp_rev_a")
                with cc2:
                    idx_b = st.selectbox("Revised Quote (New / Rev B):", options=range(len(comp_labels)), index=0, format_func=lambda i: comp_labels[i], key="comp_rev_b")

                diff_df, comp_meta = compare_revisions(history_list[idx_a], history_list[idx_b])

                cm1, cm2, cm3, cm4, cm5 = st.columns(5)
                cm1.metric("Baseline Grand Total", f"₹ {comp_meta['sum_a']['grand_total_incl_gst']:,.2f}")
                cm2.metric("Revised Grand Total", f"₹ {comp_meta['sum_b']['grand_total_incl_gst']:,.2f}", delta=f"₹ {comp_meta['grand_delta']:+,.2f} ({comp_meta['pct_delta']:+.1f}%)")
                cm3.metric("🟢 Added Items", comp_meta["counts"]["Added"])
                cm4.metric("🔴 Removed Items", comp_meta["counts"]["Removed"])
                cm5.metric("🟡 Modified Items", comp_meta["counts"]["Modified"])

                only_changed = st.checkbox("Show only changed items (hide Unchanged rows)", value=False)
                display_diff = diff_df[diff_df["Status"] != "Unchanged"] if only_changed else diff_df
                st.dataframe(display_diff, use_container_width=True, hide_index=True)

                ex1, ex2 = st.columns(2)
                fname_base = f"Comparison_{history_list[idx_a]['annexure_no']}_Rev{history_list[idx_a]['revision']}_vs_Rev{history_list[idx_b]['revision']}"
                with ex1:
                    st.download_button(
                        "⬇️ Download Comparison Report (Excel .xlsx)",
                        data=build_comparison_excel(display_diff, comp_meta),
                        file_name=f"{fname_base}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                with ex2:
                    if HAS_FPDF:
                        st.download_button(
                            "⬇️ Download Comparison Report (PDF)",
                            data=build_comparison_pdf(display_diff, comp_meta, logo_bytes=logo_bytes),
                            file_name=f"{fname_base}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )


# ============================================================
# TAB 5: CLOUD SYNC & 1-CLICK STREAMLIT SECRETS TOML GENERATOR
# ============================================================
with tab_cloud:
    st.subheader("☁️ Live Google Sheets Sync & Streamlit Cloud Secrets Helper")

    if not HAS_GSPREAD:
        st.warning("Install Google Sheets libraries first: `pip install gspread google-auth`")

    cfg = load_cloud_config()
    if cfg.get("source") == "streamlit_secrets":
        st.success("✅ **Active on Streamlit Cloud:** Connected securely via `st.secrets` (Google Sheets Auto-Sync is ON).")
    else:
        st.markdown("#### Step 1: Connect Your Google Sheet")
        gs_url = st.text_input("Google Sheet URL", value=cfg.get("gsheet_url", ""), placeholder="https://docs.google.com/spreadsheets/d/.../edit")
        sa_file = st.file_uploader("Upload Google Service Account JSON Key", type=["json"], key="sa_uploader")
        auto_sync_chk = st.checkbox("Auto-push changes to Google Sheets on every Save (Local PC)", value=bool(cfg.get("auto_sync", False)))

        sa_data = cfg.get("service_account")
        if sa_file is not None:
            try:
                sa_data = json.load(sa_file)
                st.success(f"✅ JSON Key Loaded! Share your Google Sheet as **Editor** with: `{sa_data.get('client_email', '')}`")
            except Exception as e:
                st.error(f"Invalid JSON file: {e}")

        c_save, c_toml = st.columns(2)
        with c_save:
            if st.button("💾 Save Connection Locally", type="primary", use_container_width=True):
                if not gs_url.strip() or not sa_data:
                    st.error("Please provide both the Google Sheet URL and Service Account JSON key.")
                else:
                    save_cloud_config({"enabled": True, "auto_sync": auto_sync_chk, "gsheet_url": gs_url.strip(), "service_account": sa_data})
                    st.success("✅ Saved locally! You can now Push/Pull below.")
                    st.rerun()

        with c_toml:
            if st.button("🛠️ Generate Streamlit Cloud Secrets TOML", use_container_width=True):
                if not gs_url.strip() or not sa_data:
                    st.error("Paste your Google Sheet URL and upload your JSON key first.")
                else:
                    st.session_state["generated_toml"] = generate_streamlit_toml(gs_url.strip(), sa_data)

        if st.session_state.get("generated_toml"):
            st.markdown("#### 📋 Copy & Paste Into Streamlit Cloud (`App Settings → Secrets`)")
            st.caption("Click the copy icon in the top-right corner of the box below, then paste directly into Streamlit Community Cloud → Settings → Secrets → Save.")
            st.code(st.session_state["generated_toml"], language="toml")

    st.divider()
    st.markdown("#### Manual Sync Controls")
    s1, s2, s3 = st.columns(3)
    with s1:
        if st.button("⬆️ Push Local Data → Google Sheet", use_container_width=True):
            try:
                cfg_now = load_cloud_config()
                push_df_to_gsheet(load_mastersheet(), "Mastersheet", cfg_now)
                push_df_to_gsheet(load_aliases(), "Alias_Memory", cfg_now)
                push_history_to_gsheet(load_history(), cfg_now)
                st.success("✅ Pushed Mastersheet, Alias Memory, and Quotation History to Google Sheets!")
            except Exception as e:
                st.error(f"Push failed: {e}")
    with s2:
        if st.button("⬇️ Pull Google Sheet → Local App", use_container_width=True):
            try:
                cfg_now = load_cloud_config()
                m_cloud = pull_df_from_gsheet("Mastersheet", MASTER_COLS, cfg_now)
                if not m_cloud.empty:
                    save_mastersheet(m_cloud, sync_cloud=False)
                a_cloud = pull_df_from_gsheet("Alias_Memory", ALIAS_COLS, cfg_now)
                if not a_cloud.empty:
                    save_aliases(a_cloud, sync_cloud=False)
                h_cloud = pull_history_from_gsheet(cfg_now)
                if h_cloud:
                    save_history(h_cloud, sync_cloud=False)
                st.success("✅ Pulled latest live data from Google Sheets!")
                st.rerun()
            except Exception as e:
                st.error(f"Pull failed: {e}")
    with s3:
        if st.button("🔌 Disconnect Cloud (Stay Offline)", use_container_width=True):
            if os.path.exists(CLOUD_CFG_FILE):
                os.remove(CLOUD_CFG_FILE)
            st.info("Switched back to local offline mode.")
            st.rerun()

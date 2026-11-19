import io
import os
import re
from difflib import SequenceMatcher

import openpyxl
import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
# PAGE CONFIGURATION & CUSTOM CSS (NO MORE TEXT TRUNCATION)
# ============================================================
st.set_page_config(
    page_title="SundryQuote Pro — BOM to Quotation System",
    page_icon="📊",
    layout="wide",
)

DB_FILE = "mastersheet_db.csv"
INITIAL_SEED_FILE = "final_table_checked.csv"
MASTER_COLS = ["main_material", "sub_material", "purchase_price", "discount"]

st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(90deg, #1F4E78 0%, #2B6CB0 100%);
        padding: 1.2rem 1.5rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 1.2rem;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; color: white; }
    .main-header p { margin: 0.25rem 0 0 0; opacity: 0.92; font-size: 0.95rem; }

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
        font-size: 0.93rem;
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
        font-size: 0.92rem;
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
# DATABASE FUNCTIONS (MASTERSHEET PERSISTENCE)
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
    df = df[df["sub_material"] != ""].reset_index(drop=True)
    return df


def load_mastersheet() -> pd.DataFrame:
    if os.path.exists(DB_FILE):
        return clean_master_df(pd.read_csv(DB_FILE, encoding="utf-8-sig"))
    if os.path.exists(INITIAL_SEED_FILE):
        df = clean_master_df(pd.read_csv(INITIAL_SEED_FILE, encoding="utf-8-sig"))
        save_mastersheet(df)
        return df
    starter = pd.DataFrame(
        [
            {"main_material": "CABLE TIE", "sub_material": "CABLE TIE 100MM (100 TIES /PACK)", "purchase_price": 52.0, "discount": 50.0},
            {"main_material": "CABLE TIE", "sub_material": "CABLE TIE 200MM (100 TIES /PACK)", "purchase_price": 153.0, "discount": 50.0},
            {"main_material": "BLIND - SIGNAL", "sub_material": 'BLIND FOR CABLE GLAND 1/2" NPT', "purchase_price": 133.0, "discount": 0.0},
        ]
    )
    save_mastersheet(starter)
    return starter


def save_mastersheet(df: pd.DataFrame) -> None:
    cleaned = clean_master_df(df)
    cleaned.to_csv(DB_FILE, index=False, encoding="utf-8-sig")


# ============================================================
# STRICT MATCHING & NUMBER-SAFE CLOSE-MATCH ENGINE
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
        # Penalize mismatched sizes/numbers (e.g. 6MM2 vs 10MM2)
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
# MULTI-FORMAT BOM PARSER (WITH HIDDEN EXCEL ROW/SHEET SKIP)
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
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=["main_material", "sub_material", "qty", "source"]), ""

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
        return pd.DataFrame(columns=["main_material", "sub_material", "qty", "source"]), ""

    job_no = ""
    for idx in range(header_idx):
        row_text = " ".join([str(v) for v in df_raw.iloc[idx].tolist() if pd.notna(v)])
        m = re.search(r"(?:JOB\s*NO|ANNEXURE|QUOTATION\s*NO)[^\d]*(\d{4,})", row_text, re.IGNORECASE)
        if m:
            job_no = m.group(1)
            break

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
                items.append(
                    {
                        "main_material": current_main,
                        "sub_material": desc_val,
                        "qty": extract_qty(qty_raw),
                        "source": row_source,
                    }
                )
        else:
            if sr_val and not sr_val.replace(".", "").isdigit() and not desc_val:
                current_main = sr_val
                continue
            if not sr_val and desc_val and (pd.isna(qty_raw) or str(qty_raw).strip() == ""):
                current_main = desc_val
                continue
            if desc_val:
                items.append(
                    {
                        "main_material": current_main,
                        "sub_material": desc_val,
                        "qty": extract_qty(qty_raw),
                        "source": row_source,
                    }
                )

    return pd.DataFrame(items), job_no


def read_excel_visible_only(raw_bytes: bytes, fname: str, ignore_hidden: bool = True):
    """Reads .xlsx/.xlsm via openpyxl so hidden sheets, hidden rows, and hidden columns are skipped."""
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
            parsed, job_no = parse_bom_dataframe(
                df_raw, source_label=f"{fname} [{sheet_name}]", row_numbers=orig_row_nums
            )
            if not parsed.empty:
                results[f"{fname} → {sheet_name}"] = (parsed, job_no)
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
        results, hidden_rows_skipped, hidden_sheets_skipped = read_excel_visible_only(
            raw_bytes, fname, ignore_hidden=ignore_hidden
        )

    elif ext == ".xls":
        xls = pd.ExcelFile(io.BytesIO(raw_bytes), engine="xlrd")
        for sheet in xls.sheet_names:
            try:
                df_raw = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet, header=None, engine="xlrd")
                parsed, job_no = parse_bom_dataframe(df_raw, source_label=f"{fname} [{sheet}]")
                if not parsed.empty:
                    results[f"{fname} → {sheet}"] = (parsed, job_no)
            except Exception:
                continue

    elif ext in (".csv", ".tsv", ".txt"):
        df_raw = None
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                sep = "\t" if ext == ".tsv" else None
                df_raw = pd.read_csv(
                    io.BytesIO(raw_bytes),
                    header=None,
                    encoding=enc,
                    sep=sep,
                    engine="python",
                    on_bad_lines="skip",
                )
                break
            except Exception:
                continue
        if df_raw is not None:
            parsed, job_no = parse_bom_dataframe(df_raw, source_label=fname)
            if not parsed.empty:
                results[fname] = (parsed, job_no)

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
            parsed, job_no = parse_bom_dataframe(df_raw, source_label=fname)
            if not parsed.empty:
                results[fname] = (parsed, job_no)

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
            parsed, job_no = parse_bom_dataframe(df_raw, source_label=fname)
            if not parsed.empty:
                results[fname] = (parsed, job_no)

    m = re.search(r"(\d{6,})", fname)
    fname_job = m.group(1) if m else ""
    for k, (df_p, jn) in results.items():
        if not jn and fname_job:
            results[k] = (df_p, fname_job)

    return results, hidden_rows_skipped, hidden_sheets_skipped


# ============================================================
# EXCEL QUOTATION GENERATOR (EXACT FORMAT OF quotation 6500389)
# ============================================================
def build_quotation_excel(
    final_df: pd.DataFrame,
    annexure_no: str,
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
    font_total = Font(name="Calibri", size=11, bold=True, color="000000")

    fill_cat = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_profit = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_total = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    thin = Side(style="thin", color="BFBFBF")
    med = Side(style="medium", color="1F4E78")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_total = Border(top=med, bottom=med, left=thin, right=thin)

    ws.cell(row=1, column=2, value=f"Annexure {annexure_no}").font = font_title
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

            headers = {
                3: "Make",
                4: "QTY",
                5: "Amnt",
                6: "Discount",
                7: "Total Amnt",
                9: "Purchase",
                11: "Profit",
            }
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

            c1 = ws.cell(row=r, column=1, value=sr_no)
            c1.alignment = Alignment(horizontal="center", vertical="center")

            c2 = ws.cell(row=r, column=2, value=item["sub_material"])
            c2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            c3 = ws.cell(row=r, column=3, value=default_make)
            c3.alignment = Alignment(horizontal="center", vertical="center")

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

            for col_i in (1, 2, 3, 4, 5, 6, 7, 9):
                ws.cell(row=r, column=col_i).font = font_data
                ws.cell(row=r, column=col_i).border = border_cell

            current_row += 1

    last_data_row = current_row - 1
    if first_data_row is not None:
        ws.cell(row=current_row, column=5, value="TOTAL").font = font_total
        ws.cell(row=current_row, column=5).alignment = Alignment(horizontal="right", vertical="center")

        c_tot = ws.cell(row=current_row, column=7, value=f"=SUM(G{first_data_row}:G{last_data_row})")
        c_tot.font = font_total
        c_tot.number_format = "#,##0.00"
        c_tot.alignment = Alignment(horizontal="right", vertical="center")

        for col_i in range(1, 10):
            ws.cell(row=current_row, column=col_i).fill = fill_total
            ws.cell(row=current_row, column=col_i).border = border_total
        ws.row_dimensions[current_row].height = 24

    widths = {1: 10, 2: 68, 3: 12, 4: 10, 5: 13, 6: 11, 7: 16, 8: 3, 9: 13, 10: 3, 11: 10, 12: 8}
    for col_i, w in widths.items():
        ws.column_dimensions[get_column_letter(col_i)].width = w
    ws.freeze_panes = "A3"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

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
# MAIN UI HEADER
# ============================================================
st.markdown(
    """
    <div class="main-header">
        <h1>📊 SundryQuote Pro — Automated BOM Quotation System</h1>
        <p>Visible-Only Excel Reader • Size-Safe Close-Match Shield • Editable Live-Formula Quotations</p>
    </div>
    """,
    unsafe_allow_html=True,
)

master_df = load_mastersheet()
tab_quote, tab_master = st.tabs(["🚀 Generate Quotation", "🗄️ Mastersheet Database"])


# ============================================================
# TAB 1: QUOTATION GENERATOR
# ============================================================
with tab_quote:
    if master_df.empty:
        st.warning("Your Mastersheet is currently empty. Add or import items in the **Mastersheet Database** tab first.")
    else:
        col_left, col_right = st.columns([2, 1.1])
        with col_left:
            uploaded_boms = st.file_uploader(
                "Upload one or more BOM files (Excel .xlsx/.xlsm/.xls, CSV, PDF, or Word .docx)",
                type=["xlsx", "xlsm", "xls", "csv", "tsv", "txt", "pdf", "docx"],
                accept_multiple_files=True,
            )
        with col_right:
            profit_pct = st.number_input("Profit Margin % (Cell L2)", min_value=0.0, max_value=500.0, value=12.0, step=0.5)
            default_make = st.text_input("Default Make", value="Reputed")
            ignore_hidden = st.checkbox("Ignore hidden Excel rows & sheets", value=True, help="Skips hidden/filtered rows and hidden tabs inside .xlsx/.xlsm workbooks.")
            c_opt1, c_opt2 = st.columns(2)
            with c_opt1:
                combine_dupes = st.checkbox("Combine duplicate items", value=True)
            with c_opt2:
                skip_zero_qty = st.checkbox("Skip 0 / NA qty", value=True)

        if uploaded_boms:
            all_sources = {}
            detected_jobs = []
            total_hidden_rows = 0
            total_hidden_sheets = 0

            for f in uploaded_boms:
                parsed_map, h_rows, h_sheets = read_any_bom_file(f, ignore_hidden=ignore_hidden)
                total_hidden_rows += h_rows
                total_hidden_sheets += h_sheets
                for label, (df_p, jn) in parsed_map.items():
                    all_sources[label] = df_p
                    if jn:
                        detected_jobs.append(jn)

            if ignore_hidden and (total_hidden_rows > 0 or total_hidden_sheets > 0):
                st.success(
                    f"🧹 Automatically skipped **{total_hidden_rows} hidden row(s)** and **{total_hidden_sheets} hidden sheet(s)** from your Excel file(s)."
                )

            if not all_sources:
                st.error(
                    "Could not detect a BOM table with Description and QTY columns in the uploaded file(s). "
                    "Make sure the file has column headers like `Description` (or `Sub Material`) and `QTY`."
                )
            else:
                selected_sources = st.multiselect(
                    "Detected BOM Sheets / Files (deselect any you want to ignore):",
                    options=list(all_sources.keys()),
                    default=list(all_sources.keys()),
                )

                default_annex = detected_jobs[0] if detected_jobs else "6500389"
                annexure_no = st.text_input("Annexure / Job Number (Row 1 Header)", value=default_annex)

                if selected_sources:
                    bom_raw_combined = pd.concat([all_sources[s] for s in selected_sources], ignore_index=True)
                    zero_qty_count = int((bom_raw_combined["qty"] <= 0).sum())

                    if skip_zero_qty:
                        bom_combined = bom_raw_combined[bom_raw_combined["qty"] > 0].reset_index(drop=True)
                        if zero_qty_count > 0:
                            st.caption(f"ℹ️ Filtered out **{zero_qty_count}** row(s) that had `0`, blank, or `NA` quantity.")
                    else:
                        bom_combined = bom_raw_combined.copy()

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

                    exact_rows = []
                    unmatched_rows = []

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

                    m1, m2, m3 = st.columns(3)
                    m1.metric("Active BOM Items", len(bom_combined))
                    m2.metric("✅ Auto-Matched (Exact Alphanumeric)", len(exact_rows))
                    m3.metric("⚠️ Needs Your Yes / No Review", len(unmatched_rows))

                    # --------------------------------------------------------
                    # FULL-WIDTH UNTRUNCATED CLOSE-MATCH REVIEW CARDS
                    # --------------------------------------------------------
                    approved_close_rows = []
                    excluded_rows = []

                    if unmatched_rows:
                        st.subheader("🔍 Review Unmatched BOM Items (Close-Match Approval)")
                        st.info(
                            "Every card below shows the **complete unwrapped BOM name**, **source sheet & row**, and **complete Mastersheet match**. "
                            "Items with mismatched sizes/numbers show a red warning and are never auto-accepted."
                        )

                        b1, b2, _ = st.columns([1.8, 1.2, 2.0])
                        with b1:
                            if st.button("✅ Accept Safe High-Confidence (≥ 80% & No Size Mismatch)"):
                                for u in unmatched_rows:
                                    if u["suggestions"]:
                                        top0 = u["suggestions"][0]
                                        if top0["score"] >= 80.0 and not top0["num_mismatch"]:
                                            st.session_state[f"yn_{u['uid']}"] = "Yes (Include)"
                        with b2:
                            if st.button("❌ Reject All Unmatched"):
                                for u in unmatched_rows:
                                    st.session_state[f"yn_{u['uid']}"] = "No (Exclude)"

                        all_master_records = master_df.to_dict(orient="records")

                        for i, u in enumerate(unmatched_rows, start=1):
                            top_sugg = u["suggestions"]

                            with st.container(border=True):
                                # Top bar: Item number, full BOM description, metadata, and Yes/No switch
                                c_top_l, c_top_r = st.columns([3.8, 1.2])
                                with c_top_l:
                                    st.markdown(
                                        f'<div class="bom-name-box">{i}. BOM Item: {u["bom_sub"]}</div>',
                                        unsafe_allow_html=True,
                                    )
                                    st.caption(f"Category: **{u['bom_main']}**  •  QTY: **{u['qty']:g}**  •  Source: `{u['source']}`")
                                with c_top_r:
                                    decision = st.radio(
                                        "Decision",
                                        options=["No (Exclude)", "Yes (Include)"],
                                        key=f"yn_{u['uid']}",
                                        label_visibility="collapsed",
                                    )

                                # Mode toggle: Top 5 Closest vs Search Full Mastersheet
                                search_all = st.checkbox(
                                    "Search all Mastersheet items instead of Top 5",
                                    key=f"all_{u['uid']}",
                                )

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
                                        "Suggested Close Match (click dropdown to compare Top 5):",
                                        options=range(len(option_labels)),
                                        format_func=lambda x, lbls=option_labels: lbls[x],
                                        key=f"sel_{u['uid']}",
                                    )
                                    chosen_match = top_sugg[selected_opt]

                                # Full, never-truncated display of the selected Mastersheet item
                                st.markdown(
                                    f'<div class="master-name-box">'
                                    f'<b>Selected Mastersheet Item ({chosen_match["score"]}% match):</b><br>'
                                    f'{chosen_match["sub_material"]}<br>'
                                    f'<span style="font-size:0.85rem;color:#4A5568;">'
                                    f'Group: <b>{chosen_match["main_material"]}</b> &nbsp;|&nbsp; '
                                    f'Purchase Rate: <b>₹{chosen_match["purchase_price"]:,.2f}</b> &nbsp;|&nbsp; '
                                    f'Discount: <b>{chosen_match["discount"]:g}%</b></span>'
                                    f'</div>',
                                    unsafe_allow_html=True,
                                )

                                if chosen_match["num_mismatch"]:
                                    b_diff = ", ".join(chosen_match["bom_only_nums"]) or "None"
                                    m_diff = ", ".join(chosen_match["master_only_nums"]) or "None"
                                    st.markdown(
                                        f'<div class="warn-badge">⚠️ Size / Number Mismatch — BOM has [{b_diff}] vs Mastersheet has [{m_diff}]. Verify carefully!</div>',
                                        unsafe_allow_html=True,
                                    )
                                else:
                                    st.markdown(
                                        '<div class="safe-badge">✅ All numbers & sizes match between BOM and Mastersheet</div>',
                                        unsafe_allow_html=True,
                                    )

                                # Quick inline expander to add brand-new BOM item straight into Mastersheet
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
                                                [
                                                    {
                                                        "main_material": add_main.strip() or "GENERAL",
                                                        "sub_material": u["bom_sub"],
                                                        "purchase_price": float(add_price),
                                                        "discount": float(add_disc),
                                                    }
                                                ]
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

                    # --------------------------------------------------------
                    # EDITABLE FINAL QUOTATION PREVIEW & EXCEL EXPORT
                    # --------------------------------------------------------
                    all_included = exact_rows + approved_close_rows
                    st.divider()
                    st.subheader("📑 Final Quotation Preview (Editable Rates) & Export")

                    if not all_included:
                        st.warning("No items are currently included in the quotation. Approve close matches above or update your Mastersheet.")
                    else:
                        final_df = pd.DataFrame(all_included)
                        if combine_dupes:
                            final_df = (
                                final_df.groupby(["main_material", "sub_material"], sort=False, as_index=False)
                                .agg(
                                    qty=("qty", "sum"),
                                    purchase_price=("purchase_price", "first"),
                                    discount=("discount", "first"),
                                    match_status=("match_status", lambda x: " + ".join(dict.fromkeys(x))),
                                )
                            )

                        st.caption("Tip: You can double-click **QTY**, **Purchase ₹**, or **Discount %** right below to fix any ₹0.00 prices before downloading.")
                        edited_quote = st.data_editor(
                            final_df[["main_material", "sub_material", "qty", "purchase_price", "discount"]],
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

                        # Sync edited prices back to Mastersheet button
                        if st.button("💾 Save Any Edited Purchase Prices / Discounts Back to Mastersheet"):
                            updated_master = master_df.copy()
                            updated_master["_k"] = updated_master["sub_material"].apply(strict_alnum_key)
                            for _, er in edited_quote.iterrows():
                                ek = strict_alnum_key(er["sub_material"])
                                mask = updated_master["_k"] == ek
                                if mask.any():
                                    updated_master.loc[mask, "purchase_price"] = float(er["purchase_price"])
                                    updated_master.loc[mask, "discount"] = float(er["discount"])
                            save_mastersheet(updated_master.drop(columns=["_k"]))
                            st.success("✅ Updated Mastersheet with your edited rates!")
                            st.rerun()

                        # Recompute totals using edited values
                        edited_quote["Amnt"] = edited_quote["purchase_price"] * (1.0 + profit_pct / 100.0)
                        edited_quote["Total Amnt"] = edited_quote["qty"] * edited_quote["Amnt"] * (1.0 - edited_quote["discount"] / 100.0)
                        edited_quote["Total Purchase"] = edited_quote["qty"] * edited_quote["purchase_price"]

                        grand_total = edited_quote["Total Amnt"].sum()
                        total_purchase = edited_quote["Total Purchase"].sum()
                        zero_rate_count = int((edited_quote["purchase_price"] == 0).sum())

                        k1, k2, k3, k4 = st.columns(4)
                        k1.metric("Included Line Items", len(edited_quote))
                        k2.metric("Grand Quoted Total", f"₹ {grand_total:,.2f}")
                        k3.metric("Total Purchase Cost", f"₹ {total_purchase:,.2f}")
                        k4.metric("Items with ₹0 Price", zero_rate_count)

                        if zero_rate_count > 0:
                            st.warning(f"⚠️ {zero_rate_count} item(s) still have a **₹0.00 Purchase Price**. Edit them directly in the table above.")

                        audit_records = [
                            {
                                "BOM Main Material": r["bom_main"],
                                "BOM Sub Material": r["bom_sub"],
                                "QTY": r["qty"],
                                "Source": r["source"],
                                "Matched Mastersheet Item": r["sub_material"],
                                "Status": r["match_status"],
                            }
                            for r in all_included
                        ] + excluded_rows
                        audit_df = pd.DataFrame(audit_records)

                        excel_buffer = build_quotation_excel(
                            final_df=edited_quote,
                            annexure_no=annexure_no,
                            profit_pct=profit_pct,
                            default_make=default_make,
                            audit_df=audit_df,
                        )

                        dl1, dl2 = st.columns([1.5, 1.5])
                        with dl1:
                            st.download_button(
                                label=f"⬇️ Generate & Download Quotation ({annexure_no}.xlsx)",
                                data=excel_buffer,
                                file_name=f"quotation_{annexure_no}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True,
                            )
                        with dl2:
                            st.download_button(
                                label="⬇️ Download Quotation as CSV",
                                data=edited_quote.to_csv(index=False).encode("utf-8-sig"),
                                file_name=f"quotation_{annexure_no}.csv",
                                mime="text/csv",
                                use_container_width=True,
                            )


# ============================================================
# TAB 2: MASTERSHEET DATABASE (ADD / EDIT / DELETE / IMPORT)
# ============================================================
with tab_master:
    st.subheader("🗄️ Mastersheet Rate & Discount Database")
    st.caption(f"Total items stored: **{len(master_df)}** across **{master_df['main_material'].nunique()}** Main Material groups.")

    m_tab_edit, m_tab_add, m_tab_io = st.tabs(["✏️ View, Edit & Delete", "➕ Add New Item", "📤 Import / Export Backup"])

    with m_tab_edit:
        search_q = st.text_input("🔎 Quick Search (filter view only — clear search to edit/delete safely):", placeholder="Type e.g. M20, LUGS, CABLE TIE...")
        if search_q.strip():
            mask = (
                master_df["main_material"].str.contains(search_q, case=False, na=False)
                | master_df["sub_material"].str.contains(search_q, case=False, na=False)
            )
            st.info(f"Showing {mask.sum()} matching item(s). Clear the search box to edit or delete rows.")
            st.dataframe(master_df[mask], use_container_width=True, hide_index=True)
        else:
            st.caption("Double-click any cell to edit. Click `+` at the bottom to add. Select row checkboxes and press **Delete** to remove. Click **Save Changes** when done.")
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
                    new_row = pd.DataFrame(
                        [
                            {
                                "main_material": final_main,
                                "sub_material": new_sub.strip(),
                                "purchase_price": float(new_price),
                                "discount": float(new_disc),
                            }
                        ]
                    )
                    save_mastersheet(pd.concat([master_df, new_row], ignore_index=True))
                    st.success(f"✅ Added `{new_sub.strip()}` under `{final_main}`!")
                    st.rerun()

    with m_tab_io:
        st.markdown("#### Upload Mastersheet (CSV or Excel)")
        up_master = st.file_uploader("Upload Mastersheet File", type=["csv", "xlsx", "xls"], key="master_upload")
        import_mode = st.radio(
            "Import Mode:",
            ["Merge (Add new items & update existing prices)", "Replace Entire Mastersheet"],
            horizontal=True,
        )
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
                    st.success(f"✅ Merged successfully! Mastersheet now has {len(combined)} items.")
                st.rerun()
            except Exception as e:
                st.error(f"Import failed: {e}")

        st.divider()
        st.markdown("#### Download Mastersheet Backup")
        e1, e2 = st.columns(2)
        with e1:
            st.download_button(
                "⬇️ Export as CSV",
                data=master_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="mastersheet_backup.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with e2:
            xl_buf = io.BytesIO()
            with pd.ExcelWriter(xl_buf, engine="openpyxl") as writer:
                master_df.to_excel(writer, index=False, sheet_name="Mastersheet")
            xl_buf.seek(0)
            st.download_button(
                "⬇️ Export as Excel (.xlsx)",
                data=xl_buf,
                file_name="mastersheet_backup.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
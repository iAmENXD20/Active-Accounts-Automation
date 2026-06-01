import streamlit as st
from io import BytesIO
import pandas as pd
import msoffcrypto
import openpyxl
import re

st.set_page_config(page_title="AA Automation - Excel Processor", layout="wide")

st.title("📊 AA Automation: Excel Data Extractor")
st.markdown("Upload your Excel file to view all account records and their data.")


# ── helpers ────────────────────────────────────────────────────────────────────

def decrypt_file(uploaded_file, password: str) -> BytesIO:
    buf = BytesIO(uploaded_file.read())
    try:
        office = msoffcrypto.OfficeFile(buf)
        office.load_key(password=password)
        out = BytesIO()
        office.decrypt(out)
        out.seek(0)
        return out
    except Exception:
        buf.seek(0)
        return buf


def get_sheet_names(file_bytes: BytesIO) -> list[str]:
    file_bytes.seek(0)
    wb = openpyxl.load_workbook(file_bytes, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def parse_date_from_filename(filename: str) -> str:
    """
    Extract a date from the filename.
    Supports formats: MMDDYY (6-digit) or MMDDYYYY (8-digit).
    Returns formatted MM/DD/YYYY string or empty string if not found.
    """
    # Strip extension, grab all digit sequences
    stem = filename.rsplit(".", 1)[0]
    matches = re.findall(r"\b(\d{6}|\d{8})\b", stem)
    for m in matches:
        try:
            if len(m) == 6:
                # MMDDYY
                mm, dd, yy = m[0:2], m[2:4], m[4:6]
                yyyy = "20" + yy
            else:
                # MMDDYYYY
                mm, dd, yyyy = m[0:2], m[2:4], m[4:8]
            # Basic sanity check
            if 1 <= int(mm) <= 12 and 1 <= int(dd) <= 31:
                return f"{mm}/{dd}/{yyyy}"
        except Exception:
            continue
    return "N/A"


def find_header_row(file_bytes: BytesIO, sheet: str, max_scan: int = 20) -> int:
    """Return 0-based index of the row with the most non-empty string cells."""
    file_bytes.seek(0)
    raw = pd.read_excel(
        file_bytes, sheet_name=sheet, header=None, nrows=max_scan, engine="openpyxl"
    )
    best_row, best_score = 0, -1
    for i, row in raw.iterrows():
        score = sum(1 for v in row.dropna() if isinstance(v, str) and v.strip())
        if score > best_score:
            best_score, best_row = score, i
    return int(best_row)


def clean_columns(cols: list) -> list:
    """Strip, replace blanks, and deduplicate column names."""
    result = []
    seen: dict[str, int] = {}
    for i, c in enumerate(cols):
        name = str(c).strip() if c is not None else ""
        if name.lower().startswith("unnamed:") or name in ("nan", "", "None"):
            name = f"_col_{i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        result.append(name)
    return result


def load_sheet(file_bytes: BytesIO, sheet: str, header_row: int) -> pd.DataFrame:
    """
    Load sheet via openpyxl iter_rows, but only up to the last column
    that actually has a non-empty header — eliminating phantom empty columns.
    """
    file_bytes.seek(0)
    wb = openpyxl.load_workbook(file_bytes, read_only=True, data_only=True)
    ws = wb[sheet]

    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not all_rows or header_row >= len(all_rows):
        return pd.DataFrame()

    raw_headers = list(all_rows[header_row])

    # ── key fix: trim trailing empty/None headers ──────────────────────────
    last_valid = len(raw_headers) - 1
    while last_valid >= 0 and (
        raw_headers[last_valid] is None
        or str(raw_headers[last_valid]).strip() in ("", "nan", "None")
    ):
        last_valid -= 1
    raw_headers = raw_headers[: last_valid + 1]
    n_cols = len(raw_headers)
    # ───────────────────────────────────────────────────────────────────────

    data_rows = all_rows[header_row + 1:]

    # Pad/trim every data row to match trimmed header width
    padded = [
        list(row[:n_cols]) + [None] * (n_cols - len(row[:n_cols]))
        for row in data_rows
    ]

    df = pd.DataFrame(padded, columns=clean_columns(raw_headers))
    df = df.dropna(how="all")
    df = df.reset_index(drop=True)

    # Format specific date columns as MM/DD/YYYY strings (no time)
    DATE_COLUMNS = {"ENDO_DATE", "BIRTHDATE", "LOAN FUNDING DATE"}
    for col in df.columns:
        if col.strip().upper() in DATE_COLUMNS:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%m/%d/%Y")
            df[col] = df[col].where(df[col].notna(), other=None)

    return df


# ── sidebar ────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Settings")
    password = st.text_input("Excel Password", value="30PL2026", type="password")
    auto_detect = st.checkbox("Auto-detect header row", value=True)


# ── main ───────────────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader(
    "Drag and drop your Excel file here", type=["xlsx", "xlsm", "xls"]
)

if uploaded_file is not None:
    try:
        with st.spinner("Decrypting file…"):
            file_bytes = decrypt_file(uploaded_file, password)

        # ── WL DATE from filename ────────────────────────────────────────────
        wl_date_value = parse_date_from_filename(uploaded_file.name)

        # ── sheet selection (default to first sheet) ─────────────────────────
        sheet_names = get_sheet_names(file_bytes)
        selected_sheet = sheet_names[0]

        # ── header detection ─────────────────────────────────────────────────
        if auto_detect:
            with st.spinner("Detecting header row…"):
                header_row_0based = find_header_row(file_bytes, selected_sheet)
        else:
            header_row_0based = 0

        st.caption(f"Using header row: **{header_row_0based + 1}** (1-based)")

        # ── load data ────────────────────────────────────────────────────────
        with st.spinner("Reading data…"):
            df = load_sheet(file_bytes, selected_sheet, header_row_0based)

        # ── summary metrics ──────────────────────────────────────────────────
        m1, m2, m3 = st.columns(3)
        m1.metric("WL DATE", wl_date_value)
        m2.metric("No. of Active", f"{len(df):,}")
        m3.metric("Total Columns", len(df.columns))

        st.subheader("All Account Records")
        st.dataframe(df, use_container_width=True, height=500)

        # ── download ─────────────────────────────────────────────────────────
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Accounts")

        st.download_button(
            label="⬇️ Download as Excel",
            data=output.getvalue(),
            file_name="accounts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.exception(e)
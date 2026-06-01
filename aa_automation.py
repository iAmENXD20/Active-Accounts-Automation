import streamlit as st
from io import BytesIO
import pandas as pd
import msoffcrypto
import openpyxl
import re

st.set_page_config(page_title="AA Automation - Excel Processor", layout="wide")

st.title("📊 AA Automation: Excel Data Extractor")
st.markdown("Upload one or more worklist Excel files to summarize WL DATE and No. of Active accounts.")


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
    stem = filename.rsplit(".", 1)[0]
    matches = re.findall(r"\b(\d{6}|\d{8})\b", stem)
    for m in matches:
        try:
            if len(m) == 6:
                mm, dd, yy = m[0:2], m[2:4], m[4:6]
                yyyy = "20" + yy
            else:
                mm, dd, yyyy = m[0:2], m[2:4], m[4:8]
            if 1 <= int(mm) <= 12 and 1 <= int(dd) <= 31:
                return f"{mm}/{dd}/{yyyy}"
        except Exception:
            continue
    return "N/A"


def find_header_row(file_bytes: BytesIO, sheet: str, max_scan: int = 20) -> int:
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


def count_active_rows(file_bytes: BytesIO, sheet: str, header_row: int) -> int:
    file_bytes.seek(0)
    wb = openpyxl.load_workbook(file_bytes, read_only=True, data_only=True)
    ws = wb[sheet]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not all_rows or header_row >= len(all_rows):
        return 0

    data_rows = all_rows[header_row + 1:]
    # Count rows that are not completely empty
    count = sum(1 for row in data_rows if any(cell is not None for cell in row))
    return count


def build_summary_excel(records: list[dict]) -> bytes:
    """
    Build the output Excel with this layout:
      Row 1 : 30DPD Active Accounts  (merged A1:B1)
      Row 2 : WL DATE | No. of Active
      Row 3+: data rows
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"

    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    # ── Row 1: title ────────────────────────────────────────────────────────
    ws.merge_cells("A1:B1")
    title_cell = ws["A1"]
    title_cell.value = "30DPD Active Accounts"
    title_cell.font = Font(bold=True, size=13)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill("solid", fgColor="1F4E79")
    title_cell.font = Font(bold=True, size=13, color="FFFFFF")
    ws.row_dimensions[1].height = 22

    # ── Row 2: column headers ────────────────────────────────────────────────
    header_fill = PatternFill("solid", fgColor="2E75B6")
    header_font = Font(bold=True, color="FFFFFF")
    header_border = Border(
        bottom=Side(style="thin", color="FFFFFF"),
        right=Side(style="thin", color="FFFFFF"),
    )
    for col_idx, label in enumerate(["WL DATE", "No. of Active"], start=1):
        cell = ws.cell(row=2, column=col_idx, value=label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = header_border
    ws.row_dimensions[2].height = 18

    # ── Row 3+: data ─────────────────────────────────────────────────────────
    alt_fill = PatternFill("solid", fgColor="DEEAF1")
    for row_idx, rec in enumerate(records, start=3):
        # WL DATE
        d_cell = ws.cell(row=row_idx, column=1, value=rec["wl_date"])
        d_cell.alignment = Alignment(horizontal="center")
        # No. of Active
        n_cell = ws.cell(row=row_idx, column=2, value=rec["no_of_active"])
        n_cell.alignment = Alignment(horizontal="center")
        # Alternating row color
        if row_idx % 2 == 0:
            d_cell.fill = alt_fill
            n_cell.fill = alt_fill

    # ── column widths ────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# ── sidebar ────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Settings")
    password = st.text_input("Excel Password", value="30PL2026", type="password")
    auto_detect = st.checkbox("Auto-detect header row", value=True)

# ── main ───────────────────────────────────────────────────────────────────────

uploaded_files = st.file_uploader(
    "Drag and drop your Excel files here",
    type=["xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    records = []

    for uploaded_file in uploaded_files:
        try:
            with st.spinner(f"Processing {uploaded_file.name}…"):
                file_bytes = decrypt_file(uploaded_file, password)

                wl_date = parse_date_from_filename(uploaded_file.name)

                sheet_names = get_sheet_names(file_bytes)
                # Auto-pick first sheet
                selected_sheet = sheet_names[0]

                if auto_detect:
                    header_row_0based = find_header_row(file_bytes, selected_sheet)
                else:
                    header_row_0based = 0

                no_of_active = count_active_rows(file_bytes, selected_sheet, header_row_0based)

                records.append({
                    "wl_date": wl_date,
                    "no_of_active": no_of_active,
                    "filename": uploaded_file.name,
                })

        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")

    if records:
        # ── per-file checker cards ───────────────────────────────────────────
        st.subheader("📋 Summary")
        for rec in records:
            st.caption(f"📄 {rec['filename']}")
            c1, c2 = st.columns(2)
            c1.metric("WL DATE", rec["wl_date"])
            c2.metric("No. of Active", f"{rec['no_of_active']:,}")
            st.divider()

        # ── download ─────────────────────────────────────────────────────────
        excel_bytes = build_summary_excel(records)
        st.download_button(
            label="⬇️ Download Summary as Excel",
            data=excel_bytes,
            file_name="30DPD_Active_Accounts_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
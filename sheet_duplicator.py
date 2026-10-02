import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials
import io

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(
    page_title="Sheet Tools",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ------------------------------------------------------------
# IST HELPERS
# ------------------------------------------------------------
def _ist_now():
    ist = timezone(timedelta(hours=5, minutes=30))
    return datetime.now(ist)

def _ist_today():
    return _ist_now().date()

# ============================================================
# LOGIN
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div style="text-align:center; padding: 40px 0;">
        <h1 style="color:#1a73e8; font-size: 48px;">📋 Sheet Tools</h1>
        <p style="color:#888; font-size: 16px;">Please log in to continue</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            u = st.text_input("Username", placeholder="Enter your username")
            p = st.text_input("Password", type="password", placeholder="Enter your password")
            submitted = st.form_submit_button("🔐 Log In", use_container_width=True, type="primary")

        if submitted:
            try:
                correct_user = st.secrets.get("MY_USERNAME", "")
                correct_pass = st.secrets.get("MY_PASSWORD", "")
            except Exception:
                correct_user = ""
                correct_pass = ""

            if u and p and u == correct_user and p == correct_pass:
                st.session_state.authenticated = True
                st.session_state.logged_in_user = u
                st.rerun()
            else:
                st.error("❌ Invalid username or password")

    st.stop()

current_user = st.session_state.get("logged_in_user", "user")

# ============================================================
# GOOGLE SHEETS
# ============================================================
@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_resource
def get_spreadsheet():
    return get_gspread_client().open_by_key(st.secrets["spreadsheet_id"])

# ------------------------------------------------------------
# CELL MAPPING (Duplicate Sheet)
# ------------------------------------------------------------
CELL_MAP = {
    "challan_no": "F59",
    "challan_date": "G59",
    "invoice_date": "H59",
    "vehicle_no": "I59",
    "house1st": "G62",
    "age1st": "G63",
    "birds1st": "G64",
    "house2nd": "J62",
    "age2nd": "J63",
    "birds2nd": "J64",
}

EXCEL_SOURCE_RANGE = "C2:F38"
GS_DEST_RANGE = "A57:D93"

# ============================================================
# READ HELPERS
# ============================================================
@st.cache_data(ttl=15, show_spinner=False)
def list_sheet_names():
    try:
        ss = get_spreadsheet()
        return [ws.title for ws in ss.worksheets()]
    except Exception as e:
        st.error(f"Could not fetch sheets: {e}")
        return []

@st.cache_data(ttl=15, show_spinner=False)
def get_last_sheet_prefill():
    try:
        ss = get_spreadsheet()
        sheets = ss.worksheets()
        if not sheets:
            return {}
        last = sheets[-1]
        prefill = {"sheet_name": last.title}
        for key, cell in CELL_MAP.items():
            try:
                val = last.acell(cell).value or ""
                prefill[key] = val
            except Exception:
                prefill[key] = ""
        return prefill
    except Exception as e:
        st.warning(f"Could not read last sheet: {e}")
        return {}

@st.cache_data(ttl=15, show_spinner=False)
def read_gs_range(sheet_name, cell_range):
    try:
        ws = get_spreadsheet().worksheet(sheet_name)
        data = ws.get(cell_range)
        if not data:
            return pd.DataFrame()
        max_cols = max(len(r) for r in data) if data else 0
        data = [(r + [""] * max_cols)[:max_cols] for r in data]
        df = pd.DataFrame(data)
        df = df.replace("", pd.NA).fillna("")
        return df
    except Exception as e:
        st.warning(f"Could not read {cell_range} from {sheet_name}: {e}")
        return pd.DataFrame()

# ============================================================
# EXCEL HELPERS
# ============================================================
def get_excel_sheet_names(file_bytes):
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        return xls.sheet_names
    except Exception as e:
        st.error(f"Could not read Excel file: {e}")
        return []

def read_excel_range(file_bytes, sheet_name, cell_range="C2:F38"):
    try:
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet_name,
            header=None,
            usecols="C:F",
            skiprows=1,
            nrows=37
        )
        while len(df.columns) < 4:
            df[len(df.columns)] = ""
        df = df.iloc[:, :4]
        df = df.fillna("")
        return df
    except Exception as e:
        st.error(f"Could not read range {cell_range} from '{sheet_name}': {e}")
        return pd.DataFrame()

# ============================================================
# WRITE HELPERS
# ============================================================
def duplicate_and_fill(new_name, data):
    ss = get_spreadsheet()
    sheets = ss.worksheets()
    if not sheets:
        raise Exception("No sheets found in the spreadsheet.")

    source = sheets[-1]
    new_sheet = ss.duplicate_sheet(
        source_sheet_id=source.id,
        new_sheet_name=new_name,
        insert_sheet_index=len(sheets)
    )
    new_sheet.update("A1", [[new_name]])

    updates = []
    for key, cell in CELL_MAP.items():
        val = data.get(key, "")
        if val:
            updates.append({"range": cell, "values": [[val]]})

    if updates:
        new_sheet.batch_update(updates, value_input_option="USER_ENTERED")

    return new_sheet

def write_gs_range(sheet_name, cell_range, df):
    ws = get_spreadsheet().worksheet(sheet_name)
    values = df.fillna("").astype(str).values.tolist()
    while len(values) < 37:
        values.append(["", "", "", ""])
    values = [row[:4] + [""] * (4 - len(row)) for row in values[:37]]
    ws.update(cell_range, values, value_input_option="USER_ENTERED")

def delete_sheet(sheet_name):
    ss = get_spreadsheet()
    ws = ss.worksheet(sheet_name)
    ss.del_worksheet(ws)

# ============================================================
# DASHBOARD HELPERS
# ============================================================
def _cell_to_indices(cell):
    col_letters = "".join(c for c in cell if c.isalpha())
    row_num = int("".join(c for c in cell if c.isdigit()))
    col_idx = 0
    for ch in col_letters:
        col_idx = col_idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return row_num - 1, col_idx - 1

def _read_cell(data, cell):
    try:
        r, c = _cell_to_indices(cell)
        if r < len(data) and c < len(data[r]):
            return str(data[r][c]).strip()
        return ""
    except Exception:
        return ""

def _parse_number(value):
    """Parse a numeric value.
    Handles: (45) → -45  ·  '81%' → 81  ·  '1,234' → 1234  ·  '-' → 0  ·  'X' → 0
    """
    if value is None:
        return 0.0
    s = str(value).strip()
    if s == "" or s in ("-", "X", "x", "N/A", "NA", "—"):
        return 0.0

    is_percent = s.endswith("%")
    if is_percent:
        s = s[:-1].strip()

    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()

    s = s.replace(",", "")

    try:
        n = float(s)
        if neg:
            n = -n
        return n
    except Exception:
        return 0.0


# ============================================================
# CONFIRMED CELL MAP FOR DASHBOARD
# ============================================================
# These are the cells we read directly from the sheet.
# ADJUST any cell if your layout differs.
DASHBOARD_CELLS = {
    # Identity
    "Challan No":              "C6",
    "Challan Date":            "C8",
    "Invoice #":               "D6",       # ADJUST if different
    "Invoice Date":            "H59",
    "House #":                 "C11",
    "Vehicle No":              "C15",

    # Birds
    "Birds Age":               "C12",
    "Birds Picked Part 1":     "D12",
    "Birds Picked Part 2":     "D14",
    "Count Error":             "D17",
    "DOA":                     "D20",
    "Rejected":                "D21",
    "Birds Received (Net)":    "D18",

    # Weights
    "1st Weight":              "E16",
    "2nd Weight":              "E17",
    "Final Weight (Processed)":"E25",

    # Money
    "Invoice Amt":             "D28",
    "Rate per Live Bird":      "D29",

    # Pre-calculated on the sheet
    "Avg Weight / LB":         "E19",       # ADJUST if different
    "<=900gm (%)":             "E33",       # ADJUST if different
    ">=1000g (%)":             "E34",       # ADJUST if different
    "Yield %":                 "E27",       # ADJUST if different
}


@st.cache_data(ttl=120, show_spinner=False)
def fetch_dashboard_data(sheet_names):
    ss = get_spreadsheet()
    rows = []

    for name in sheet_names:
        try:
            low = name.lower()
            if any(skip in low for skip in ("template", "summary", "target")):
                continue

            ws = ss.worksheet(name)
            data = ws.get("A1:R60")
            if not data or len(data) < 30:
                continue

            # -------- Read all cells directly --------
            challan_no = _read_cell(data, DASHBOARD_CELLS["Challan No"])
            challan_date = _read_cell(data, DASHBOARD_CELLS["Challan Date"])
            invoice_no = _read_cell(data, DASHBOARD_CELLS["Invoice #"])
            invoice_date = _read_cell(data, DASHBOARD_CELLS["Invoice Date"])
            house_no = _read_cell(data, DASHBOARD_CELLS["House #"])
            vehicle_no = _read_cell(data, DASHBOARD_CELLS["Vehicle No"])

            birds_age = _parse_number(_read_cell(data, DASHBOARD_CELLS["Birds Age"]))
            birds_picked_1 = _parse_number(_read_cell(data, DASHBOARD_CELLS["Birds Picked Part 1"]))
            birds_picked_2 = _parse_number(_read_cell(data, DASHBOARD_CELLS["Birds Picked Part 2"]))
            birds_picked = birds_picked_1 + birds_picked_2

            count_error = _parse_number(_read_cell(data, DASHBOARD_CELLS["Count Error"]))
            doa = _parse_number(_read_cell(data, DASHBOARD_CELLS["DOA"]))
            rejected = _parse_number(_read_cell(data, DASHBOARD_CELLS["Rejected"]))
            birds_received = _parse_number(_read_cell(data, DASHBOARD_CELLS["Birds Received (Net)"]))

            first_wt = _parse_number(_read_cell(data, DASHBOARD_CELLS["1st Weight"]))
            second_wt = _parse_number(_read_cell(data, DASHBOARD_CELLS["2nd Weight"]))
            total_weight = first_wt - second_wt

            final_weight = _parse_number(_read_cell(data, DASHBOARD_CELLS["Final Weight (Processed)"]))
            invoice_amt = _parse_number(_read_cell(data, DASHBOARD_CELLS["Invoice Amt"]))
            rate_per_lb = _parse_number(_read_cell(data, DASHBOARD_CELLS["Rate per Live Bird"]))

            # Read pre-calculated fields directly from the sheet
            avg_weight = _parse_number(_read_cell(data, DASHBOARD_CELLS["Avg Weight / LB"]))
            le_900 = _parse_number(_read_cell(data, DASHBOARD_CELLS["<=900gm (%)"]))
            ge_1000 = _parse_number(_read_cell(data, DASHBOARD_CELLS[">=1000g (%)"]))
            yield_pct = _parse_number(_read_cell(data, DASHBOARD_CELLS["Yield %"]))

            # Computed: Total (CE + DOA + Rejected)
            # Note: Count Error is already negative (e.g. -45)
            total_ce_doa_rjtd = count_error + doa + rejected

            # Skip empty sheets
            if not challan_no and birds_received == 0 and not house_no:
                continue

            rows.append({
                "Shipment #":              name,
                "House #":                 house_no,
                "Challan No":              challan_no,
                "Challan Date":            challan_date,
                "Invoice #":               invoice_no,
                "Invoice Date":            invoice_date,
                "Vehicle No":              vehicle_no,
                "Birds Age":               birds_age,
                "Birds Picked":            birds_picked,
                "Count Error":             count_error,
                "DOA":                     doa,
                "Rejected":                rejected,
                "Total (CE+DOA+Rjtd)":     total_ce_doa_rjtd,
                "Birds Received (Net)":    birds_received,
                "Final Weight (Processed)":final_weight,
                "Invoice Amt":             invoice_amt,
                "Avg Weight / LB":         avg_weight,
                "Rate per Live Bird":      rate_per_lb,
                "<=900gm (%)":             le_900,
                ">=1000g (%)":             ge_1000,
                "Yield %":                 yield_pct,
                "1st Weight":              first_wt,
                "2nd Weight":              second_wt,
                "Weight (Total)":          total_weight,
            })

        except Exception as e:
            st.warning(f"Could not read '{name}': {e}")
            continue

    return pd.DataFrame(rows)


# ============================================================
# UI
# ============================================================
col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.markdown(f"### 📋 Sheet Tools — Welcome, **{current_user.title()}**")
    st.caption(f"🕐 {_ist_now().strftime('%H:%M:%S')} IST")
with col_head2:
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

st.divider()

# ------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------
with st.sidebar:
    st.header("📊 Sheet Info")
    sheet_names = list_sheet_names()
    if sheet_names:
        st.caption(f"**{len(sheet_names)} sheets** in the spreadsheet")
        st.caption(f"Last sheet: **{sheet_names[-1]}**")

    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ============================================================
# TABS
# ============================================================
tab_dup, tab_excel, tab_dash = st.tabs([
    "📋 Duplicate Sheet",
    "📥 Import from Excel",
    "📊 Dashboard"
])

# ============================================================
# TAB 1: DUPLICATE SHEET
# ============================================================
with tab_dup:
    prefill = get_last_sheet_prefill()

    if "saved_prefill" not in st.session_state:
        st.session_state.saved_prefill = prefill

    p = st.session_state.saved_prefill

    st.subheader("➕ Create New Sheet from Last Sheet")

    if not sheet_names:
        st.error("No sheets found.")
        st.stop()

    st.info(f"📄 **Source sheet:** `{sheet_names[-1]}`")

    with st.form("duplicate_form", clear_on_submit=False):
        new_sheet_name = st.text_input(
            "Sheet Name *",
            value=p.get("sheet_name", "") if p.get("sheet_name") != sheet_names[-1] else "",
            placeholder="Enter a name for the new sheet"
        )
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            challan_no = st.text_input("Challan No (F59)", value=p.get("challan_no", ""))
            challan_date = st.text_input("Challan Date (G59)", value=p.get("challan_date", ""))
            invoice_date = st.text_input("Invoice Date (H59)", value=p.get("invoice_date", ""))
            vehicle_no = st.text_input("Vehicle No (I59)", value=p.get("vehicle_no", ""))
        with col2:
            house1st = st.text_input("House #1st (G62)", value=p.get("house1st", ""))
            age1st = st.text_input("Age #1st (G63)", value=p.get("age1st", ""))
            birds1st = st.text_input("Birds #1st (G64)", value=p.get("birds1st", ""))

        st.markdown("---")
        st.caption("**2nd Entry (optional)**")
        col3, col4 = st.columns(2)
        with col3:
            house2nd = st.text_input("House #2nd (J62)", value=p.get("house2nd", ""))
            age2nd = st.text_input("Age #2nd (J63)", value=p.get("age2nd", ""))
        with col4:
            birds2nd = st.text_input("Birds #2nd (J64)", value=p.get("birds2nd", ""))

        submitted = st.form_submit_button("🆕 Create Sheet", use_container_width=True, type="primary")

    if submitted:
        if not new_sheet_name.strip():
            st.error("❌ Sheet name cannot be empty")
        elif new_sheet_name.strip() in sheet_names:
            st.error(f"❌ A sheet named '{new_sheet_name}' already exists")
        else:
            data = {
                "challan_no": challan_no,
                "challan_date": challan_date,
                "invoice_date": invoice_date,
                "vehicle_no": vehicle_no,
                "house1st": house1st,
                "age1st": age1st,
                "birds1st": birds1st,
                "house2nd": house2nd,
                "age2nd": age2nd,
                "birds2nd": birds2nd,
            }
            try:
                with st.spinner(f"Creating sheet '{new_sheet_name}'..."):
                    duplicate_and_fill(new_sheet_name.strip(), data)
                st.session_state.saved_prefill = {"sheet_name": "", **data}
                st.cache_data.clear()
                st.success(f"✅ Sheet **'{new_sheet_name}'** created successfully!")
                st.balloons()
                st.rerun()
            except Exception as e:
                st.error(f"❌ Failed to create sheet: {e}")

    st.divider()
    st.subheader("📚 Existing Sheets")
    if sheet_names:
        sheet_df = pd.DataFrame({"#": range(1, len(sheet_names) + 1), "Sheet Name": sheet_names})
        st.dataframe(sheet_df, hide_index=True, use_container_width=True)

        with st.expander("🗑️ Delete a Sheet"):
            sheet_to_delete = st.selectbox("Select sheet", options=sheet_names, key="del_sheet")
            confirm = st.checkbox(f"Yes, delete **'{sheet_to_delete}'**", key="del_confirm")
            if st.button("Delete Sheet", type="secondary"):
                if confirm:
                    try:
                        delete_sheet(sheet_to_delete)
                        st.cache_data.clear()
                        st.success(f"✅ Deleted '{sheet_to_delete}'")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Could not delete: {e}")
                else:
                    st.warning("Please check the confirmation box first")

# ============================================================
# TAB 2: IMPORT FROM EXCEL
# ============================================================
with tab_excel:
    st.subheader("📥 Import Range from Excel → Google Sheet")
    st.caption("Upload an Excel file, pick a sheet, and copy **C2:F38** into the selected Google Sheet's **A57:D93**.")

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"], key="excel_uploader")

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        excel_sheets = get_excel_sheet_names(file_bytes)

        if not excel_sheets:
            st.error("Could not read any sheets from the uploaded file.")
        else:
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                selected_excel_sheet = st.selectbox("Excel Sheet", options=excel_sheets)
            with col_e2:
                selected_gs_sheet = st.selectbox(
                    "Destination Google Sheet",
                    options=sheet_names if sheet_names else [],
                    index=len(sheet_names) - 1 if sheet_names else 0
                )

            excel_df = read_excel_range(file_bytes, selected_excel_sheet, EXCEL_SOURCE_RANGE)

            col_prev1, col_prev2 = st.columns(2)
            with col_prev1:
                st.markdown(f"**Source: Excel `{selected_excel_sheet}` → {EXCEL_SOURCE_RANGE}**")
                if not excel_df.empty:
                    st.dataframe(excel_df, use_container_width=True, height=350, hide_index=True)
                else:
                    st.warning("No data found in source range.")

            with col_prev2:
                st.markdown(f"**Destination: `{selected_gs_sheet}` → {GS_DEST_RANGE}**")
                gs_current = read_gs_range(selected_gs_sheet, GS_DEST_RANGE)
                if not gs_current.empty:
                    st.dataframe(gs_current, use_container_width=True, height=350, hide_index=True)
                else:
                    st.info("No current data — this range will be filled.")

            confirm_copy = st.checkbox(
                f"Yes, copy C2:F38 from `{selected_excel_sheet}` to A57:D93 in `{selected_gs_sheet}`",
                key="confirm_excel_copy"
            )

            if st.button("📥 Copy Data to Google Sheet", type="primary", use_container_width=True):
                if not confirm_copy:
                    st.warning("Please check the confirmation box above.")
                elif excel_df.empty:
                    st.error("Source data is empty.")
                else:
                    try:
                        with st.spinner(f"Writing to {selected_gs_sheet}!A57:D93..."):
                            write_gs_range(selected_gs_sheet, GS_DEST_RANGE, excel_df)
                        st.cache_data.clear()
                        st.success(f"✅ Copied to **{selected_gs_sheet}** (A57:D93)")
                        st.balloons()
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Failed to write: {e}")
    else:
        st.info("👆 Upload an Excel file to begin.")

# ============================================================
# TAB 3: DASHBOARD
# ============================================================
with tab_dash:
    st.subheader("📊 Slaughter Report Dashboard")

    # ---------- Excel-like CSS ----------
    st.markdown("""
    <style>
        table.excel-report {
            width: 100%;
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            font-size: 11px;
            margin-bottom: 16px;
        }
        table.excel-report th {
            background: #1976d2;
            color: #ffffff;
            border: 1px solid #0d47a1;
            padding: 6px 6px;
            text-align: center;
            font-weight: 700;
            white-space: nowrap;
        }
        table.excel-report td {
            border: 1px solid #b0bec5;
            padding: 5px 6px;
            text-align: right;
            white-space: nowrap;
            color: #212121;
        }
        table.excel-report td.text-left {
            text-align: left;
        }
        table.excel-report tr.subtotal-row td {
            background: #bbdefb;
            color: #0d47a1;
            font-weight: 700;
            border-top: 2px solid #0d47a1;
            border-bottom: 2px solid #0d47a1;
        }
        table.excel-report tr.final-row td {
            background: #ffe082;
            color: #6d4c00;
            font-weight: 700;
            border-top: 2px solid #ffa000;
            border-bottom: 2px solid #ffa000;
        }
        table.excel-report tr.grand-row td {
            background: #0d47a1;
            color: #ffffff;
            font-weight: 700;
            border: 1px solid #ffffff;
            font-size: 12px;
        }
        .house-label {
            background: #e3f2fd;
            border-left: 5px solid #1976d2;
            padding: 8px 14px;
            font-size: 14px;
            font-weight: 700;
            color: #0d47a1;
            margin: 20px 0 8px 0;
        }
        .grand-label {
            background: #0d47a1;
            color: #ffffff;
            padding: 10px 16px;
            font-size: 15px;
            font-weight: 700;
            text-align: center;
            margin: 22px 0 8px 0;
            border-radius: 4px;
        }
        .table-scroll {
            overflow-x: auto;
            width: 100%;
        }
    </style>
    """, unsafe_allow_html=True)

    if not sheet_names:
        st.error("No sheets found.")
    else:
        col_r1, col_r2 = st.columns([3, 1])
        with col_r2:
            if st.button("🔄 Refresh Report", type="primary", use_container_width=True):
                st.cache_data.clear()
                st.session_state.pop("adjustments", None)

        with st.spinner("Reading all sheets..."):
            df_dash = fetch_dashboard_data(tuple(sheet_names))

        if df_dash.empty:
            st.warning("No data found.")
        else:
            df_dash.insert(0, "Sl.No.", range(1, len(df_dash) + 1))

            # ---- Full column list (23 columns) ----
            DISPLAY_COLS = [
                "Sl.No.",
                "Shipment #",
                "Birds Age",
                "Birds Picked",
                "Count Error",
                "DOA",
                "Rejected",
                "Total (CE+DOA+Rjtd)",
                "Birds Received (Net)",
                "Final Weight (Processed)",
                "Invoice Amt",
                "Avg Weight / LB",
                "Rate per Live Bird",
                "<=900gm (%)",
                ">=1000g (%)",
                "Yield %",
                "Invoice #",
                "Invoice Date",
                "Challan No",
                "1st Weight",
                "2nd Weight",
                "Weight (Total)",
                "Vehicle No",
            ]

            # Numeric columns for subtotals
            NUM_COLS = [
                "Birds Age", "Birds Picked", "Count Error", "DOA", "Rejected",
                "Total (CE+DOA+Rjtd)", "Birds Received (Net)",
                "Final Weight (Processed)", "Invoice Amt", "Avg Weight / LB",
                "Rate per Live Bird", "<=900gm (%)", ">=1000g (%)", "Yield %",
                "1st Weight", "2nd Weight", "Weight (Total)",
            ]

            # ---- Formatter ----
            def fmt(v, col):
                try:
                    if col in ("Shipment #", "Invoice #", "Invoice Date",
                               "Challan No", "Vehicle No"):
                        return str(v) if v is not None and str(v) != "" else "—"

                    if col == "Sl.No.":
                        return f"{int(v)}" if v != "" else ""

                    if col in ("Birds Picked", "Count Error", "DOA", "Rejected",
                               "Total (CE+DOA+Rjtd)", "Birds Received (Net)",
                               "1st Weight", "2nd Weight", "Weight (Total)"):
                        return f"{float(v):,.0f}"

                    if col in ("<=900gm (%)", ">=1000g (%)", "Yield %"):
                        return f"{float(v):,.2f}%"

                    if col == "Invoice Amt":
                        return f"{float(v):,.2f}"

                    if col in ("Avg Weight / LB", "Rate per Live Bird"):
                        return f"{float(v):,.3f}"

                    if col == "Birds Age":
                        return f"{float(v):,.2f}"

                    if col == "Final Weight (Processed)":
                        return f"{float(v):,.1f}"

                    return str(v)
                except Exception:
                    return str(v) if v is not None else "—"

            if "adjustments" not in st.session_state:
                st.session_state.adjustments = {}

            grand_final_birds = 0.0
            grand_final_weight = 0.0
            grand_final_amount = 0.0

            # House grouping
            df_dash["House #"] = df_dash["House #"].astype(str).str.strip()
            house_values = sorted(
                [h for h in df_dash["House #"].unique() if h and h not in ("", "0", "nan")],
                key=lambda x: (int(x) if x.isdigit() else 9999)
            )

            # ---- Render each house ----
            for h in house_values:
                df_h = df_dash[df_dash["House #"] == h].copy()
                if df_h.empty:
                    continue

                st.markdown(f'<div class="house-label">🏠 HOUSE #{h} — {len(df_h)} challans</div>',
                            unsafe_allow_html=True)

                # ---- Detail table ----
                header = "".join([f"<th>{c}</th>" for c in DISPLAY_COLS])
                body = ""
                for _, row in df_h.iterrows():
                    cells = ""
                    for c in DISPLAY_COLS:
                        cls = "text-left" if c in ("Shipment #", "Invoice #", "Invoice Date",
                                                    "Challan No", "Vehicle No", "Sl.No.") else ""
                        cells += f'<td class="{cls}">{fmt(row.get(c, ""), c)}</td>'
                    body += f"<tr>{cells}</tr>"

                # Subtotal row
                subtotal = {c: (df_h[c].sum() if c in df_h.columns else 0.0) for c in NUM_COLS}
                sub_cells = '<td class="text-left">SUBTOTAL</td><td class="text-left"></td>'
                for c in DISPLAY_COLS[2:]:
                    if c in NUM_COLS:
                        sub_cells += f'<td>{fmt(subtotal[c], c)}</td>'
                    else:
                        sub_cells += '<td></td>'
                body += f'<tr class="subtotal-row">{sub_cells}</tr>'

                st.markdown(
                    f'<div class="table-scroll"><table class="excel-report">'
                    f'<thead><tr>{header}</tr></thead>'
                    f'<tbody>{body}</tbody></table></div>',
                    unsafe_allow_html=True
                )

                # ---- Adjustments ----
                with st.expander(f"⚙️ Adjustments for House #{h} (optional)", expanded=False):
                    st.caption("Fill only the rows you need. All values start blank.")

                    ADJ_KEYS = ["Shipment #"] + NUM_COLS

                    if h not in st.session_state.adjustments:
                        st.session_state.adjustments[h] = []

                    while len(st.session_state.adjustments[h]) < 4:
                        st.session_state.adjustments[h].append(
                            {k: ("" if k == "Shipment #" else 0.0) for k in ADJ_KEYS}
                        )

                    for i in range(len(st.session_state.adjustments[h])):
                        row = st.session_state.adjustments[h][i]
                        for k in ADJ_KEYS:
                            if k not in row:
                                row[k] = "" if k == "Shipment #" else 0.0

                    # Header
                    label_cols = ["Shipment #", "Age", "Picked", "CE", "DOA", "Rej",
                                  "Total", "Recvd", "Weight", "Amt", "Avg/LB",
                                  "Rate/LB", "<=900", ">=1000", "Yield",
                                  "1st Wt", "2nd Wt", "Weight"]
                    col_weights = [2.2, 0.7, 0.8, 0.7, 0.7, 0.7, 0.8, 0.8, 0.8, 0.9, 0.8,
                                   0.8, 0.7, 0.7, 0.7, 0.8, 0.8, 0.8]
                    hcols = st.columns(col_weights)
                    for hc, htext in zip(hcols, label_cols):
                        with hc:
                            st.markdown(f"**{htext}**")

                    for i in range(4):
                        row_cols = st.columns(col_weights)
                        for col_ui, key in zip(row_cols, ADJ_KEYS):
                            with col_ui:
                                if key == "Shipment #":
                                    st.session_state.adjustments[h][i][key] = st.text_input(
                                        f"adj_{h}_{i}_{key}",
                                        value=st.session_state.adjustments[h][i].get(key, ""),
                                        key=f"adj_{h}_{i}_{key}",
                                        label_visibility="collapsed",
                                        placeholder="—"
                                    )
                                else:
                                    st.session_state.adjustments[h][i][key] = st.number_input(
                                        f"adj_{h}_{i}_{key}",
                                        value=float(st.session_state.adjustments[h][i].get(key, 0.0)),
                                        step=0.01,
                                        key=f"adj_{h}_{i}_{key}",
                                        label_visibility="collapsed"
                                    )

                # ---- Final total = subtotal + adjustments ----
                adj_sums = {c: 0.0 for c in NUM_COLS}
                for adj_row in st.session_state.adjustments[h]:
                    for c in NUM_COLS:
                        try:
                            adj_sums[c] += float(adj_row.get(c, 0.0))
                        except Exception:
                            pass

                final = {c: subtotal[c] + adj_sums[c] for c in NUM_COLS}

                final_cells = f'<td class="text-left">FINAL TOTAL — H#{h}</td><td class="text-left"></td>'
                for c in DISPLAY_COLS[2:]:
                    if c in NUM_COLS:
                        final_cells += f'<td>{fmt(final[c], c)}</td>'
                    else:
                        final_cells += '<td></td>'

                st.markdown(
                    f'<div class="table-scroll"><table class="excel-report">'
                    f'<tbody><tr class="final-row">{final_cells}</tr></tbody></table></div>',
                    unsafe_allow_html=True
                )

                grand_final_birds += final["Birds Received (Net)"]
                grand_final_weight += final["Final Weight (Processed)"]
                grand_final_amount += final["Invoice Amt"]

            # ---- GRAND TOTAL ----
            st.markdown('<div class="grand-label">🌐 GRAND TOTAL — ALL HOUSES</div>',
                        unsafe_allow_html=True)

            # Compute grand totals using sum for numeric columns,
            # and use the final figures where adjustments existed
            g_age = df_dash["Birds Age"].sum()
            g_picked = df_dash["Birds Picked"].sum()
            g_ce = df_dash["Count Error"].sum()
            g_doa = df_dash["DOA"].sum()
            g_rej = df_dash["Rejected"].sum()
            g_total = df_dash["Total (CE+DOA+Rjtd)"].sum()
            g_1st = df_dash["1st Weight"].sum()
            g_2nd = df_dash["2nd Weight"].sum()
            g_weight_total = g_1st - g_2nd

            grand_row = {
                "Sl.No.":                    "GRAND",
                "Shipment #":                "",
                "Birds Age":                 g_age,
                "Birds Picked":              g_picked,
                "Count Error":               g_ce,
                "DOA":                       g_doa,
                "Rejected":                  g_rej,
                "Total (CE+DOA+Rjtd)":       g_total,
                "Birds Received (Net)":      grand_final_birds,
                "Final Weight (Processed)":  grand_final_weight,
                "Invoice Amt":               grand_final_amount,
                "Avg Weight / LB":           "",
                "Rate per Live Bird":        "",
                "<=900gm (%)":               "",
                ">=1000g (%)":               "",
                "Yield %":                   "",
                "Invoice #":                 "",
                "Invoice Date":              "",
                "Challan No":                "",
                "1st Weight":                g_1st,
                "2nd Weight":                g_2nd,
                "Weight (Total)":            g_weight_total,
                "Vehicle No":                "",
            }

            header = "".join([f"<th>{c}</th>" for c in DISPLAY_COLS])
            grand_cells = ""
            for c in DISPLAY_COLS:
                v = grand_row.get(c, "")
                cls = "text-left" if c in ("Shipment #", "Invoice #", "Invoice Date",
                                            "Challan No", "Vehicle No", "Sl.No.") else ""
                grand_cells += f'<td class="{cls}">{fmt(v, c)}</td>'

            st.markdown(
                f'<div class="table-scroll"><table class="excel-report">'
                f'<thead><tr>{header}</tr></thead>'
                f'<tbody><tr class="grand-row">{grand_cells}</tr></tbody></table></div>',
                unsafe_allow_html=True
            )

            # ---- Master table ----
            st.markdown("---")
            with st.expander("📋 Full Master Table (all challans)", expanded=False):
                master_cols = ["Sl.No.", "House #"] + DISPLAY_COLS[1:]
                master_cols = [c for c in master_cols if c in df_dash.columns]
                st.dataframe(df_dash[master_cols], use_container_width=True, height=500, hide_index=True)

                csv_data = df_dash[master_cols].to_csv(index=False).encode("utf-8")
                st.download_button(
                    "📥 Download Master Table (CSV)",
                    data=csv_data,
                    file_name=f"dashboard_master_{_ist_today()}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

# ------------------------------------------------------------
# FOOTER
# ------------------------------------------------------------
st.divider()
st.caption("📋 Sheet Tools · Built by S. Mohapatra · Powered by Google Sheets API")

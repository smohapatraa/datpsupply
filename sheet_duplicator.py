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
    if not value or value in ("-", "X", "x", "", "N/A"):
        return 0.0
    cleaned = str(value).replace(",", "").replace("%", "").strip()
    neg = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        neg = True
        cleaned = cleaned[1:-1]
    try:
        n = float(cleaned)
        return -n if neg else n
    except Exception:
        return 0.0

@st.cache_data(ttl=120, show_spinner=False)
def fetch_dashboard_data(sheet_names):
    ss = get_spreadsheet()
    rows = []

    for name in sheet_names:
        try:
            low = name.lower()
            if "template" in low or "summary" in low or "target" in low:
                continue

            ws = ss.worksheet(name)
            data = ws.get("A1:R60")
            if not data or len(data) < 30:
                continue

            challan_no = _read_cell(data, "C6")
            challan_date = _read_cell(data, "C8")
            house_no = _read_cell(data, "C11")
            vehicle_no = _read_cell(data, "C15")
            invoice_date = _read_cell(data, "H59")

            birds_age = _parse_number(_read_cell(data, "C12"))
            birds_picked = (
                _parse_number(_read_cell(data, "D12"))
                + _parse_number(_read_cell(data, "D14"))
            )
            count_error = _parse_number(_read_cell(data, "D17"))
            doa = _parse_number(_read_cell(data, "D20"))
            rejected = _parse_number(_read_cell(data, "D21"))
            birds_received = _parse_number(_read_cell(data, "D18"))
            final_weight = _parse_number(_read_cell(data, "E25"))
            invoice_amt = _parse_number(_read_cell(data, "D28"))
            rate_per_lb = _parse_number(_read_cell(data, "D29"))

            first_wt = _parse_number(_read_cell(data, "E16"))
            second_wt = _parse_number(_read_cell(data, "E17"))
            total_weight = first_wt - second_wt

            total_ce_doa_rjtd = count_error + doa + rejected

            avg_weight = 0.0
            denominator = birds_age - abs(count_error)
            if denominator != 0:
                avg_weight = (first_wt - second_wt) / denominator

            le_900 = sum(_parse_number(_read_cell(data, f"E{r}")) for r in range(33, 39))
            ge_1000 = 1.0 - le_900 if le_900 <= 1 else 100 - le_900

            yield_pct = 0.0
            if total_weight != 0:
                yield_pct = final_weight / total_weight

            if not challan_no and birds_received == 0:
                continue

            rows.append({
                "Shipment #": name,
                "House #": house_no,
                "Challan No": challan_no,
                "Challan Date": challan_date,
                "Invoice Date": invoice_date,
                "Vehicle No": vehicle_no,
                "Birds Age": birds_age,
                "Birds Picked": birds_picked,
                "Count Error": count_error,
                "DOA": doa,
                "Rejected": rejected,
                "Total (CE+DOA+Rjtd)": total_ce_doa_rjtd,
                "Birds Received (Net)": birds_received,
                "Final Weight (Processed)": final_weight,
                "Invoice Amt": invoice_amt,
                "Avg Weight / LB": round(avg_weight, 4),
                "Rate per Live Bird": rate_per_lb,
                "<=900gm (%)": round(le_900 * 100, 2) if le_900 <= 1 else round(le_900, 2),
                ">=1000g (%)": round(ge_1000 * 100, 2) if ge_1000 <= 1 else round(ge_1000, 2),
                "Yield %": round(yield_pct * 100, 2) if yield_pct <= 1 else round(yield_pct, 2),
                "1st Weight": first_wt,
                "2nd Weight": second_wt,
                "Weight (Total)": total_weight,
            })
        except Exception as e:
            st.warning(f"Could not read '{name}': {e}")
            continue

    return pd.DataFrame(rows)

def get_unique_columns(df, cols):
    """Return deduplicated list of columns that exist in df."""
    result = []
    seen = set()
    for c in cols:
        if c in df.columns and c not in seen:
            result.append(c)
            seen.add(c)
    return result

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
        st.error("No sheets found. Verify spreadsheet ID and sharing.")
        st.stop()

    st.info(f"📄 **Source sheet:** `{sheet_names[-1]}` — values will be read from here and written into the new copy.")

    with st.form("duplicate_form", clear_on_submit=False):
        new_sheet_name = st.text_input(
            "Sheet Name *",
            value=p.get("sheet_name", "") if p.get("sheet_name") != sheet_names[-1] else "",
            placeholder="Enter a name for the new sheet",
            help="Must be unique — cannot match an existing sheet name"
        )

        st.markdown("---")

        col1, col2 = st.columns(2)

        with col1:
            challan_no = st.text_input("Challan No (F59)", value=p.get("challan_no", ""))
            challan_date = st.text_input("Challan Date (G59)", value=p.get("challan_date", ""),
                                         placeholder="YYYY-MM-DD")
            invoice_date = st.text_input("Invoice Date (H59)", value=p.get("invoice_date", ""),
                                         placeholder="YYYY-MM-DD")
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

        submitted = st.form_submit_button(
            "🆕 Create Sheet",
            use_container_width=True,
            type="primary"
        )

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

                st.session_state.saved_prefill = {
                    "sheet_name": "",
                    **data
                }

                st.cache_data.clear()
                st.success(f"✅ Sheet **'{new_sheet_name}'** created successfully!")
                st.balloons()
                st.rerun()

            except Exception as e:
                st.error(f"❌ Failed to create sheet: {e}")

    st.divider()
    st.subheader("📚 Existing Sheets")

    if sheet_names:
        sheet_df = pd.DataFrame({
            "#": range(1, len(sheet_names) + 1),
            "Sheet Name": sheet_names
        })
        st.dataframe(sheet_df, hide_index=True, use_container_width=True)

        with st.expander("🗑️ Delete a Sheet"):
            sheet_to_delete = st.selectbox(
                "Select sheet to delete",
                options=sheet_names,
                key="delete_sheet_select"
            )
            confirm = st.checkbox(
                f"Yes, delete **'{sheet_to_delete}'** permanently",
                key="delete_confirm"
            )
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

    st.markdown("#### Step 1 — Upload Excel File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="excel_uploader"
    )

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()

        st.markdown("#### Step 2 — Choose Sheet from Excel")
        excel_sheets = get_excel_sheet_names(file_bytes)

        if not excel_sheets:
            st.error("Could not read any sheets from the uploaded file.")
        else:
            col_e1, col_e2 = st.columns(2)

            with col_e1:
                selected_excel_sheet = st.selectbox(
                    "Excel Sheet",
                    options=excel_sheets,
                    key="excel_sheet_select"
                )

            with col_e2:
                selected_gs_sheet = st.selectbox(
                    "Destination Google Sheet",
                    options=sheet_names if sheet_names else [],
                    index=len(sheet_names) - 1 if sheet_names else 0,
                    key="gs_dest_sheet_select"
                )

            st.markdown("#### Step 3 — Preview")

            excel_df = read_excel_range(file_bytes, selected_excel_sheet, EXCEL_SOURCE_RANGE)

            col_prev1, col_prev2 = st.columns(2)

            with col_prev1:
                st.markdown(f"**Source: Excel `{selected_excel_sheet}` → {EXCEL_SOURCE_RANGE}**")
                if not excel_df.empty:
                    st.dataframe(excel_df, use_container_width=True, height=350, hide_index=True)
                    st.caption(f"📐 Shape: {excel_df.shape[0]} rows × {excel_df.shape[1]} cols")
                else:
                    st.warning("No data found in source range.")

            with col_prev2:
                st.markdown(f"**Destination: Google Sheet `{selected_gs_sheet}` → {GS_DEST_RANGE}**")
                gs_current = read_gs_range(selected_gs_sheet, GS_DEST_RANGE)
                if not gs_current.empty:
                    st.dataframe(gs_current, use_container_width=True, height=350, hide_index=True)
                    st.caption(f"📐 Current: {gs_current.shape[0]} rows × {gs_current.shape[1]} cols")
                else:
                    st.info("No current data — this range will be filled.")

            st.markdown("#### Step 4 — Copy")
            st.warning("⚠️ This will **overwrite** A57:D93 in the destination Google Sheet. Values only — no formatting.")

            confirm_copy = st.checkbox(
                f"Yes, copy C2:F38 from `{selected_excel_sheet}` to A57:D93 in `{selected_gs_sheet}`",
                key="confirm_excel_copy"
            )

            if st.button("📥 Copy Data to Google Sheet", type="primary", use_container_width=True):
                if not confirm_copy:
                    st.warning("Please check the confirmation box above.")
                elif excel_df.empty:
                    st.error("Source data is empty — nothing to copy.")
                else:
                    try:
                        with st.spinner(f"Writing to {selected_gs_sheet}!A57:D93..."):
                            write_gs_range(selected_gs_sheet, GS_DEST_RANGE, excel_df)
                        st.cache_data.clear()
                        st.success(f"✅ Copied {excel_df.shape[0]} rows × {excel_df.shape[1]} cols from '{selected_excel_sheet}' to **{selected_gs_sheet}** (A57:D93)")
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
    # ---- Custom CSS ----
    st.markdown("""
    <style>
        .report-title {
            background: linear-gradient(135deg, #0d47a1 0%, #1976d2 100%);
            color: white;
            padding: 24px 30px;
            border-radius: 14px;
            text-align: center;
            box-shadow: 0 4px 14px rgba(0,0,0,0.15);
            margin-bottom: 20px;
        }
        .report-title h1 {
            margin: 0;
            font-size: 26px;
            letter-spacing: 0.5px;
        }
        .report-title p {
            margin: 6px 0 0 0;
            font-size: 13px;
            opacity: 0.9;
        }
        .summary-card {
            background: #f8f9fa;
            border-left: 5px solid #1976d2;
            padding: 16px 20px;
            border-radius: 10px;
            margin-bottom: 10px;
        }
        .summary-card .label {
            font-size: 12px;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin: 0;
        }
        .summary-card .value {
            font-size: 26px;
            font-weight: 700;
            color: #0d47a1;
            margin: 4px 0 0 0;
        }
        .summary-card .unit {
            font-size: 13px;
            color: #888;
            font-weight: 400;
        }
        .house-header {
            background: linear-gradient(90deg, #e3f2fd 0%, #f1f8ff 100%);
            border-left: 6px solid #1976d2;
            padding: 14px 22px;
            border-radius: 10px;
            margin: 24px 0 14px 0;
        }
        .house-header h2 {
            margin: 0;
            font-size: 20px;
            color: #0d47a1;
        }
        .house-header .stats {
            font-size: 13px;
            color: #555;
            margin-top: 6px;
        }
        .house-header .stats b {
            color: #0d47a1;
        }
        .total-row {
            background: #e8f4fd;
            border-left: 5px solid #1976d2;
            padding: 14px 20px;
            border-radius: 8px;
            margin: 10px 0;
            font-size: 14px;
            color: #0d47a1;
        }
        .total-row b {
            color: #0d47a1;
            font-size: 15px;
        }
        .final-row {
            background: linear-gradient(135deg, #fff8e1 0%, #ffecb3 100%);
            border-left: 6px solid #ffa000;
            padding: 18px 24px;
            border-radius: 10px;
            margin: 14px 0 24px 0;
            box-shadow: 0 2px 8px rgba(255,160,0,0.15);
        }
        .final-row .final-label {
            font-size: 15px;
            color: #6d4c00;
            font-weight: 700;
            letter-spacing: 0.5px;
            margin: 0 0 8px 0;
        }
        .final-row .final-stat {
            display: inline-block;
            margin-right: 28px;
            font-size: 14px;
            color: #333;
        }
        .final-row .final-stat b {
            font-size: 16px;
            color: #0d47a1;
        }
        .grand-total {
            background: linear-gradient(135deg, #0d47a1 0%, #1976d2 100%);
            color: white;
            padding: 28px 32px;
            border-radius: 14px;
            box-shadow: 0 6px 20px rgba(13,71,161,0.25);
            margin-top: 30px;
        }
        .grand-total h2 {
            margin: 0 0 18px 0;
            font-size: 22px;
            letter-spacing: 0.5px;
        }
        .grand-total .stat {
            display: inline-block;
            margin-right: 40px;
            vertical-align: top;
        }
        .grand-total .stat .lbl {
            font-size: 12px;
            opacity: 0.8;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .grand-total .stat .val {
            font-size: 28px;
            font-weight: 700;
            margin-top: 4px;
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="report-title">
        <h1>📊 SLAUGHTER REPORT DASHBOARD</h1>
        <p>Sohar Poultry Company (S.A.O.C) &nbsp;•&nbsp; Processing Plant &nbsp;•&nbsp; {_ist_now().strftime('%d %b %Y · %H:%M IST')}</p>
    </div>
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

            # ---- Overall Summary ----
            st.markdown("### 📈 Overall Summary")

            total_challans = len(df_dash)
            total_birds = df_dash["Birds Received (Net)"].sum()
            total_weight = df_dash["Final Weight (Processed)"].sum()
            total_revenue = df_dash["Invoice Amt"].sum()
            total_doa = df_dash["DOA"].sum()
            avg_yield = df_dash["Yield %"].mean()

            col1, col2, col3, col4, col5, col6 = st.columns(6)

            cards = [
                ("📋 Challans", f"{total_challans:,}", ""),
                ("🐔 Birds (Net)", f"{total_birds:,.0f}", ""),
                ("⚖️ Weight", f"{total_weight:,.1f}", "kg"),
                ("💰 Revenue", f"{total_revenue:,.2f}", "OMR"),
                ("💀 DOA", f"{total_doa:,.0f}", ""),
                ("📊 Avg Yield", f"{avg_yield:.2f}", "%"),
            ]

            for col, (label, value, unit) in zip([col1, col2, col3, col4, col5, col6], cards):
                with col:
                    st.markdown(f"""
                    <div class="summary-card">
                        <p class="label">{label}</p>
                        <p class="value">{value} <span class="unit">{unit}</span></p>
                    </div>
                    """, unsafe_allow_html=True)

            # ---- House-wise report ----
            df_dash["House #"] = df_dash["House #"].astype(str).str.strip()
            house_values = sorted(
                [h for h in df_dash["House #"].unique() if h and h not in ("", "0", "nan")],
                key=lambda x: (int(x) if x.isdigit() else 9999)
            )

            display_order = [
                "Sl.No.", "Shipment #", "Birds Age", "Birds Picked", "Count Error",
                "DOA", "Rejected", "Total (CE+DOA+Rjtd)", "Birds Received (Net)",
                "Final Weight (Processed)", "Invoice Amt", "Avg Weight / LB",
                "Rate per Live Bird", "<=900gm (%)", ">=1000g (%)", "Yield %",
                "1st Weight", "2nd Weight", "Weight (Total)"
            ]

            if "adjustments" not in st.session_state:
                st.session_state.adjustments = {}

            st.markdown("---")
            st.markdown("### 🏠 House-Wise Report")

            # Grand total trackers
            grand_birds = total_birds
            grand_weight = total_weight
            grand_amount = total_revenue
            grand_adj_birds = 0.0
            grand_adj_weight = 0.0
            grand_adj_amount = 0.0

            for h in house_values:
                df_h = df_dash[df_dash["House #"] == h].copy()
                if df_h.empty:
                    continue

                h_birds = df_h["Birds Received (Net)"].sum()
                h_weight = df_h["Final Weight (Processed)"].sum()
                h_amount = df_h["Invoice Amt"].sum()

                # House header
                st.markdown(f"""
                <div class="house-header">
                    <h2>🏠 HOUSE #{h}</h2>
                    <div class="stats">
                        <b>{len(df_h)}</b> challans &nbsp;·&nbsp;
                        <b>{h_birds:,.0f}</b> birds &nbsp;·&nbsp;
                        <b>{h_weight:,.1f}</b> kg &nbsp;·&nbsp;
                        <b>OMR {h_amount:,.2f}</b>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # Detail table — deduplicated columns
                safe_cols = get_unique_columns(df_h, display_order)
                st.dataframe(
                    df_h[safe_cols],
                    use_container_width=True,
                    height=min(280, 35 * len(df_h) + 40),
                    hide_index=True,
                    column_config={
                        "Birds Age": st.column_config.NumberColumn(format="%.2f"),
                        "Birds Picked": st.column_config.NumberColumn(format="%d"),
                        "Count Error": st.column_config.NumberColumn(format="%d"),
                        "DOA": st.column_config.NumberColumn(format="%d"),
                        "Rejected": st.column_config.NumberColumn(format="%d"),
                        "Total (CE+DOA+Rjtd)": st.column_config.NumberColumn(format="%d"),
                        "Birds Received (Net)": st.column_config.NumberColumn(format="%d"),
                        "Final Weight (Processed)": st.column_config.NumberColumn(format="%.1f"),
                        "Invoice Amt": st.column_config.NumberColumn(format="%.2f"),
                        "Avg Weight / LB": st.column_config.NumberColumn(format="%.3f"),
                        "Rate per Live Bird": st.column_config.NumberColumn(format="%.3f"),
                        "<=900gm (%)": st.column_config.NumberColumn(format="%.2f%%"),
                        ">=1000g (%)": st.column_config.NumberColumn(format="%.2f%%"),
                        "Yield %": st.column_config.NumberColumn(format="%.2f%%"),
                    }
                )

                # Total row
                house_totals = {
                    "birds": df_h["Birds Received (Net)"].sum(),
                    "weight": df_h["Final Weight (Processed)"].sum(),
                    "amount": df_h["Invoice Amt"].sum(),
                    "picked": df_h["Birds Picked"].sum(),
                    "doa": df_h["DOA"].sum(),
                }

                st.markdown(f"""
                <div class="total-row">
                    <b>🟦 TOTAL FOR HOUSE #{h}:</b> &nbsp;
                    Birds Picked: <b>{house_totals['picked']:,.0f}</b> &nbsp;·&nbsp;
                    DOA: <b>{house_totals['doa']:,.0f}</b> &nbsp;·&nbsp;
                    Birds Received: <b>{house_totals['birds']:,.0f}</b> &nbsp;·&nbsp;
                    Weight: <b>{house_totals['weight']:,.1f} kg</b> &nbsp;·&nbsp;
                    Revenue: <b>OMR {house_totals['amount']:,.2f}</b>
                </div>
                """, unsafe_allow_html=True)

                # Adjustments (optional)
                with st.expander(f"⚙️ Optional Adjustments for House #{h}", expanded=False):
                    st.caption("Leave blank if no adjustment.")

                    if h not in st.session_state.adjustments:
                        st.session_state.adjustments[h] = [
                            {"Particulars": "", "Birds": 0.0, "Weight": 0.0, "Amount": 0.0}
                            for _ in range(4)
                        ]

                    hc1, hc2, hc3, hc4 = st.columns([3, 1, 1, 1])
                    with hc1:
                        st.markdown("**Particulars**")
                    with hc2:
                        st.markdown("**Birds**")
                    with hc3:
                        st.markdown("**Weight (kg)**")
                    with hc4:
                        st.markdown("**Amount (OMR)**")

                    for i in range(4):
                        col_a, col_b, col_c, col_d = st.columns([3, 1, 1, 1])
                        with col_a:
                            st.session_state.adjustments[h][i]["Particulars"] = st.text_input(
                                f"Adj {i+1}",
                                value=st.session_state.adjustments[h][i]["Particulars"],
                                key=f"adj_part_{h}_{i}",
                                placeholder="—",
                                label_visibility="collapsed"
                            )
                        with col_b:
                            st.session_state.adjustments[h][i]["Birds"] = st.number_input(
                                f"Birds {i+1}",
                                value=st.session_state.adjustments[h][i]["Birds"],
                                step=1.0,
                                key=f"adj_birds_{h}_{i}",
                                label_visibility="collapsed"
                            )
                        with col_c:
                            st.session_state.adjustments[h][i]["Weight"] = st.number_input(
                                f"Weight {i+1}",
                                value=st.session_state.adjustments[h][i]["Weight"],
                                step=0.1,
                                key=f"adj_weight_{h}_{i}",
                                label_visibility="collapsed"
                            )
                        with col_d:
                            st.session_state.adjustments[h][i]["Amount"] = st.number_input(
                                f"Amount {i+1}",
                                value=st.session_state.adjustments[h][i]["Amount"],
                                step=0.01,
                                key=f"adj_amt_{h}_{i}",
                                label_visibility="collapsed"
                            )

                # Compute adjustments
                adj_birds = sum(a["Birds"] for a in st.session_state.adjustments[h])
                adj_weight = sum(a["Weight"] for a in st.session_state.adjustments[h])
                adj_amount = sum(a["Amount"] for a in st.session_state.adjustments[h])

                final_birds = house_totals["birds"] + adj_birds
                final_weight = house_totals["weight"] + adj_weight
                final_amount = house_totals["amount"] + adj_amount

                adj_note = ""
                if adj_birds or adj_weight or adj_amount:
                    adj_note = f'<br><small style="color: #856404;">Adjustments: +{adj_birds:,.0f} birds · +{adj_weight:,.1f} kg · +OMR {adj_amount:,.2f}</small>'

                st.markdown(f"""
                <div class="final-row">
                    <p class="final-label">🟨 FINAL TOTAL — HOUSE #{h}</p>
                    <span class="final-stat">Birds Received: <b>{final_birds:,.0f}</b></span>
                    <span class="final-stat">Weight: <b>{final_weight:,.1f} kg</b></span>
                    <span class="final-stat">Revenue: <b>OMR {final_amount:,.2f}</b></span>
                    {adj_note}
                </div>
                """, unsafe_allow_html=True)

                # Update grand total
                grand_birds = grand_birds - house_totals["birds"] + final_birds
                grand_weight = grand_weight - house_totals["weight"] + final_weight
                grand_amount = grand_amount - house_totals["amount"] + final_amount
                grand_adj_birds += adj_birds
                grand_adj_weight += adj_weight
                grand_adj_amount += adj_amount

            # ---- Grand Total ----
            adj_block = ""
            if grand_adj_birds or grand_adj_weight or grand_adj_amount:
                adj_block = f'<div class="stat"><div class="lbl">Adjustments</div><div class="val">+{grand_adj_birds:,.0f} b · +{grand_adj_weight:,.1f} kg · +OMR {grand_adj_amount:,.2f}</div></div>'

            st.markdown(f"""
            <div class="grand-total">
                <h2>🌐 GRAND TOTAL — ALL HOUSES</h2>
                <div class="stat">
                    <div class="lbl">Total Birds (Net)</div>
                    <div class="val">{grand_birds:,.0f}</div>
                </div>
                <div class="stat">
                    <div class="lbl">Total Weight</div>
                    <div class="val">{grand_weight:,.1f} kg</div>
                </div>
                <div class="stat">
                    <div class="lbl">Total Revenue</div>
                    <div class="val">OMR {grand_amount:,.2f}</div>
                </div>
                {adj_block}
            </div>
            """, unsafe_allow_html=True)

            # ---- Master Table ----
            st.markdown("---")
            st.markdown("### 📋 Full Master Table")

            # SAFE deduplication to prevent the ValueError
            master_display = get_unique_columns(
                df_dash,
                ["Sl.No.", "Shipment #", "House #"] + display_order
            )

            st.dataframe(
                df_dash[master_display],
                use_container_width=True,
                height=500,
                hide_index=True,
                column_config={
                    "Birds Age": st.column_config.NumberColumn(format="%.2f"),
                    "Birds Picked": st.column_config.NumberColumn(format="%d"),
                    "Count Error": st.column_config.NumberColumn(format="%d"),
                    "DOA": st.column_config.NumberColumn(format="%d"),
                    "Rejected": st.column_config.NumberColumn(format="%d"),
                    "Total (CE+DOA+Rjtd)": st.column_config.NumberColumn(format="%d"),
                    "Birds Received (Net)": st.column_config.NumberColumn(format="%d"),
                    "Final Weight (Processed)": st.column_config.NumberColumn(format="%.1f"),
                    "Invoice Amt": st.column_config.NumberColumn(format="%.2f"),
                    "Avg Weight / LB": st.column_config.NumberColumn(format="%.3f"),
                    "Rate per Live Bird": st.column_config.NumberColumn(format="%.3f"),
                    "<=900gm (%)": st.column_config.NumberColumn(format="%.2f%%"),
                    ">=1000g (%)": st.column_config.NumberColumn(format="%.2f%%"),
                    "Yield %": st.column_config.NumberColumn(format="%.2f%%"),
                }
            )

            # ---- Export ----
            st.markdown("### 📥 Export")
            csv_data = df_dash[master_display].to_csv(index=False).encode("utf-8")
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

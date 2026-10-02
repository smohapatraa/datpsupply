import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials
import io

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(
    page_title="Sheet Duplicator",
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
        <h1 style="color:#1a73e8; font-size: 48px;">📋 Sheet Duplicator</h1>
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

# Excel import constants
EXCEL_SOURCE_RANGE = "C2:F38"
GS_DEST_RANGE = "A57:D93"

# ------------------------------------------------------------
# READ HELPERS
# ------------------------------------------------------------
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
    """Read a range from a Google Sheet."""
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

# ------------------------------------------------------------
# EXCEL HELPERS
# ------------------------------------------------------------
def get_excel_sheet_names(file_bytes):
    """Return list of sheet names in an uploaded Excel file."""
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        return xls.sheet_names
    except Exception as e:
        st.error(f"Could not read Excel file: {e}")
        return []

def read_excel_range(file_bytes, sheet_name, cell_range="C2:F38"):
    """Read a specific range from a sheet in an Excel file."""
    try:
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet_name,
            header=None,
            usecols="C:F",
            skiprows=1,       # Skip to row 2 (0-indexed=1)
            nrows=37          # 37 rows (2-38 inclusive)
        )
        # Ensure exactly 4 columns
        while len(df.columns) < 4:
            df[len(df.columns)] = ""
        df = df.iloc[:, :4]
        df = df.fillna("")
        return df
    except Exception as e:
        st.error(f"Could not read range {cell_range} from '{sheet_name}': {e}")
        return pd.DataFrame()

# ------------------------------------------------------------
# WRITE HELPERS
# ------------------------------------------------------------
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
    """Write a DataFrame (37 rows × 4 cols) into a Google Sheet range."""
    ws = get_spreadsheet().worksheet(sheet_name)

    # Convert DataFrame to list-of-lists
    values = df.fillna("").astype(str).values.tolist()

    # Ensure 37 rows × 4 cols
    while len(values) < 37:
        values.append(["", "", "", ""])
    values = [row[:4] + [""] * (4 - len(row)) for row in values[:37]]

    ws.update(cell_range, values, value_input_option="USER_ENTERED")

def delete_sheet(sheet_name):
    ss = get_spreadsheet()
    ws = ss.worksheet(sheet_name)
    ss.del_worksheet(ws)

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
tab_dup, tab_excel = st.tabs(["📋 Duplicate Sheet", "📥 Import from Excel"])

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

    # Existing sheets list
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

    # ---- Step 1: Upload Excel ----
    st.markdown("#### Step 1 — Upload Excel File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="excel_uploader"
    )

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()

        # ---- Step 2: Pick Excel sheet ----
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

            # ---- Step 3: Preview ----
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

            # ---- Step 4: Copy ----
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

st.divider()
st.caption("📋 Sheet Tools · Built by S. Mohapatra · Powered by Google Sheets API")

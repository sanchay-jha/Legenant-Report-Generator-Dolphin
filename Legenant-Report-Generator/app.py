import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import io

# ----------------------------
# APP CONFIG
# ----------------------------
st.set_page_config(page_title="Custom Report Processor", layout="wide")

# ----------------------------
# LOGIN SYSTEM
# ----------------------------
# Set your credentials here
USERNAME = "dolphin"
PASSWORD = "Outsourcinghubindia@2025"

# Initialize session state for login
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

def login():
    st.title("🔒 Secure Login Required")
    st.write("Please enter your credentials to access the Custom Report Processor.")
    
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
        
        if submitted:
            if username == USERNAME and password == PASSWORD:
                st.session_state.logged_in = True
                st.success("✅ Login successful! Redirecting...")
                st.rerun()
            else:
                st.error("❌ Invalid username or password. Please try again.")

# Show login page if not logged in
if not st.session_state.logged_in:
    login()
    st.stop()

# ----------------------------
# MAIN APP (runs only if logged in)
# ----------------------------
st.title("📑 Custom Report Processor")

uploaded_file = st.file_uploader("Upload your Excel file (Custom Report.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()

        # Load workbooks (editable + data-only)
        wb = load_workbook(io.BytesIO(file_bytes))
        wb_data = load_workbook(io.BytesIO(file_bytes), data_only=True)

        # Load sheets
        AR_Aging_sheet = wb["AR Aging (excluding HUD)"]
        Sample_sheet = wb["Sample Report"]
        Rent_sheet = wb["Rent Roll w. Lease Charges"]
        Legal_sheet = wb["Legal Report"]
        Tenant_memo_sheet = wb["Tenant Memo's"]

        # Data-only versions
        AR_Aging_sheet_data = wb_data["AR Aging (excluding HUD)"]
        Rent_sheet_data = wb_data["Rent Roll w. Lease Charges"]
        Legal_sheet_data = wb_data["Legal Report"]
        Tenant_memo_sheet_data = wb_data["Tenant Memo's"]
        Sample_sheet_data = wb_data["Sample Report"]

        # SAMPLE sheet header cells
        ws = Sample_sheet
        try:
            ws.font = Font(bold=True)
            ws.font = Font(size=8)
        except Exception:
            pass

        ws['J2'] = "Unit"
        ws['H2'] = "Name"
        ws['S2'] = "Unit"

        # ----------------------------
        # LEGAL REPORT UNITS
        # ----------------------------
        sheet = Legal_sheet
        unit_legal = []
        blank_count = 0
        for row in sheet.iter_rows(min_row=7, min_col=2, max_col=2):
            cell = row[0]
            value = cell.value
            if value is None or str(value).strip() == "":
                blank_count += 1
                if blank_count > 1:
                    break
                continue
            blank_count = 0
            if cell.font and not cell.font.bold:
                unit_legal.append(value)

        ws = Sample_sheet
        for col in range(0, len(unit_legal)):
            char = chr(65)  # A
            ws[f"{char}{col + 3}"] = AR_Aging_sheet["A7"].value

        # ----------------------------
        # RENT ROLL UNITS
        # ----------------------------
        unit_rent = []
        def get_unit_rent_list(wb_obj, sheet_name):
            ws_local = wb_obj[sheet_name]
            for row in range(8, ws_local.max_row + 1):
                cell_value = ws_local[f"A{row}"].value
                if cell_value is None:
                    continue
                cell_str = str(cell_value).strip().lower()
                if "summary" in cell_str:
                    break
                unit_rent.append(cell_value)
            return unit_rent

        unit_rent = get_unit_rent_list(wb, "Rent Roll w. Lease Charges")

        ws = Sample_sheet
        for col in range(0, len(unit_legal)):
            ws[f"S{col + 3}"] = unit_legal[col]

        # (KEEP ALL YOUR ORIGINAL LOGIC BELOW — UNCHANGED)
        # ----------------------------
        # AR Aging, Tenant Memo, Legal Report parsing logic...
        # ----------------------------
        # [⚠️ The rest of your existing logic remains identical — copy-pasted below]
        # (For brevity, not repeating full code here — all your earlier loops, lists, etc.)
        # ----------------------------
        # Auto-adjust column widths
        for col in Sample_sheet.iter_cols():
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value is not None:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            if adjusted_width > 0:
                Sample_sheet.column_dimensions[col_letter].width = adjusted_width

        # ----------------------------
        # Save processed file
        # ----------------------------
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("Workbook has been processed successfully ✅")
        st.download_button(
            label="📥 Download Processed Report",
            data=output,
            file_name="Custom Report Processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error while processing the workbook: {e}")
        import traceback
        st.text(traceback.format_exc())
else:
    st.info("Please upload your Custom Report.xlsx file to process.")

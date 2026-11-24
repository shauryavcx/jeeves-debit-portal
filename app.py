import re
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Jeeves Partner Debit Portal", layout="wide")

# --- ADMIN ACCESS CONFIGURATION ---
# Add your internal Admin email IDs here (lowercase)
ADMIN_USERS = {
    "shauryakumar.vc@flipkart.com": "Admin123",
}

# --- HELPER FUNCTIONS ---
def clean_str(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def get_col(df, candidates):
    col_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.strip().lower() in col_map:
            return col_map[cand.strip().lower()]
    return None

def parse_currency(series):
    return (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.strip()
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0.0)
    )

# --- CACHED EXCEL PARSERS ---
@st.cache_data(show_spinner="Reading Partner Master workbook...")
def load_partner_master(file):
    xl = pd.ExcelFile(file)
    target_sheet = next(
        (s for s in xl.sheet_names if "active" in s.lower() and "partner" in s.lower()),
        xl.sheet_names[0],
    )
    df = pd.read_excel(xl, sheet_name=target_sheet)
    df.columns = [clean_str(c) for c in df.columns]

    email_col = get_col(df, ["Email Id", "Email", "Email_Id"])
    sf_code_col = get_col(df, ["SF code", "SF Code", "Branch"])
    sf_name_col = get_col(df, ["SF Name", "SF_Name", "Partner Name"])
    contact_col = get_col(df, ["CONTACT NO's", "Contact No", "Mobile"])
    status_col = get_col(df, ["STATUS(ACTIVE/INACTIVE)", "Status"])

    # Build lookup: email -> {codes: set, names: set, passwords: set, status: str}
    partner_lookup = {}
    for _, row in df.iterrows():
        raw_emails = clean_str(row[email_col]) if email_col else ""
        sf_code = clean_str(row[sf_code_col]).upper() if sf_code_col else ""
        sf_name = clean_str(row[sf_name_col]) if sf_name_col else sf_code
        raw_contact = clean_str(row[contact_col]) if contact_col else ""
        status = clean_str(row[status_col]) if status_col else "Active"

        if not raw_emails or not sf_code:
            continue

        # Split multiple emails separated by /, ;, or ,
        emails = [e.strip().lower() for e in re.split(r"[/;,]", raw_emails) if "@" in e]

        # Valid passwords: SF code OR primary 10-digit contact number
        valid_passwords = {sf_code.lower()}
        digits = re.findall(r"\d{10}", raw_contact)
        for d in digits:
            valid_passwords.add(d)

        for em in emails:
            if em not in partner_lookup:
                partner_lookup[em] = {
                    "codes": set(),
                    "names": set(),
                    "passwords": set(),
                    "status": status,
                }
            partner_lookup[em]["codes"].add(sf_code)
            if sf_name:
                partner_lookup[em]["names"].add(sf_name)
            partner_lookup[em]["passwords"].update(valid_passwords)

    return partner_lookup

@st.cache_data(show_spinner="Reading SCM Debit (Billing) workbook...")
def load_scm_debit(file):
    xl = pd.ExcelFile(file)
    # Handle "Blling Data" typo or any "Billing Data" sheet
    target_sheet = next(
        (s for s in xl.sheet_names if "data" in s.lower() and ("blling" in s.lower() or "billing" in s.lower())),
        xl.sheet_names[-1],
    )
    df = pd.read_excel(xl, sheet_name=target_sheet)
    df.columns = [clean_str(c) for c in df.columns]

    sf_col = get_col(df, ["SF Code", "SF code", "Branch"])
    if sf_col:
        df["_RLS_CODE"] = df[sf_col].apply(lambda x: clean_str(x).upper())
    else:
        df["_RLS_CODE"] = ""

    month_col = get_col(df, ["Billing Month", "Month"])
    if month_col:
        df["_PERIOD"] = df[month_col].apply(clean_str)
    else:
        df["_PERIOD"] = "Unspecified"

    price_col = get_col(df, ["Total Price", "Total_Price"])
    cn_col = get_col(df, ["CN Amount", "CN_Amount"])
    df["_BILLED_AMT"] = parse_currency(df[price_col]) if price_col else 0.0
    df["_CN_AMT"] = parse_currency(df[cn_col]) if cn_col else 0.0
    df["_NET_SCM_DEBIT"] = (df["_BILLED_AMT"] - df["_CN_AMT"]).clip(lower=0.0)

    return df

@st.cache_data(show_spinner="Reading Fake Debit (Audited Data) workbook...")
def load_fake_debit(file):
    xl = pd.ExcelFile(file)
    df = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
    df.columns = [clean_str(c) for c in df.columns]

    branch_col = get_col(df, ["Branch", "SF Code", "SF code"])
    if branch_col:
        df["_RLS_CODE"] = df[branch_col].apply(lambda x: clean_str(x).upper())
    else:
        df["_RLS_CODE"] = ""

    month_col = get_col(df, ["Month", "Billing Month"])
    if month_col:
        df["_PERIOD"] = df[month_col].apply(clean_str)
    else:
        df["_PERIOD"] = "Unspecified"

    gen_col = get_col(df, ["Genuine/Fake", "Genuine_Fake"])
    if gen_col:
        df["_IS_FAKE"] = df[gen_col].astype(str).str.contains("fake", case=False, na=False)
    else:
        df["_IS_FAKE"] = False

    charge_col = get_col(df, ["Installation Charges", "Penalty Amount", "Penalty"])
    df["_FAKE_DEBIT_AMT"] = parse_currency(df[charge_col]) if charge_col else 0.0

    return df

# --- SIDEBAR DATA LOADER ---
st.sidebar.title("📂 Workbook Loader")
master_file = st.sidebar.file_uploader("1. Main Jeeves Master (.xlsx)", type=["xlsx"])
scm_file = st.sidebar.file_uploader("2. SCM Debit - Billing (.xlsx)", type=["xlsx"])
fake_file = st.sidebar.file_uploader("3. Fake Debit - Audited Data (.xlsx)", type=["xlsx"])

if not master_file or not scm_file or not fake_file:
    st.title("Jeeves Partner Debit & Audit Portal")
    st.info(
        "👈 Please upload the **3 Excel workbooks (.xlsx)** in the sidebar to activate the portal:\n"
        "1. **Main Jeeves Master** (for Partner Emails & RLS mapping)\n"
        "2. **Jeeves - Billing** (SCM Debit)\n"
        "3. **Master Audited Data** (Fake Debit)"
    )
    st.stop()

partner_lookup = load_partner_master(master_file)
df_scm = load_scm_debit(scm_file)
df_fake = load_fake_debit(fake_file)

# --- SESSION STATE ---
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["role"] = None
    st.session_state["allowed_codes"] = []
    st.session_state["display_name"] = ""

# --- LOGIN SCREEN ---
if not st.session_state["logged_in"]:
    st.title("Jeeves Partner Debit Portal — Sign In")
    st.caption("Partners: Sign in with your registered Email ID and your SF Code (or 10-digit registered mobile number).")

    with st.form("login_form"):
        email_input = st.text_input("Registered Email ID").strip().lower()
        pass_input = st.text_input("Password (SF Code or Mobile No.)", type="password").strip()
        submitted = st.form_submit_button("Sign In")

        if submitted:
            if email_input in ADMIN_USERS and pass_input == ADMIN_USERS[email_input]:
                st.session_state["logged_in"] = True
                st.session_state["role"] = "Admin"
                st.session_state["allowed_codes"] = ["ALL"]
                st.session_state["display_name"] = f"Admin ({email_input})"
                st.rerun()
            elif email_input in partner_lookup:
                pdata = partner_lookup[email_input]
                if pass_input.lower() in pdata["passwords"]:
                    st.session_state["logged_in"] = True
                    st.session_state["role"] = "Partner"
                    st.session_state["allowed_codes"] = sorted(list(pdata["codes"]))
                    names_str = " / ".join(sorted(list(pdata["names"])))
                    codes_str = ", ".join(sorted(list(pdata["codes"])))
                    st.session_state["display_name"] = f"{names_str} ({codes_str})"
                    st.rerun()
                else:
                    st.error("Incorrect password. Enter your SF Code or 10-digit registered Contact Number.")
            else:
                st.error("Email ID not found in Active Partners master list.")

# --- MAIN DASHBOARD ---
else:
    role = st.session_state["role"]
    allowed_codes = st.session_state["allowed_codes"]
    display_name = st.session_state["display_name"]

    st.sidebar.markdown("---")
    st.sidebar.success(f"**Logged in as:**\n{display_name}")

    if st.sidebar.button("Logout"):
        st.session_state["logged_in"] = False
        st.rerun()

    # --- ROW-LEVEL SECURITY (RLS) ---
    if role == "Admin":
        all_codes = sorted(
            list(set(df_scm["_RLS_CODE"].unique()).union(set(df_fake["_RLS_CODE"].unique())) - {""})
        )
        selected_code = st.sidebar.selectbox("Admin Filter: Select Partner / SF Code", ["ALL"] + all_codes)
        if selected_code == "ALL":
            scm_view = df_scm.copy()
            fake_view = df_fake.copy()
        else:
            scm_view = df_scm[df_scm["_RLS_CODE"] == selected_code].copy()
            fake_view = df_fake[df_fake["_RLS_CODE"] == selected_code].copy()
    else:
        # Strict RLS lock for Partners (supports multi-branch Common Partners)
        if len(allowed_codes) > 1:
            selected_code = st.sidebar.selectbox("Your Assigned Branches", ["ALL MY BRANCHES"] + allowed_codes)
            if selected_code == "ALL MY BRANCHES":
                scm_view = df_scm[df_scm["_RLS_CODE"].isin(allowed_codes)].copy()
                fake_view = df_fake[df_fake["_RLS_CODE"].isin(allowed_codes)].copy()
            else:
                scm_view = df_scm[df_scm["_RLS_CODE"] == selected_code].copy()
                fake_view = df_fake[df_fake["_RLS_CODE"] == selected_code].copy()
        else:
            st.sidebar.info(f"Locked to SF Code: **{allowed_codes[0]}**")
            scm_view = df_scm[df_scm["_RLS_CODE"].isin(allowed_codes)].copy()
            fake_view = df_fake[df_fake["_RLS_CODE"].isin(allowed_codes)].copy()

    # --- GLOBAL PERIOD / MONTH FILTER ---
    all_periods = sorted(
        list(set(scm_view["_PERIOD"].unique()).union(set(fake_view["_PERIOD"].unique())) - {"", "Unspecified"})
    )
    selected_periods = st.sidebar.multiselect("Filter by Month / Period", options=all_periods, default=[])
    if selected_periods:
        scm_view = scm_view[scm_view["_PERIOD"].isin(selected_periods)]
        fake_view = fake_view[fake_view["_PERIOD"].isin(selected_periods)]

    # --- HEADER & COMBINED DEBIT SUMMARY ---
    st.title("Partner Debit & Performance Insights")
    st.caption(f"Account: **{display_name}**")

    total_scm_net = scm_view["_NET_SCM_DEBIT"].sum()
    total_fake_amt = fake_view.loc[fake_view["_IS_FAKE"], "_FAKE_DEBIT_AMT"].sum()
    combined_debit = total_scm_net + total_fake_amt

    c1, c2, c3 = st.columns(3)
    c1.metric("Net SCM Debit (Pending CN)", f"₹{total_scm_net:,.2f}")
    c2.metric("Fake Cancellation Debit", f"₹{total_fake_amt:,.2f}")
    c3.metric("Total Combined Debit", f"₹{combined_debit:,.2f}")

    st.markdown("---")

    # --- 3 CATEGORY FILTER TABS ---
    tab_scm, tab_fake, tab_brand = st.tabs(["📦 SCM Debit", "🚨 Fake Debit", "🏷️ Brand Debit"])

    # ================= TAB 1: SCM DEBIT =================
    with tab_scm:
        st.subheader("SCM Debit — Spare Parts Billing & Credit Note (CN) Status")
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Total Billed Cases", f"{len(scm_view):,}")
        s2.metric("Gross Billed Amount", f"₹{scm_view['_BILLED_AMT'].sum():,.2f}")
        s3.metric("CN Credited Amount", f"₹{scm_view['_CN_AMT'].sum():,.2f}")
        s4.metric("Net SCM Debit", f"₹{total_scm_net:,.2f}")

        if not scm_view.empty:
            reason_col = get_col(scm_view, ["Billing reason", "CN issue reason"])
            if reason_col:
                st.markdown("**Net SCM Debit by Billing Reason**")
                reason_grp = scm_view.groupby(reason_col)["_NET_SCM_DEBIT"].sum().sort_values(ascending=False).head(8)
                fig1, ax1 = plt.subplots(figsize=(9, 3.5))
                reason_grp.plot(kind="bar", ax=ax1, color="#2563eb")
                ax1.set_ylabel("Amount (₹)")
                ax1.set_xlabel("")
                plt.xticks(rotation=35, ha="right")
                st.pyplot(fig1)

            display_scm_cols = [
                c for c in [
                    "Billing Month", "Billing Date", "SF Code", "SF Name", "Case ID",
                    "Brand", "Product", "Desc", "Total Price", "CN Amount",
                    "Billing reason", "SCM Head Approval Status", "Final Status"
                ] if c in scm_view.columns
            ]
            st.dataframe(scm_view[display_scm_cols], use_container_width=True)
            st.download_button(
                "⬇️ Download SCM Debit Report (CSV)",
                data=scm_view[display_scm_cols].to_csv(index=False).encode("utf-8"),
                file_name="SCM_Debit_Report.csv",
                mime="text/csv",
            )
        else:
            st.info("No SCM Debit records found for the selected filter.")

    # ================= TAB 2: FAKE DEBIT =================
    with tab_fake:
        st.subheader("Fake Debit — Audited Cancellations & Penalty Breakdown")
        fake_only_df = fake_view[fake_view["_IS_FAKE"]]

        f1, f2, f3, f4 = st.columns(4)
        f1.metric("Total Audited Cases", f"{len(fake_view):,}")
        f2.metric("Genuine Cases", f"{len(fake_view) - len(fake_only_df):,}")
        f3.metric("Fake Cancellations", f"{len(fake_only_df):,}")
        f4.metric("Fake Debit Amount", f"₹{total_fake_amt:,.2f}")

        if not fake_view.empty:
            rem_col = get_col(fake_view, ["Remarks L1", "Product_Service_Subcat_New"])
            if rem_col and not fake_only_df.empty:
                st.markdown("**Fake Cancellations by Audit Remark (Remarks L1)**")
                rem_grp = fake_only_df[rem_col].value_counts().head(8)
                fig2, ax2 = plt.subplots(figsize=(9, 3.5))
                rem_grp.plot(kind="bar", ax=ax2, color="#dc2626")
                ax2.set_ylabel("Case Count")
                ax2.set_xlabel("")
                plt.xticks(rotation=35, ha="right")
                st.pyplot(fig2)

            display_fake_cols = [
                c for c in [
                    "Month", "Week", "Audited Date", "Branch", "SF_Name", "Case ID",
                    "Order ID", "Genuine/Fake", "Remarks L1", "Remarks L2",
                    "Product_Service_Subcat_New", "Installation Charges"
                ] if c in fake_view.columns
            ]
            st.dataframe(fake_view[display_fake_cols], use_container_width=True)
            st.download_button(
                "⬇️ Download Fake Debit Report (CSV)",
                data=fake_view[display_fake_cols].to_csv(index=False).encode("utf-8"),
                file_name="Fake_Debit_Report.csv",
                mime="text/csv",
            )
        else:
            st.info("No Audited records found for the selected filter.")

    # ================= TAB 3: BRAND DEBIT =================
    with tab_brand:
        st.subheader("Brand Debit")
        st.info(
            "🚧 **Brand Debit module is ready for integration.**\n\n"
            "Once your Brand Debit workbook is available, it will automatically plug into this tab "
            "using the exact same **SF Code / Branch** Row-Level Security and Month filters."
        )

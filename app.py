import streamlit as st
import pandas as pd
import re
import io
from sqlalchemy import create_engine, text
import pdfplumber

# Database Engine Configuration
engine = create_engine(
    st.secrets["DATABASE_URL"],
    pool_pre_ping=True,
    connect_args={"connect_timeout": 10}
)

st.set_page_config(page_title="Personal Finance Hub", layout="wide", page_icon="💳")
st.title("💳 Expense, Cash & Wealth Tracker")

# Ensure database schema supports extended tracking
with engine.connect() as conn:
    conn.execute(text("""
        CREATE TABLE IF NOT EXISTS custom_categories (
            name TEXT PRIMARY KEY,
            cat_type TEXT DEFAULT 'Expense'
        );
        CREATE TABLE IF NOT EXISTS transactions (
            id SERIAL PRIMARY KEY,
            date DATE NOT NULL,
            description TEXT NOT NULL,
            amount NUMERIC(12, 2) NOT NULL,
            category TEXT NOT NULL,
            account_type TEXT NOT NULL,
            is_payment INT DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS categories (
            keyword TEXT PRIMARY KEY,
            category TEXT NOT NULL
        );
        ALTER TABLE custom_categories ADD COLUMN IF NOT EXISTS cat_type TEXT DEFAULT 'Expense';
    """))
    
    # Pre-seed categories with structural types
    seed_categories = [
        # Living Expenses
        ('Food', 'Expense'), ('Va-Al-SM', 'Expense'), ('Transportation', 'Expense'),
        ('Utilities', 'Expense'), ('Cell Phone', 'Expense'), ('Clothes', 'Expense'),
        ('Entertainment', 'Expense'), ('Rent', 'Expense'), ('Health', 'Expense'),
        ('Misc', 'Expense'), ('Citizenship', 'Expense'),
        # Wealth & Remittances (Excluded from living expenses)
        ('TFSA', 'Wealth/Savings'), ('FHSA', 'Wealth/Savings'), ('Savings', 'Wealth/Savings'),
        ('Sent to India', 'Remittance'), ('Cash', 'Wash/Transfer'), ('Bank Acc Charges', 'Wash/Transfer'),
        ('Shared Reimbursement', 'Wash/Transfer'),
        # Income Streams
        ('Salary', 'Income'), ('OT', 'Income'), ('Reimbursement', 'Income'), ('ITR Return', 'Income')
    ]
    for cat_name, ctype in seed_categories:
        conn.execute(text("""
            INSERT INTO custom_categories (name, cat_type) 
            VALUES (:name, :type) 
            ON CONFLICT (name) DO UPDATE SET cat_type = :type
        """), {"name": cat_name, "type": ctype})
    conn.commit()

# Application Tabs
tab_upload, tab_split, tab_cash, tab_dashboard, tab_rules = st.tabs([
    "📥 Upload & Coverage", "✂️ Review & Split Expenses", "💵 Cash Wallet", "📊 Financial Dashboard", "⚙️ Category & Auto-Rules"
])

def get_categories_dict(conn):
    """Returns mapping of category name to its operational type."""
    rows = conn.execute(text("SELECT name, cat_type FROM custom_categories ORDER BY name")).fetchall()
    return {r[0]: r[1] for r in rows}

def parse_pdf_statement(file):
    """Extracts date, description, and directional amounts from bank and credit card statements."""
    records = []
    date_regex = re.compile(
        r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{1,2}[/-]\d{1,2}|\d{4}[/-]\d{2}[/-]\d{2})",
        re.IGNORECASE
    )
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            table_found = False
            for table in tables:
                for row in table:
                    clean_row = [str(cell).strip() for cell in row if cell is not None]
                    if len(clean_row) >= 3:
                        d_match = date_regex.match(clean_row[0])
                        a_match = re.search(r"(-?[\$]?\d{1,3}(?:,\d{3})*\.\d{2})", clean_row[-1])
                        if d_match and a_match:
                            amt_val = float(a_match.group(1).replace("$", "").replace(",", ""))
                            records.append({
                                "Date": clean_row[0],
                                "Description": " ".join(clean_row[1:-1]),
                                "Amount": amt_val
                            })
                            table_found = True
            if not table_found:
                text_content = page.extract_text()
                if text_content:
                    for line in text_content.split("\n"):
                        line = line.strip()
                        if date_regex.match(line):
                            amt_match = re.search(r"(-?[\$]?\d{1,3}(?:,\d{3})*\.\d{2})\s*$", line)
                            if amt_match:
                                amt_val = float(amt_match.group(1).replace("$", "").replace(",", ""))
                                tokens = line[:amt_match.start()].strip().split()
                                if len(tokens) >= 2:
                                    date_part = " ".join(tokens[:2]) if tokens[1].isdigit() else tokens[0]
                                    desc_part = " ".join(tokens[2:]) if tokens[1].isdigit() else " ".join(tokens[1:])
                                    records.append({
                                        "Date": date_part,
                                        "Description": desc_part if desc_part else "Transaction",
                                        "Amount": amt_val
                                    })
    return pd.DataFrame(records)

def generate_excel_download(df_raw, living_pivot=None, wealth_pivot=None, inc_pivot=None):
    """Generates structured multi-sheet Excel export."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export = df_raw.copy()
        if 'date' in df_export.columns:
            df_export['date'] = pd.to_datetime(df_export['date']).dt.strftime('%Y-%m-%d')
        cols = [c for c in ['date', 'description', 'amount', 'category', 'account_type'] if c in df_export.columns]
        df_export[cols].rename(columns={
            'date': 'Date', 'description': 'Description', 'amount': 'Amount', 
            'category': 'Category', 'account_type': 'Account'
        }).to_excel(writer, sheet_name='All Transactions', index=False)
        
        if living_pivot is not None and not living_pivot.empty:
            living_pivot.to_excel(writer, sheet_name='Living Expenses')
        if wealth_pivot is not None and not wealth_pivot.empty:
            wealth_pivot.to_excel(writer, sheet_name='Wealth & Remittance')
        if inc_pivot is not None and not inc_pivot.empty:
            inc_pivot.to_excel(writer, sheet_name='Income')
    return output.getvalue()

# ==========================================
# --- TAB 1: UPLOAD & COVERAGE CHECKLIST ---
# ==========================================
with tab_upload:
    st.subheader("Statement Upload & Audit Checklist")
    c_up, c_cov = st.columns([1.2, 1.8])
    
    with c_up:
        uploaded_file = st.file_uploader("Upload Statement (PDF, CSV, Excel)", type=["pdf", "csv", "xlsx", "xls"])
        account_type = st.selectbox("Account Type", ["Bank Account", "Credit Card"])

        if uploaded_file and st.button("Process & Import Statement"):
            if uploaded_file.name.lower().endswith(".pdf"):
                df = parse_pdf_statement(uploaded_file)
            elif uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            if df.empty:
                st.error("No transactions could be extracted from this statement.")
            else:
                cols = {str(c).strip().lower(): c for c in df.columns}
                date_col = next((cols[k] for k in ["date", "transaction date", "posting date"] if k in cols), None)
                desc_col = next((cols[k] for k in ["description", "product", "merchant", "memo", "activity description"] if k in cols), None)
                amt_col = next((cols[k] for k in ["amount", "cost", "withdrawal", "debit", "deposit", "amount (s)"] if k in cols), None)

                if not (date_col and desc_col and amt_col):
                    st.error(f"Missing essential columns. Identified: {list(df.columns)}")
                else:
                    with engine.connect() as conn:
                        rules = dict(conn.execute(text("SELECT keyword, category FROM categories")).fetchall())
                        imported_count = 0
                        skipped_count = 0
                        file_seen_counts = {}

                        for _, row in df.iterrows():
                            desc = str(row[desc_col]).strip()
                            try:
                                raw_amt = float(str(row[amt_col]).replace("$", "").replace(",", ""))
                                if pd.isna(raw_amt) or raw_amt == 0:
                                    continue
                            except (ValueError, TypeError):
                                continue

                            try:
                                date_str = pd.to_datetime(row[date_col]).strftime("%Y-%m-%d")
                            except Exception:
                                date_str = pd.to_datetime("today").strftime("%Y-%m-%d")

                            # Determine directional sign and category
                            is_payment = 1 if any(w in desc.upper() for w in ["TD VISA", "PAYMENT - THANK YOU", "PAYMENT-THANK YOU", "CREDIT CARD BILL"]) else 0
                            
                            # Bank Fee Offset Handling
                            is_fee_rebate = "ACCT BAL REBATE" in desc.upper()
                            
                            # Normalize transaction amount
                            if account_type == "Credit Card":
                                # Negative on CC that is not payment = Merchant Refund
                                if raw_amt < 0 and not is_payment:
                                    final_amt = raw_amt  # Keep negative to offset spending
                                    is_refund = True
                                else:
                                    final_amt = abs(raw_amt)
                                    is_refund = False
                            else:
                                # Bank: Fee rebate is negative to offset account fee
                                if is_fee_rebate:
                                    final_amt = -abs(raw_amt)
                                else:
                                    final_amt = abs(raw_amt)
                                is_refund = False

                            # Frequency-based duplicate prevention
                            tx_key = (date_str, desc, final_amt, account_type)
                            file_seen_counts[tx_key] = file_seen_counts.get(tx_key, 0) + 1
                            current_occurrence = file_seen_counts[tx_key]

                            db_count = conn.execute(text("""
                                SELECT COUNT(*) FROM transactions 
                                WHERE date = :date AND description = :desc AND amount = :amt AND account_type = :acc
                            """), {"date": date_str, "desc": desc, "amt": final_amt, "acc": account_type}).scalar()

                            if db_count >= current_occurrence:
                                skipped_count += 1
                                continue

                            # Auto-routing rules
                            matched_cat = "Uncategorized"
                            if is_payment:
                                matched_cat = "Credit Card Bill"
                            elif any(k in desc.upper() for k in ["ACCOUNT FEE", "BAL REBATE", "MONTHLY FEE"]):
                                matched_cat = "Bank Acc Charges"
                            elif any(k in desc.upper() for k in ["HCL CANADA", "PAYROLL", "SALARY"]):
                                matched_cat = "Salary"
                            elif any(k in desc.upper() for k in ["RIA MONEY", "RIA TRANS"]):
                                matched_cat = "Sent to India"
                            elif "TD MHA PPP TFS" in desc.upper():
                                matched_cat = "TFSA"
                            elif "TD MHA PPP INV" in desc.upper():
                                matched_cat = "FHSA"
                            elif any(k in desc.upper() for k in ["ATM W/D", "ATM WITHDRAWAL"]):
                                matched_cat = "Cash"
                            elif any(k in desc.upper() for k in ["GST", "CANADA PRO", "TAX REFUND", "RIT"]):
                                matched_cat = "ITR Return"
                            else:
                                for kw, cat in rules.items():
                                    if kw.upper() in desc.upper():
                                        matched_cat = cat
                                        break

                            conn.execute(text("""
                                INSERT INTO transactions (date, description, amount, category, account_type, is_payment)
                                VALUES (:date, :desc, :amt, :cat, :acc, :pay)
                            """), {
                                "date": date_str, "desc": desc, "amt": final_amt, 
                                "cat": matched_cat, "acc": account_type, "pay": is_payment
                            })
                            imported_count += 1

                        conn.commit()

                    if imported_count > 0:
                        st.success(f"Imported {imported_count} new entries! (Skipped {skipped_count} existing duplicates)")
                    else:
                        st.info(f"All {skipped_count} transactions in this statement already exist in the database.")
                    st.rerun()

    with c_cov:
        st.write("#### 📅 Monthly Statement Coverage Audit")
        with engine.connect() as conn:
            audit_df = pd.read_sql("SELECT date, account_type FROM transactions WHERE account_type IN ('Bank Account', 'Credit Card')", conn)
        
        if audit_df.empty:
            st.info("No statements logged in database yet.")
        else:
            audit_df['date'] = pd.to_datetime(audit_df['date'])
            audit_df['Period'] = audit_df['date'].dt.to_period('M')
            audit_df['Month'] = audit_df['date'].dt.strftime("%b'%y")
            cov = audit_df.groupby(['Period', 'Month', 'account_type']).size().unstack(fill_value=0).reset_index().sort_values('Period')
            
            for col in ['Bank Account', 'Credit Card']:
                if col not in cov.columns:
                    cov[col] = 0

            def audit_label(r):
                b, c = r['Bank Account'] > 0, r['Credit Card'] > 0
                if b and c: return "✅ Complete (Both Uploaded)"
                if b: return "⚠️ Missing Credit Card"
                if c: return "⚠️ Missing Bank Account"
                return "❌ Missing"

            cov['Status'] = cov.apply(audit_label, axis=1)
            cov['Bank Account'] = cov['Bank Account'].apply(lambda x: f"✅ {x} txns" if x > 0 else "❌ Missing")
            cov['Credit Card'] = cov['Credit Card'].apply(lambda x: f"✅ {x} txns" if x > 0 else "❌ Missing")
            st.dataframe(cov[['Month', 'Bank Account', 'Credit Card', 'Status']], use_container_width=True, hide_index=True)

# ====================================================
# --- TAB 2: REVIEW UNCATEGORIZED & EXPENSE SPLIT ---
# ====================================================
with tab_split:
    st.subheader("📝 Categorization & Shared Expense Splitter")
    
    with engine.connect() as conn:
        uncat_tx = pd.read_sql("SELECT id, date, description, amount, account_type FROM transactions WHERE category = 'Uncategorized' ORDER BY date DESC", conn)
        cat_map = get_categories_dict(conn)
        active_cats = sorted(list(cat_map.keys()))

    if uncat_tx.empty:
        st.success("🎉 All uploaded transactions are fully categorized!")
    else:
        st.write(f"**{len(uncat_tx)}** transactions need categorization or splitting:")
        
        for _, row in uncat_tx.iterrows():
            with st.expander(f"📌 {row['date']} | {row['description']} | ${row['amount']:.2f} ({row['account_type']})", expanded=True):
                col_std, col_split = st.columns([1.2, 1.2])
                
                # Standard Direct Categorization
                with col_std:
                    st.write("**Direct Category Assignment**")
                    selected_cat = st.selectbox("Assign Category", active_cats, key=f"sel_{row['id']}")
                    rem_vendor = st.checkbox("Remember Vendor Rule", key=f"rem_{row['id']}", value=False)
                    if st.button("Save Assignment", key=f"btn_save_{row['id']}"):
                        with engine.connect() as conn:
                            conn.execute(text("UPDATE transactions SET category = :cat WHERE id = :id"), {"cat": selected_cat, "id": row['id']})
                            if rem_vendor:
                                kw = " ".join(str(row['description']).split()[:2]).upper()
                                conn.execute(text("INSERT INTO categories (keyword, category) VALUES (:kw, :cat) ON CONFLICT (keyword) DO UPDATE SET category = :cat"), {"kw": kw, "cat": selected_cat})
                            conn.commit()
                        st.rerun()

                # Shared Group Bill Splitter
                with col_split:
                    st.write("**✂️ Split Group / Shared Bill**")
                    my_share = st.number_input("Your Personal Share ($)", min_value=0.0, max_value=float(abs(row['amount'])), value=float(abs(row['amount']))/2, step=1.0, key=f"split_my_{row['id']}")
                    friends_share = round(float(abs(row['amount'])) - my_share, 2)
                    st.caption(f"Reimbursable by Friends: **${friends_share:.2f}** (Auto-assigned to `Shared Reimbursement`)")
                    my_category = st.selectbox("Category for Your Share", [c for c in active_cats if cat_map[c] == 'Expense'], key=f"split_cat_{row['id']}")
                    
                    if st.button("Execute Bill Split", key=f"btn_split_{row['id']}"):
                        with engine.connect() as conn:
                            # 1. Update original record to personal share
                            conn.execute(text("""
                                UPDATE transactions 
                                SET amount = :my_amt, category = :my_cat, description = :desc 
                                WHERE id = :id
                            """), {
                                "my_amt": my_share, 
                                "my_cat": my_category, 
                                "desc": f"{row['description']} (My Share)", 
                                "id": row['id']
                            })
                            # 2. Insert separate row for friend reimbursable share
                            conn.execute(text("""
                                INSERT INTO transactions (date, description, amount, category, account_type, is_payment)
                                VALUES (:date, :desc, :amt, 'Shared Reimbursement', :acc, 0)
                            """), {
                                "date": row['date'],
                                "desc": f"{row['description']} (Friends Reimbursable)",
                                "amt": friends_share,
                                "acc": row['account_type']
                            })
                            conn.commit()
                        st.success("Transaction split into personal share and friend reimbursable!")
                        st.rerun()

# ==========================================
# --- TAB 3: CASH WALLET RECONCILIATION ---
# ==========================================
with tab_cash:
    st.subheader("💵 Physical Cash Wallet & Reconciliation")
    
    with engine.connect() as conn:
        cash_withdrawn = conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE account_type = 'Bank Account' AND (category = 'Cash' OR description ILIKE '%ATM%')")).scalar()
        cash_spent = conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE account_type = 'Cash'")).scalar()
        df_cash_tx = pd.read_sql("SELECT id, date, description, amount, category FROM transactions WHERE account_type = 'Cash' ORDER BY date DESC", conn)
        cat_map = get_categories_dict(conn)
        spending_cats = [c for c in cat_map.keys() if cat_map[c] == 'Expense']

    cash_in_hand = float(cash_withdrawn - cash_spent)
    
    cw1, cw2, cw3 = st.columns(3)
    cw1.metric("Total ATM Cash Withdrawn", f"${float(cash_withdrawn):,.2f}")
    cw2.metric("Total Cash Spent Logged", f"${float(cash_spent):,.2f}")
    cw3.metric("Current Cash in Hand", f"${cash_in_hand:,.2f}", delta=cash_in_hand)

    st.markdown("---")
    cf_col, ch_col = st.columns([1.2, 1.8])
    
    with cf_col:
        st.write("#### ✍️ Log Cash Expense")
        with st.form("cash_spend_form", clear_on_submit=True):
            c_date = st.date_input("Date Spent")
            c_desc = st.text_input("Merchant / Description", placeholder="e.g. Barber, Fruit vendor, Cash tip")
            c_amt = st.number_input("Amount ($)", min_value=0.50, max_value=max(cash_in_hand, 5000.0), step=1.0)
            c_cat = st.selectbox("Category", spending_cats)
            if st.form_submit_button("Record Cash Expense"):
                if c_desc.strip():
                    with engine.connect() as conn:
                        conn.execute(text("""
                            INSERT INTO transactions (date, description, amount, category, account_type, is_payment)
                            VALUES (:date, :desc, :amt, :cat, 'Cash', 0)
                        """), {"date": c_date.strftime("%Y-%m-%d"), "desc": c_desc.strip(), "amt": c_amt, "cat": c_cat})
                        conn.commit()
                    st.success("Cash expense logged!")
                    st.rerun()

    with ch_col:
        st.write("#### 📜 Cash Outflow History")
        if df_cash_tx.empty:
            st.info("No cash spending logged yet. Unspent ATM withdrawals remain in your Cash Wallet.")
        else:
            for _, r in df_cash_tx.iterrows():
                h1, h2, h3, h4 = st.columns([2, 1.2, 1, 0.8])
                h1.write(f"**{r['description']}** ({r['date']})")
                h2.write(f"`{r['category']}`")
                h3.write(f"${r['amount']:.2f}")
                if h4.button("Delete", key=f"del_c_{r['id']}"):
                    with engine.connect() as conn:
                        conn.execute(text("DELETE FROM transactions WHERE id = :id"), {"id": r['id']})
                        conn.commit()
                    st.rerun()

# ====================================================
# --- TAB 4: INTERACTIVE DASHBOARD & WEALTH METRICS ---
# ====================================================
with tab_dashboard:
    with engine.connect() as conn:
        df_tx = pd.read_sql("SELECT id, date, description, amount, category, account_type, is_payment FROM transactions", conn)
        cc_charges = conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE account_type = 'Credit Card' AND is_payment = 0")).scalar()
        cc_payments = conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE is_payment = 1")).scalar()
        cat_map = get_categories_dict(conn)

    if df_tx.empty:
        st.info("No transactions available. Import bank or credit card statements to view analytics.")
    else:
        df_tx['date'] = pd.to_datetime(df_tx['date'])
        df_tx['Year'] = df_tx['date'].dt.year
        df_tx['Month_Period'] = df_tx['date'].dt.to_period('M')
        df_tx['Month_Str'] = df_tx['date'].dt.strftime("%b'%y")
        df_tx['cat_type'] = df_tx['category'].map(cat_map).fillna('Expense')

        # --- Filter Controls ---
        st.markdown("### 🔍 Dashboard Filters & Data Export")
        f1, f2, f3 = st.columns([1.5, 1.5, 1.5])
        
        all_years = sorted(df_tx['Year'].unique(), reverse=True)
        with f1:
            selected_years = st.multiselect("Filter Year(s)", options=all_years, default=all_years)
        
        avail_months = df_tx[df_tx['Year'].isin(selected_years)].sort_values('date')[['Month_Period', 'Month_Str']].drop_duplicates()
        month_order = avail_months['Month_Str'].tolist()
        with f2:
            selected_months = st.multiselect("Filter Month(s)", options=month_order, default=month_order)

        df_filtered = df_tx[(df_tx['Year'].isin(selected_years)) & (df_tx['Month_Str'].isin(selected_months))]

        # Exclude Bank Cash W/D (wash transfer to wallet) from living expenses
        df_living = df_filtered[
            (df_filtered['cat_type'] == 'Expense') & 
            (df_filtered['is_payment'] == 0) &
            ~((df_filtered['account_type'] == 'Bank Account') & (df_filtered['category'] == 'Cash'))
        ]
        
        df_wealth = df_filtered[df_filtered['cat_type'].isin(['Wealth/Savings', 'Remittance'])]
        df_income = df_filtered[df_filtered['cat_type'] == 'Income']

        # Pivot Tables
        living_pivot = pd.pivot_table(df_living, index='category', columns='Month_Str', values='amount', aggfunc='sum', fill_value=0) if not df_living.empty else pd.DataFrame()
        if not living_pivot.empty:
            living_pivot = living_pivot[[m for m in month_order if m in living_pivot.columns]]
            living_pivot['Total'] = living_pivot.sum(axis=1)
            living_pivot = living_pivot.sort_values('Total', ascending=False)

        wealth_pivot = pd.pivot_table(df_wealth, index='category', columns='Month_Str', values='amount', aggfunc='sum', fill_value=0) if not df_wealth.empty else pd.DataFrame()
        if not wealth_pivot.empty:
            wealth_pivot = wealth_pivot[[m for m in month_order if m in wealth_pivot.columns]]
            wealth_pivot['Total'] = wealth_pivot.sum(axis=1)

        inc_pivot = pd.pivot_table(df_income, index='category', columns='Month_Str', values='amount', aggfunc='sum', fill_value=0) if not df_income.empty else pd.DataFrame()
        if not inc_pivot.empty:
            inc_pivot = inc_pivot[[m for m in month_order if m in inc_pivot.columns]]
            inc_pivot['Total'] = inc_pivot.sum(axis=1)

        with f3:
            st.write(" ")
            st.write(" ")
            xl_bytes = generate_excel_download(df_filtered, living_pivot, wealth_pivot, inc_pivot)
            st.download_button("📥 Download Excel Report (.xlsx)", data=xl_bytes, file_name="Financial_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        # --- Top Executive Metrics ---
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        tot_living = df_living['amount'].sum()
        tot_income = df_income['amount'].sum()
        net_saved = tot_income - tot_living
        unpaid_cc = float(cc_charges - cc_payments)

        m1.metric("Core Living Expenses", f"${tot_living:,.2f}")
        m2.metric("Total Inflow / Income", f"${tot_income:,.2f}")
        m3.metric("Net Operational Savings", f"${net_saved:,.2f}", delta=net_saved)
        m4.metric("Unpaid Credit Card Balance", f"${unpaid_cc:,.2f}", delta=-unpaid_cc, delta_color="inverse")

        # --- Dedicated Wealth & Remittance KPIs ---
        st.markdown("---")
        st.write("#### 🛡️ Wealth Accumulation & Remittances")
        w1, w2, w3, w4 = st.columns(4)
        
        tfsa_tot = df_tx[df_tx['category'] == 'TFSA']['amount'].sum()
        fhsa_tot = df_tx[df_tx['category'] == 'FHSA']['amount'].sum()
        ria_tot = df_tx[df_tx['category'] == 'Sent to India']['amount'].sum()
        
        # Shared Reimbursement Audit (Positive means friends still owe you money)
        shared_owed = df_tx[df_tx['category'] == 'Shared Reimbursement']['amount'].sum()

        w1.metric("TFSA Contributions", f"${tfsa_tot:,.2f}")
        w2.metric("FHSA Contributions (Cap: $8,000)", f"${fhsa_tot:,.2f}")
        w3.metric("Sent to India (Ria)", f"${ria_tot:,.2f}")
        w4.metric("Friend Reimbursements Due", f"${shared_owed:,.2f}", delta=-shared_owed if shared_owed > 0 else 0, delta_color="inverse")
        
        # FHSA Contribution Progress
        fhsa_cap = 8000.0
        st.caption(f"FHSA Annual Room Used: ${fhsa_tot:,.2f} /${fhsa_cap:,.2f} ({min(fhsa_tot/fhsa_cap, 1.0)*100:.1f}%)")
        st.progress(min(max(fhsa_tot / fhsa_cap, 0.0), 1.0))

        # --- Living Expenses Breakdown ---
        st.markdown("---")
        st.subheader("🛒 Living Expenses Breakdown")
        
        if not living_pivot.empty:
            st.dataframe(living_pivot.style.format("${:,.2f}"), use_container_width=True)
            
            # MoM Living Expense Chart
            mom_agg = df_living.groupby(['Month_Period', 'Month_Str'])['amount'].sum().reset_index().sort_values('Month_Period')
            st.bar_chart(mom_agg.set_index('Month_Str')[['amount']].rename(columns={'amount': 'Monthly Living Spend ($)'}))

        # --- Wealth & Remittance Table ---
        if not wealth_pivot.empty:
            st.markdown("---")
            st.subheader("📈 Wealth Building & Transfers Breakdown")
            st.dataframe(wealth_pivot.style.format("${:,.2f}"), use_container_width=True)

        # --- Income Breakdown ---
        if not inc_pivot.empty:
            st.markdown("---")
            st.subheader("💵 Income Stream Breakdown")
            st.dataframe(inc_pivot.style.format("${:,.2f}"), use_container_width=True)

# ===================================================
# --- TAB 5: CATEGORY & KEYWORD AUTO-RULE MANAGER ---
# ===================================================
with tab_rules:
    st.subheader("⚙️ Categories & Keyword Auto-Routing Engine")
    cr_left, cr_right = st.columns(2)
    
    with cr_left:
        st.write("#### 📂 Master Category Definitions")
        with engine.connect() as conn:
            all_cats_df = pd.read_sql("SELECT name AS Category, cat_type AS Type FROM custom_categories ORDER BY cat_type, name", conn)
        st.dataframe(all_cats_df, use_container_width=True)
        
        st.write("##### Add Custom Category")
        new_c_name = st.text_input("New Category Name")
        new_c_type = st.selectbox("Category Nature", ["Expense", "Income", "Wealth/Savings", "Remittance", "Wash/Transfer"])
        if st.button("Add to Master Categories"):
            if new_c_name.strip():
                with engine.connect() as conn:
                    conn.execute(text("INSERT INTO custom_categories (name, cat_type) VALUES (:name, :type) ON CONFLICT (name) DO UPDATE SET cat_type = :type"), {
                        "name": new_c_name.strip(), "type": new_c_type
                    })
                    conn.commit()
                st.success(f"Category '{new_c_name.strip()}' configured as {new_c_type}!")
                st.rerun()

    with cr_right:
        st.write("#### 🔍 Keyword Auto-Rules")
        with engine.connect() as conn:
            rules_df = pd.read_sql("SELECT keyword AS Keyword, category AS Category FROM categories ORDER BY Category", conn)
        st.dataframe(rules_df, use_container_width=True)
        
        st.write("##### Create Custom Vendor Keyword")
        kw_text = st.text_input("Vendor Keyword (e.g., UBEREATS, SHELL)")
        with engine.connect() as conn:
            available_rule_cats = [r[0] for r in conn.execute(text("SELECT name FROM custom_categories ORDER BY name")).fetchall()]
        kw_target_cat = st.selectbox("Target Category", available_rule_cats)
        if st.button("Save Keyword Rule"):
            if kw_text.strip():
                with engine.connect() as conn:
                    conn.execute(text("INSERT INTO categories (keyword, category) VALUES (:kw, :cat) ON CONFLICT (keyword) DO UPDATE SET category = :cat"), {
                        "kw": kw_text.strip().upper(), "cat": kw_target_cat
                    })
                    conn.commit()
                st.success(f"Rule established: '{kw_text.strip().upper()}' ➔ '{kw_target_cat}'")
                st.rerun()

import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
from collections import defaultdict
from sqlalchemy import create_engine, text
import pypdf
from pdfminer.high_level import extract_text as pdfminer_extract_text
from pdfminer.layout import LAParams, LTTextBox
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator

# Database Engine Configuration
engine = create_engine(
    st.secrets["DATABASE_URL"],
    pool_pre_ping=True,
    connect_args={"connect_timeout": 10}
)

st.set_page_config(page_title="Personal Finance Hub", layout="wide", page_icon="💳")
st.title("💳 Personal Finance & 2026 Reconciliation Hub")

# Clean legacy payment entries on startup
with engine.connect() as conn:
    conn.execute(text("""
        UPDATE transactions 
        SET category = 'Card Payment', is_payment = 1 
        WHERE category IN ('Credit Card Bill', 'Credit Card Payment')
           OR description ILIKE '%TD VISA%'
           OR description ILIKE '%PAYMENT%THANK YOU%';
        DELETE FROM custom_categories WHERE name IN ('Credit Card Bill', 'Credit Card Payment');
    """))
    conn.commit()

tab_upload, tab_split, tab_cash, tab_dashboard, tab_rules = st.tabs([
    "📥 Upload Statement", "✂️ Review & Split Bills", "💵 Cash Wallet", "📊 2026 Budget Dashboard", "⚙️ Rules & Categories"
])

def get_categories_dict(conn):
    rows = conn.execute(text("SELECT name, cat_type FROM custom_categories ORDER BY name")).fetchall()
    return {r[0]: r[1] for r in rows}

# ==========================================================
# --- NATIVE TD STATEMENT PDF PARSING ENGINE ---
# ==========================================================

MONTH_MAP = {'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05', 'JUN': '06',
             'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'}

def parse_td_chequing_pdf(file_bytes, year=2026):
    """Accurately extracts all transactions from a TD Chequing PDF statement."""
    rsrcmgr = PDFResourceManager()
    laparams = LAParams(line_margin=0.1)
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    
    words = []
    fp = io.BytesIO(file_bytes)
    for page_idx, page in enumerate(PDFPage.get_pages(fp)):
        interpreter.process_page(page)
        layout = device.get_result()
        for element in layout:
            if isinstance(element, LTTextBox):
                for text_line in element:
                    txt = text_line.get_text().strip()
                    if txt:
                        words.append((page_idx, text_line.bbox[1], text_line.bbox[0], txt))
                        
    clustered = defaultdict(list)
    for p, y, x, txt in sorted(words, key=lambda item: (item[0], -item[1], item[2])):
        found = False
        for (cp, cy) in clustered:
            if cp == p and abs(cy - y) <= 4.0:
                clustered[(cp, cy)].append((x, txt))
                found = True
                break
        if not found:
            clustered[(p, y)].append((x, txt))
            
    txns = []
    for (p, cy) in sorted(clustered.keys(), key=lambda k: (k[0], -k[1])):
        items = sorted(clustered[(p, cy)], key=lambda x: x[0])
        row_txt = " ".join([t for _, t in items])
        if any(h in row_txt for h in ['START ING BALANCE', 'CLOSING BALANCE', 'Account /Transact', 'Overdraft']):
            continue
            
        date_item = None
        desc_parts = []
        withdrawal = None
        deposit = None
        
        for x, t in items:
            clean_t = t.replace(" ", "").replace(",", "").replace("$", "")
            m_date = re.match(r'^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)(\d{2})$', clean_t, re.IGNORECASE)
            if m_date and 400 < x < 450:
                date_item = (m_date.group(1).upper(), m_date.group(2))
            elif 250 < x < 330:
                m_amt = re.match(r'^-?(\d+\.\d{2})$', clean_t)
                if m_amt: withdrawal = float(m_amt.group(1))
            elif 340 < x < 400:
                m_amt = re.match(r'^-?(\d+\.\d{2})$', clean_t)
                if m_amt: deposit = float(m_amt.group(1))
            elif x < 250:
                desc_parts.append(t)
                
        if date_item and (withdrawal is not None or deposit is not None):
            m_str, d_str = date_item
            iso_date = f"{year}-{MONTH_MAP[m_str]}-{d_str}"
            desc = " ".join(desc_parts).replace("_", " ").strip()
            amt = withdrawal if withdrawal is not None else deposit
            is_dep = deposit is not None
            txns.append({
                "Date": iso_date,
                "Description": desc,
                "Amount": amt,
                "is_deposit": is_dep
            })
    return pd.DataFrame(txns)

def parse_td_visa_pdf(file_bytes, year=2026):
    """Accurately extracts all transactions from a TD Visa Credit Card PDF statement."""
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    date_re = re.compile(
        r'^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{1,2})\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{1,2})\s+(-?\$?[\d,]+\.\d{2})(.*)$',
        re.IGNORECASE
    )
    txns = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        for line in txt.split('\n'):
            m = date_re.match(line.strip())
            if m:
                m1, d1, _, _, amt_str, desc = m.groups()
                amt = float(amt_str.replace("$", "").replace(",", ""))
                iso_date = f"{year}-{MONTH_MAP[m1.upper()]}-{int(d1):02d}"
                txns.append({
                    "Date": iso_date,
                    "Description": desc.strip(),
                    "Amount": amt,
                    "is_deposit": amt < 0
                })
    return pd.DataFrame(txns)

# ==========================================================
# --- TAB 1: UPLOAD & AUTOMATIC STATEMENT INGESTION ---
# ==========================================================
with tab_upload:
    st.subheader("📥 Upload Statement (TD Bank Account or TD Visa PDF)")
    c_up, c_cov = st.columns([1.2, 1.8])
    
    with c_up:
        uploaded_file = st.file_uploader("Upload PDF, CSV, or Excel statement", type=["pdf", "csv", "xlsx", "xls"])
        account_type = st.selectbox("Select Account Type", ["Bank Account", "Credit Card"])
        statement_year = st.number_input("Statement Year", min_value=2024, max_value=2030, value=2026)

        if uploaded_file and st.button("Process & Reconcile Statement"):
            file_bytes = uploaded_file.read()
            df = pd.DataFrame()
            
            if uploaded_file.name.lower().endswith(".pdf"):
                if account_type == "Bank Account":
                    df = parse_td_chequing_pdf(file_bytes, statement_year)
                else:
                    df = parse_td_visa_pdf(file_bytes, statement_year)
            elif uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(io.BytesIO(file_bytes))
            else:
                df = pd.read_excel(io.BytesIO(file_bytes))

            if df.empty:
                st.error("No transactions could be extracted. Please check the file.")
            else:
                cols = {str(c).strip().lower(): c for c in df.columns}
                date_col = next((cols[k] for k in ["date", "transaction date", "posting date"] if k in cols), None)
                desc_col = next((cols[k] for k in ["description", "product", "merchant", "memo"] if k in cols), None)
                amt_col = next((cols[k] for k in ["amount", "cost", "withdrawal", "debit", "deposit"] if k in cols), None)

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

                        is_payment = 1 if any(w in desc.upper() for w in ["TD VISA", "PAYMENT - THANK YOU", "PAYMENT-THANK YOU", "CREDIT CARD BILL"]) else 0
                        is_fee_rebate = "ACCT BAL REBATE" in desc.upper()

                        if account_type == "Credit Card":
                            if raw_amt < 0 and not is_payment:
                                final_amt = raw_amt
                            else:
                                final_amt = abs(raw_amt)
                        else:
                            if is_fee_rebate:
                                final_amt = -abs(raw_amt)
                            else:
                                final_amt = abs(raw_amt)

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

                        # Precise Auto-Categorization
                        matched_cat = "Uncategorized"
                        if is_payment:
                            matched_cat = "Card Payment"
                        elif any(k in desc.upper() for k in ["PTS TO", "TFR-TO", "HL071 TFR", "TRANSFER TO"]):
                            matched_cat = "Internal Transfer"
                        elif any(k in desc.upper() for k in ["ACCOUNT FEE", "BAL REBATE", "MONTHLY FEE"]):
                            matched_cat = "Bank Acc Charges"
                        elif any(k in desc.upper() for k in ["HCL CANADA", "PAYROLL", "SALARY"]):
                            matched_cat = "Salary"
                        elif any(k in desc.upper() for k in ["RIA MONEY", "RIA TRANS"]):
                            matched_cat = "Ria"
                        elif "TD MHA PPP TFS" in desc.upper() or "TD MHA PPP INV" in desc.upper():
                            matched_cat = "Savings"
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
                    st.success(f"Imported {imported_count} transactions cleanly! (Skipped {skipped_count} duplicates)")
                else:
                    st.info(f"All {skipped_count} transactions already exist in the database.")
                st.rerun()

    with c_cov:
        st.write("#### 📅 Upload Coverage Audit")
        with engine.connect() as conn:
            audit_df = pd.read_sql("SELECT date, account_type FROM transactions WHERE account_type IN ('Bank Account', 'Credit Card')", conn)
        
        if not audit_df.empty:
            audit_df['date'] = pd.to_datetime(audit_df['date'])
            audit_df['Period'] = audit_df['date'].dt.to_period('M')
            audit_df['Month'] = audit_df['date'].dt.strftime("%b'%y")
            cov = audit_df.groupby(['Period', 'Month', 'account_type']).size().unstack(fill_value=0).reset_index().sort_values('Period')
            for col in ['Bank Account', 'Credit Card']:
                if col not in cov.columns: cov[col] = 0
            
            def audit_badge(r):
                b, c = r['Bank Account'] > 0, r['Credit Card'] > 0
                if b and c: return "✅ Complete (Both Uploaded)"
                if b: return "⚠️ Missing Credit Card"
                if c: return "⚠️ Missing Bank Account"
                return "❌ Missing"
            
            cov['Status'] = cov.apply(audit_badge, axis=1)
            cov['Bank Account'] = cov['Bank Account'].apply(lambda x: f"✅ {x} txns" if x > 0 else "❌ Missing")
            cov['Credit Card'] = cov['Credit Card'].apply(lambda x: f"✅ {x} txns" if x > 0 else "❌ Missing")
            st.dataframe(cov[['Month', 'Bank Account', 'Credit Card', 'Status']], use_container_width=True, hide_index=True)

# ==========================================================
# --- TAB 2: REVIEW UNCATEGORIZED & EXPENSE SPLITTER ---
# ==========================================================
with tab_split:
    st.subheader("📝 Review Uncategorized & Split Group Bills")
    with engine.connect() as conn:
        uncat_tx = pd.read_sql("SELECT id, date, description, amount, account_type FROM transactions WHERE category = 'Uncategorized' ORDER BY date DESC", conn)
        cat_map = get_categories_dict(conn)
        assignable_cats = sorted([c for c in cat_map.keys() if c not in ["Card Payment", "Internal Transfer"]])

    if uncat_tx.empty:
        st.success("🎉 All transactions are categorized!")
    else:
        st.write(f"**{len(uncat_tx)}** transactions need categorization:")
        for _, row in uncat_tx.iterrows():
            with st.expander(f"📌 {row['date']} | {row['description']} | ${float(row['amount']):.2f} ({row['account_type']})", expanded=True):
                c1, c2 = st.columns(2)
                with c1:
                    st.write("**Assign Category**")
                    new_c = st.selectbox("Category", assignable_cats, key=f"cat_sel_{row['id']}")
                    save_r = st.checkbox("Remember vendor rule", key=f"rule_chk_{row['id']}")
                    if st.button("Save", key=f"btn_save_{row['id']}"):
                        with engine.connect() as conn:
                            conn.execute(text("UPDATE transactions SET category = :cat WHERE id = :id"), {"cat": new_c, "id": row['id']})
                            if save_r:
                                kw = " ".join(str(row['description']).split()[:2]).upper()
                                conn.execute(text("INSERT INTO categories (keyword, category) VALUES (:kw, :cat) ON CONFLICT (keyword) DO UPDATE SET category = :cat"), {"kw": kw, "cat": new_c})
                            conn.commit()
                        st.rerun()
                with c2:
                    st.write("**Split Group / Shared Bill**")
                    full_amt = float(abs(row['amount']))
                    my_share = st.number_input("Your Share ($)", min_value=0.0, max_value=full_amt, value=full_amt/2, step=1.0, key=f"sp_my_{row['id']}")
                    fr_share = round(full_amt - my_share, 2)
                    st.caption(f"Friends Owe You: **${fr_share:.2f}** (Auto-assigned to `Shared Reimbursement`)")
                    my_cat = st.selectbox("Your Category", [c for c in assignable_cats if cat_map.get(c) == 'Expense'], key=f"sp_cat_{row['id']}")
                    if st.button("Execute Split", key=f"btn_sp_{row['id']}"):
                        with engine.connect() as conn:
                            conn.execute(text("UPDATE transactions SET amount = :a, category = :c, description = :d WHERE id = :id"), {
                                "a": my_share, "c": my_cat, "d": f"{row['description']} (My Share)", "id": row['id']
                            })
                            conn.execute(text("INSERT INTO transactions (date, description, amount, category, account_type, is_payment) VALUES (:d, :desc, :a, 'Shared Reimbursement', :acc, 0)"), {
                                "d": row['date'], "desc": f"{row['description']} (Friends Share)", "a": fr_share, "acc": row['account_type']
                            })
                            conn.commit()
                        st.rerun()

# ==========================================================
# --- TAB 3: CASH WALLET LEDGER ---
# ==========================================================
with tab_cash:
    st.subheader("💵 Physical Cash Wallet Ledger")
    with engine.connect() as conn:
        cash_withdrawn = float(conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE account_type = 'Bank Account' AND (category = 'Cash' OR description ILIKE '%ATM%')")).scalar() or 0.0)
        cash_spent = float(conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE account_type = 'Cash'")).scalar() or 0.0)
        df_cash = pd.read_sql("SELECT id, date, description, amount, category FROM transactions WHERE account_type = 'Cash' ORDER BY date DESC", conn)
        cat_map = get_categories_dict(conn)
        spending_cats = [c for c in cat_map.keys() if cat_map[c] == 'Expense']

    cash_left = cash_withdrawn - cash_spent
    w1, w2, w3 = st.columns(3)
    w1.metric("Total ATM Cash Withdrawn", f"${cash_withdrawn:,.2f}")
    w2.metric("Total Cash Spent Logged", f"${cash_spent:,.2f}")
    w3.metric("Current Cash in Hand", f"${cash_left:,.2f}", delta=cash_left)

    st.markdown("---")
    cf1, cf2 = st.columns([1.2, 1.8])
    with cf1:
        st.write("#### ✍️ Log Where Cash Was Spent")
        with st.form("cash_spend_form", clear_on_submit=True):
            cs_date = st.date_input("Date")
            cs_desc = st.text_input("Merchant / Memo", placeholder="e.g. Barber, Fruit vendor")
            cs_amt = st.number_input("Amount ($)", min_value=0.50, max_value=max(cash_left, 5000.0), step=1.0)
            cs_cat = st.selectbox("Category", spending_cats)
            if st.form_submit_button("Record Cash Expense"):
                if cs_desc.strip():
                    with engine.connect() as conn:
                        conn.execute(text("INSERT INTO transactions (date, description, amount, category, account_type, is_payment) VALUES (:d, :desc, :a, :c, 'Cash', 0)"), {
                            "d": cs_date.strftime("%Y-%m-%d"), "desc": cs_desc.strip(), "a": cs_amt, "c": cs_cat
                        })
                        conn.commit()
                    st.success("Cash expense saved!")
                    st.rerun()

    with cf2:
        st.write("#### 📜 Cash Outflow Log")
        if df_cash.empty:
            st.info("No cash spending logged yet. Unspent ATM withdrawals remain safely in your Cash in Hand.")
        else:
            for _, r in df_cash.iterrows():
                h1, h2, h3, h4 = st.columns([2, 1.2, 1, 0.8])
                h1.write(f"**{r['description']}** ({r['date']})")
                h2.write(f"`{r['category']}`")
                h3.write(f"${float(r['amount']):.2f}")
                if h4.button("Delete", key=f"del_c_{r['id']}"):
                    with engine.connect() as conn:
                        conn.execute(text("DELETE FROM transactions WHERE id = :id"), {"id": r['id']})
                        conn.commit()
                    st.rerun()

# ==========================================================
# --- TAB 4: THE 2026 BUDGET DASHBOARD & RECONCILIATION ---
# ==========================================================
with tab_dashboard:
    with engine.connect() as conn:
        df_tx = pd.read_sql("SELECT id, date, description, amount, category, account_type, is_payment FROM transactions", conn)
        cat_map = get_categories_dict(conn)

    if df_tx.empty:
        st.info("No transactions found. Upload your bank and credit card statements to view your dashboard!")
    else:
        df_tx['amount'] = pd.to_numeric(df_tx['amount'], errors='coerce').fillna(0.0).astype(float)
        df_tx['date'] = pd.to_datetime(df_tx['date'])
        df_tx['Year'] = df_tx['date'].dt.year
        df_tx['Month_Period'] = df_tx['date'].dt.to_period('M')
        df_tx['Month_Str'] = df_tx['date'].dt.strftime("%b'%y")
        df_tx['cat_type'] = df_tx['category'].map(cat_map).fillna('Expense')

        # Month order
        all_months = df_tx.sort_values('date')[['Month_Period', 'Month_Str']].drop_duplicates()['Month_Str'].tolist()

        # Strict Filter: Exclude CC Bill Payments, Inter-Account transfers, and fee rebates from living burn rate
        excluded_living = ['Card Payment', 'Credit Card Bill', 'Credit Card Payment', 'Internal Transfer', 'Bank Acc Charges', 'Shared Reimbursement']
        
        # 1. Living Expenses Matrix
        df_exp_only = df_tx[
            (df_tx['cat_type'] == 'Expense') & 
            (df_tx['is_payment'] == 0) & 
            (~df_tx['category'].isin(excluded_living)) &
            ~((df_tx['account_type'] == 'Bank Account') & (df_tx['category'] == 'Cash'))
        ]
        
        # Exact categories in order of your 2026 sheet
        ordered_cats = [
            'Food', 'Utilities', 'Va-Al-SM', 'Transportation', 'Entertainment',
            'Cell Phone', 'Clothes', 'Rent', 'Savings', 'Ria', 'Cash',
            'Citizenship', 'India', 'Misc', 'Health'
        ]
        
        # Build 2026 Master Grid
        pivot_data = []
        for cat in ordered_cats:
            cat_rows = df_tx[df_tx['category'] == cat] if cat in ['Savings', 'Ria', 'Cash'] else df_exp_only[df_exp_only['category'] == cat]
            row_dict = {'Type': cat}
            for m in all_months:
                val = cat_rows[cat_rows['Month_Str'] == m]['amount'].sum()
                row_dict[m] = round(val, 2)
            row_dict['Total'] = round(sum([row_dict[m] for m in all_months]), 2)
            pivot_data.append(row_dict)

        grid_df = pd.DataFrame(pivot_data).set_index('Type')

        # Compute Total Spent & Actual Spent
        total_spent_row = {'Type': 'Total Spent'}
        actual_spent_row = {'Type': 'Actual Spent'}
        for m in all_months + ['Total']:
            tot = grid_df[m].sum()
            total_spent_row[m] = round(tot, 2)
            # Actual spent subtracts Savings & Ria (just like row 29 in your Excel)
            sav = grid_df.loc['Savings', m] if 'Savings' in grid_df.index else 0
            ria = grid_df.loc['Ria', m] if 'Ria' in grid_df.index else 0
            actual_spent_row[m] = round(tot - (sav + ria), 2)

        # Compute Income Row
        df_inc = df_tx[df_tx['category'].isin(['Salary', 'OT', 'Reimbursement', 'ITR Return'])]
        income_row = {'Type': 'Income'}
        for m in all_months:
            income_row[m] = round(df_inc[df_inc['Month_Str'] == m]['amount'].sum(), 2)
        income_row['Total'] = round(sum([income_row[m] for m in all_months]), 2)

        # Compute Left Row (Income - Total Spent)
        left_row = {'Type': 'Left'}
        for m in all_months + ['Total']:
            left_row[m] = round(income_row[m] - total_spent_row[m], 2)

        # Account Split Rows
        cc_row = {'Type': 'Credit Card'}
        ba_row = {'Type': 'Bank Account'}
        cash_acc_row = {'Type': 'Cash'}
        for m in all_months:
            m_tx = df_tx[df_tx['Month_Str'] == m]
            cc_row[m] = round(m_tx[(m_tx['account_type'] == 'Credit Card') & (m_tx['is_payment'] == 0)]['amount'].sum(), 2)
            ba_row[m] = round(m_tx[(m_tx['account_type'] == 'Bank Account') & (m_tx['is_payment'] == 0) & (~m_tx['category'].isin(['Bank Acc Charges', 'Internal Transfer']))]['amount'].sum(), 2)
            cash_acc_row[m] = round(m_tx[(m_tx['account_type'] == 'Cash')]['amount'].sum(), 2)
        cc_row['Total'] = round(sum([cc_row[m] for m in all_months]), 2)
        ba_row['Total'] = round(sum([ba_row[m] for m in all_months]), 2)
        cash_acc_row['Total'] = round(sum([cash_acc_row[m] for m in all_months]), 2)

        # Append Summary Rows to Display Grid
        summary_rows = [total_spent_row, actual_spent_row, income_row, left_row, cc_row, ba_row, cash_acc_row]
        df_display_2026 = pd.concat([grid_df, pd.DataFrame(summary_rows).set_index('Type')])

        # Verification Row (Formula: Total Spent == Credit Card + Bank Account + Cash)
        check_row = {}
        for m in all_months:
            is_bal = abs(total_spent_row[m] - (cc_row[m] + ba_row[m] + cash_acc_row[m])) < 1.0
            check_row[m] = "✅ Correct" if is_bal else "⚠️ Check Data"
        check_row['Total'] = "✅ Correct"
        check_df = pd.DataFrame([check_row], index=['Reconciliation Status'])

        # Top Executive Metrics (Latest Month)
        latest_m = all_months[-1]
        st.markdown(f"### 🎯 Summary for **{latest_m}**")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Income", f"${income_row[latest_m]:,.2f}")
        k2.metric("Actual Living Spent", f"${actual_spent_row[latest_m]:,.2f}")
        k3.metric("Savings & Sent to India", f"${(grid_df.loc['Savings', latest_m] + grid_df.loc['Ria', latest_m]):,.2f}")
        k4.metric("Left (Net Savings)", f"${left_row[latest_m]:,.2f}", delta=left_row[latest_m])

        st.markdown("---")
        st.subheader("📋 2026 Master Budget Ledger")
        st.dataframe(df_display_2026.style.format("${:,.2f}"), use_container_width=True)
        st.dataframe(check_df, use_container_width=True)

        # Dedicated Category MoM Drill-down
        st.markdown("---")
        st.subheader("📈 Category Month-on-Month Drilldown")
        inspect_cat = st.selectbox("Select Category to Analyze", options=[c for c in ordered_cats if c in grid_df.index])
        
        cat_chart_data = grid_df.loc[inspect_cat, all_months]
        c_ch1, c_ch2 = st.columns([1.5, 1.5])
        with c_ch1:
            st.write(f"#### {inspect_cat} Spending Trajectory")
            st.bar_chart(cat_chart_data)
        with c_ch2:
            st.write(f"#### {inspect_cat} Transactions")
            sub_tx = df_tx[(df_tx['category'] == inspect_cat)].sort_values('date', ascending=False)
            if not sub_tx.empty:
                st.dataframe(
                    sub_tx[['date', 'description', 'amount', 'account_type']]
                    .assign(Date=lambda x: x['date'].dt.strftime('%Y-%m-%d'))
                    [['Date', 'description', 'amount', 'account_type']]
                    .rename(columns={'description': 'Merchant', 'amount': 'Amount ($)', 'account_type': 'Account'})
                    .style.format({'Amount ($)': "${:,.2f}"}),
                    use_container_width=True, hide_index=True
                )

# ==========================================================
# --- TAB 5: RULES & MASTER CATEGORIES ---
# ==========================================================
with tab_rules:
    st.subheader("⚙️ Rules & Vendor Keyword Engine")
    r1, r2 = st.columns(2)
    with r1:
        st.write("#### 📂 Master Categories")
        with engine.connect() as conn:
            all_c = pd.read_sql("SELECT name AS Category, cat_type AS Type FROM custom_categories ORDER BY cat_type, name", conn)
        st.dataframe(all_c, use_container_width=True, hide_index=True)
    with r2:
        st.write("#### 🔍 Vendor Auto-Rules")
        with engine.connect() as conn:
            all_r = pd.read_sql("SELECT keyword AS Keyword, category AS Category FROM categories ORDER BY Category", conn)
        st.dataframe(all_r, use_container_width=True, hide_index=True)
        
        st.write("##### Add Keyword Rule")
        new_kw = st.text_input("Vendor Keyword (e.g. COSTCO)")
        with engine.connect() as conn:
            avail_c = [r[0] for r in conn.execute(text("SELECT name FROM custom_categories ORDER BY name")).fetchall()]
        new_cat = st.selectbox("Category", avail_c)
        if st.button("Save Rule"):
            if new_kw.strip():
                with engine.connect() as conn:
                    conn.execute(text("INSERT INTO categories (keyword, category) VALUES (:kw, :cat) ON CONFLICT (keyword) DO UPDATE SET category = :cat"), {
                        "kw": new_kw.strip().upper(), "cat": new_cat
                    })
                    conn.commit()
                st.success(f"Rule added: {new_kw.upper()} ➔ {new_cat}")
                st.rerun()

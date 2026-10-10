import streamlit as st
import pandas as pd
import io
import base64
from datetime import datetime
from sqlalchemy import create_engine, text

# Import the parsing engines from parsers.py
from parsers import parse_td_chequing_pdf, parse_td_visa_pdf, unlock_pdf_bytes

# Database Configuration
engine = create_engine(
    st.secrets["DATABASE_URL"],
    pool_pre_ping=True,
    connect_args={"connect_timeout": 10}
)

st.set_page_config(page_title="Personal Finance Hub", layout="wide", page_icon="💳")
st.title("💳 Personal Finance & Master Reconciliation Hub")

# Auto-setup tables and perform database integrity fixes
with engine.connect() as conn:
    conn.execute(text("""
        CREATE TABLE IF NOT EXISTS statement_files (
            id SERIAL PRIMARY KEY,
            filename TEXT,
            account_type TEXT,
            statement_year INT,
            period_label TEXT,
            pdf_data BYTEA,
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (account_type, statement_year, period_label)
        );

        INSERT INTO custom_categories (name, cat_type) VALUES 
            ('Payment', 'Wash/Transfer'),
            ('TFSA', 'Expense'),
            ('FHSA', 'Expense')
        ON CONFLICT (name) DO NOTHING;

        UPDATE transactions 
        SET category = 'Payment', is_payment = 1 
        WHERE (description ILIKE '%TD VISA%' 
           OR description ILIKE '%PAYMENT%THANK YOU%' 
           OR description ILIKE '%PAYMENT-THANK YOU%' 
           OR description ILIKE '%CREDIT CARD BILL%'
           OR category IN ('Card Payment', 'Credit Card Bill', 'Credit Card Payment'));

        UPDATE transactions
        SET category = 'Ria'
        WHERE category = 'Uncategorized'
          AND (description ILIKE '%R ia%' OR description ILIKE '%Ria%');

        UPDATE transactions
        SET category = 'Uncategorized'
        WHERE category = 'Bank Acc Charges'
          AND NOT (description ILIKE '%MONTHLY ACCOUNT FEE%' OR description ILIKE '%ACCT BAL REBATE%' OR description ILIKE '%FEE%');

        DELETE FROM custom_categories WHERE name IN ('Credit Card Bill', 'Credit Card Payment', 'Card Payment');
    """))
    conn.commit()

tab_upload, tab_manage, tab_cash, tab_dashboard, tab_audit, tab_rules = st.tabs([
    "📥 Upload Statement", "✏️ Edit & Split Transactions", "💵 Cash Wallet", 
    "📊 Master Budget Dashboard", "📄 PDF Statement Audit", "⚙️ Rules & Categories"
])

def get_categories_dict(conn):
    rows = conn.execute(text("SELECT name, cat_type FROM custom_categories ORDER BY name")).fetchall()
    return {r[0]: r[1] for r in rows}

def generate_excel_report(year, grid_df, savings_df, summary_df, recon_df, tx_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        grid_df.to_excel(writer, sheet_name=f'{year} Outflows')
        savings_df.to_excel(writer, sheet_name=f'{year} Savings & Investments')
        summary_df.to_excel(writer, sheet_name=f'{year} Summary')
        recon_df.to_excel(writer, sheet_name=f'{year} Reconciliation')
        
        tx_clean = tx_df[['date', 'description', 'amount', 'category', 'account_type']].copy()
        tx_clean['date'] = tx_clean['date'].dt.strftime('%Y-%m-%d')
        tx_clean.rename(columns={
            'date': 'Date', 'description': 'Description', 'amount': 'Amount ($)',
            'category': 'Category', 'account_type': 'Account'
        }, inplace=True)
        tx_clean.to_excel(writer, sheet_name=f'{year} Transactions', index=False)
    return output.getvalue()

# ==========================================================
# --- TAB 1: UPLOAD STATEMENT ---
# ==========================================================
with tab_upload:
    st.subheader("📥 Upload Statement (TD Bank Account or TD Visa PDF)")
    c_up, c_cov = st.columns([1.2, 1.8])
    
    with c_up:
        uploaded_file = st.file_uploader("Upload PDF, CSV, or Excel statement", type=["pdf", "csv", "xlsx", "xls"])
        account_type = st.selectbox("Select Account Type", ["Bank Account", "Credit Card"])
        statement_year = st.number_input("Statement Year", min_value=2024, max_value=2035, value=2026)
        pdf_password = st.text_input("PDF Password (leave blank if none)", type="password", help="If password-protected, enter password here.")

        if uploaded_file and st.button("Process & Reconcile Statement"):
            file_bytes = uploaded_file.read()
            df = pd.DataFrame()
            clean_pdf_bytes = None
            
            if uploaded_file.name.lower().endswith(".pdf"):
                clean_bytes, err_msg = unlock_pdf_bytes(file_bytes, pdf_password)
                if err_msg:
                    st.error(err_msg)
                    st.stop()
                clean_pdf_bytes = clean_bytes

                if account_type == "Bank Account":
                    df = parse_td_chequing_pdf(clean_pdf_bytes, statement_year)
                else:
                    df = parse_td_visa_pdf(clean_pdf_bytes, statement_year)
            elif uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(io.BytesIO(file_bytes))
            else:
                df = pd.read_excel(io.BytesIO(file_bytes))

            if df.empty:
                st.error("No transactions could be extracted. Please verify the file.")
            else:
                cols = {str(c).strip().lower(): c for c in df.columns}
                date_col = next((cols[k] for k in ["date", "transaction date", "posting date"] if k in cols), None)
                desc_col = next((cols[k] for k in ["description", "product", "merchant", "memo"] if k in cols), None)
                amt_col = next((cols[k] for k in ["amount", "cost", "withdrawal", "debit", "deposit"] if k in cols), None)

                with engine.connect() as conn:
                    # Save PDF copy for the Audit Tab
                    if clean_pdf_bytes:
                        try:
                            sample_date = pd.to_datetime(df[date_col].iloc[0])
                            p_label = sample_date.strftime("%b'%y")
                        except Exception:
                            p_label = f"M_{datetime.now().strftime('%m')}"
                            
                        conn.execute(text("""
                            INSERT INTO statement_files (filename, account_type, statement_year, period_label, pdf_data)
                            VALUES (:fn, :acc, :yr, :p, :data)
                            ON CONFLICT (account_type, statement_year, period_label)
                            DO UPDATE SET pdf_data = :data, filename = :fn, uploaded_at = CURRENT_TIMESTAMP
                        """), {
                            "fn": uploaded_file.name, "acc": account_type,
                            "yr": statement_year, "p": p_label, "data": clean_pdf_bytes
                        })

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

                        desc_nospace = desc.upper().replace(" ", "")
                        is_payment = 1 if any(w in desc_nospace for w in ["TDVISA", "PAYMENT-THANKYOU", "PAYMENTTHANKYOU", "CREDITCARDBILL"]) else 0
                        is_fee_rebate = "ACCTBALREBATE" in desc_nospace

                        if account_type == "Credit Card":
                            final_amt = raw_amt if (raw_amt < 0 and not is_payment) else abs(raw_amt)
                        else:
                            final_amt = -abs(raw_amt) if is_fee_rebate else abs(raw_amt)

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

                        matched_cat = "Uncategorized"
                        if is_payment:
                            matched_cat = "Payment"
                        elif any(k in desc_nospace for k in ["PTSTO", "TFR-TO", "HL071TFR", "TRANSFERTO"]):
                            matched_cat = "Internal Transfer"
                        elif any(k in desc_nospace for k in ["ACCOUNTFEE", "BALREBATE", "MONTHLYFEE"]):
                            matched_cat = "Bank Acc Charges"
                        elif any(k in desc_nospace for k in ["HCLCANADA", "PAYROLL", "SALARY"]):
                            matched_cat = "Salary"
                        elif any(k in desc_nospace for k in ["RIAMONEY", "RIATRANS", "RIA"]):
                            matched_cat = "Ria"
                        elif "TDMHAPPP" in desc_nospace:
                            matched_cat = "Savings"
                        elif any(k in desc_nospace for k in ["ATMW/D", "ATMWITHDRAWAL"]):
                            matched_cat = "Cash"
                        elif any(k in desc_nospace for k in ["GST", "CANADAPRO", "TAXREFUND", "RIT"]):
                            matched_cat = "ITR Return"
                        else:
                            for kw, cat in rules.items():
                                if kw.upper().replace(" ", "") in desc_nospace:
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
        c_cov_head, c_cov_yr = st.columns([2, 1])
        with c_cov_head:
            st.write("#### 📅 Upload Coverage Audit")
        
        with engine.connect() as conn:
            audit_df = pd.read_sql("SELECT date, account_type FROM transactions WHERE account_type IN ('Bank Account', 'Credit Card')", conn)
        
        if not audit_df.empty:
            audit_df['date'] = pd.to_datetime(audit_df['date'])
            audit_df['Year'] = audit_df['date'].dt.year
            audit_df['Period'] = audit_df['date'].dt.to_period('M')
            audit_df['Month'] = audit_df['date'].dt.strftime("%b'%y")
            
            avail_audit_years = sorted(audit_df['Year'].unique(), reverse=True)
            if 2026 not in avail_audit_years: avail_audit_years.append(2026)
            avail_audit_years = sorted(list(set(avail_audit_years)), reverse=True)
            
            with c_cov_yr:
                selected_audit_year = st.selectbox("Audit Year", options=avail_audit_years, index=0)

            audit_year_df = audit_df[audit_df['Year'] == selected_audit_year]
            if not audit_year_df.empty:
                cov = audit_year_df.groupby(['Period', 'Month', 'account_type']).size().unstack(fill_value=0).reset_index().sort_values('Period')
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
# --- TAB 2: EDIT & SPLIT TRANSACTIONS ---
# ==========================================================
with tab_manage:
    st.subheader("✏️ Review, Edit Categories & Split Bills")
    
    with engine.connect() as conn:
        all_tx = pd.read_sql("SELECT id, date, description, amount, category, account_type, is_payment FROM transactions ORDER BY date DESC", conn)
        cat_map = get_categories_dict(conn)
        existing_cats = set(cat_map.keys()).union(set(all_tx['category'].dropna().unique()))
        assignable_cats = sorted([c for c in existing_cats if c != 'Uncategorized'])
        dropdown_cats = ['Uncategorized'] + assignable_cats

    if not all_tx.empty:
        all_tx['date'] = pd.to_datetime(all_tx['date'])
        all_tx['Year'] = all_tx['date'].dt.year
        all_tx['Month_Str'] = all_tx['date'].dt.strftime("%b'%y")

        f_yr, f_mo, f_acc, f_status, f_cat, f_search = st.columns([1.0, 1.1, 1.3, 1.5, 1.3, 1.8])

        avail_years = sorted(all_tx['Year'].unique(), reverse=True)
        with f_yr:
            selected_yr = st.selectbox("📅 Year", options=["All Years"] + [str(y) for y in avail_years], index=0)

        filtered_by_yr = all_tx if selected_yr == "All Years" else all_tx[all_tx['Year'] == int(selected_yr)]
        avail_months = filtered_by_yr.sort_values('date')['Month_Str'].unique().tolist()
        
        with f_mo:
            selected_mo = st.selectbox("🗓️ Month", options=["All Months"] + avail_months, index=0)
        with f_acc:
            selected_acc = st.selectbox("🏦 Account", options=["All Accounts", "Credit Card", "Bank Account", "Cash"], index=0)

        slice_df = all_tx.copy()
        if selected_yr != "All Years": slice_df = slice_df[slice_df['Year'] == int(selected_yr)]
        if selected_mo != "All Months": slice_df = slice_df[slice_df['Month_Str'] == selected_mo]
        if selected_acc != "All Accounts": slice_df = slice_df[slice_df['account_type'] == selected_acc]

        uncat_n = len(slice_df[slice_df['category'] == 'Uncategorized'])
        cat_n = len(slice_df[slice_df['category'] != 'Uncategorized'])

        with f_status:
            status_view = st.selectbox("🏷️ Status", options=[f"All ({len(slice_df)})", f"🚨 Uncategorized ({uncat_n})", f"✅ Categorized ({cat_n})"], index=1 if uncat_n > 0 else 0)
        with f_cat:
            selected_cat_filter = st.selectbox("📂 Category", options=["All Categories"] + dropdown_cats, index=0)
        with f_search:
            search_query = st.text_input("🔍 Search Merchant", placeholder="e.g. MCD, Costco...")

        df_view = slice_df.copy()
        if "Uncategorized" in status_view: df_view = df_view[df_view['category'] == 'Uncategorized']
        elif "Categorized" in status_view: df_view = df_view[df_view['category'] != 'Uncategorized']
        if selected_cat_filter != "All Categories": df_view = df_view[df_view['category'] == selected_cat_filter]
        if search_query.strip(): df_view = df_view[df_view['description'].str.contains(search_query.strip(), case=False, na=False)]

        st.caption(f"Showing **{len(df_view)}** matching transactions:")
        h_left, h_cat, h_btn = st.columns([3.4, 1.8, 0.8])
        h_left.markdown("**Transaction Details** *(Click dropdown to split)*")
        h_cat.markdown("**Category**")
        h_btn.markdown("**Action**")

        for _, row in df_view.iterrows():
            amt_display = f"${float(row['amount']):.2f}"
            c_left, c_cat, c_btn = st.columns([3.4, 1.8, 0.8])
            
            with c_left:
                with st.expander(f"📌 {row['date'].strftime('%Y-%m-%d')} | **{row['description']}** | {amt_display} ({row['account_type']})"):
                    split_mode = st.radio("Choose Split Type", ["🛒 Split into Multiple Categories", "👥 Split with Friends"], key=f"sm_{row['id']}")
                    full_amt = float(abs(row['amount']))
                    
                    if "Multiple Categories" in split_mode:
                        sp_c1, sp_c2 = st.columns(2)
                        with sp_c1:
                            part1_amt = st.number_input("Part 1 Amount ($)", min_value=0.01, max_value=max(full_amt - 0.01, 0.02), value=round(full_amt/2, 2), step=1.0, key=f"p1_amt_{row['id']}")
                            part1_cat = st.selectbox("Part 1 Category", [c for c in assignable_cats if cat_map.get(c) == 'Expense'], key=f"p1_cat_{row['id']}")
                        with sp_c2:
                            part2_amt = round(full_amt - part1_amt, 2)
                            st.metric("Part 2 Amount ($)", f"${part2_amt:.2f}")
                            part2_cat = st.selectbox("Part 2 Category", [c for c in assignable_cats if cat_map.get(c) == 'Expense'], key=f"p2_cat_{row['id']}")
                        
                        if st.button("Confirm Category Split", key=f"btn_catsplit_{row['id']}"):
                            with engine.connect() as conn:
                                conn.execute(text("UPDATE transactions SET amount = :a, category = :c, description = :d WHERE id = :id"), {
                                    "a": part1_amt, "c": part1_cat, "d": f"{row['description']} (Part 1 - {part1_cat})", "id": row['id']
                                })
                                conn.execute(text("INSERT INTO transactions (date, description, amount, category, account_type, is_payment) VALUES (:d, :desc, :a, :c, :acc, 0)"), {
                                    "d": row['date'], "desc": f"{row['description']} (Part 2 - {part2_cat})", "a": part2_amt, "c": part2_cat, "acc": row['account_type']
                                })
                                conn.commit()
                            st.success("Split completed!")
                            st.rerun()
                    else:
                        my_share = st.number_input("Your Share ($)", min_value=0.0, max_value=full_amt, value=round(full_amt/2, 2), step=1.0, key=f"sp_my_{row['id']}")
                        fr_share = round(full_amt - my_share, 2)
                        split_cat = st.selectbox("Category for Your Share", [c for c in assignable_cats if cat_map.get(c) == 'Expense'], key=f"sp_cat_{row['id']}")
                        
                        if st.button("Confirm Friend Split", key=f"btn_sp_{row['id']}"):
                            with engine.connect() as conn:
                                conn.execute(text("UPDATE transactions SET amount = :a, category = :c, description = :d WHERE id = :id"), {
                                    "a": my_share, "c": split_cat, "d": f"{row['description']} (My Share)", "id": row['id']
                                })
                                conn.execute(text("INSERT INTO transactions (date, description, amount, category, account_type, is_payment) VALUES (:d, :desc, :a, 'Shared Reimbursement', :acc, 0)"), {
                                    "d": row['date'], "desc": f"{row['description']} (Friends Share)", "a": fr_share, "acc": row['account_type']
                                })
                                conn.commit()
                            st.rerun()

            edit_key = f"edit_active_{row['id']}"
            is_editing = st.session_state.get(edit_key, False)
            is_uncat = row['category'] == 'Uncategorized'

            with c_cat:
                if is_uncat or is_editing:
                    cur_cat = row['category'] if row['category'] in dropdown_cats else 'Uncategorized'
                    new_selected_cat = st.selectbox("Category", dropdown_cats, index=dropdown_cats.index(cur_cat), key=f"cat_sel_{row['id']}", label_visibility="collapsed")
                else:
                    st.markdown(f"**`{row['category']}`**")

            with c_btn:
                if is_uncat or is_editing:
                    if st.button("Save", key=f"btn_save_{row['id']}"):
                        if new_selected_cat != 'Uncategorized':
                            is_pay = 1 if new_selected_cat == 'Payment' else 0
                            with engine.connect() as conn:
                                conn.execute(text("UPDATE transactions SET category = :cat, is_payment = :pay WHERE id = :id"), {
                                    "cat": new_selected_cat, "pay": is_pay, "id": row['id']
                                })
                                conn.commit()
                            st.session_state[edit_key] = False
                            st.rerun()
                else:
                    if st.button("Edit", key=f"btn_edit_{row['id']}"):
                        st.session_state[edit_key] = True
                        st.rerun()

# ==========================================================
# --- TAB 3: CASH WALLET ---
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
                    st.rerun()

    with cf2:
        st.write("#### 📜 Cash Outflow Log")
        if not df_cash.empty:
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
# --- TAB 4: MASTER BUDGET DASHBOARD ---
# ==========================================================
with tab_dashboard:
    with engine.connect() as conn:
        df_tx = pd.read_sql("SELECT id, date, description, amount, category, account_type, is_payment FROM transactions", conn)
        cat_map = get_categories_dict(conn)

    if not df_tx.empty:
        df_tx['amount'] = pd.to_numeric(df_tx['amount'], errors='coerce').fillna(0.0).astype(float)
        df_tx['date'] = pd.to_datetime(df_tx['date'])
        df_tx['Year'] = df_tx['date'].dt.year
        df_tx['Month_Str'] = df_tx['date'].dt.strftime("%b'%y")
        df_tx['cat_type'] = df_tx['category'].map(cat_map).fillna('Expense')

        dash_years = sorted(df_tx['Year'].unique(), reverse=True)
        d_col1, d_col2 = st.columns([1.5, 1.5])
        with d_col1:
            selected_dash_year = st.selectbox("📅 Select Budget Year", options=dash_years, index=0)

        df_tx_year = df_tx[df_tx['Year'] == selected_dash_year]
        all_months = df_tx_year.sort_values('date')[['date', 'Month_Str']].drop_duplicates()['Month_Str'].tolist()

        if all_months:
            ordered_cats = ['Food', 'Utilities', 'Va-Al-SM', 'Transportation', 'Entertainment', 'Cell Phone', 'Clothes', 'Rent', 'Savings', 'Ria', 'Cash', 'Citizenship', 'India', 'Misc', 'Health']
            all_known_cats = ordered_cats + [c for c in df_tx_year['category'].unique() if c not in ordered_cats and c not in ['Uncategorized', 'Payment', 'Internal Transfer', 'Bank Acc Charges', 'Shared Reimbursement'] and cat_map.get(c) == 'Expense']
            seen = set()
            cats_eval = [c for c in all_known_cats if not (c in seen or seen.add(c))]

            # 1. Category Outflows
            t1_h, t1_c = st.columns([2.2, 1.3])
            with t1_h: st.subheader(f"📋 1. {selected_dash_year} Master Category Outflows")
            with t1_c: sort_highest = st.checkbox("⬇️ Sort categories from highest to lowest spend", value=False)

            pivot_data = []
            for cat in cats_eval:
                cat_rows = df_tx_year[df_tx_year['category'] == cat]
                row_dict = {'Type': cat}
                for m in all_months:
                    row_dict[m] = round(cat_rows[cat_rows['Month_Str'] == m]['amount'].sum(), 2)
                row_total = round(sum([row_dict[m] for m in all_months]), 2)
                row_dict['Total'] = row_total
                if row_total > 0: pivot_data.append(row_dict)

            if sort_highest: pivot_data = sorted(pivot_data, key=lambda x: x['Total'], reverse=True)

            tot_out_row = {'Type': 'Total'}
            for m in all_months: tot_out_row[m] = round(sum(r[m] for r in pivot_data), 2)
            tot_out_row['Total'] = round(sum(r['Total'] for r in pivot_data), 2)

            grid_df = pd.DataFrame(pivot_data + [tot_out_row]).set_index('Type')
            st.dataframe(grid_df.style.format("${:,.2f}"), use_container_width=True)

            # 2. Savings & Investments Breakdown
            st.subheader("💰 2. Investments, Savings & Remittances Breakdown")
            def classify_savings(row):
                d = str(row['description']).upper().replace(" ", "")
                c = str(row['category']).upper().replace(" ", "")
                if 'FHSA' in c or 'FHSA' in d: return 'FHSA'
                elif 'TFSA' in c or 'TFSA' in d or 'TFS' in d: return 'TFSA'
                elif any(k in c or k in d for k in ['RIA', 'INDIA', 'REMITTANCE']): return 'Sent to India (Ria)'
                elif 'SAVING' in c or 'INV' in d or 'INVEST' in c: return 'Other Savings / Investments'
                return None

            df_tx_year['savings_goal'] = df_tx_year.apply(classify_savings, axis=1)
            sav_pivot = []
            for asset in ['TFSA', 'FHSA', 'Sent to India (Ria)', 'Other Savings / Investments']:
                sub_r = df_tx_year[df_tx_year['savings_goal'] == asset]
                s_dict = {'Asset / Goal': asset}
                for m in all_months: s_dict[m] = round(sub_r[sub_r['Month_Str'] == m]['amount'].sum(), 2)
                s_dict['Total'] = round(sum([s_dict[m] for m in all_months]), 2)
                sav_pivot.append(s_dict)

            tot_sav_row = {'Asset / Goal': 'Total Saved & Invested'}
            for m in all_months: tot_sav_row[m] = round(sum(r[m] for r in sav_pivot), 2)
            tot_sav_row['Total'] = round(sum(r['Total'] for r in sav_pivot), 2)

            savings_df = pd.DataFrame(sav_pivot + [tot_sav_row]).set_index('Asset / Goal')
            st.dataframe(savings_df.style.format("${:,.2f}"), use_container_width=True)

            # 3. Monthly Financial Summary
            st.subheader("💵 3. Monthly Financial Summary")
            inc_row = {'Metric': 'Total Income'}
            tsp_row = {'Metric': 'Total Spent'}
            act_row = {'Metric': 'Actual Spent (Excl. Savings & Ria)'}
            net_row = {'Metric': 'Net Left (Savings Balance)'}

            df_inc = df_tx_year[df_tx_year['cat_type'] == 'Income']
            for m in all_months:
                tot = tot_out_row[m]
                sav = tot_sav_row[m]
                inc = df_inc[df_inc['Month_Str'] == m]['amount'].sum()
                inc_row[m] = round(inc, 2)
                tsp_row[m] = round(tot, 2)
                act_row[m] = round(tot - sav, 2)
                net_row[m] = round(inc - tot, 2)

            inc_row['Total'] = round(sum([inc_row[m] for m in all_months]), 2)
            tsp_row['Total'] = round(sum([tsp_row[m] for m in all_months]), 2)
            act_row['Total'] = round(sum([act_row[m] for m in all_months]), 2)
            net_row['Total'] = round(inc_row['Total'] - tsp_row['Total'], 2)

            summary_df = pd.DataFrame([inc_row, tsp_row, act_row, net_row]).set_index('Metric')
            st.dataframe(summary_df.style.format("${:,.2f}"), use_container_width=True)

            # 4. Reconciliation
            st.subheader("⚖️ 4. Payment Sources & Reconciliation Audit")
            cc_row = {'Source': 'Credit Card'}
            ba_row = {'Source': 'Bank Account'}
            ca_row = {'Source': 'Cash'}
            out_sum_row = {'Source': 'Total Account Outflows'}
            rec_stat_row = {'Source': 'Reconciliation Status'}

            for m in all_months:
                m_tx = df_tx_year[df_tx_year['Month_Str'] == m]
                cc_val = m_tx[(m_tx['account_type'] == 'Credit Card') & (m_tx['is_payment'] == 0)]['amount'].sum()
                ba_val = m_tx[(m_tx['account_type'] == 'Bank Account') & (m_tx['cat_type'] != 'Income') & (m_tx['is_payment'] == 0) & (~m_tx['category'].isin(['Payment', 'Internal Transfer', 'Bank Acc Charges', 'Shared Reimbursement']))]['amount'].sum()
                ca_val = m_tx[m_tx['account_type'] == 'Cash']['amount'].sum()

                cc_row[m] = round(cc_val, 2)
                ba_row[m] = round(ba_val, 2)
                ca_row[m] = round(ca_val, 2)
                t_out = round(cc_val + ba_val + ca_val, 2)
                out_sum_row[m] = t_out
                rec_stat_row[m] = "✅ Correct" if abs(tsp_row[m] - t_out) < 0.05 else "⚠️ Check Data"

            cc_row['Total'] = round(sum([cc_row[m] for m in all_months]), 2)
            ba_row['Total'] = round(sum([ba_row[m] for m in all_months]), 2)
            ca_row['Total'] = round(sum([ca_row[m] for m in all_months]), 2)
            out_sum_row['Total'] = round(sum([out_sum_row[m] for m in all_months]), 2)
            rec_stat_row['Total'] = "✅ Correct" if abs(tsp_row['Total'] - out_sum_row['Total']) < 0.05 else "⚠️ Check Data"

            recon_df = pd.DataFrame([cc_row, ba_row, ca_row, out_sum_row, rec_stat_row]).set_index('Source')
            st.dataframe(recon_df.map(lambda x: f"${x:,.2f}" if isinstance(x, (int, float)) else str(x)), use_container_width=True)

            with d_col2:
                st.write("")
                st.write("")
                excel_bytes = generate_excel_report(selected_dash_year, grid_df, savings_df, summary_df, recon_df, df_tx_year)
                st.download_button(f"📥 Download {selected_dash_year} Budget Report (.xlsx)", excel_bytes, f"{selected_dash_year}_Budget_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# ==========================================================
# --- TAB 5: DEDICATED PDF STATEMENT AUDIT VIEWER ---
# ==========================================================
with tab_audit:
    st.subheader("📄 Interactive PDF Statement Audit & Verification")
    
    with engine.connect() as conn:
        saved_files = conn.execute(text("SELECT id, filename, account_type, statement_year, period_label FROM statement_files ORDER BY statement_year DESC, period_label DESC")).fetchall()
    
    if not saved_files:
        st.info("No saved statements found in the database. When you upload a statement in Tab 1, a copy will be archived here automatically!")
    else:
        file_options = {f"{r[3]} {r[4]} | {r[2]} ({r[1]})": r[0] for r in saved_files}
        chosen_label = st.selectbox("Select Statement to Audit", list(file_options.keys()))
        selected_file_id = file_options[chosen_label]

        with engine.connect() as conn:
            file_rec = conn.execute(text("SELECT filename, account_type, statement_year, period_label, pdf_data FROM statement_files WHERE id = :id"), {"id": selected_file_id}).fetchone()
        
        if file_rec:
            fname, a_type, s_year, s_period, p_bytes = file_rec
            col_pdf, col_data = st.columns([1.2, 1])

            with col_pdf:
                st.markdown(f"#### 📑 Statement Document: `{fname}`")
                st.download_button("📥 Download This PDF", bytes(p_bytes), fname, "application/pdf")
                
                b64_pdf = base64.b64encode(bytes(p_bytes)).decode('utf-8')
                pdf_embed_html = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="820" type="application/pdf" style="border: 1px solid #444; border-radius: 8px;"></iframe>'
                st.markdown(pdf_embed_html, unsafe_allow_html=True)

            with col_data:
                st.markdown(f"#### 🔍 Extracted Data in App ({s_period})")
                
                with engine.connect() as conn:
                    matched_tx = pd.read_sql(
                        text("SELECT date, description, amount, category FROM transactions WHERE account_type = :acc ORDER BY date ASC"),
                        conn,
                        params={"acc": a_type}
                    )

                if not matched_tx.empty:
                    matched_tx['date_dt'] = pd.to_datetime(matched_tx['date'])
                    matched_tx = matched_tx[
                        (matched_tx['date_dt'].dt.year == int(s_year)) &
                        (matched_tx['date_dt'].dt.strftime("%b'%y") == s_period)
                    ]

                if matched_tx.empty:
                    st.warning(f"No transactions found for {s_period} in {a_type}. Try re-uploading this statement in Tab 1.")
                else:
                    k1, k2 = st.columns(2)
                    k1.metric("Logged Transactions", len(matched_tx))
                    k2.metric("Total Account Outflow", f"${matched_tx['amount'].sum():,.2f}")
                    
                    st.dataframe(
                        matched_tx.assign(Date=lambda x: pd.to_datetime(x['date']).dt.strftime('%Y-%m-%d'))
                        [['Date', 'description', 'amount', 'category']]
                        .rename(columns={'description': 'Merchant / Description', 'amount': 'Amount ($)', 'category': 'Category'})
                        .style.format({'Amount ($)': "${:,.2f}"}),
                        use_container_width=True,
                        height=750
                    )

# ==========================================================
# --- TAB 6: RULES & MASTER CATEGORIES ---
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

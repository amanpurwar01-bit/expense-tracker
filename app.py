import streamlit as st
import pandas as pd
import re
from sqlalchemy import create_engine, text
import pdfplumber

# Connect to cloud database via Streamlit Secrets
engine = create_engine(
    st.secrets["DATABASE_URL"],
    pool_pre_ping=True,
    connect_args={"connect_timeout": 10}
)

st.set_page_config(page_title="Expense Tracker", layout="wide")
st.title("💳 Expense & Credit Card Tracker")

tab_upload, tab_review, tab_dashboard, tab_rules = st.tabs([
    "📥 Upload Statement", "📝 Review Uncategorized", "📊 Interactive Dashboard", "⚙️ Manage Categories & Rules"
])

def get_categories(conn):
    """Fetch active custom categories from Supabase."""
    cats = [r[0] for r in conn.execute(text("SELECT name FROM custom_categories ORDER BY name")).fetchall()]
    if not cats:
        cats = ["Food", "Va-Al-SM", "Transportation", "Utilities", "Cell Phone", "Clothes", 
                "Entertainment", "Rent", "Health", "Misc", "Savings", "Sent to India", "Salary", "Bank Acc Charges"]
    return cats

def parse_pdf_statement(file):
    """Extracts transaction rows from bank and credit card statement PDFs."""
    records = []
    date_regex = re.compile(
        r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{1,2}[/-]\d{1,2}|\d{4}[/-]\d{2}[/-]\d{2})",
        re.IGNORECASE
    )
    amount_regex = re.compile(r"[\$]?(\d{1,3}(?:,\d{3})*\.\d{2})\s*$")

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            table_found = False
            for table in tables:
                for row in table:
                    clean_row = [str(cell).strip() for cell in row if cell is not None]
                    if len(clean_row) >= 3:
                        d_match = date_regex.match(clean_row[0])
                        a_match = re.search(r"(\d+\.\d{2})", clean_row[-1])
                        if d_match and a_match:
                            records.append({
                                "Date": clean_row[0],
                                "Description": " ".join(clean_row[1:-1]),
                                "Amount": float(a_match.group(1).replace(",", ""))
                            })
                            table_found = True

            if not table_found:
                text_content = page.extract_text()
                if text_content:
                    for line in text_content.split("\n"):
                        line = line.strip()
                        if date_regex.match(line):
                            amt_match = amount_regex.search(line)
                            if amt_match:
                                amt_val = float(amt_match.group(1).replace(",", ""))
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

# --- Tab 1: Upload Statements ---
with tab_upload:
    st.subheader("Upload Bank or Credit Card Statement")
    uploaded_file = st.file_uploader("Upload PDF, CSV, or Excel file", type=["pdf", "csv", "xlsx", "xls"])
    account_type = st.selectbox("Account Type", ["Bank Account", "Credit Card"])

    if uploaded_file and st.button("Process & Import"):
        if uploaded_file.name.lower().endswith(".pdf"):
            df = parse_pdf_statement(uploaded_file)
        elif uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        if df.empty:
            st.error("No valid transactions could be extracted from this file.")
        else:
            cols = {str(c).strip().lower(): c for c in df.columns}
            date_col = next((cols[k] for k in ["date", "transaction date", "posting date"] if k in cols), None)
            desc_col = next((cols[k] for k in ["description", "product", "merchant", "memo"] if k in cols), None)
            amt_col = next((cols[k] for k in ["amount", "cost", "withdrawal", "debit", "deposit"] if k in cols), None)

            if not (date_col and desc_col and amt_col):
                st.error(f"Could not map columns. Found: {list(df.columns)}")
            else:
                with engine.connect() as conn:
                    rules = dict(conn.execute(text("SELECT keyword, category FROM categories")).fetchall())
                    imported_count = 0
                    
                    for _, row in df.iterrows():
                        desc = str(row[desc_col]).strip()
                        try:
                            amt = abs(float(row[amt_col]))
                            if pd.isna(amt) or amt == 0:
                                continue
                        except (ValueError, TypeError):
                            continue
                        
                        try:
                            date_str = pd.to_datetime(row[date_col]).strftime("%Y-%m-%d")
                        except Exception:
                            date_str = pd.to_datetime("today").strftime("%Y-%m-%d")
                        
                        is_payment = 1 if any(w in desc.upper() for w in ["TD VISA", "PAYMENT - THANK YOU", "CREDIT CARD BILL", "PAYMENT-THANK YOU"]) else 0
                        
                        matched_cat = "Uncategorized"
                        if not is_payment:
                            if any(k in desc.upper() for k in ["HCL CANADA", "PAYROLL", "SALARY"]):
                                matched_cat = "Salary"
                            elif any(k in desc.upper() for k in ["ACCOUNT FEE", "BAL REBATE", "MONTHLY FEE"]):
                                matched_cat = "Bank Acc Charges"
                            elif any(k in desc.upper() for k in ["GST", "CANADA PRO", "TAX REFUND", "RIT"]):
                                matched_cat = "ITR Return"
                            else:
                                for kw, cat in rules.items():
                                    if kw.upper() in desc.upper():
                                        matched_cat = cat
                                        break
                        else:
                            matched_cat = "Credit Card Bill"
                        
                        conn.execute(text("""
                            INSERT INTO transactions (date, description, amount, category, account_type, is_payment)
                            VALUES (:date, :desc, :amt, :cat, :acc, :pay)
                        """), {
                            "date": date_str, 
                            "desc": desc, 
                            "amt": amt, 
                            "cat": matched_cat, 
                            "acc": account_type, 
                            "pay": is_payment
                        })
                        imported_count += 1
                    
                    conn.commit()
                st.success(f"Successfully processed and imported {imported_count} transactions!")

# --- Tab 2: Review Uncategorized ---
with tab_review:
    st.subheader("Review & Assign Categories")
    with engine.connect() as conn:
        uncat_tx = pd.read_sql("SELECT id, date, description, amount, account_type FROM transactions WHERE category = 'Uncategorized' ORDER BY date DESC", conn)
        active_categories = get_categories(conn)

    with st.expander("➕ Add a New Category on the fly"):
        new_quick_cat = st.text_input("New Category Name")
        is_inc = st.checkbox("Is this an Income category?", value=False)
        if st.button("Add Category"):
            if new_quick_cat.strip():
                with engine.connect() as conn:
                    conn.execute(text("INSERT INTO custom_categories (name, is_income) VALUES (:name, :inc) ON CONFLICT (name) DO NOTHING"), {
                        "name": new_quick_cat.strip(), "inc": 1 if is_inc else 0
                    })
                    conn.commit()
                st.success(f"Category '{new_quick_cat.strip()}' added!")
                st.rerun()

    if uncat_tx.empty:
        st.info("No uncategorized transactions pending!")
    else:
        st.write(f"**{len(uncat_tx)}** transactions need categorization:")
        for idx, row in uncat_tx.iterrows():
            c1, c2, c3, c4 = st.columns([3, 1, 2, 1.5])
            c1.write(f"**{row['description']}** ({row['date']})")
            c2.write(f"${row['amount']:.2f}")
            new_cat = c3.selectbox("Category", active_categories, key=f"cat_{row['id']}")
            
            with c4:
                save_rule = st.checkbox("Remember vendor", key=f"rule_{row['id']}", value=False)
                if st.button("Save", key=f"btn_{row['id']}"):
                    with engine.connect() as conn:
                        conn.execute(text("UPDATE transactions SET category = :cat WHERE id = :id"), {
                            "cat": new_cat, "id": row['id']
                        })
                        if save_rule:
                            keyword = " ".join(str(row['description']).split()[:2]).upper()
                            conn.execute(text("""
                                INSERT INTO categories (keyword, category) 
                                VALUES (:kw, :cat) 
                                ON CONFLICT (keyword) DO UPDATE SET category = :cat
                            """), {"kw": keyword, "cat": new_cat})
                        conn.commit()
                    st.rerun()

# --- Tab 3: Interactive Dashboard ---
with tab_dashboard:
    with engine.connect() as conn:
        df_tx = pd.read_sql("SELECT id, date, description, amount, category, account_type, is_payment FROM transactions", conn)
        cc_charges = conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE account_type = 'Credit Card'")).scalar()
        cc_payments = conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE is_payment = 1")).scalar()
        income_cats = [r[0] for r in conn.execute(text("SELECT name FROM custom_categories WHERE is_income = 1")).fetchall()]

    if df_tx.empty:
        st.info("No transactions found in the database. Upload statements to view insights!")
    else:
        # Pre-process dates
        df_tx['date'] = pd.to_datetime(df_tx['date'])
        df_tx['Year'] = df_tx['date'].dt.year
        df_tx['Month_Period'] = df_tx['date'].dt.to_period('M')
        df_tx['Month_Str'] = df_tx['date'].dt.strftime("%b'%y")

        # Exclude internal non-spending transactions from default expense calculations
        excluded_defaults = ["Credit Card Bill", "Transfer", "Bank Acc Charges"] + income_cats

        # --- Interactive Filter Bar ---
        st.markdown("### 🔍 Filters & Settings")
        f_col1, f_col2 = st.columns(2)
        
        all_years = sorted(df_tx['Year'].unique(), reverse=True)
        with f_col1:
            selected_years = st.multiselect("Select Year(s)", options=all_years, default=all_years)
        
        # Filter months by selected years
        available_months = (
            df_tx[df_tx['Year'].isin(selected_years)]
            .sort_values('date')[['Month_Period', 'Month_Str']]
            .drop_duplicates()
        )
        month_options = available_months['Month_Str'].tolist()
        with f_col2:
            selected_months = st.multiselect("Select Month(s)", options=month_options, default=month_options)

        # Category inclusion/exclusion filter
        available_categories = sorted([c for c in df_tx['category'].unique() if c not in ["Credit Card Bill", "Transfer", "Bank Acc Charges"] and c not in income_cats])
        selected_categories = st.multiselect(
            "Include / Remove Spending Categories from Dashboard",
            options=available_categories,
            default=available_categories,
            help="Uncheck any category (e.g. Sent to India, Rent) to see your variable living expenses without them."
        )

        # Apply active filters
        df_filtered = df_tx[
            (df_tx['Year'].isin(selected_years)) &
            (df_tx['Month_Str'].isin(selected_months))
        ]
        
        df_spending_filtered = df_filtered[
            (df_filtered['category'].isin(selected_categories)) & 
            (df_filtered['is_payment'] == 0)
        ]
        
        df_income_filtered = df_filtered[df_filtered['category'].isin(income_cats)]

        # --- Summary Metric Cards ---
        st.markdown("---")
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        
        total_income = df_income_filtered['amount'].sum()
        total_spending = df_spending_filtered['amount'].sum()
        net_saved = total_income - total_spending
        unpaid_cc = float(cc_charges - cc_payments)

        m_col1.metric("Selected Spending", f"${total_spending:,.2f}")
        m_col2.metric("Selected Income", f"${total_income:,.2f}")
        m_col3.metric("Net Cash Flow", f"${net_saved:,.2f}", delta=net_saved)
        m_col4.metric("Unpaid Credit Card Balance", f"${unpaid_cc:,.2f}", delta=-unpaid_cc, delta_color="inverse")

        # --- Month-on-Month Trends & Comparison ---
        st.markdown("---")
        st.subheader("📈 Month-on-Month Trends")
        
        if not df_spending_filtered.empty:
            # Aggregate by month chronologically
            monthly_agg = (
                df_spending_filtered.groupby(['Month_Period', 'Month_Str'])['amount']
                .sum()
                .reset_index()
                .sort_values('Month_Period')
            )
            chart_data = monthly_agg.set_index('Month_Str')[['amount']].rename(columns={'amount': 'Total Spending ($)'})
            st.bar_chart(chart_data)

            # Month-by-Month Category Pivot Table
            st.write("#### 📊 Category Breakdown Table")
            cat_pivot = df_spending_filtered.pivot_table(
                index='category', 
                columns='Month_Str', 
                values='amount', 
                aggfunc='sum', 
                fill_value=0
            )
            # Reorder columns chronologically
            ordered_cols = [m for m in month_options if m in cat_pivot.columns]
            cat_pivot = cat_pivot[ordered_cols]
            cat_pivot['Total'] = cat_pivot.sum(axis=1)
            cat_pivot = cat_pivot.sort_values('Total', ascending=False)
            st.dataframe(cat_pivot.style.format("${:,.2f}"), use_container_width=True)

        # --- Single Category Inspector ---
        st.markdown("---")
        st.subheader("🎯 Deep-Dive by Specific Category")
        
        inspect_cat = st.selectbox("Choose a category to drill into", options=["All"] + available_categories)
        if inspect_cat != "All":
            cat_tx = df_filtered[df_filtered['category'] == inspect_cat].sort_values('date', ascending=False)
            st.write(f"Showing **{len(cat_tx)}** transactions for **{inspect_cat}** totaling **${cat_tx['amount'].sum():,.2f}**:")
            
            # Show category MoM trend chart
            cat_mom = cat_tx.groupby(['Month_Period', 'Month_Str'])['amount'].sum().reset_index().sort_values('Month_Period')
            if len(cat_mom) > 1:
                st.line_chart(cat_mom.set_index('Month_Str')['amount'])
                
            st.dataframe(
                cat_tx[['date', 'description', 'amount', 'account_type']]
                .rename(columns={'date': 'Date', 'description': 'Description', 'amount': 'Amount ($)', 'account_type': 'Account'})
                .assign(Date=lambda x: x['Date'].dt.strftime("%Y-%m-%d")),
                use_container_width=True
            )

        # --- Income Breakdown Table ---
        if not df_income_filtered.empty:
            st.markdown("---")
            st.subheader("💵 Income Stream Summary")
            inc_pivot = df_income_filtered.pivot_table(
                index='category', 
                columns='Month_Str', 
                values='amount', 
                aggfunc='sum', 
                fill_value=0
            )
            inc_ordered = [m for m in month_options if m in inc_pivot.columns]
            inc_pivot = inc_pivot[inc_ordered]
            inc_pivot['Total'] = inc_pivot.sum(axis=1)
            st.dataframe(inc_pivot.style.format("${:,.2f}"), use_container_width=True)

# --- Tab 4: Category & Keyword Rules Management ---
with tab_rules:
    st.subheader("⚙️ Manage Categories & Keyword Rules")
    
    c_left, c_right = st.columns(2)
    
    with c_left:
        st.write("#### 📂 Master Categories")
        with engine.connect() as conn:
            all_cats_df = pd.read_sql("SELECT name AS Category, CASE WHEN is_income = 1 THEN 'Income' ELSE 'Expense/Savings' END AS Type FROM custom_categories ORDER BY name", conn)
        st.dataframe(all_cats_df, use_container_width=True)
        
        st.write("##### Add New Category")
        new_cat_name = st.text_input("Category Name", key="new_cat_input")
        cat_is_inc = st.checkbox("Is Income?", key="new_cat_is_inc")
        if st.button("Add to Master List"):
            if new_cat_name.strip():
                with engine.connect() as conn:
                    conn.execute(text("INSERT INTO custom_categories (name, is_income) VALUES (:name, :inc) ON CONFLICT (name) DO NOTHING"), {
                        "name": new_cat_name.strip(), "inc": 1 if cat_is_inc else 0
                    })
                    conn.commit()
                st.success(f"Category '{new_cat_name.strip()}' created!")
                st.rerun()

    with c_right:
        st.write("#### 🔍 Keyword Auto-Rules")
        with engine.connect() as conn:
            rules_df = pd.read_sql("SELECT keyword AS Keyword, category AS Category FROM categories ORDER BY Category", conn)
        st.dataframe(rules_df, use_container_width=True)
        
        st.write("##### Add Custom Auto-Rule")
        kw_input = st.text_input("Vendor Keyword (e.g. GYM)")
        with engine.connect() as conn:
            rule_cats = get_categories(conn)
        kw_cat = st.selectbox("Assign to Category", rule_cats, key="kw_cat_select")
        if st.button("Save Auto-Rule"):
            if kw_input.strip():
                with engine.connect() as conn:
                    conn.execute(text("INSERT INTO categories (keyword, category) VALUES (:kw, :cat) ON CONFLICT (keyword) DO UPDATE SET category = :cat"), {
                        "kw": kw_input.strip().upper(), "cat": kw_cat
                    })
                    conn.commit()
                st.success(f"Rule added: '{kw_input.strip().upper()}' ➔ '{kw_cat}'")
                st.rerun()

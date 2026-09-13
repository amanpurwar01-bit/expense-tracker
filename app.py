import streamlit as st
import pandas as pd
import re
from sqlalchemy import create_engine, text
import pdfplumber

# Connect to cloud database via Streamlit Secrets
engine = create_engine(st.secrets["DATABASE_URL"])

st.set_page_config(page_title="Expense Tracker", layout="wide")
st.title("💳 Expense & Credit Card Tracker")

tab_upload, tab_review, tab_dashboard, tab_rules = st.tabs([
    "📥 Upload Statement", "📝 Review Uncategorized", "📊 Monthly Summary", "⚙️ Category Rules"
])

CATEGORIES = [
    "Food", "Va-Al-SM", "Transportation", "Utilities", 
    "Cell Phone", "Clothes", "Entertainment", "Rent", 
    "Savings", "Sent to India", "Health", "Misc"
]

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
    account_type = st.selectbox("Account Type", ["Credit Card", "Bank Account"])

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
            amt_col = next((cols[k] for k in ["amount", "cost", "withdrawal", "debit"] if k in cols), None)

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
                        
                        is_payment = 1 if any(w in desc.upper() for w in ["TD VISA", "PAYMENT - THANK YOU", "CREDIT CARD BILL"]) else 0
                        
                        matched_cat = "Uncategorized"
                        if not is_payment:
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

    if uncat_tx.empty:
        st.info("No uncategorized transactions pending!")
    else:
        st.write(f"**{len(uncat_tx)}** transactions need categorization:")
        for idx, row in uncat_tx.iterrows():
            c1, c2, c3, c4 = st.columns([3, 1, 2, 1.5])
            c1.write(f"**{row['description']}** ({row['date']})")
            c2.write(f"${row['amount']:.2f}")
            new_cat = c3.selectbox("Category", CATEGORIES, key=f"cat_{row['id']}")
            
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

# --- Tab 3: Monthly Summary & Balance ---
with tab_dashboard:
    st.subheader("📊 Expense & Cash Flow Overview")
    with engine.connect() as conn:
        cc_charges = conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE account_type = 'Credit Card'")).scalar()
        cc_payments = conn.execute(text("SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE is_payment = 1")).scalar()
        df_tx = pd.read_sql("SELECT date, amount, category FROM transactions WHERE category NOT IN ('Credit Card Bill', 'Transfer') AND is_payment = 0", conn)

    unpaid_balance = float(cc_charges - cc_payments)
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Card Charges Tracked", f"${float(cc_charges):,.2f}")
    col2.metric("Total Payments Sent to Card", f"${float(cc_payments):,.2f}")
    col3.metric("Unpaid Credit Card Balance", f"${unpaid_balance:,.2f}", delta=-unpaid_balance, delta_color="inverse")

    st.markdown("---")
    st.write("### Spending by Category and Month")
    if not df_tx.empty:
        df_tx['date'] = pd.to_datetime(df_tx['date'])
        df_tx['Month'] = df_tx['date'].dt.strftime("%b'%y")
        pivot = df_tx.pivot_table(index='category', columns='Month', values='amount', aggfunc='sum', fill_value=0)
        st.dataframe(pivot.style.format("${:,.2f}"), use_container_width=True)

# --- Tab 4: Category Keyword Rules ---
with tab_rules:
    st.subheader("⚙️ Managed Keyword Rules")
    with engine.connect() as conn:
        rules_df = pd.read_sql("SELECT keyword AS Keyword, category AS Category FROM categories ORDER BY Category", conn)
    st.dataframe(rules_df, use_container_width=True)

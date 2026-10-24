"""
Personal Expense Tracker (Streamlit)
=====================================
A single-user daily/monthly expense tracker with charts.

Storage (automatic):
  - If a Postgres connection string is provided -> uses Postgres (data is permanent)
  - Otherwise                                   -> uses a local SQLite file (fine for testing on your laptop)

Where the connection string comes from, in order of priority:
  1. st.secrets["DATABASE_URL"]      <- used on Streamlit Community Cloud
  2. environment variable DATABASE_URL
  3. none -> local SQLite file "expenses.db"

Run locally:
    pip install -r requirements.txt
    streamlit run app.py
"""

import os
from datetime import date, datetime
from calendar import monthrange

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------
# Config: the categories Asad asked for
# ----------------------------------------------------------------------
DAILY_CATEGORIES = ["InDrive to office", "Lunch", "Dinner"]
MONTHLY_CATEGORIES = ["Rent", "Internet bill", "Mobile package"]
MISC_CATEGORY = "Miscellaneous"
ALL_CATEGORIES = DAILY_CATEGORIES + MONTHLY_CATEGORIES + [MISC_CATEGORY]

CURRENCY = "PKR"

st.set_page_config(page_title="My Expense Tracker", page_icon="💸", layout="wide")


# ----------------------------------------------------------------------
# Figure out which database to use
# ----------------------------------------------------------------------
def get_database_url():
    # 1. Streamlit secrets (used on Streamlit Cloud)
    try:
        if "DATABASE_URL" in st.secrets:
            return st.secrets["DATABASE_URL"]
    except Exception:
        pass
    # 2. environment variable
    return os.environ.get("DATABASE_URL")


DATABASE_URL = get_database_url()
USE_POSTGRES = bool(DATABASE_URL)
SQLITE_PATH = "expenses.db"


def get_conn():
    if USE_POSTGRES:
        import psycopg2
        return psycopg2.connect(DATABASE_URL, sslmode="require")
    else:
        import sqlite3
        conn = sqlite3.connect(SQLITE_PATH)
        return conn


def q(sql):
    """SQLite uses ? placeholders; Postgres uses %s. Write with %s, translate for SQLite."""
    return sql if USE_POSTGRES else sql.replace("%s", "?")


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    if USE_POSTGRES:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS expenses (
                id SERIAL PRIMARY KEY,
                category TEXT NOT NULL,
                amount NUMERIC NOT NULL,
                description TEXT DEFAULT '',
                spent_on DATE NOT NULL,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
    else:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category TEXT NOT NULL,
                amount REAL NOT NULL,
                description TEXT DEFAULT '',
                spent_on TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
    conn.commit()
    cur.close()
    conn.close()


# ----------------------------------------------------------------------
# Settings helpers (budget, salary, salary day)
# ----------------------------------------------------------------------
def get_setting(key, default=None):
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(q("SELECT value FROM settings WHERE key = %s"), (key,))
        row = cur.fetchone()
        cur.close()
        conn.close()
        return row[0] if row else default
    except Exception:
        return default


def set_setting(key, value):
    conn = get_conn()
    cur = conn.cursor()
    if USE_POSTGRES:
        cur.execute(
            "INSERT INTO settings (key, value) VALUES (%s, %s) "
            "ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value",
            (key, str(value)),
        )
    else:
        cur.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
            (key, str(value)),
        )
    conn.commit()
    cur.close()
    conn.close()


# ----------------------------------------------------------------------
# Expense helpers
# ----------------------------------------------------------------------
def add_expense(category, amount, description, spent_on):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        q("INSERT INTO expenses (category, amount, description, spent_on) VALUES (%s, %s, %s, %s)"),
        (category, float(amount), description, str(spent_on)),
    )
    conn.commit()
    cur.close()
    conn.close()


def delete_expense(expense_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(q("DELETE FROM expenses WHERE id = %s"), (int(expense_id),))
    conn.commit()
    cur.close()
    conn.close()


def load_expenses(month=None):
    conn = get_conn()
    cur = conn.cursor()
    if month:
        cur.execute(
            q("SELECT id, category, amount, description, spent_on FROM expenses "
              "WHERE CAST(spent_on AS TEXT) LIKE %s ORDER BY spent_on DESC, id DESC"),
            (month + "%",),
        )
    else:
        cur.execute(
            q("SELECT id, category, amount, description, spent_on FROM expenses "
              "ORDER BY spent_on DESC, id DESC")
        )
    rows = cur.fetchall()
    cur.close()
    conn.close()
    df = pd.DataFrame(rows, columns=["id", "category", "amount", "description", "spent_on"])
    if not df.empty:
        df["amount"] = df["amount"].astype(float)
        df["spent_on"] = pd.to_datetime(df["spent_on"]).dt.date
    return df


def money(x):
    return f"{CURRENCY} {x:,.0f}"


# ----------------------------------------------------------------------
# App
# ----------------------------------------------------------------------
init_db()

st.title("💸 My Expense Tracker")

# ---- Sidebar: budget & salary settings ----
with st.sidebar:
    st.header("⚙️ Setup")
    st.caption("Set these once. They are saved automatically.")

    current_budget = float(get_setting("monthly_budget", 0) or 0)
    current_salary = float(get_setting("monthly_salary", 0) or 0)
    current_sal_day = int(get_setting("salary_day", 1) or 1)

    monthly_salary = st.number_input(
        f"Monthly salary ({CURRENCY})", min_value=0.0, value=current_salary, step=1000.0
    )
    monthly_budget = st.number_input(
        f"Monthly budget ({CURRENCY})", min_value=0.0, value=current_budget, step=1000.0,
        help="How much you plan to spend per month."
    )
    salary_day = st.number_input(
        "Salary date (day of month)", min_value=1, max_value=31, value=current_sal_day, step=1,
        help="The day your salary arrives, e.g. 1 for the 1st."
    )

    if st.button("💾 Save setup", use_container_width=True):
        set_setting("monthly_salary", monthly_salary)
        set_setting("monthly_budget", monthly_budget)
        set_setting("salary_day", salary_day)
        st.success("Saved!")

    st.divider()
    storage_label = "Postgres (permanent)" if USE_POSTGRES else "Local SQLite (testing only)"
    st.caption(f"Storage: {storage_label}")


# ---- Which month are we looking at ----
today = date.today()
col_a, col_b = st.columns([1, 3])
with col_a:
    years = list(range(today.year - 2, today.year + 1))
    sel_year = st.selectbox("Year", years, index=len(years) - 1)
    sel_month_num = st.selectbox(
        "Month", list(range(1, 13)),
        index=today.month - 1,
        format_func=lambda m: datetime(2000, m, 1).strftime("%B"),
    )
month_key = f"{sel_year:04d}-{sel_month_num:02d}"

# ---- Add expense form ----
st.subheader("➕ Add an expense")
with st.form("add_expense_form", clear_on_submit=True):
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        category = st.selectbox("Category", ALL_CATEGORIES)
    with c2:
        amount = st.number_input(f"Amount ({CURRENCY})", min_value=0.0, step=50.0)
    with c3:
        spent_on = st.date_input("Date", value=today)
    description = st.text_input(
        "Description (optional, useful for Miscellaneous)",
        placeholder="e.g. Haircut, medicine, gift",
    )
    submitted = st.form_submit_button("Add expense", use_container_width=True)
    if submitted:
        if amount <= 0:
            st.error("Amount must be greater than 0.")
        else:
            add_expense(category, amount, description, spent_on)
            st.success(f"Added {money(amount)} for {category}.")

# ---- Load this month's data ----
df = load_expenses(month=month_key)
total_spent = float(df["amount"].sum()) if not df.empty else 0.0

budget = float(get_setting("monthly_budget", 0) or 0)
salary = float(get_setting("monthly_salary", 0) or 0)
sal_day = int(get_setting("salary_day", 1) or 1)

# ---- Top metrics ----
st.subheader(f"📊 Overview for {datetime(sel_year, sel_month_num, 1).strftime('%B %Y')}")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total spent", money(total_spent))
m2.metric("Monthly budget", money(budget),
          delta=money(budget - total_spent) + " left" if budget else None,
          delta_color="normal")
m3.metric("Salary", money(salary),
          delta=money(salary - total_spent) + " left" if salary else None,
          delta_color="normal")

# days until next salary
try:
    last_day = monthrange(today.year, today.month)[1]
    pay_day = min(sal_day, last_day)
    this_month_payday = date(today.year, today.month, pay_day)
    if today <= this_month_payday:
        days_to_salary = (this_month_payday - today).days
    else:
        nxt_month = today.month + 1 if today.month < 12 else 1
        nxt_year = today.year if today.month < 12 else today.year + 1
        nxt_last = monthrange(nxt_year, nxt_month)[1]
        next_payday = date(nxt_year, nxt_month, min(sal_day, nxt_last))
        days_to_salary = (next_payday - today).days
    m4.metric("Days to salary", f"{days_to_salary} days")
except Exception:
    m4.metric("Days to salary", "-")

# ---- Budget progress bar ----
if budget > 0:
    pct = min(total_spent / budget, 1.0)
    st.progress(pct, text=f"Budget used: {money(total_spent)} / {money(budget)} ({pct*100:.0f}%)")
    if total_spent > budget:
        st.warning(f"You are over budget by {money(total_spent - budget)} this month.")

st.divider()

# ---- Charts ----
if df.empty:
    st.info("No expenses yet for this month. Add one above to see your charts.")
else:
    left, right = st.columns(2)

    with left:
        st.markdown("**Spending by category**")
        by_cat = df.groupby("category")["amount"].sum().reindex(ALL_CATEGORIES).fillna(0)
        by_cat = by_cat[by_cat > 0]
        st.bar_chart(by_cat)

    with right:
        st.markdown("**Daily spending this month**")
        by_day = df.groupby("spent_on")["amount"].sum().sort_index()
        st.line_chart(by_day)

    # category share table
    st.markdown("**Category breakdown**")
    breakdown = (
        df.groupby("category")["amount"].sum().reindex(ALL_CATEGORIES).fillna(0)
    )
    breakdown_df = breakdown[breakdown > 0].reset_index()
    breakdown_df.columns = ["Category", "Amount"]
    breakdown_df["Share"] = (breakdown_df["Amount"] / total_spent * 100).round(1).astype(str) + "%"
    breakdown_df["Amount"] = breakdown_df["Amount"].apply(money)
    st.dataframe(breakdown_df, use_container_width=True, hide_index=True)

    st.divider()

    # ---- All expenses with delete ----
    st.subheader("🧾 This month's expenses")
    st.caption("Tick the box next to a row and press Delete to remove it.")

    display = df.copy()
    display["delete?"] = False
    display = display[["delete?", "spent_on", "category", "amount", "description", "id"]]
    display["amount"] = display["amount"].apply(money)

    edited = st.data_editor(
        display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "delete?": st.column_config.CheckboxColumn("Delete", default=False),
            "spent_on": "Date",
            "category": "Category",
            "amount": "Amount",
            "description": "Description",
            "id": None,  # hide the id column
        },
        disabled=["spent_on", "category", "amount", "description"],
        key="expense_editor",
    )

    to_delete = edited[edited["delete?"]]["id"].tolist()
    if to_delete:
        if st.button(f"🗑️ Delete {len(to_delete)} selected", type="primary"):
            for eid in to_delete:
                delete_expense(eid)
            st.success("Deleted. Refreshing...")
            st.rerun()

    # ---- Download as CSV ----
    csv = df[["spent_on", "category", "amount", "description"]].to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download this month as CSV",
        csv,
        file_name=f"expenses_{month_key}.csv",
        mime="text/csv",
    )

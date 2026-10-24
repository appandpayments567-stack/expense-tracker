"""
Personal Expense Tracker (Streamlit)
=====================================
Single-user daily/monthly expense tracker with charts, full CRUD,
and a formatted Excel export. Default currency: PKR.

Storage (automatic):
  - If a Postgres connection string is provided -> Postgres (permanent)
  - Otherwise                                   -> local SQLite file (testing)

Connection string is read from, in order:
  1. st.secrets["DATABASE_URL"]   (Streamlit Community Cloud)
  2. env var DATABASE_URL
  3. none -> local SQLite "expenses.db"

Run locally:
    pip install -r requirements.txt
    streamlit run app.py
"""

import os
from io import BytesIO
from datetime import date, datetime
from calendar import monthrange

import pandas as pd
import altair as alt
import streamlit as st

# ----------------------------------------------------------------------
# Categories
# ----------------------------------------------------------------------
COMMUTE_CATEGORIES = ["InDrive to office", "InDrive back home"]
FOOD_CATEGORIES = ["Breakfast", "Lunch", "Dinner"]
DAILY_CATEGORIES = COMMUTE_CATEGORIES + FOOD_CATEGORIES
MONTHLY_CATEGORIES = ["Rent", "Internet bill", "Mobile package"]
MISC_CATEGORY = "Miscellaneous"
ALL_CATEGORIES = DAILY_CATEGORIES + MONTHLY_CATEGORIES + [MISC_CATEGORY]

CURRENCY = "PKR"

st.set_page_config(page_title="My Expense Tracker", page_icon="💸", layout="wide")

# ----------------------------------------------------------------------
# A little styling polish
# ----------------------------------------------------------------------
st.markdown(
    """
    <style>
      .block-container { padding-top: 2rem; }
      div[data-testid="stMetricValue"] { font-size: 1.6rem; }
      div[data-testid="stMetric"] {
          background: rgba(108, 92, 231, 0.08);
          border: 1px solid rgba(108, 92, 231, 0.20);
          border-radius: 12px;
          padding: 14px 16px;
      }
      h1, h2, h3 { letter-spacing: -0.3px; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------
# Database
# ----------------------------------------------------------------------
def get_database_url():
    try:
        if "DATABASE_URL" in st.secrets:
            return st.secrets["DATABASE_URL"]
    except Exception:
        pass
    return os.environ.get("DATABASE_URL")


DATABASE_URL = get_database_url()
USE_POSTGRES = bool(DATABASE_URL)
SQLITE_PATH = "expenses.db"


def get_conn():
    if USE_POSTGRES:
        import psycopg2
        return psycopg2.connect(DATABASE_URL, sslmode="require")
    import sqlite3
    return sqlite3.connect(SQLITE_PATH)


def q(sql):
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
        cur.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
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
        cur.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
    conn.commit()
    cur.close()
    conn.close()


# ----- settings -----
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
        cur.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, str(value)))
    conn.commit()
    cur.close()
    conn.close()


# ----- CRUD -----
def create_expense(category, amount, description, spent_on):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        q("INSERT INTO expenses (category, amount, description, spent_on) VALUES (%s, %s, %s, %s)"),
        (category, float(amount), description or "", str(spent_on)),
    )
    conn.commit()
    cur.close()
    conn.close()


def update_expense(expense_id, category, amount, description, spent_on):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        q("UPDATE expenses SET category=%s, amount=%s, description=%s, spent_on=%s WHERE id=%s"),
        (category, float(amount), description or "", str(spent_on), int(expense_id)),
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
        df["description"] = df["description"].fillna("")
    return df


def money(x):
    try:
        return f"{CURRENCY} {float(x):,.0f}"
    except Exception:
        return f"{CURRENCY} 0"


# ----------------------------------------------------------------------
# Excel export (formatted)
# ----------------------------------------------------------------------
def build_excel(df, month_label, salary, budget, total_spent):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    header_fill = PatternFill("solid", fgColor="6C5CE7")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    title_font = Font(bold=True, size=14, color="1A1A2E")
    label_font = Font(bold=True, size=11)
    money_fmt = '#,##0'
    thin = Side(style="thin", color="D9D9E3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    # ---- Sheet 1: Expenses ----
    ws = wb.active
    ws.title = "Expenses"
    ws["A1"] = f"Expenses - {month_label}"
    ws["A1"].font = title_font
    ws.merge_cells("A1:D1")

    headers = ["Date", "Category", "Amount (PKR)", "Description"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=col, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = border

    r = 4
    data = df.sort_values("spent_on") if not df.empty else df
    for _, row in data.iterrows():
        ws.cell(row=r, column=1, value=pd.to_datetime(row["spent_on"]).strftime("%Y-%m-%d")).border = border
        ws.cell(row=r, column=2, value=row["category"]).border = border
        amt = ws.cell(row=r, column=3, value=float(row["amount"]))
        amt.number_format = money_fmt
        amt.border = border
        ws.cell(row=r, column=4, value=row["description"]).border = border
        r += 1

    # total row
    tcell = ws.cell(row=r, column=2, value="TOTAL")
    tcell.font = label_font
    tsum = ws.cell(row=r, column=3, value=float(total_spent))
    tsum.number_format = money_fmt
    tsum.font = label_font

    widths = [14, 22, 16, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    if r > 4:
        ws.auto_filter.ref = f"A3:D{r-1}"

    # ---- Sheet 2: Summary ----
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = f"Summary - {month_label}"
    ws2["A1"].font = title_font
    ws2.merge_cells("A1:C1")

    rows_info = [
        ("Monthly salary", salary),
        ("Monthly budget", budget),
        ("Total spent", total_spent),
        ("Left from budget", (budget - total_spent) if budget else 0),
        ("Left from salary", (salary - total_spent) if salary else 0),
    ]
    rr = 3
    for label, val in rows_info:
        lc = ws2.cell(row=rr, column=1, value=label)
        lc.font = label_font
        vc = ws2.cell(row=rr, column=2, value=float(val))
        vc.number_format = money_fmt
        rr += 1

    # category breakdown
    rr += 1
    ws2.cell(row=rr, column=1, value="Category").fill = header_fill
    ws2.cell(row=rr, column=1).font = header_font
    ws2.cell(row=rr, column=2, value="Amount (PKR)").fill = header_fill
    ws2.cell(row=rr, column=2).font = header_font
    ws2.cell(row=rr, column=3, value="Share").fill = header_fill
    ws2.cell(row=rr, column=3).font = header_font
    rr += 1

    if not df.empty:
        by_cat = df.groupby("category")["amount"].sum().reindex(ALL_CATEGORIES).fillna(0)
        by_cat = by_cat[by_cat > 0]
        for cat, amt in by_cat.items():
            ws2.cell(row=rr, column=1, value=cat)
            ac = ws2.cell(row=rr, column=2, value=float(amt))
            ac.number_format = money_fmt
            share = (amt / total_spent * 100) if total_spent else 0
            ws2.cell(row=rr, column=3, value=f"{share:.1f}%")
            rr += 1

    for i, w in enumerate([22, 16, 10], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ======================================================================
# App
# ======================================================================
init_db()
today = date.today()

st.title("💸 My Expense Tracker")

# ---- Sidebar: setup ----
with st.sidebar:
    st.header("⚙️ Setup")
    st.caption("Set these once. Saved automatically.")

    cur_salary = float(get_setting("monthly_salary", 0) or 0)
    cur_budget = float(get_setting("monthly_budget", 0) or 0)
    cur_sal_day = int(get_setting("salary_day", 1) or 1)

    salary_in = st.number_input("Monthly salary (PKR)", min_value=0, value=int(cur_salary),
                                step=1000, format="%d")
    budget_in = st.number_input("Monthly budget (PKR)", min_value=0, value=int(cur_budget),
                                step=1000, format="%d", help="How much you plan to spend per month.")
    salday_in = st.number_input("Salary date (day of month)", min_value=1, max_value=31,
                                value=cur_sal_day, step=1, help="Day your salary arrives.")

    if st.button("💾 Save setup", use_container_width=True):
        set_setting("monthly_salary", salary_in)
        set_setting("monthly_budget", budget_in)
        set_setting("salary_day", salday_in)
        st.success("Saved!")
        st.rerun()

    st.divider()
    st.caption("Storage: " + ("Postgres (permanent)" if USE_POSTGRES else "Local SQLite (testing only)"))

# ---- Month selector ----
ctop1, ctop2, _ = st.columns([1, 1, 3])
with ctop1:
    years = list(range(today.year - 2, today.year + 1))
    sel_year = st.selectbox("Year", years, index=len(years) - 1)
with ctop2:
    sel_month = st.selectbox("Month", list(range(1, 13)), index=today.month - 1,
                             format_func=lambda m: datetime(2000, m, 1).strftime("%B"))
month_key = f"{sel_year:04d}-{sel_month:02d}"
month_label = datetime(sel_year, sel_month, 1).strftime("%B %Y")

df = load_expenses(month=month_key)
total_spent = float(df["amount"].sum()) if not df.empty else 0.0
salary = float(get_setting("monthly_salary", 0) or 0)
budget = float(get_setting("monthly_budget", 0) or 0)
sal_day = int(get_setting("salary_day", 1) or 1)

tab_dash, tab_add, tab_manage, tab_export = st.tabs(
    ["📊 Dashboard", "➕ Add", "🧾 Manage (edit / delete)", "⬇️ Export"]
)

# ======================================================================
# TAB: Dashboard
# ======================================================================
with tab_dash:
    st.subheader(f"Overview for {month_label}")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total spent", money(total_spent))
    m2.metric("Monthly budget", money(budget),
              delta=(money(budget - total_spent) + " left") if budget else None)
    m3.metric("Salary", money(salary),
              delta=(money(salary - total_spent) + " left") if salary else None)

    # days to salary
    try:
        last_day = monthrange(today.year, today.month)[1]
        this_payday = date(today.year, today.month, min(sal_day, last_day))
        if today <= this_payday:
            days_to = (this_payday - today).days
        else:
            nm = today.month + 1 if today.month < 12 else 1
            ny = today.year if today.month < 12 else today.year + 1
            nl = monthrange(ny, nm)[1]
            days_to = (date(ny, nm, min(sal_day, nl)) - today).days
        m4.metric("Days to salary", f"{days_to} days")
    except Exception:
        m4.metric("Days to salary", "-")

    if budget > 0:
        pct = min(total_spent / budget, 1.0)
        st.progress(pct, text=f"Budget used: {money(total_spent)} / {money(budget)} ({pct*100:.0f}%)")
        if total_spent > budget:
            st.warning(f"Over budget by {money(total_spent - budget)} this month.")

    st.divider()

    if df.empty:
        st.info("No expenses yet for this month. Add one in the ➕ Add tab.")
    else:
        left, right = st.columns(2)
        with left:
            st.markdown("**Spending by category**")
            by_cat = (df.groupby("category")["amount"].sum()
                        .reindex(ALL_CATEGORIES).fillna(0))
            by_cat = by_cat[by_cat > 0].reset_index()
            by_cat.columns = ["Category", "Amount"]
            chart = (
                alt.Chart(by_cat)
                .mark_bar(cornerRadiusEnd=4, color="#6C5CE7")
                .encode(
                    x=alt.X("Amount:Q", title="PKR", axis=alt.Axis(format=",.0f")),
                    y=alt.Y("Category:N", sort="-x", title=None),
                    tooltip=[alt.Tooltip("Category:N"),
                             alt.Tooltip("Amount:Q", format=",.0f", title="PKR")],
                )
                .properties(height=280)
            )
            st.altair_chart(chart, use_container_width=True)

        with right:
            st.markdown("**Daily spending this month**")
            by_day = (df.groupby("spent_on")["amount"].sum().sort_index().reset_index())
            by_day.columns = ["Date", "Amount"]
            line = (
                alt.Chart(by_day)
                .mark_line(point=True, color="#00B894")
                .encode(
                    x=alt.X("Date:T", title=None),
                    y=alt.Y("Amount:Q", title="PKR", axis=alt.Axis(format=",.0f")),
                    tooltip=[alt.Tooltip("Date:T"),
                             alt.Tooltip("Amount:Q", format=",.0f", title="PKR")],
                )
                .properties(height=280)
            )
            st.altair_chart(line, use_container_width=True)

        st.markdown("**Category breakdown**")
        bd = df.groupby("category")["amount"].sum().reindex(ALL_CATEGORIES).fillna(0)
        bd = bd[bd > 0].reset_index()
        bd.columns = ["Category", "Amount"]
        bd["Share"] = (bd["Amount"] / total_spent * 100).round(1).astype(str) + "%"
        bd["Amount"] = bd["Amount"].apply(money)
        st.dataframe(bd, use_container_width=True, hide_index=True)

# ======================================================================
# TAB: Add
# ======================================================================
with tab_add:
    st.subheader("Add an expense")
    with st.form("add_form", clear_on_submit=True):
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            cat = st.selectbox("Category", ALL_CATEGORIES)
        with c2:
            amt = st.number_input("Amount (PKR)", min_value=0, step=50, format="%d")
        with c3:
            when = st.date_input("Date", value=today)
        desc = st.text_input("Description (optional, useful for Miscellaneous)",
                             placeholder="e.g. Haircut, medicine, gift")
        ok = st.form_submit_button("Add expense", use_container_width=True, type="primary")
        if ok:
            if amt <= 0:
                st.error("Amount must be greater than 0.")
            else:
                create_expense(cat, amt, desc, when)
                st.success(f"Added {money(amt)} for {cat}.")
                st.rerun()

# ======================================================================
# TAB: Manage (edit / delete)  -> full CRUD
# ======================================================================
with tab_manage:
    st.subheader(f"Manage expenses - {month_label}")
    if df.empty:
        st.info("Nothing to manage for this month yet.")
    else:
        st.caption("Edit any cell directly. Use the trash icon on the left of a row to delete it. "
                   "Then press Save changes.")

        editable = df[["id", "spent_on", "category", "amount", "description"]].copy()

        edited = st.data_editor(
            editable,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "id": None,
                "spent_on": st.column_config.DateColumn("Date", format="YYYY-MM-DD"),
                "category": st.column_config.SelectboxColumn("Category", options=ALL_CATEGORIES, required=True),
                "amount": st.column_config.NumberColumn("Amount (PKR)", min_value=0, step=50, format="%d"),
                "description": st.column_config.TextColumn("Description"),
            },
            key="editor",
        )

        if st.button("💾 Save changes", type="primary"):
            orig_by_id = {int(r["id"]): r for _, r in df.iterrows()}
            seen_ids = set()
            changed = added = removed = 0

            for _, row in edited.iterrows():
                rid = row.get("id")
                has_id = pd.notna(rid)
                amount = row.get("amount")
                category = row.get("category")
                spent = row.get("spent_on")
                description = row.get("description") or ""

                # skip blank new rows
                if not has_id and (pd.isna(amount) or not category or not spent):
                    continue

                if has_id:
                    rid = int(rid)
                    seen_ids.add(rid)
                    o = orig_by_id.get(rid)
                    if o is not None:
                        if (str(o["spent_on"]) != str(spent) or o["category"] != category
                                or float(o["amount"]) != float(amount or 0)
                                or (o["description"] or "") != description):
                            update_expense(rid, category, amount or 0, description, spent)
                            changed += 1
                else:
                    if amount and float(amount) > 0:
                        create_expense(category, amount, description, spent)
                        added += 1

            # anything in original not seen = deleted in the grid
            for rid in orig_by_id:
                if rid not in seen_ids:
                    delete_expense(rid)
                    removed += 1

            st.success(f"Saved. {changed} edited, {added} added, {removed} deleted.")
            st.rerun()

# ======================================================================
# TAB: Export
# ======================================================================
with tab_export:
    st.subheader("Download your data")
    if df.empty:
        st.info("No data for this month to export yet.")
    else:
        excel_bytes = build_excel(df, month_label, salary, budget, total_spent)
        st.download_button(
            "⬇️ Download Excel (formatted)",
            data=excel_bytes,
            file_name=f"expenses_{month_key}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        csv = df[["spent_on", "category", "amount", "description"]].to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download CSV (plain)",
            data=csv,
            file_name=f"expenses_{month_key}.csv",
            mime="text/csv",
            use_container_width=True,
        )
        st.caption("The Excel file has two sheets: a formatted Expenses list with a total, "
                   "and a Summary with salary, budget, and category breakdown.")

"""
Personal Expense Tracker (Streamlit)
=====================================
Single-user daily/monthly expense tracker with charts, full CRUD,
a formatted Excel export, and smart salary-date handling. Currency: PKR.

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
from datetime import date, datetime, timedelta
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
DEFAULT_SALARY_DAY = 25

st.set_page_config(page_title="My Expense Tracker", page_icon="💸", layout="wide")

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
# Database connection (reused, so the app stays fast)
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


@st.cache_resource
def _pg_conn():
    """One Postgres connection, reused across reruns. This is the speed fix."""
    import psycopg2
    conn = psycopg2.connect(
        DATABASE_URL,
        sslmode="require",
        keepalives=1,
        keepalives_idle=30,
        keepalives_interval=10,
        keepalives_count=5,
    )
    conn.autocommit = True
    return conn


def q(sql):
    return sql if USE_POSTGRES else sql.replace("%s", "?")


def run(sql, params=(), fetch=False):
    """Run one statement. Reuses the cached connection on Postgres."""
    sql = q(sql)
    if USE_POSTGRES:
        import psycopg2
        for attempt in range(2):
            try:
                conn = _pg_conn()
                cur = conn.cursor()
                cur.execute(sql, params)
                data = cur.fetchall() if fetch else None
                cur.close()
                return data
            except (psycopg2.InterfaceError, psycopg2.OperationalError):
                # connection went stale (Supabase closed it) -> rebuild once
                _pg_conn.clear()
                if attempt == 1:
                    raise
    else:
        import sqlite3
        conn = sqlite3.connect(SQLITE_PATH)
        try:
            cur = conn.cursor()
            cur.execute(sql, params)
            data = cur.fetchall() if fetch else None
            conn.commit()
            return data
        finally:
            conn.close()


def init_db():
    if USE_POSTGRES:
        run("""CREATE TABLE IF NOT EXISTS expenses (
                id SERIAL PRIMARY KEY, category TEXT NOT NULL, amount NUMERIC NOT NULL,
                description TEXT DEFAULT '', spent_on DATE NOT NULL,
                created_at TIMESTAMP DEFAULT NOW())""")
    else:
        run("""CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT, category TEXT NOT NULL, amount REAL NOT NULL,
                description TEXT DEFAULT '', spent_on TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
    run("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")


# run schema check only once per session, not on every rerun
if "db_ready" not in st.session_state:
    init_db()
    st.session_state["db_ready"] = True


# ----- settings (loaded in one query per rerun) -----
def load_settings():
    rows = run("SELECT key, value FROM settings", fetch=True) or []
    return {k: v for k, v in rows}


def set_setting(key, value):
    if USE_POSTGRES:
        run("INSERT INTO settings (key, value) VALUES (%s, %s) "
            "ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value", (key, str(value)))
    else:
        run("INSERT OR REPLACE INTO settings (key, value) VALUES (%s, %s)", (key, str(value)))


# ----- CRUD -----
def create_expense(category, amount, description, spent_on):
    run("INSERT INTO expenses (category, amount, description, spent_on) VALUES (%s, %s, %s, %s)",
        (category, float(amount), description or "", str(spent_on)))


def update_expense(expense_id, category, amount, description, spent_on):
    run("UPDATE expenses SET category=%s, amount=%s, description=%s, spent_on=%s WHERE id=%s",
        (category, float(amount), description or "", str(spent_on), int(expense_id)))


def delete_expense(expense_id):
    run("DELETE FROM expenses WHERE id = %s", (int(expense_id),))


def load_expenses(month=None):
    if month:
        rows = run("SELECT id, category, amount, description, spent_on FROM expenses "
                   "WHERE CAST(spent_on AS TEXT) LIKE %s ORDER BY spent_on DESC, id DESC",
                   (month + "%",), fetch=True)
    else:
        rows = run("SELECT id, category, amount, description, spent_on FROM expenses "
                   "ORDER BY spent_on DESC, id DESC", fetch=True)
    rows = rows or []
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
# Salary date logic
# ----------------------------------------------------------------------
def estimated_payday(year, month, normal_day):
    """The normal_day of the month, moved back to Friday if it lands on a weekend."""
    last = monthrange(year, month)[1]
    d = date(year, month, min(normal_day, last))
    wd = d.weekday()  # Mon=0 ... Sat=5, Sun=6
    if wd == 5:        # Saturday -> Friday
        d -= timedelta(days=1)
    elif wd == 6:      # Sunday -> Friday
        d -= timedelta(days=2)
    return d


def payday_for_month(settings, year, month, normal_day):
    """Manual override for this month if set, else the weekend-adjusted estimate."""
    key = f"payday_{year:04d}-{month:02d}"
    if settings.get(key):
        try:
            return date.fromisoformat(settings[key])
        except Exception:
            pass
    return estimated_payday(year, month, normal_day)


def next_payday(settings, normal_day, from_date):
    y, m = from_date.year, from_date.month
    pd_ = estimated_payday(y, m, normal_day)
    for _ in range(4):
        pd_ = payday_for_month(settings, y, m, normal_day)
        if pd_ >= from_date:
            return pd_
        m += 1
        if m > 12:
            m = 1
            y += 1
    return pd_


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
    money_fmt = "#,##0"
    thin = Side(style="thin", color="D9D9E3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    ws = wb.active
    ws.title = "Expenses"
    ws["A1"] = f"Expenses - {month_label}"
    ws["A1"].font = title_font
    ws.merge_cells("A1:D1")

    for col, h in enumerate(["Date", "Category", "Amount (PKR)", "Description"], start=1):
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

    ws.cell(row=r, column=2, value="TOTAL").font = label_font
    ts = ws.cell(row=r, column=3, value=float(total_spent))
    ts.number_format = money_fmt
    ts.font = label_font

    for i, w in enumerate([14, 22, 16, 40], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    if r > 4:
        ws.auto_filter.ref = f"A3:D{r-1}"

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = f"Summary - {month_label}"
    ws2["A1"].font = title_font
    ws2.merge_cells("A1:C1")

    rr = 3
    for label, val in [
        ("Monthly salary", salary),
        ("Monthly budget", budget),
        ("Total spent", total_spent),
        ("Left from budget", (budget - total_spent) if budget else 0),
        ("Left from salary", (salary - total_spent) if salary else 0),
    ]:
        ws2.cell(row=rr, column=1, value=label).font = label_font
        vc = ws2.cell(row=rr, column=2, value=float(val))
        vc.number_format = money_fmt
        rr += 1

    rr += 1
    for col, h in enumerate(["Category", "Amount (PKR)", "Share"], start=1):
        c = ws2.cell(row=rr, column=col, value=h)
        c.fill = header_fill
        c.font = header_font
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
today = date.today()
settings = load_settings()

st.title("💸 My Expense Tracker")

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

normal_day = int(settings.get("salary_day", DEFAULT_SALARY_DAY) or DEFAULT_SALARY_DAY)
salary = float(settings.get("monthly_salary", 0) or 0)
budget = float(settings.get("monthly_budget", 0) or 0)

# ---- Sidebar ----
with st.sidebar:
    st.header("⚙️ Setup")
    st.caption("Set these once. Saved automatically.")

    salary_in = st.number_input("Monthly salary (PKR)", min_value=0, value=int(salary),
                                step=1000, format="%d")
    budget_in = st.number_input("Monthly budget (PKR)", min_value=0, value=int(budget),
                                step=1000, format="%d", help="How much you plan to spend per month.")
    salday_in = st.number_input("Normal salary day", min_value=1, max_value=31, value=normal_day,
                                step=1, help="Usually 25. If it falls on a weekend the app moves it to the Friday before.")
    if st.button("💾 Save setup", use_container_width=True):
        set_setting("monthly_salary", salary_in)
        set_setting("monthly_budget", budget_in)
        set_setting("salary_day", salday_in)
        st.success("Saved!")
        st.rerun()

    st.divider()
    st.subheader("📅 Salary date")
    st.caption(f"For {month_label}. The app guesses it from your normal day. "
               "If a holiday moved it (like the 23rd this month), set the real date here.")
    default_pd = payday_for_month(settings, sel_year, sel_month, normal_day)
    actual_pd = st.date_input("Actual salary date", value=default_pd, key=f"pd_{month_key}")
    colp1, colp2 = st.columns(2)
    with colp1:
        if st.button("Save date", use_container_width=True):
            set_setting(f"payday_{month_key}", actual_pd.isoformat())
            st.success("Saved.")
            st.rerun()
    with colp2:
        if st.button("Reset", use_container_width=True, help="Go back to the automatic guess."):
            set_setting(f"payday_{month_key}", "")
            st.success("Reset.")
            st.rerun()

    st.divider()
    st.caption("Storage: " + ("Postgres (permanent)" if USE_POSTGRES else "Local SQLite (testing only)"))

# ---- Load month data ----
df = load_expenses(month=month_key)
total_spent = float(df["amount"].sum()) if not df.empty else 0.0

tab_dash, tab_add, tab_manage, tab_export = st.tabs(
    ["📊 Dashboard", "➕ Add", "🧾 Manage (edit / delete)", "⬇️ Export"]
)

# ======================================================================
# Dashboard
# ======================================================================
with tab_dash:
    st.subheader(f"Overview for {month_label}")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total spent", money(total_spent))
    m2.metric("Monthly budget", money(budget),
              delta=(money(budget - total_spent) + " left") if budget else None)
    m3.metric("Salary", money(salary),
              delta=(money(salary - total_spent) + " left") if salary else None)

    np_ = next_payday(settings, normal_day, today)
    days_to = (np_ - today).days
    m4.metric("Next salary", np_.strftime("%d %b"),
              delta=f"in {days_to} days" if days_to > 0 else "today",
              delta_color="off")

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
            bc = df.groupby("category")["amount"].sum().reindex(ALL_CATEGORIES).fillna(0)
            bc = bc[bc > 0].reset_index()
            bc.columns = ["Category", "Amount"]
            st.altair_chart(
                alt.Chart(bc).mark_bar(cornerRadiusEnd=4, color="#6C5CE7").encode(
                    x=alt.X("Amount:Q", title="PKR", axis=alt.Axis(format=",.0f")),
                    y=alt.Y("Category:N", sort="-x", title=None),
                    tooltip=[alt.Tooltip("Category:N"), alt.Tooltip("Amount:Q", format=",.0f", title="PKR")],
                ).properties(height=280),
                use_container_width=True,
            )
        with right:
            st.markdown("**Daily spending this month**")
            bd = df.groupby("spent_on")["amount"].sum().sort_index().reset_index()
            bd.columns = ["Date", "Amount"]
            st.altair_chart(
                alt.Chart(bd).mark_line(point=True, color="#00B894").encode(
                    x=alt.X("Date:T", title=None),
                    y=alt.Y("Amount:Q", title="PKR", axis=alt.Axis(format=",.0f")),
                    tooltip=[alt.Tooltip("Date:T"), alt.Tooltip("Amount:Q", format=",.0f", title="PKR")],
                ).properties(height=280),
                use_container_width=True,
            )

        st.markdown("**Category breakdown**")
        bd2 = df.groupby("category")["amount"].sum().reindex(ALL_CATEGORIES).fillna(0)
        bd2 = bd2[bd2 > 0].reset_index()
        bd2.columns = ["Category", "Amount"]
        bd2["Share"] = (bd2["Amount"] / total_spent * 100).round(1).astype(str) + "%"
        bd2["Amount"] = bd2["Amount"].apply(money)
        st.dataframe(bd2, use_container_width=True, hide_index=True)

# ======================================================================
# Add
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
        if st.form_submit_button("Add expense", use_container_width=True, type="primary"):
            if amt <= 0:
                st.error("Amount must be greater than 0.")
            else:
                create_expense(cat, amt, desc, when)
                st.success(f"Added {money(amt)} for {cat}.")
                st.rerun()

# ======================================================================
# Manage (full CRUD)
# ======================================================================
with tab_manage:
    st.subheader(f"Manage expenses - {month_label}")
    if df.empty:
        st.info("Nothing to manage for this month yet.")
    else:
        st.caption("Edit any cell directly. Use the trash icon to delete a row. Then press Save changes.")
        editable = df[["id", "spent_on", "category", "amount", "description"]].copy()
        edited = st.data_editor(
            editable, use_container_width=True, hide_index=True, num_rows="dynamic",
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
            seen = set()
            changed = added = removed = 0
            for _, row in edited.iterrows():
                rid = row.get("id")
                has_id = pd.notna(rid)
                amount = row.get("amount")
                category = row.get("category")
                spent = row.get("spent_on")
                description = row.get("description") or ""
                if not has_id and (pd.isna(amount) or not category or not spent):
                    continue
                if has_id:
                    rid = int(rid)
                    seen.add(rid)
                    o = orig_by_id.get(rid)
                    if o is not None and (
                        str(o["spent_on"]) != str(spent) or o["category"] != category
                        or float(o["amount"]) != float(amount or 0)
                        or (o["description"] or "") != description
                    ):
                        update_expense(rid, category, amount or 0, description, spent)
                        changed += 1
                else:
                    if amount and float(amount) > 0:
                        create_expense(category, amount, description, spent)
                        added += 1
            for rid in orig_by_id:
                if rid not in seen:
                    delete_expense(rid)
                    removed += 1
            st.success(f"Saved. {changed} edited, {added} added, {removed} deleted.")
            st.rerun()

# ======================================================================
# Export
# ======================================================================
with tab_export:
    st.subheader("Download your data")
    if df.empty:
        st.info("No data for this month to export yet.")
    else:
        st.download_button(
            "⬇️ Download Excel (formatted)",
            data=build_excel(df, month_label, salary, budget, total_spent),
            file_name=f"expenses_{month_key}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
        st.download_button(
            "⬇️ Download CSV (plain)",
            data=df[["spent_on", "category", "amount", "description"]].to_csv(index=False).encode("utf-8"),
            file_name=f"expenses_{month_key}.csv",
            mime="text/csv", use_container_width=True,
        )
        st.caption("Excel has two sheets: a formatted Expenses list with a total, "
                   "and a Summary with salary, budget, and category breakdown.")

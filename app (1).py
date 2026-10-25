"""
Personal Expense Tracker (Streamlit)
=====================================
Single-user expense tracker with charts, full CRUD, custom categories,
a formatted Excel export, smart salary dates, and a pay-cycle view.
Currency: PKR.

"Month" grouping:
  - Pay cycle (default): salary date to the day before the next salary.
    The salary received end of October funds the "November" cycle
    (e.g. 23 Oct to 24 Nov).
  - Calendar month: plain 1st to end of month.

Storage (automatic):
  - Postgres if a connection string is provided (permanent)
  - else local SQLite file (testing)
Read from st.secrets["DATABASE_URL"], then env DATABASE_URL, else SQLite.

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
# Built-in categories & groups
# ----------------------------------------------------------------------
BUILTIN_CATEGORIES = [
    "InDrive to office", "InDrive back home",
    "Breakfast", "Lunch", "Dinner",
    "Rent", "Internet bill", "Mobile package",
    "Miscellaneous",
]
BUILTIN_GROUPS = {
    "InDrive to office": "Commute", "InDrive back home": "Commute",
    "Breakfast": "Food", "Lunch": "Food", "Dinner": "Food",
    "Rent": "Fixed monthly", "Internet bill": "Fixed monthly", "Mobile package": "Fixed monthly",
    "Miscellaneous": "Other",
}
GROUP_OPTIONS = ["Commute", "Food", "Fixed monthly", "Entertainment", "Shopping", "Health", "Other"]

CURRENCY = "PKR"
DEFAULT_SALARY_DAY = 25

st.set_page_config(page_title="My Expense Tracker", page_icon="💸", layout="wide")

st.markdown(
    """
    <style>
      .block-container { padding-top: 2rem; }
      div[data-testid="stMetricValue"] { font-size: 1.5rem; }
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
# Database (one reused connection = fast)
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
    import psycopg2
    conn = psycopg2.connect(
        DATABASE_URL, sslmode="require",
        keepalives=1, keepalives_idle=30, keepalives_interval=10, keepalives_count=5,
    )
    conn.autocommit = True
    return conn


def q(sql):
    return sql if USE_POSTGRES else sql.replace("%s", "?")


def run(sql, params=(), fetch=False):
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
    run("CREATE TABLE IF NOT EXISTS categories (name TEXT PRIMARY KEY, grp TEXT)")


if "db_ready" not in st.session_state:
    init_db()
    st.session_state["db_ready"] = True


def load_settings():
    rows = run("SELECT key, value FROM settings", fetch=True) or []
    return {k: v for k, v in rows}


def set_setting(key, value):
    if USE_POSTGRES:
        run("INSERT INTO settings (key, value) VALUES (%s, %s) "
            "ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value", (key, str(value)))
    else:
        run("INSERT OR REPLACE INTO settings (key, value) VALUES (%s, %s)", (key, str(value)))


# ----- custom categories -----
def load_custom_categories():
    rows = run("SELECT name, grp FROM categories ORDER BY name", fetch=True) or []
    return [(r[0], r[1]) for r in rows]


def add_category(name, grp):
    if USE_POSTGRES:
        run("INSERT INTO categories (name, grp) VALUES (%s, %s) "
            "ON CONFLICT (name) DO UPDATE SET grp = EXCLUDED.grp", (name, grp))
    else:
        run("INSERT OR REPLACE INTO categories (name, grp) VALUES (%s, %s)", (name, grp))


def delete_category(name):
    run("DELETE FROM categories WHERE name = %s", (name,))


# ----- CRUD -----
def create_expense(category, amount, description, spent_on):
    run("INSERT INTO expenses (category, amount, description, spent_on) VALUES (%s, %s, %s, %s)",
        (category, float(amount), description or "", str(spent_on)))


def update_expense(expense_id, category, amount, description, spent_on):
    run("UPDATE expenses SET category=%s, amount=%s, description=%s, spent_on=%s WHERE id=%s",
        (category, float(amount), description or "", str(spent_on), int(expense_id)))


def delete_expense(expense_id):
    run("DELETE FROM expenses WHERE id = %s", (int(expense_id),))


def load_expenses_range(start, end):
    rows = run("SELECT id, category, amount, description, spent_on FROM expenses "
               "WHERE spent_on >= %s AND spent_on <= %s ORDER BY spent_on DESC, id DESC",
               (str(start), str(end)), fetch=True) or []
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
# Month / salary-date helpers
# ----------------------------------------------------------------------
def prev_month(y, m):
    return (y - 1, 12) if m == 1 else (y, m - 1)


def next_month(y, m):
    return (y + 1, 1) if m == 12 else (y, m + 1)


def estimated_payday(year, month, normal_day):
    last = monthrange(year, month)[1]
    d = date(year, month, min(normal_day, last))
    wd = d.weekday()
    if wd == 5:
        d -= timedelta(days=1)
    elif wd == 6:
        d -= timedelta(days=2)
    return d


def payday_for_month(settings, year, month, normal_day):
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
        y, m = next_month(y, m)
    return pd_


def cycle_bounds(settings, normal_day, label_year, label_month):
    py, pm = prev_month(label_year, label_month)
    start = payday_for_month(settings, py, pm, normal_day)
    end = payday_for_month(settings, label_year, label_month, normal_day) - timedelta(days=1)
    return start, end


def current_cycle_label(settings, normal_day, today):
    pay_this = payday_for_month(settings, today.year, today.month, normal_day)
    if today >= pay_this:
        return next_month(today.year, today.month)
    return (today.year, today.month)


# ----------------------------------------------------------------------
# Excel export
# ----------------------------------------------------------------------
def build_excel(df, label, salary, budget, total_spent, all_categories):
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
    ws["A1"] = f"Expenses - {label}"
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
    ws2["A1"] = f"Summary - {label}"
    ws2["A1"].font = title_font
    ws2.merge_cells("A1:C1")

    rr = 3
    for lbl, val in [
        ("Salary", salary), ("Budget", budget), ("Total spent", total_spent),
        ("Left from budget", (budget - total_spent) if budget else 0),
        ("Left from salary", (salary - total_spent) if salary else 0),
    ]:
        ws2.cell(row=rr, column=1, value=lbl).font = label_font
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
        by_cat = df.groupby("category")["amount"].sum().sort_values(ascending=False)
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
custom_cats = load_custom_categories()

# dynamic category list + group map (built-in + custom)
ALL_CATEGORIES = BUILTIN_CATEGORIES + [n for n, _ in custom_cats if n not in BUILTIN_CATEGORIES]
GROUPS = dict(BUILTIN_GROUPS)
for n, g in custom_cats:
    GROUPS[n] = g or "Other"

normal_day = int(settings.get("salary_day", DEFAULT_SALARY_DAY) or DEFAULT_SALARY_DAY)
salary = float(settings.get("monthly_salary", 0) or 0)
budget = float(settings.get("monthly_budget", 0) or 0)
use_cycle = (settings.get("view_mode", "cycle") == "cycle")

st.title("💸 My Expense Tracker")

# ---- Period selector ----
cur_label_y, cur_label_m = current_cycle_label(settings, normal_day, today) if use_cycle else (today.year, today.month)
years = list(range(today.year - 2, today.year + 2))
ctop1, ctop2, ctop3 = st.columns([1, 1, 2])
with ctop1:
    sel_year = st.selectbox("Year", years, index=years.index(cur_label_y))
with ctop2:
    sel_month = st.selectbox(
        "Salary month" if use_cycle else "Month",
        list(range(1, 13)), index=cur_label_m - 1,
        format_func=lambda m: datetime(2000, m, 1).strftime("%B"),
    )
with ctop3:
    mode_choice = st.radio(
        "View", ["Pay cycle", "Calendar month"],
        index=0 if use_cycle else 1, horizontal=True,
        help="Pay cycle = salary date to next salary date. Calendar = 1st to month end.",
    )
    new_mode = "cycle" if mode_choice == "Pay cycle" else "calendar"
    if new_mode != settings.get("view_mode", "cycle"):
        set_setting("view_mode", new_mode)
        st.rerun()

label = datetime(sel_year, sel_month, 1).strftime("%B %Y")

if use_cycle:
    period_start, period_end = cycle_bounds(settings, normal_day, sel_year, sel_month)
    period_key = f"cycle_{sel_year:04d}-{sel_month:02d}"
else:
    last = monthrange(sel_year, sel_month)[1]
    period_start, period_end = date(sel_year, sel_month, 1), date(sel_year, sel_month, last)
    period_key = f"{sel_year:04d}-{sel_month:02d}"
range_text = f"{period_start.strftime('%d %b')} to {period_end.strftime('%d %b %Y')}"

# ---- Sidebar ----
with st.sidebar:
    st.header("⚙️ Setup")
    st.caption("Set these once. Saved automatically.")
    salary_in = st.number_input("Salary per cycle (PKR)", min_value=0, value=int(salary),
                                step=1000, format="%d")
    budget_in = st.number_input("Budget per cycle (PKR)", min_value=0, value=int(budget),
                                step=1000, format="%d", help="How much you plan to spend per cycle.")
    salday_in = st.number_input("Normal salary day", min_value=1, max_value=31, value=normal_day,
                                step=1, help="Usually 25. If it falls on a weekend the app uses the Friday before.")
    if st.button("💾 Save setup", use_container_width=True):
        set_setting("monthly_salary", salary_in)
        set_setting("monthly_budget", budget_in)
        set_setting("salary_day", salday_in)
        st.success("Saved!")
        st.rerun()

    st.divider()
    st.subheader("📅 Salary date")
    if use_cycle:
        st.caption(f"Date the salary for the {label} cycle arrived (starts this cycle). "
                   "Set the real date if a holiday moved it.")
        py, pm = prev_month(sel_year, sel_month)
        default_pd = payday_for_month(settings, py, pm, normal_day)
        pd_key_target = f"payday_{py:04d}-{pm:02d}"
    else:
        st.caption(f"Date the salary arrived in {label}. Set the real date if a holiday moved it.")
        default_pd = payday_for_month(settings, sel_year, sel_month, normal_day)
        pd_key_target = f"payday_{sel_year:04d}-{sel_month:02d}"

    actual_pd = st.date_input("Actual salary date", value=default_pd, key=f"pd_{period_key}")
    cpd1, cpd2 = st.columns(2)
    with cpd1:
        if st.button("Save date", use_container_width=True):
            set_setting(pd_key_target, actual_pd.isoformat())
            st.success("Saved.")
            st.rerun()
    with cpd2:
        if st.button("Reset", use_container_width=True, help="Back to the automatic guess."):
            set_setting(pd_key_target, "")
            st.success("Reset.")
            st.rerun()

    st.divider()
    st.caption("Storage: " + ("Postgres (permanent)" if USE_POSTGRES else "Local SQLite (testing only)"))

# ---- Load period data ----
df = load_expenses_range(period_start, period_end)
total_spent = float(df["amount"].sum()) if not df.empty else 0.0
budget_left = budget - total_spent
salary_left = salary - total_spent

tab_dash, tab_add, tab_notes, tab_manage, tab_cats, tab_export = st.tabs(
    ["📊 Dashboard", "➕ Add", "📝 Notes", "🧾 Manage", "🏷️ Categories", "⬇️ Export"]
)

# ======================================================================
# Dashboard
# ======================================================================
with tab_dash:
    title_word = "cycle" if use_cycle else "month"
    st.subheader(f"Overview for {label} {title_word}")
    st.caption(f"Covers {range_text}")

    # ---- clear money tiles: spent, budget left, salary left, next salary ----
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Spent", money(total_spent))

    if budget > 0:
        m2.metric("Budget left", money(budget_left),
                  delta=f"of {money(budget)}", delta_color="off")
    else:
        m2.metric("Budget left", "set budget", help="Set a budget in the sidebar.")

    if salary > 0:
        m3.metric("Salary left", money(salary_left),
                  delta=f"of {money(salary)}", delta_color="off")
    else:
        m3.metric("Salary left", "set salary", help="Set your salary in the sidebar.")

    np_ = next_payday(settings, normal_day, today)
    days_to = (np_ - today).days
    m4.metric("Next salary", np_.strftime("%d %b"),
              delta=(f"in {days_to} days" if days_to > 0 else "today"), delta_color="off")

    # ---- budget bar + safe-to-spend ----
    if budget > 0:
        pct = min(total_spent / budget, 1.0)
        st.progress(pct, text=f"Budget used: {money(total_spent)} / {money(budget)} ({pct*100:.0f}%)")
        if total_spent > budget:
            st.error(f"Over budget by {money(total_spent - budget)} this {title_word}.")
        else:
            st.success(f"You have {money(budget_left)} left to spend this {title_word}.")
            # safe-to-spend per day, only meaningful for the cycle you are currently in
            if period_start <= today <= period_end:
                days_left = (period_end - today).days + 1
                if days_left > 0:
                    st.caption(f"That is about {money(budget_left / days_left)} per day "
                               f"for the {days_left} days left until your next salary.")

    st.divider()
    if df.empty:
        st.info(f"No expenses yet for this {title_word}. Add one in the ➕ Add tab.")
    else:
        bc = df.groupby("category")["amount"].sum().sort_values(ascending=False).reset_index()
        bc.columns = ["Category", "Amount"]

        grp = df.copy()
        grp["Group"] = grp["category"].map(lambda c: GROUPS.get(c, "Other"))
        gg = grp.groupby("Group")["amount"].sum().reset_index()
        gg.columns = ["Group", "Amount"]

        r1c1, r1c2 = st.columns(2)
        with r1c1:
            st.markdown("**Spending by category**")
            st.altair_chart(
                alt.Chart(bc).mark_bar(cornerRadiusEnd=4, color="#6C5CE7").encode(
                    x=alt.X("Amount:Q", title="PKR", axis=alt.Axis(format=",.0f")),
                    y=alt.Y("Category:N", sort="-x", title=None),
                    tooltip=[alt.Tooltip("Category:N"), alt.Tooltip("Amount:Q", format=",.0f", title="PKR")],
                ).properties(height=300),
                use_container_width=True,
            )
        with r1c2:
            st.markdown("**Daily spending**")
            bd = df.groupby("spent_on")["amount"].sum().sort_index().reset_index()
            bd.columns = ["Date", "Amount"]
            st.altair_chart(
                alt.Chart(bd).mark_line(point=True, color="#00B894").encode(
                    x=alt.X("Date:T", title=None),
                    y=alt.Y("Amount:Q", title="PKR", axis=alt.Axis(format=",.0f")),
                    tooltip=[alt.Tooltip("Date:T"), alt.Tooltip("Amount:Q", format=",.0f", title="PKR")],
                ).properties(height=300),
                use_container_width=True,
            )

        r2c1, r2c2 = st.columns(2)
        with r2c1:
            st.markdown("**Category share (pie)**")
            st.altair_chart(
                alt.Chart(bc).mark_arc().encode(
                    theta=alt.Theta("Amount:Q", stack=True),
                    color=alt.Color("Category:N", scale=alt.Scale(scheme="tableau20"),
                                    legend=alt.Legend(title=None, orient="bottom", columns=2)),
                    tooltip=[alt.Tooltip("Category:N"), alt.Tooltip("Amount:Q", format=",.0f", title="PKR")],
                ).properties(height=320),
                use_container_width=True,
            )
        with r2c2:
            st.markdown("**Fixed vs flexible (donut)**")
            st.altair_chart(
                alt.Chart(gg).mark_arc(innerRadius=70).encode(
                    theta=alt.Theta("Amount:Q", stack=True),
                    color=alt.Color("Group:N", scale=alt.Scale(scheme="set2"),
                                    legend=alt.Legend(title=None, orient="bottom")),
                    tooltip=[alt.Tooltip("Group:N"), alt.Tooltip("Amount:Q", format=",.0f", title="PKR")],
                ).properties(height=320),
                use_container_width=True,
            )

        st.markdown("**Category breakdown**")
        tbl = bc.copy()
        tbl["Share"] = (tbl["Amount"] / total_spent * 100).round(1).astype(str) + "%"
        tbl["Amount"] = tbl["Amount"].apply(money)
        st.dataframe(tbl, use_container_width=True, hide_index=True)

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
        desc = st.text_input("Description (optional)",
                             placeholder="e.g. Movie with friends, medicine, gift")
        if st.form_submit_button("Add expense", use_container_width=True, type="primary"):
            if amt <= 0:
                st.error("Amount must be greater than 0.")
            else:
                create_expense(cat, amt, desc, when)
                st.success(f"Added {money(amt)} for {cat}.")
                st.rerun()
    st.caption("Need a category that is not listed? Add it in the 🏷️ Categories tab.")

# ======================================================================
# Notes (view descriptions)
# ======================================================================
with tab_notes:
    st.subheader(f"Notes & descriptions - {label}")
    if df.empty:
        st.info("No expenses this period yet.")
    else:
        with_notes = df[df["description"].str.strip() != ""].copy()
        cats_present = ["All categories"] + sorted(df["category"].unique().tolist())
        pick = st.selectbox("Filter by category", cats_present)
        view = with_notes if pick == "All categories" else with_notes[with_notes["category"] == pick]

        if view.empty:
            st.info("No descriptions written for this selection. "
                    "Add a description when you log an expense and it will show here.")
        else:
            show = view[["spent_on", "category", "amount", "description"]].copy()
            show = show.sort_values("spent_on", ascending=False)
            show["amount"] = show["amount"].apply(money)
            show.columns = ["Date", "Category", "Amount", "Description"]
            st.dataframe(
                show, use_container_width=True, hide_index=True,
                column_config={"Description": st.column_config.TextColumn(width="large")},
            )
            st.caption(f"{len(view)} expense(s) with a description.")

# ======================================================================
# Manage (full CRUD)
# ======================================================================
with tab_manage:
    st.subheader(f"Manage expenses - {label}")
    if df.empty:
        st.info("Nothing to manage for this period yet.")
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
                "description": st.column_config.TextColumn("Description", width="large"),
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
# Categories (custom)
# ======================================================================
with tab_cats:
    st.subheader("Your categories")
    st.caption("Built-in categories are always available. Add your own below, "
               "each in a group so it shows correctly in the donut chart.")

    st.markdown("**Add a new category**")
    with st.form("add_cat_form", clear_on_submit=True):
        cc1, cc2 = st.columns([2, 1])
        with cc1:
            new_name = st.text_input("Category name", placeholder="e.g. Movie nights, InDrive (other), Gym")
        with cc2:
            new_grp = st.selectbox("Group", GROUP_OPTIONS, index=GROUP_OPTIONS.index("Entertainment"))
        if st.form_submit_button("Add category", type="primary"):
            name = (new_name or "").strip()
            if not name:
                st.error("Please type a category name.")
            elif name in ALL_CATEGORIES:
                st.error("That category already exists.")
            else:
                add_category(name, new_grp)
                st.success(f"Added '{name}' in group {new_grp}.")
                st.rerun()

    st.divider()
    st.markdown("**Built-in categories**")
    builtin_df = pd.DataFrame(
        [(c, BUILTIN_GROUPS[c]) for c in BUILTIN_CATEGORIES], columns=["Category", "Group"]
    )
    st.dataframe(builtin_df, use_container_width=True, hide_index=True)

    st.markdown("**Your custom categories**")
    if not custom_cats:
        st.info("None yet. Add one above.")
    else:
        for name, grp in custom_cats:
            coln, colg, cold = st.columns([3, 2, 1])
            coln.write(name)
            colg.write(grp or "Other")
            if cold.button("Delete", key=f"delcat_{name}"):
                delete_category(name)
                st.success(f"Deleted '{name}'. Existing expenses keep their category.")
                st.rerun()
        st.caption("Deleting a category does not delete past expenses logged under it.")

# ======================================================================
# Export
# ======================================================================
with tab_export:
    st.subheader("Download your data")
    if df.empty:
        st.info("No data for this period to export yet.")
    else:
        st.download_button(
            "⬇️ Download Excel (formatted)",
            data=build_excel(df, label, salary, budget, total_spent, ALL_CATEGORIES),
            file_name=f"expenses_{period_key}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
        st.download_button(
            "⬇️ Download CSV (plain)",
            data=df[["spent_on", "category", "amount", "description"]].to_csv(index=False).encode("utf-8"),
            file_name=f"expenses_{period_key}.csv",
            mime="text/csv", use_container_width=True,
        )
        st.caption("Excel has two sheets: a formatted Expenses list with a total, "
                   "and a Summary with salary, budget, and category breakdown.")

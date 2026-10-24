# Expense Tracker: Setup and Deploy Guide

This is a Streamlit app. You will get it running on your laptop first, then put it online for free so you can open it from your phone too.

There are two parts:
- **Part A: Run it on your laptop** (5 minutes, no database needed)
- **Part B: Put it online + add a permanent database** (20 minutes)

---

## Part A: Run it on your laptop

1. Make sure you have Python installed. Open a terminal and type `python --version`. If you see a version number like 3.11, you are good. If not, install it from python.org.

2. Open a terminal and go into the project folder:
   ```
   cd expense-tracker
   ```

3. Install the requirements (one time only):
   ```
   pip install -r requirements.txt
   ```

4. Start the app:
   ```
   streamlit run app.py
   ```

5. Your browser opens automatically at `http://localhost:8501`. That is your tracker. Add a few expenses to try it.

At this stage your data is saved in a file called `expenses.db` on your laptop. That is fine for testing. For the online version you will use a proper database so nothing gets lost.

---

## Part B: Put it online for free

You need 3 free accounts: GitHub (to hold the code), Supabase (the database), and Streamlit Cloud (the hosting). Do them in this order.

### Step 1: Create the free database on Supabase

1. Go to supabase.com and click **Start your project**. Sign up (Google sign-in is fine).
2. Click **New project**. Give it any name like `expense-tracker`. Set a database password and **save that password somewhere**, you will need it in a moment.
3. Pick the region closest to you (Singapore or Mumbai are closest to Pakistan). Click **Create new project** and wait about 2 minutes.
4. On the left sidebar click the **gear icon (Project Settings)**, then click **Database**.
5. Scroll to **Connection string**. Click the **URI** tab. You will see a line that starts with `postgresql://`.
6. Copy that whole line. It has the word `[YOUR-PASSWORD]` inside it. Replace that part with the database password you saved in step 2.

   It should end up looking like:
   ```
   postgresql://postgres.abcxyz:MyRealPassword@aws-0-ap-south-1.pooler.supabase.com:5432/postgres
   ```
   Keep this line handy for Step 3.

### Step 2: Put the code on GitHub

1. Go to github.com and sign up or log in.
2. Click the **+** at the top right, then **New repository**.
3. Name it `expense-tracker`, keep it **Private**, and click **Create repository**.
4. The easiest way to upload without commands: on the new repo page click **uploading an existing file**, then drag in all the files from your `expense-tracker` folder (app.py, requirements.txt, the `.streamlit` folder, etc). Click **Commit changes**.

   Note: the `.gitignore` file makes sure your local database and secrets never get uploaded. That is intentional and safe.

### Step 3: Deploy on Streamlit Community Cloud

1. Go to share.streamlit.io and click **Sign in with GitHub**. Allow access.
2. Click **Create app**, then **Deploy a public app from GitHub** (your repo can still be private).
3. Fill in:
   - **Repository**: pick your `expense-tracker` repo
   - **Branch**: `main`
   - **Main file path**: `app.py`
4. Before clicking Deploy, click **Advanced settings**. In the **Secrets** box, paste this single line (use your real connection string from Step 1):
   ```
   DATABASE_URL = "postgresql://postgres.abcxyz:MyRealPassword@aws-0-ap-south-1.pooler.supabase.com:5432/postgres"
   ```
5. Click **Save**, then **Deploy**. Wait 2 to 3 minutes.
6. You get a public link like `https://your-app-name.streamlit.app`. Open it on your phone and bookmark it to your home screen.

That is it. Your expenses now save to Supabase and will never disappear, and you can open the tracker from any device with that link.

---

## How to use it day to day

- **First time**: open the sidebar (arrow at top left), type your salary, your monthly budget, and your salary date, then press **Save setup**.
- **Every day**: pick a category (InDrive, Lunch, Dinner, or one of the monthly ones, or Miscellaneous), type the amount, press **Add expense**. For Miscellaneous, add a short description.
- **Charts** update automatically: spending by category, daily spending, and how much budget is left.
- Use the **Year/Month** selectors at the top to look back at an earlier month.

---

## Common questions

**Do I have to pay anything?** No. Supabase and Streamlit Cloud both have free tiers that are plenty for one person.

**My app "goes to sleep" on Streamlit.** Free apps sleep after not being used for a while. Just open the link and it wakes up in a few seconds. Your data is safe in Supabase the whole time.

**I did not set up Supabase, is that ok?** The app still runs, but on Streamlit Cloud it would lose data when the app restarts. For a tracker you will use for months, do Step 1. It is worth the 5 minutes.

**I want to change the categories later.** Open `app.py` and edit the lists near the top called `DAILY_CATEGORIES`, `MONTHLY_CATEGORIES`. Re-upload to GitHub and Streamlit redeploys automatically.

# Faculty Workload Suite — Web App
## Deploy in 10 minutes, free forever

---

## What you get
A permanent web link (e.g. `https://faculty-workload.streamlit.app`)
that you open in any browser — no Python, no installation, no desktop app needed.

---

## Step 1 — Create a free GitHub account
Go to https://github.com and sign up (free).

---

## Step 2 — Create a new repository
1. Click the **＋** button → **New repository**
2. Name it: `faculty-workload-suite`
3. Set it to **Public**
4. Click **Create repository**

---

## Step 3 — Upload the app files
In your new repository, click **Add file → Upload files** and upload ALL of these:

```
app.py
requirements.txt
pages/
  __init__.py
  cleaner.py
  workload.py
  audit.py
```

> Tip: You can drag the entire `FacultyApp` folder into the upload window.

Click **Commit changes**.

---

## Step 4 — Deploy on Streamlit Cloud
1. Go to https://streamlit.io/cloud and sign up with your GitHub account (free)
2. Click **New app**
3. Select your repository: `faculty-workload-suite`
4. Main file path: `app.py`
5. Click **Deploy!**

⏳ Wait ~2 minutes while it builds.

---

## Step 5 — Share your link
You'll get a permanent URL like:
```
https://your-username-faculty-workload-suite-app-xxxxx.streamlit.app
```
Bookmark it. Share it with colleagues. Open it from any device.

---

## Updating the app later
If you want to change anything:
1. Edit the files on GitHub (click any file → pencil icon → edit → commit)
2. The app re-deploys automatically in ~1 minute

---

## The three tools

| Tab | What it does |
|-----|-------------|
| 🧹 Faculty Cleaner | Upload course file → removes duplicates → adds live SUMIF total |
| 📊 Workload Report | Upload student data → generates block-formula workload report |
| 🔗 Audit Merge | Upload Audit + Book5 → merges supervision WL by Faculty ID |

---

## Troubleshooting

**App shows error on startup**
→ Check that all 4 files (app.py + pages/*.py) are uploaded correctly

**"Module not found" error**
→ Check requirements.txt is present and contains: streamlit, pandas, openpyxl

**File upload fails**
→ Streamlit free tier has a 200MB file upload limit (more than enough for Excel files)

**App goes to sleep**
→ Free tier sleeps after ~7 days of no use. Just open the URL and it wakes in ~30 seconds.
   Upgrade to "Community Cloud" (still free) to keep it always awake.

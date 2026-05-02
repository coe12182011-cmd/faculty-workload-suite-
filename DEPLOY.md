# Faculty Workload Suite — Deploy Guide
## Get a permanent web link in ~10 minutes, completely free

---

## STEP 1 — Create a free GitHub account
Go to https://github.com and sign up.

---

## STEP 2 — Create a new repository
1. Click **＋ → New repository**
2. Repository name: `faculty-workload-suite`
3. Visibility: **Public**
4. Click **Create repository**

---

## STEP 3 — Upload the 2 files
In your new repository, click **Add file → Upload files**

Upload BOTH of these files (they are in this zip):
```
app.py
requirements.txt
```

Click **Commit changes**.

> ✅ That's all you need — just 2 files!

---

## STEP 4 — Deploy on Streamlit Cloud
1. Go to https://streamlit.io/cloud
2. Sign in with your GitHub account (free)
3. Click **New app**
4. Fill in:
   - Repository: `your-username/faculty-workload-suite`
   - Branch: `main`
   - Main file path: `app.py`
5. Click **Deploy!**

⏳ Wait about 2 minutes while it installs and starts.

---

## STEP 5 — Your permanent link is ready!
You'll get a URL like:
```
https://your-username-faculty-workload-suite-app.streamlit.app
```
**Bookmark it.** Open it from any browser, any device, anywhere.
No Python. No installation. No desktop app needed.

---

## Updating the app
If you want to change anything:
1. Go to your repository on GitHub
2. Click the file → pencil icon (edit) → make your change → **Commit changes**
3. The app automatically re-deploys in ~1 minute

---

## Notes
- **Free tier** sleeps after 7 days of no use → wakes in ~30 seconds when you open the link
- **File upload limit**: 200MB per file (more than enough for Excel files)
- **No data is stored** — files are processed in memory and immediately discarded

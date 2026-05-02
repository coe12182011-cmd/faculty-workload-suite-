"""
Faculty Workload Suite — Streamlit Web Application
===================================================
Three tools in one web app:
  1. 🧹 Faculty Cleaner   — remove duplicate course rows + live SUMIF total
  2. 📊 Workload Report   — Capstone / Internship / Thesis block-formula reports
  3. 🔗 Audit Merge       — merge supervision WL into COE Payment Audit file

Deploy free on https://streamlit.io/cloud
"""

import streamlit as st

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Faculty Workload Suite",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── SIDEBAR NAV ───────────────────────────────────────────────────────────────
st.sidebar.image(
    "https://img.icons8.com/color/96/graduation-cap.png", width=60
)
st.sidebar.title("Faculty Workload Suite")
st.sidebar.caption("Abu Dhabi University — COE")
st.sidebar.markdown("---")

page = st.sidebar.radio(
    "Choose a tool:",
    options=[
        "🧹  Faculty Cleaner",
        "📊  Workload Report",
        "🔗  Audit Merge",
    ],
)

st.sidebar.markdown("---")
st.sidebar.markdown(
    """
**How to use:**
1. Select a tool above
2. Upload your Excel file(s)
3. Click the action button
4. Download the result
"""
)
st.sidebar.markdown("---")
st.sidebar.caption("Built with Streamlit · Free hosting on streamlit.io/cloud")

# ── ROUTE TO PAGES ────────────────────────────────────────────────────────────
if   page == "🧹  Faculty Cleaner":
    from pages import cleaner;  cleaner.show()
elif page == "📊  Workload Report":
    from pages import workload; workload.show()
elif page == "🔗  Audit Merge":
    from pages import audit;    audit.show()

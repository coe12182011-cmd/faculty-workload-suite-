"""
Faculty Workload Suite — Single-file Streamlit Web App
=======================================================
Three tools:
  🧹 Faculty Cleaner   — remove duplicate course rows + live SUMIF total
  📊 Workload Report   — Capstone / Internship / Thesis block-formula reports
  🔗 Audit Merge       — merge supervision WL into COE Payment Audit file

Deploy free: https://streamlit.io/cloud
No Python installation needed — runs entirely in the browser.
"""

import io, math
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═════════════════════════════════════════════════════════════════════════════
# SHARED CONSTANTS & HELPERS
# ═════════════════════════════════════════════════════════════════════════════

NAVY="1B3A6B"; TEAL="0E7C7B"; WHITE="FFFFFF"; LIGHT="EEF2FA"
LIGHT2="F5F8FF"; GOLD="F5A623"; GRAY="95A5A6"; DARK="2C3E50"
RED="E74C3C"; ORANGE="E67E22"; GREEN="27AE60"
TOTAL_BG="0A3D62"; SUP_HDR="0A5C52"; SUP_VAL="D4F0EA"

TERM_LABELS = {
    "2501":"Fall 2501","2502":"Winter 2502",
    "2503":"Spring 2503","2504":"Summer 2504",
}
TERM_PALETTE = ["1B3A6B","0E7C7B","7B3A1B","2E5944","5C4033","3D6B9E"]

TYPE_PRIORITY = {
    "LEC":9,"MIX":8,"LAB":7,"STU":6,"VC":5,
    "FLD":4,"PRA":3,"RSC":2,"THE":1,"NON":0,
}
TERM_COL  = {"2501":"1B3A6B","2502":"0E7C7B","2503":"7B3A1B","2504":"2E5944"}
ROLE_COL  = {"Primary Instructor":"1B3A6B","Secondary Instructor":"0E7C7B",
              "Teaching Assistant":"5C4033"}
TYPE_COL  = {"LEC":"2E4057","MIX":"0E7C7B","LAB":"5C4033","STU":"6B3A7B",
              "VC":"3A6B1B","NON":GRAY,"RSC":"7B5A1B","FLD":"1B6B5C","THE":"4A1B6B"}
COL_W = {
    "ID":10,"Name":26,"Term":6,"Session":8,"Campus":8,"Subject":8,
    "Course Career":9,"Descr":8,"Course ID":10,"Section":7,"Class Nbr":10,
    "Facil ID":10,"Mtg Start":9,"Mtg End":9,"Start Date":12,"End Date":12,
    "Cap Enrl":7,"Tot Enrl":7,"Max Units":8,"Pat":6,"Assign Type":9,
    "Role":18,"Assignment %":11,"Long Title":36,"Description":22,"Email":28,
    "Total Assignment %":14,"Faculty Name":28,"Faculty ID":11,
    "Capstone\nWL":10,"UG Internship\nWL":11,"Thesis\nWL":9,
    "Total\nSupervision WL":12,"Supervisor":32,
}

def H(ws, r, c, v, bg=NAVY, fg=WHITE, bold=True, sz=9, ha="center", wrap=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font      = Font(name="Arial", bold=bold, color=fg, size=sz)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
    return cell

def D(ws, r, c, v, bg=WHITE, bold=False, sz=9, ha="center", fmt=None, fg="111111"):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font      = Font(name="Arial", bold=bold, color=fg, size=sz)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=ha, vertical="center")
    if fmt and isinstance(v, (int, float)):
        cell.number_format = fmt
    return cell

def total_colour(v):
    if   v >= 30: return "E74C3C", WHITE
    elif v >= 18: return "E67E22", WHITE
    elif v >=  9: return "27AE60", WHITE
    else:         return "DDEEFF", TOTAL_BG


# ═════════════════════════════════════════════════════════════════════════════
# TOOL 1 — FACULTY CLEANER
# ═════════════════════════════════════════════════════════════════════════════

def cleaner_process(df_raw):
    """Remove duplicate rows, keep best per (Name, Class Nbr, Section)."""
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df["Name"].notna()].copy()
    df["Name"]         = df["Name"].astype(str).str.strip()
    df["Assignment %"] = pd.to_numeric(df["Assignment %"], errors="coerce").fillna(0)
    df["Class Nbr"]    = pd.to_numeric(df["Class Nbr"],    errors="coerce")
    df["Section"]      = pd.to_numeric(df["Section"],      errors="coerce")
    df["_pri"] = df["Assign Type"].apply(
        lambda t: TYPE_PRIORITY.get(str(t).strip().upper(), 0))
    s = df.sort_values(
        ["Name","Class Nbr","Section","Assignment %","_pri"],
        ascending=[True,True,True,False,False])
    clean   = s.drop_duplicates(subset=["Name","Class Nbr","Section"],
                                keep="first").drop(columns=["_pri"])
    removed = df.loc[list(set(df.index) - set(clean.index))].drop(
        columns=["_pri"]).sort_values(["Name","Class Nbr","Section"])
    sort_c  = [c for c in ["Name","Term","Class Nbr"] if c in clean.columns]
    clean   = clean.sort_values(sort_c).reset_index(drop=True)
    return clean, removed


def cleaner_excel(df_clean, df_removed, source=""):
    orig  = list(df_clean.columns)
    cols  = orig + ["Total Assignment %"]
    nc    = len(cols)
    DS    = 6
    MR    = DS + len(df_clean) - 1
    NL    = get_column_letter(orig.index("Name") + 1)
    AL    = get_column_letter(orig.index("Assignment %") + 1)
    TL    = get_column_letter(nc)

    def sf(r):
        return (f"=SUMIF(${NL}${DS}:${NL}${MR},"
                f"{NL}{r},${AL}${DS}:${AL}${MR})")
    def sf2(fr):
        return (f"=SUMIF('Cleaned Data'!${NL}${DS}:"
                f"'Cleaned Data'!${NL}${MR},"
                f"'Cleaned Data'!{NL}{fr},"
                f"'Cleaned Data'!${AL}${DS}:"
                f"'Cleaned Data'!${AL}${MR})")

    totals = df_clean.groupby("Name")["Assignment %"].sum().to_dict()
    wb     = Workbook()

    # ── Sheet 1: Cleaned Data ─────────────────────────────────────────────
    ws = wb.active; ws.title = "Cleaned Data"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 6

    ws.merge_cells(f"A2:{TL}2")
    c = ws["A2"]
    c.value = (f"FACULTY COURSE LIST — CLEANED"
               + (f"  |  {source}" if source else "")
               + f"  ({len(df_clean):,} records | {df_clean['Name'].nunique()} faculty)")
    c.font  = Font(name="Arial", bold=True, size=13, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 30

    ws.merge_cells(f"A3:{TL}3")
    c = ws["A3"]
    c.value = ("  Duplicates removed — kept highest Assignment % per "
               "(Faculty · Class Nbr · Section)  |  Sorted A→Z by Faculty Name  |  "
               "⚡ 'Total Assignment %' = SUMIF — auto-updates when rows are deleted")
    c.font  = Font(name="Arial", size=9, italic=True, color="CCDDEE")
    c.fill  = PatternFill("solid", fgColor=TEAL)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 6

    for j, col in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = COL_W.get(col, 12)
        is_t = (col == "Total Assignment %")
        c = ws.cell(row=5, column=j, value=col)
        c.font      = Font(name="Arial", bold=True, size=9, color=WHITE)
        c.fill      = PatternFill("solid", fgColor=TOTAL_BG if is_t else NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = Border(bottom=Side(style="medium",
                                         color="00CCFF" if is_t else GOLD))
    ws.row_dimensions[5].height = 28

    prev = None; fi = -1; bgs = [LIGHT, LIGHT2]; nfr = {}
    for i, (_, row) in enumerate(df_clean.iterrows()):
        r = i + DS; nm = row["Name"]
        is_new = (nm != prev)
        if is_new: fi += 1; prev = nm; nfr[nm] = r
        bg = bgs[fi % 2]

        for j, col in enumerate(cols, 1):
            cell = ws.cell(row=r, column=j)
            is_t = (col == "Total Assignment %")

            if is_t:
                cell.value = sf(r); cell.number_format = "0.00"
                tbg, tfg = total_colour(totals.get(nm, 0))
                cell.font  = Font(name="Arial", bold=True, size=10, color=tfg)
                cell.fill  = PatternFill("solid", fgColor=tbg)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(
                    left=Side(style="medium",  color=TOTAL_BG),
                    right=Side(style="medium", color=TOTAL_BG),
                    top=Side(style="thin",
                             color="9AAAC4" if is_new else "CCDDFF"))
                continue

            val = row[col]
            if col == "Name":
                cell.value = str(val) if pd.notna(val) else ""
                cell.font  = Font(name="Arial", size=9, bold=is_new, color=DARK)
                cell.fill  = PatternFill("solid",
                                          fgColor="DCE8FF" if is_new else bg)
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col == "Role":
                v = str(val).strip() if pd.notna(val) else ""
                cell.value = v
                cell.font  = Font(name="Arial", size=8, bold=True, color=WHITE)
                cell.fill  = PatternFill("solid",
                                          fgColor=ROLE_COL.get(v, DARK))
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == "Term":
                v = str(val).strip() if pd.notna(val) else ""
                cell.value = v
                cell.font  = Font(name="Arial", size=9, bold=True, color=WHITE)
                cell.fill  = PatternFill("solid",
                                          fgColor=TERM_COL.get(v, NAVY))
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == "Assign Type":
                v = str(val).strip() if pd.notna(val) else ""
                cell.value = v
                cell.font  = Font(name="Arial", size=9, bold=True, color=WHITE)
                cell.fill  = PatternFill("solid",
                                          fgColor=TYPE_COL.get(v, DARK))
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == "Assignment %":
                v = pd.to_numeric(val, errors="coerce")
                cell.value = float(v) if pd.notna(v) else 0
                cell.number_format = "0.0"
                cell.font  = Font(name="Arial", size=9,
                                  bold=(float(v)>0 if pd.notna(v) else False),
                                  color=NAVY if (pd.notna(v) and float(v)>0) else GRAY)
                cell.fill  = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col in ("Start Date", "End Date"):
                cell.value = pd.to_datetime(val) if pd.notna(val) else None
                if pd.notna(val): cell.number_format = "DD-MMM-YYYY"
                cell.font  = Font(name="Arial", size=8)
                cell.fill  = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col in ("Cap Enrl","Tot Enrl","Class Nbr","Course ID","ID","Section"):
                v = pd.to_numeric(val, errors="coerce")
                cell.value = int(v) if pd.notna(v) else None
                cell.font  = Font(name="Arial", size=9, color=DARK)
                cell.fill  = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col in ("Long Title", "Description", "Email"):
                cell.value = str(val).strip() if pd.notna(val) else ""
                cell.font  = Font(name="Arial", size=9,
                                  italic=(col=="Email"),
                                  color="1155CC" if col=="Email" else DARK)
                cell.fill  = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.value = str(val).strip() if pd.notna(val) else None
                cell.font  = Font(name="Arial", size=9, color=DARK)
                cell.fill  = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if is_new:
                cell.border = Border(top=Side(style="thin", color="9AAAC4"))
        ws.row_dimensions[r].height = 15
    ws.freeze_panes = "A6"

    # ── Sheet 2: Faculty Summary ──────────────────────────────────────────
    ws2 = wb.create_sheet("Faculty Summary")
    ws2.sheet_view.showGridLines = False
    ws2.row_dimensions[1].height = 6
    ws2.merge_cells("A2:G2")
    c = ws2["A2"]
    c.value = (f"FACULTY SUMMARY — {df_clean['Name'].nunique()} faculty  |  "
               f"{len(df_clean):,} sections")
    c.font  = Font(name="Arial", bold=True, size=13, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[2].height = 28

    ws2.merge_cells("A3:G3")
    c = ws2["A3"]
    c.value = ("  ⚡ 'Total Assign %' uses cross-sheet SUMIF — "
               "deleting rows in Cleaned Data auto-updates totals (Ctrl+Alt+F9)")
    c.font  = Font(name="Arial", size=9, italic=True, color="555555")
    c.fill  = PatternFill("solid", fgColor="DDEEFF")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[3].height = 16
    ws2.row_dimensions[4].height = 8

    fsum = (df_clean.groupby("Name").agg(
        Sections=("Class Nbr","count"),
        Terms   =("Term",   lambda x:", ".join(sorted(x.dropna().astype(str).unique()))),
        Courses =("Descr",  "nunique"),
        Campuses=("Campus", lambda x:", ".join(sorted(x.dropna().astype(str).unique()))),
        Roles   =("Role",   lambda x:", ".join(sorted(x.dropna().astype(str).unique()))),
    ).reset_index().sort_values("Name").reset_index(drop=True))

    sh = ["#","Faculty Name","Sections","Total Assign %","Terms","Campus(es)","Role(s)"]
    sw = [4,32,10,14,18,16,30]
    for j,(h,w) in enumerate(zip(sh,sw),1):
        ws2.column_dimensions[get_column_letter(j)].width = w
        is_t = (h == "Total Assign %")
        c = ws2.cell(row=5, column=j, value=h)
        c.font      = Font(name="Arial", bold=True, size=9, color=WHITE)
        c.fill      = PatternFill("solid", fgColor=TOTAL_BG if is_t else NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = Border(bottom=Side(style="medium",
                                         color="00CCFF" if is_t else GOLD))
    ws2.row_dimensions[5].height = 22

    for i, row in fsum.iterrows():
        r = i + 6; bg = LIGHT if r % 2 == 0 else LIGHT2; nm = row["Name"]
        c = ws2.cell(row=r,column=1,value=i+1)
        c.font=Font(name="Arial",size=9,color=GRAY)
        c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="center",vertical="center")
        c = ws2.cell(row=r,column=2,value=nm)
        c.font=Font(name="Arial",size=9,bold=True,color=DARK)
        c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="left",vertical="center")
        c = ws2.cell(row=r,column=3,value=row["Sections"])
        c.font=Font(name="Arial",size=9); c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="center",vertical="center")
        tbg, tfg = total_colour(totals.get(nm, 0))
        fr = nfr.get(nm, DS)
        c = ws2.cell(row=r,column=4,value=sf2(fr))
        c.number_format="0.00"
        c.font=Font(name="Arial",size=10,bold=True,color=tfg)
        c.fill=PatternFill("solid",fgColor=tbg)
        c.alignment=Alignment(horizontal="center",vertical="center")
        for j, col in enumerate(["Terms","Campuses","Roles"], 5):
            c = ws2.cell(row=r,column=j,value=row[col])
            c.font=Font(name="Arial",size=9,color=DARK)
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Alignment(horizontal="left",vertical="center")
        ws2.row_dimensions[r].height = 15

    tr = len(fsum) + 6
    ws2.merge_cells(f"A{tr}:C{tr}")
    c = ws2.cell(row=tr,column=1,value="GRAND TOTAL")
    c.font=Font(name="Arial",bold=True,size=11,color=WHITE)
    c.fill=PatternFill("solid",fgColor=NAVY)
    c.alignment=Alignment(horizontal="center",vertical="center")
    c = ws2.cell(row=tr,column=4,value=f"=SUM(D6:D{tr-1})")
    c.number_format="0.00"
    c.font=Font(name="Arial",bold=True,size=11,color=WHITE)
    c.fill=PatternFill("solid",fgColor=NAVY)
    c.alignment=Alignment(horizontal="center",vertical="center")
    for j in [5,6,7]:
        ws2.cell(row=tr,column=j).fill=PatternFill("solid",fgColor=NAVY)
    ws2.row_dimensions[tr].height = 22
    ws2.freeze_panes = "A6"

    # ── Sheet 3: Removed Rows Log ─────────────────────────────────────────
    ws3 = wb.create_sheet("Removed Rows Log")
    ws3.sheet_view.showGridLines = False
    ws3.row_dimensions[1].height = 6
    lc = list(df_removed.columns); nl = len(lc)
    ws3.merge_cells(f"A2:{get_column_letter(nl)}2")
    c = ws3["A2"]
    c.value = f"REMOVED ROWS LOG — {len(df_removed)} duplicate rows removed"
    c.font  = Font(name="Arial",bold=True,size=12,color=WHITE)
    c.fill  = PatternFill("solid",fgColor="E74C3C")
    c.alignment = Alignment(horizontal="center",vertical="center")
    ws3.row_dimensions[2].height = 26
    ws3.merge_cells(f"A3:{get_column_letter(nl)}3")
    c = ws3["A3"]
    c.value = ("  Removed because a higher-priority row exists for the same "
               "Faculty + Class Nbr + Section.  Rule: highest Assignment % → LEC > MIX > LAB > NON")
    c.font  = Font(name="Arial",size=9,italic=True,color="555555")
    c.fill  = PatternFill("solid",fgColor="FEE8E8")
    c.alignment = Alignment(horizontal="left",vertical="center")
    ws3.row_dimensions[3].height = 16; ws3.row_dimensions[4].height = 8
    for j, col in enumerate(lc, 1):
        ws3.column_dimensions[get_column_letter(j)].width = COL_W.get(col, 12)
        c = ws3.cell(row=5,column=j,value=col)
        c.font=Font(name="Arial",bold=True,size=9,color=WHITE)
        c.fill=PatternFill("solid",fgColor="C0392B")
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws3.row_dimensions[5].height = 22
    for i,(_, row) in enumerate(df_removed.iterrows()):
        r = i+6; bg = "FFF5F5" if i%2==0 else WHITE
        for j, col in enumerate(lc, 1):
            val = row[col]
            c = ws3.cell(row=r,column=j,
                         value=str(val).strip() if pd.notna(val) else None)
            c.font=Font(name="Arial",size=9,color="888888",italic=True)
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Alignment(horizontal="center",vertical="center")
        ws3.row_dimensions[r].height = 14
    ws3.freeze_panes = "A6"

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf


def show_cleaner():
    st.title("🧹 Faculty Course Cleaner")
    st.markdown(
        "Upload a faculty course file to remove duplicate rows and add a "
        "**live SUMIF** `Total Assignment %` column that auto-recalculates "
        "when you delete rows in Excel."
    )

    with st.expander("📋 Cleaning rules & colour key"):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
**Duplicate rule** — for each `Faculty Name + Class Nbr + Section` group:
1. Keep the row with the **highest Assignment %**
2. If tied → best type: `LEC > MIX > LAB > STU > VC > NON`
3. NON/0% rows are almost always the ones dropped

**Output sheets:**
- **Cleaned Data** — sorted A→Z by faculty, SUMIF total column
- **Faculty Summary** — one row per faculty, colour-coded
- **Removed Rows Log** — full audit trail
            """)
        with col2:
            st.markdown("**Total Assignment % colours:**")
            st.markdown("""
| Colour | Range |
|--------|-------|
| 🔵 Light blue | < 9% |
| 🟢 Green | 9 – 17% |
| 🟠 Orange | 18 – 29% |
| 🔴 Red | ≥ 30% |
            """)

    uploaded = st.file_uploader(
        "Upload faculty course file (.xlsx or .xls)",
        type=["xlsx","xls"],
        help="Must have columns: Name, Class Nbr, Section, Assignment %",
    )

    if uploaded is None:
        st.info("👆 Upload a file to get started.")
        return

    try:
        df_raw = pd.read_excel(uploaded, header=0)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
    except Exception as e:
        st.error(f"Could not read file: {e}"); return

    required = {"Name","Class Nbr","Section","Assignment %"}
    missing  = required - set(df_raw.columns)
    if missing:
        st.error(f"Missing columns: {missing}")
        st.write("Columns found:", list(df_raw.columns)); return

    df_raw = df_raw[df_raw["Name"].notna()].copy()
    c1,c2,c3 = st.columns(3)
    c1.metric("Total rows",       f"{len(df_raw):,}")
    c2.metric("Unique faculty",    f"{df_raw['Name'].astype(str).str.strip().nunique()}")
    c3.metric("Unique sections",
              f"{df_raw.groupby(['Name','Class Nbr','Section']).ngroups:,}")

    with st.expander("👀 Preview raw data (first 20 rows)"):
        st.dataframe(df_raw.head(20), use_container_width=True)

    st.markdown("---")
    if st.button("🧹  Clean & Generate Excel", type="primary",
                 use_container_width=True):
        with st.spinner("Cleaning data and building Excel…"):
            try:
                df_clean, df_removed = cleaner_process(df_raw)
                buf = cleaner_excel(df_clean, df_removed, source=uploaded.name)
            except Exception as e:
                st.error(f"Error: {e}")
                import traceback; st.code(traceback.format_exc()); return

        st.success("✅  Done!")
        r1,r2,r3,r4 = st.columns(4)
        r1.metric("Rows in",    f"{len(df_raw):,}")
        r2.metric("Rows out",   f"{len(df_clean):,}")
        r3.metric("Removed",    f"{len(df_removed)}",
                  delta=f"-{len(df_removed)}", delta_color="inverse")
        r4.metric("Faculty",    f"{df_clean['Name'].nunique()}")

        totals = (df_clean.groupby("Name")["Assignment %"]
                  .sum().reset_index()
                  .rename(columns={"Assignment %":"Total Assign %"})
                  .sort_values("Total Assign %", ascending=False)
                  .reset_index(drop=True))
        totals["Total Assign %"] = totals["Total Assign %"].round(2)

        with st.expander("📊 Faculty Summary (top 20)"):
            st.dataframe(totals.head(20), use_container_width=True)

        from datetime import datetime
        st.download_button(
            label="⬇️  Download Cleaned Excel",
            data=buf,
            file_name=f"Faculty_Cleaned_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        if len(df_removed) > 0:
            with st.expander(f"🗑️ View {len(df_removed)} removed rows"):
                st.dataframe(df_removed.reset_index(drop=True),
                             use_container_width=True)


# ═════════════════════════════════════════════════════════════════════════════
# TOOL 2 — WORKLOAD REPORT
# ═════════════════════════════════════════════════════════════════════════════

def wl_classify(row):
    title = str(row.get("Course Title","")).lower()
    code  = str(row.get("Course Code", "")).upper()
    if "thesis"     in title:                              return "Thesis"
    if "internship" in title or "399" in code or "398" in code: return "UG_Internship"
    return "Capstone"

def cap_wl(cr,n):
    n=min(n,15)
    if cr>=3: return math.ceil(n/5)*1.0 if n>=3 else (1/3)*n
    else:     return math.ceil(n/5)*(cr/3) if n>=3 else (cr/3/3)*n

def int_wl(cr,n):
    nm=min(n,18)
    if cr>=6:   return math.ceil(nm/6)*2.0 if nm>=6 else (2/6)*nm
    elif cr>=3: return math.ceil(nm/6)*1.0 if nm>=6 else (1/6)*nm
    else:       return (cr/3/6)*nm

def calc_wl(row):
    k,c,n = row["Key"],row["Course_Credits"],row["Student_Count"]
    if k=="Capstone":      return cap_wl(c,n)
    if k=="UG_Internship": return int_wl(c,n)
    if k=="Thesis":        return 0.5
    return 0.0

def wl_process_raw(df):
    df = df.copy(); df.columns=[str(c).strip() for c in df.columns]
    rename={}
    for col in df.columns:
        cl=col.lower()
        if "supervisor"  in cl:                   rename[col]="Supervisor"
        elif cl in ("id","student id"):            rename[col]="ID"
        elif "credit"    in cl:                   rename[col]="Course_Credits"
        elif "class" in cl and "nbr" in cl:        rename[col]="Class Nbr"
        elif "course" in cl and "title" in cl:     rename[col]="Course Title"
        elif "course" in cl and "code"  in cl:     rename[col]="Course Code"
        elif "enrollment" in cl and "term" in cl:  rename[col]="Enrollment Term"
        elif cl == "section":                      rename[col]="Section"
    df=df.rename(columns=rename)
    needed=["Supervisor","Course Code","Course Title","Course_Credits",
            "Class Nbr","Section","Enrollment Term"]
    missing=[c for c in needed if c not in df.columns]
    if missing: raise ValueError(f"Missing columns: {missing}")
    df=df[df["Supervisor"].notna()].copy()
    df["Supervisor"]      = df["Supervisor"].astype(str).str.strip()
    df["Course_Credits"]  = pd.to_numeric(df["Course_Credits"],  errors="coerce")
    df["Class Nbr"]       = pd.to_numeric(df["Class Nbr"],       errors="coerce")
    df["Enrollment Term"] = pd.to_numeric(df["Enrollment Term"], errors="coerce")
    df=df[df["Course_Credits"].notna()].copy()
    df=df[~df["Course Code"].isin(["Course Code","0"])].copy()
    df["Key"]=df.apply(wl_classify,axis=1)
    sec=df.groupby(["Supervisor","Key","Course Code","Course Title",
                    "Course_Credits","Class Nbr","Section","Enrollment Term"]
                  ).agg(Student_Count=("ID","count")).reset_index()
    sec["Term"]=sec["Enrollment Term"].apply(
        lambda x: TERM_LABELS.get(str(x).split(".")[0], f"Term {str(x).split('.')[0]}"))
    sec["Workload"]=sec.apply(calc_wl,axis=1)
    return sec

def wl_build_excel(df_all):
    all_terms=sorted(df_all["Term"].unique(),
                     key=lambda t: t.split()[-1] if t.split() else t)
    tc={t: TERM_PALETTE[i%len(TERM_PALETTE)] for i,t in enumerate(all_terms)}
    df_all=df_all.copy()
    df_all["_to"]=df_all["Term"].apply(
        lambda t: all_terms.index(t) if t in all_terms else 99)

    summary=df_all.groupby(["Supervisor","Key"])["Workload"].sum().unstack(fill_value=0).reset_index()
    for k in ["Capstone","UG_Internship","Thesis"]:
        if k not in summary.columns: summary[k]=0
    summary["Total"]=summary["Capstone"]+summary["UG_Internship"]+summary["Thesis"]
    summary=summary.sort_values("Total",ascending=False).reset_index(drop=True)
    summary.columns=["Supervisor","Capstone","UG_Internship","Thesis","Total"]
    tp=df_all.groupby(["Supervisor","Term"])["Workload"].sum().unstack(fill_value=0).reset_index()
    for t in all_terms:
        if t not in tp.columns: tp[t]=0
    summary=summary.merge(tp[["Supervisor"]+all_terms],on="Supervisor",how="left")
    for t in all_terms: summary[t]=summary[t].fillna(0)
    sup_sec=df_all.groupby("Supervisor")["Class Nbr"].nunique().to_dict()

    wb=Workbook(); ws1=wb.active; ws1.title="Dashboard"
    ws1.sheet_view.showGridLines=False; ws1.row_dimensions[1].height=6
    ndc=7+len(all_terms)
    ws1.merge_cells(f"A2:{get_column_letter(ndc)}2")
    H(ws1,2,1,f"FACULTY WORKLOAD CALCULATOR — {' + '.join(all_terms)}",sz=13)
    ws1.row_dimensions[2].height=36
    ws1.merge_cells(f"A3:{get_column_letter(ndc)}3")
    H(ws1,3,1,"Capstone: CEIL(n/5) blocks (max 15)  |  Internship: CEIL(n/6) blocks (max 18)  |  Thesis: 0.5 per thesis",bg=TEAL,sz=10)
    ws1.row_dimensions[3].height=20; ws1.row_dimensions[4].height=8

    kpis=[("A","Supervisors",int(len(summary)),NAVY),
          ("B","Total Students",int(df_all["Student_Count"].sum()),TEAL),
          ("C","Capstone WL",round(float(summary["Capstone"].sum()),2),"2E4057"),
          ("D","Internship WL",round(float(summary["UG_Internship"].sum()),2),"048A81"),
          ("E","Thesis WL",round(float(summary["Thesis"].sum()),2),"5C4033"),
          ("F","Grand Total",round(float(summary["Total"].sum()),2),NAVY)]
    for cl,lbl,val,bg in kpis:
        c=ws1[f"{cl}5"]; c.value=lbl
        c.font=Font(name="Arial",size=8,color="AACCEE"); c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="center",vertical="center")
        c=ws1[f"{cl}6"]; c.value=val
        c.font=Font(name="Arial",size=12,bold=True,color=WHITE); c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="center",vertical="center")
    ws1.row_dimensions[5].height=16; ws1.row_dimensions[6].height=26; ws1.row_dimensions[7].height=8

    hdrs=["#","Supervisor","Capstone\nWL","Intern.\nWL","Thesis\nWL","Total\nWL"]
    wdths=[4,42,12,13,11,12]
    for t in all_terms:
        sh=t.replace("Fall ","F").replace("Winter ","W").replace("Spring ","S")
        hdrs.append(sh+"\nWL"); wdths.append(11)
    hdrs+=["Sections","Category"]; wdths+=[9,14]
    for j,(h,w) in enumerate(zip(hdrs,wdths),1):
        ws1.column_dimensions[get_column_letter(j)].width=w
        idx=j-7; hbg=tc.get(all_terms[idx],NAVY) if 6<j<=6+len(all_terms) else NAVY
        c=ws1.cell(row=8,column=j,value=h)
        c.font=Font(name="Arial",bold=True,size=9,color=WHITE)
        c.fill=PatternFill("solid",fgColor=hbg)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border=Border(bottom=Side(style="medium",color=GOLD))
    ws1.row_dimensions[8].height=32

    tls=["E8EEFF","E8F8F5","FEF0E8","F0FFF0"]
    for i,row in summary.iterrows():
        r=i+9; bg=LIGHT if i%2==0 else "FFFFFF"; total=round(float(row["Total"]),4)
        D(ws1,r,1,i+1,bg=bg,ha="center",sz=9,fg="999999")
        D(ws1,r,2,row["Supervisor"],bg=bg,bold=(i<5),ha="left",sz=10)
        for j,col in enumerate(["Capstone","UG_Internship","Thesis"],3):
            v=round(float(row[col]),4)
            D(ws1,r,j,v if v>0 else "-",bg=bg,fmt="0.000" if v>0 else None)
        tbg,tfg=(RED,WHITE) if total>=5 else (ORANGE,WHITE) if total>=3 else (bg,NAVY)
        c=ws1.cell(row=r,column=6,value=total); c.number_format="0.000"
        c.font=Font(name="Arial",size=10,bold=True,color=tfg)
        c.fill=PatternFill("solid",fgColor=tbg)
        c.alignment=Alignment(horizontal="center",vertical="center")
        for j,term in enumerate(all_terms,7):
            v=round(float(row[term]),4); tl=tls[list(all_terms).index(term)%4]
            c=ws1.cell(row=r,column=j,value=v if v>0 else "-")
            if v>0: c.number_format="0.000"
            c.font=Font(name="Arial",size=9,color=NAVY if v>0 else GRAY)
            c.fill=PatternFill("solid",fgColor=tl if bg==LIGHT else "FFFFFF")
            c.alignment=Alignment(horizontal="center",vertical="center")
        D(ws1,r,7+len(all_terms),int(sup_sec.get(row["Supervisor"],0)),bg=bg)
        cat,cbg=("High Load",RED) if total>=5 else ("Moderate",ORANGE) if total>=3 else ("Normal",GREEN) if total>=1 else ("Low",GRAY)
        c=ws1.cell(row=r,column=8+len(all_terms),value=cat)
        c.font=Font(name="Arial",size=9,bold=True,color=WHITE)
        c.fill=PatternFill("solid",fgColor=cbg)
        c.alignment=Alignment(horizontal="center",vertical="center")
        ws1.row_dimensions[r].height=17

    tr=len(summary)+9
    ws1.merge_cells(f"A{tr}:B{tr}")
    c=ws1.cell(row=tr,column=1,value="TOTAL")
    c.font=Font(name="Arial",bold=True,size=11,color=WHITE); c.fill=PatternFill("solid",fgColor=NAVY)
    c.alignment=Alignment(horizontal="center",vertical="center")
    for j,col in enumerate(["Capstone","UG_Internship","Thesis","Total"],3):
        c=ws1.cell(row=tr,column=j,value=round(float(summary[col].sum()),4))
        c.number_format="0.000"; c.font=Font(name="Arial",bold=True,size=11,color=WHITE)
        c.fill=PatternFill("solid",fgColor=NAVY); c.alignment=Alignment(horizontal="center",vertical="center")
    for j,term in enumerate(all_terms,7):
        c=ws1.cell(row=tr,column=j,value=round(float(summary[term].sum()),4))
        c.number_format="0.000"; c.font=Font(name="Arial",bold=True,size=11,color=WHITE)
        c.fill=PatternFill("solid",fgColor=NAVY); c.alignment=Alignment(horizontal="center",vertical="center")
    ws1.row_dimensions[tr].height=22; ws1.freeze_panes="A9"

    ws2=wb.create_sheet("Section Detail"); ws2.sheet_view.showGridLines=False
    ws2.row_dimensions[1].height=6
    ws2.merge_cells("A2:J2")
    H(ws2,2,1,f"SECTION DETAIL — {' + '.join(all_terms)}",bg=TEAL,sz=12)
    ws2.row_dimensions[2].height=26; ws2.row_dimensions[3].height=6
    KCL={"Capstone":NAVY,"UG_Internship":TEAL,"Thesis":"5C4033"}
    for j,(h,w) in enumerate(zip(["Supervisor","Term","Type","Course Code",
                                    "Course Title","Credits","Class Nbr","Sec",
                                    "Students","Workload"],
                                   [34,12,15,12,34,7,10,6,10,11]),1):
        ws2.column_dimensions[get_column_letter(j)].width=w
        c=ws2.cell(row=4,column=j,value=h)
        c.font=Font(name="Arial",bold=True,size=9,color=WHITE); c.fill=PatternFill("solid",fgColor=NAVY)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws2.row_dimensions[4].height=26
    for i,row in df_all.sort_values(["Supervisor","_to","Key","Course Code"]).reset_index(drop=True).iterrows():
        r=i+5; bg=LIGHT if i%2==0 else "FFFFFF"; k=str(row.get("Key",""))
        D(ws2,r,1,row["Supervisor"],bg=bg,ha="left",sz=9)
        c=ws2.cell(row=r,column=2,value=row["Term"])
        c.font=Font(name="Arial",size=8,bold=True,color=WHITE)
        c.fill=PatternFill("solid",fgColor=tc.get(row["Term"],NAVY))
        c.alignment=Alignment(horizontal="center",vertical="center")
        c=ws2.cell(row=r,column=3,value=k)
        c.font=Font(name="Arial",size=8,bold=True,color=WHITE)
        c.fill=PatternFill("solid",fgColor=KCL.get(k,NAVY))
        c.alignment=Alignment(horizontal="center",vertical="center")
        D(ws2,r,4,row["Course Code"],bg=bg,sz=9)
        D(ws2,r,5,row["Course Title"],bg=bg,sz=9,ha="left")
        D(ws2,r,6,row["Course_Credits"],bg=bg,sz=9,fmt="0.0")
        D(ws2,r,7,row["Class Nbr"],bg=bg,sz=9)
        D(ws2,r,8,row["Section"],bg=bg,sz=9)
        D(ws2,r,9,int(row["Student_Count"]),bg=bg,sz=10,bold=True)
        c=ws2.cell(row=r,column=10,value=round(float(row["Workload"]),4))
        c.number_format="0.000"; c.font=Font(name="Arial",bold=True,size=10)
        c.fill=PatternFill("solid",fgColor=bg); c.alignment=Alignment(horizontal="center",vertical="center")
        ws2.row_dimensions[r].height=15
    ws2.freeze_panes="A5"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, {"supervisors":len(summary),"sections":len(df_all),
                 "students":int(df_all["Student_Count"].sum()),
                 "grand_total":round(float(summary["Total"].sum()),3),
                 "terms":all_terms}


def show_workload():
    st.title("📊 Workload Report")
    st.markdown(
        "Upload student data files to generate a full "
        "**Capstone / Internship / Thesis** workload report using the block-based formula."
    )

    with st.expander("📋 Formula reference"):
        st.markdown("""
| Type | Scenario | Condition | Formula |
|------|----------|-----------|---------|
| Capstone | A.1 | credits≥3, n≥3 | `CEIL(n/5) × 1.0` |
| Capstone | A.2 | credits≥3, n<3 | `(1/3) × n` |
| Capstone | B.1 | credits<3, n≥3 | `CEIL(n/5) × (credits/3)` |
| Capstone | B.2 | credits<3, n<3 | `(credits/3/3) × n` |
| Internship | Scen-A | 3cr, n<6 | `(1/6) × n` |
| Internship | Blk-A | 3cr, n≥6 | `CEIL(n/6) × 1.0` |
| Internship | Scen-B | ≥6cr, n<6 | `(2/6) × n` |
| Internship | Blk-B | ≥6cr, n≥6 | `CEIL(n/6) × 2.0` |
| Thesis | — | any | `0.5 per thesis` |
        """)

    files = st.file_uploader(
        "Upload raw student data file(s)",
        type=["xlsx","xls"], accept_multiple_files=True,
        help="Must have: Supervisor, Course Code, Course_Credits, Class Nbr, Section, Enrollment Term, ID"
    )

    if not files:
        st.info("👆 Upload one or more files. You can combine multiple terms."); return

    st.write(f"**{len(files)} file(s):** " + ", ".join(f.name for f in files))

    if st.button("📊  Generate Workload Report", type="primary", use_container_width=True):
        all_secs=[]; errors=[]
        prog=st.progress(0)
        for i,f in enumerate(files):
            with st.spinner(f"Processing {f.name}…"):
                try:
                    df_raw=pd.read_excel(f, header=0)
                    sec=wl_process_raw(df_raw)
                    fname=f.name.lower()
                    if sec["Term"].eq(sec["Term"].iloc[0]).all():
                        if   "spring" in fname: sec["Term"]="Spring 2503"
                        elif "winter" in fname: sec["Term"]="Winter 2502"
                        elif "fall"   in fname: sec["Term"]="Fall 2501"
                    all_secs.append(sec)
                except Exception as e:
                    errors.append(f"{f.name}: {e}")
            prog.progress((i+1)/len(files))
        if errors:
            for err in errors: st.error(err)
            if not all_secs: return

        df_all=pd.concat(all_secs,ignore_index=True)
        df_all=df_all[df_all["Workload"].notna()].copy()
        with st.spinner("Building Excel report…"):
            buf,stats=wl_build_excel(df_all)

        st.success("✅  Done!")
        c1,c2,c3,c4=st.columns(4)
        c1.metric("Supervisors",stats["supervisors"])
        c2.metric("Sections",   stats["sections"])
        c3.metric("Students",   stats["students"])
        c4.metric("Grand Total WL",stats["grand_total"])

        from datetime import datetime
        st.download_button(
            label="⬇️  Download Workload Report",
            data=buf,
            file_name=f"Workload_Report_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)


# ═════════════════════════════════════════════════════════════════════════════
# TOOL 3 — AUDIT MERGE
# ═════════════════════════════════════════════════════════════════════════════

def audit_load_sup(file):
    df=pd.read_excel(file, header=0)
    df.columns=[str(c).replace("\n"," ").strip() for c in df.columns]
    df=df[df["Name"].notna()].copy()
    df["Name"]       =df["Name"].astype(str).str.strip()
    df["Faculty ID"] =pd.to_numeric(df["Faculty ID"],errors="coerce")
    df["Workload Assignment"]=pd.to_numeric(df["Workload Assignment"],errors="coerce").fillna(0)
    tc=next((c for c in df.columns if "type" in c.lower() and
             any(k in c for k in ["LEC","LAB","INT","Type"])),None)
    if tc is None: raise ValueError("Cannot find Class Type column.")
    agg=df.groupby(["Faculty ID",tc])["Workload Assignment"].sum().unstack(fill_value=0).reset_index()
    for col in ["Capstone","UG_Internship","Thesis"]:
        if col not in agg.columns: agg[col]=0
    agg["Total Supervision WL"]=agg["Capstone"]+agg["UG_Internship"]+agg["Thesis"]
    for col in ["Capstone","UG_Internship","Thesis","Total Supervision WL"]:
        agg[col]=agg[col].round(4)
    return agg.set_index("Faculty ID")

def audit_load_audit(file, sheet_name=None):
    xl=pd.ExcelFile(file)
    if sheet_name and sheet_name in xl.sheet_names:
        target=sheet_name
    else:
        pref=[s for s in xl.sheet_names if "payment" in s.lower() and "(2)" in s]
        if not pref: pref=[s for s in xl.sheet_names if "payment" in s.lower()]
        target=pref[0] if pref else xl.sheet_names[0]
    raw=pd.read_excel(file,sheet_name=target,header=None)
    hr=None
    for i,row in raw.iterrows():
        vals=[str(v).replace("\n"," ").strip() for v in row.values if pd.notna(v)]
        if "Faculty Name" in vals: hr=i; break
    if hr is None: raise ValueError(f"Cannot find 'Faculty Name' header in '{target}'")
    df=pd.read_excel(file,sheet_name=target,header=hr)
    df.columns=[str(c).replace("\n"," ").strip() for c in df.columns]
    df=df[df["Faculty Name"].notna()].copy()
    df["Faculty ID"]=pd.to_numeric(df["Faculty ID"],errors="coerce")
    return df,target

def audit_merge(audit_df, sup_df):
    seen=set(); cv,iv,tv,totv,mv=[],[],[],[],[]
    for _,row in audit_df.iterrows():
        fid=row.get("Faculty ID")
        if pd.notna(fid) and int(fid) in sup_df.index and int(fid) not in seen:
            s=sup_df.loc[int(fid)]
            cv.append(round(float(s["Capstone"]),4))
            iv.append(round(float(s["UG_Internship"]),4))
            tv.append(round(float(s["Thesis"]),4))
            totv.append(round(float(s["Total Supervision WL"]),4))
            mv.append(True); seen.add(int(fid))
        else:
            cv.append(None); iv.append(None)
            tv.append(None); totv.append(None); mv.append(False)
    df=audit_df.copy()
    df["_Cap"]=cv; df["_Int"]=iv; df["_Th"]=tv; df["_Tot"]=totv; df["_m"]=mv
    return df

def audit_build_excel(merged, source=""):
    orig=[c for c in merged.columns if not c.startswith("_")]
    fni=next((i for i,c in enumerate(orig) if c=="Faculty Name"),0)
    SUP=[("Capstone\nWL","_Cap"),("UG Internship\nWL","_Int"),
         ("Thesis\nWL","_Th"),("Total\nSupervision WL","_Tot")]
    col_order=([(c,None) for c in orig[:fni+1]]
               +[(lbl,key) for lbl,key in SUP]
               +[(c,None) for c in orig[fni+1:]])
    n=len(col_order)

    wb=Workbook(); ws=wb.active; ws.title="Merged Audit"
    ws.sheet_view.showGridLines=False; ws.row_dimensions[1].height=6
    ws.merge_cells(f"A2:{get_column_letter(n)}2")
    c=ws["A2"]
    c.value="COE Faculty Workload — Audit + Supervision Merged"+(f"  ({source})" if source else "")
    c.font=Font(name="Arial",bold=True,size=12,color=WHITE)
    c.fill=PatternFill("solid",fgColor=NAVY)
    c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[2].height=28
    ws.merge_cells(f"A3:{get_column_letter(n)}3")
    c=ws["A3"]
    c.value="  Blue-green columns = supervision workload  |  Values on first row of each faculty only"
    c.font=Font(name="Arial",size=9,italic=True,color="CCDDEE")
    c.fill=PatternFill("solid",fgColor=TEAL)
    c.alignment=Alignment(horizontal="left",vertical="center")
    ws.row_dimensions[3].height=16; ws.row_dimensions[4].height=6

    for j,(lbl,key) in enumerate(col_order,1):
        ws.column_dimensions[get_column_letter(j)].width=COL_W.get(lbl,10)
        is_s=(key is not None)
        c=ws.cell(row=5,column=j,value=lbl.replace("\\n","\n"))
        c.font=Font(name="Arial",bold=True,color=WHITE,size=8)
        c.fill=PatternFill("solid",fgColor=SUP_HDR if is_s else NAVY)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border=Border(bottom=Side(style="medium",color="00E5CC" if is_s else GOLD))
    ws.row_dimensions[5].height=28

    prev=None
    for _,row in merged.iterrows():
        fid=row.get("Faculty ID"); is_f=(fid!=prev); prev=fid
        r=ws.max_row+1; bg=LIGHT if r%2==0 else "FFFFFF"
        for j,(lbl,key) in enumerate(col_order,1):
            cell=ws.cell(row=r,column=j); is_s=(key is not None)
            if is_s:
                val=row.get(key) if is_f else None
                if val is not None and isinstance(val,float):
                    cell.value=round(val,4); cell.number_format="0.000"
                    cell.font=Font(name="Arial",bold=True,color=DARK,size=9)
                    cell.fill=PatternFill("solid",fgColor=SUP_VAL)
                else:
                    cell.fill=PatternFill("solid",fgColor="F5FFFE")
                cell.alignment=Alignment(horizontal="center",vertical="center")
            else:
                raw=row.get(lbl)
                cell.value=str(raw).strip() if pd.notna(raw) else None
                cell.font=Font(name="Arial",size=9,
                               bold=(is_f and lbl=="Faculty Name"))
                cell.fill=PatternFill("solid",
                                      fgColor="E8F0FF" if is_f and lbl=="Faculty Name" else bg)
                cell.alignment=Alignment(horizontal="left" if lbl=="Faculty Name" else "center",
                                         vertical="center")
        ws.row_dimensions[r].height=15
    ws.freeze_panes="A6"

    ws2=wb.create_sheet("Supervision Summary")
    ws2.sheet_view.showGridLines=False; ws2.row_dimensions[1].height=6
    ws2.merge_cells("A2:G2")
    H(ws2,2,1,"SUPERVISION WORKLOAD SUMMARY",sz=12); ws2.row_dimensions[2].height=24
    for j,(h,w) in enumerate(zip(["#","Faculty Name","Faculty ID",
                                    "Capstone WL","UG Internship WL",
                                    "Thesis WL","Total Supervision WL"],
                                   [4,36,12,14,16,12,18]),1):
        ws2.column_dimensions[get_column_letter(j)].width=w
        c=ws2.cell(row=3,column=j,value=h)
        c.font=Font(name="Arial",bold=True,size=9,color=WHITE)
        c.fill=PatternFill("solid",fgColor=SUP_HDR)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws2.row_dimensions[3].height=22
    sub=merged[merged["_m"]==True].drop_duplicates("Faculty ID").sort_values("_Tot",ascending=False)
    for i,row in sub.reset_index(drop=True).iterrows():
        r=i+4; bg=LIGHT if i%2==0 else "FFFFFF"
        for j,v in enumerate([i+1,row.get("Faculty Name",""),
                               int(row["Faculty ID"]) if pd.notna(row["Faculty ID"]) else None,
                               row.get("_Cap"),row.get("_Int"),row.get("_Th"),row.get("_Tot")],1):
            c=ws2.cell(row=r,column=j,value=v)
            c.font=Font(name="Arial",size=9,bold=(j==7),color=DARK)
            c.fill=PatternFill("solid",fgColor=SUP_VAL if j==7 else bg)
            c.alignment=Alignment(horizontal="left" if j==2 else "center",vertical="center")
            if j in (4,5,6,7) and isinstance(v,float): c.number_format="0.000"
        ws2.row_dimensions[r].height=15

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

def show_audit():
    st.title("🔗 Audit Merge")
    st.markdown(
        "Merge **Capstone / Internship / Thesis** supervision workload "
        "into the COE FT/PT Payment Audit file. Matched by **Faculty ID**."
    )

    col1,col2=st.columns(2)
    with col1:
        audit_f=st.file_uploader("1️⃣  Audit File (COE FT/PT Payment Form)",
                                  type=["xlsx","xls"],key="audit_f")
    with col2:
        sup_f=st.file_uploader("2️⃣  Supervision Workload File (Book5 format)",
                                type=["xlsx","xls"],key="sup_f")

    sheet=st.text_input("Audit sheet name (leave blank to auto-detect)",
                         value="FT-PT Payment AY2026 (2)")

    if not audit_f or not sup_f:
        st.info("👆 Upload both files above to continue."); return

    if st.button("🔗  Merge & Generate Report",type="primary",use_container_width=True):
        with st.spinner("Loading and merging…"):
            try:
                sup=audit_load_sup(sup_f)
                audit,sheet_used=audit_load_audit(audit_f,sheet or None)
                merged=audit_merge(audit,sup)
                buf=audit_build_excel(merged,source=audit_f.name)
            except Exception as e:
                st.error(f"Error: {e}")
                import traceback; st.code(traceback.format_exc()); return

        matched=int(merged["_m"].sum()); total_f=merged["Faculty ID"].nunique()
        st.success("✅  Done!")
        c1,c2,c3=st.columns(3)
        c1.metric("Total audit rows",    f"{len(merged):,}")
        c2.metric("Unique faculty",       f"{total_f}")
        c3.metric("Faculty with sup. WL", f"{matched}",
                  delta=f"{matched}/{total_f} matched")
        st.caption(f"Sheet used: **{sheet_used}**")

        from datetime import datetime
        st.download_button(
            label="⬇️  Download Merged Audit Excel",
            data=buf,
            file_name=f"Audit_Merged_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)


# ═════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG & NAVIGATION
# ═════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Faculty Workload Suite",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.sidebar.title("🎓 Faculty Workload Suite")
st.sidebar.caption("COE — Abu Dhabi University")
st.sidebar.markdown("---")

page = st.sidebar.radio(
    "Choose a tool:",
    ["🧹  Faculty Cleaner", "📊  Workload Report", "🔗  Audit Merge"],
)

st.sidebar.markdown("---")
st.sidebar.markdown("""
**How to use:**
1. Pick a tool from the list above
2. Upload your Excel file(s)
3. Click the action button
4. Download your result instantly
""")
st.sidebar.markdown("---")
st.sidebar.caption("Built with [Streamlit](https://streamlit.io) · Free hosting available")

# Route
if   page == "🧹  Faculty Cleaner":  show_cleaner()
elif page == "📊  Workload Report":  show_workload()
elif page == "🔗  Audit Merge":      show_audit()

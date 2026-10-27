"""ACI Premio Plastics | SO Work Productivity Dashboard
Read-only executive dashboard. Place workbook at data/SO_Activity_Attendance.xlsx.
"""
from pathlib import Path
from datetime import datetime, time
import re
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="SO Productivity | ACI Premio", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
BASE = Path(__file__).resolve().parent
DATA_FILE = BASE / "data" / "SO_Activity_Attendance.xlsx"
OUTLET_DAILY_TARGET = 20
MEMO_DAILY_TARGET = 8
WORK_HOURS_TARGET = 11
LATE_CUTOFF_MIN = 8 * 60 + 45

st.markdown("""
<style>
:root{--navy:#0b1224;--panel:#111c34;--muted:#9fb0ca;--ink:#eef4ff;--cyan:#43d9ff;--violet:#9b8cff}
.stApp{background:radial-gradient(ellipse at 15% 0%,#18264a 0%,#0b1224 44%,#080d19 100%);color:var(--ink)}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#101a31,#080d19);border-right:1px solid #253454}
section[data-testid="stSidebar"] *{color:#edf3ff}
.block-container{padding-top:1.2rem;padding-bottom:2rem;max-width:1600px}
.hero{padding:25px 28px;border:1px solid #39476d;border-radius:22px;background:linear-gradient(120deg,rgba(38,62,112,.92),rgba(80,50,135,.82),rgba(12,119,147,.5));box-shadow:0 16px 40px #0005;margin-bottom:18px}
.hero h1{font-size:clamp(1.45rem,3vw,2.5rem);margin:0;color:white;letter-spacing:.2px}.hero p{color:#d6e4ff;margin:.45rem 0 0}
.kpi{padding:17px 18px;border-radius:18px;background:linear-gradient(145deg,#1b2b4b,#11192e);border:1px solid #34486c;box-shadow:0 10px 24px #0004;min-height:116px}
.kpi .label{font-size:.82rem;color:#b7c7e2;font-weight:650}.kpi .value{font-size:clamp(1.35rem,2.2vw,2rem);font-weight:800;color:#fff;margin-top:7px}.kpi .hint{font-size:.76rem;color:#8ceaff;margin-top:3px}
.section{font-size:1.28rem;font-weight:800;margin:1rem 0 .35rem;color:#f2f6ff}

div[data-testid="stTabs"] button{font-weight:800;color:#f4f7ff!important}
div[data-testid="stTabs"] button[aria-selected="true"]{background:#ffd43b!important;color:#172033!important;border-radius:9px}
div[data-testid="stTabs"] button[aria-selected="true"] *{color:#172033!important}
button[kind="secondary"],button[kind="primary"]{color:#172033}
#MainMenu{visibility:hidden}
footer{visibility:hidden}
div[data-testid="stToolbar"]{visibility:hidden;height:0;position:fixed}
div[data-testid="stDecoration"]{visibility:hidden}

div[data-testid="stDataFrame"]{border:1px solid #273858;border-radius:12px}
.stCaption{color:#aebbd2}
@media(max-width:700px){.block-container{padding-left:.7rem;padding-right:.7rem}.hero{padding:18px}.kpi{padding:13px;min-height:100px}}
</style>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def load_data(path, modified):
    xls = pd.ExcelFile(path)
    required = {"Summary Report", "SO Activity", "Attendance "}
    if not required.issubset(set(xls.sheet_names)):
        raise ValueError(f"Workbook sheets mismatch. Expected: {sorted(required)}; found: {xls.sheet_names}")
    summary = pd.read_excel(path, sheet_name="Summary Report")
    activity = pd.read_excel(path, sheet_name="SO Activity")
    attendance = pd.read_excel(path, sheet_name="Attendance ")
    for c in ["Total Outlet visitTarget","Total Number Of Outlet Visit","Total Shop Memo Target","Number Of Ordered Memo","Attendance ( Count of 21 Days)","Present Days"]:
        summary[c] = pd.to_numeric(summary[c], errors="coerce").fillna(0)
    summary["StaffKey"] = summary["Staff ID"].map(id_key)
    activity["StaffKey"] = activity["Staff ID"].map(id_key)
    attendance["StaffKey"] = attendance["Staff Id"].map(id_key)
    activity["Visit Date"] = pd.to_datetime(activity["Visit Date"], errors="coerce")
    attendance["Attendance Date"] = pd.to_datetime(attendance["Attendance Date"], errors="coerce")
    for c in ["Number Of Outlet Visit", "Number Of Ordered Shop"]:
        activity[c] = pd.to_numeric(activity[c], errors="coerce").fillna(0)
    attendance["InMinutes"] = attendance["In Time Only"].map(time_minutes)
    attendance["HoursMinutes"] = attendance.get("Total Hours", pd.Series(index=attendance.index,dtype=object)).map(hours_minutes)
    return summary, activity, attendance

def id_key(v):
    if pd.isna(v): return ""
    s = str(v).strip()
    if s.endswith(".0"): s = s[:-2]
    return s.upper()

def time_minutes(v):
    if pd.isna(v): return np.nan
    if isinstance(v, time): return v.hour*60 + v.minute + v.second/60
    try:
        t = pd.to_datetime(str(v), errors="coerce")
        return t.hour*60+t.minute+t.second/60 if pd.notna(t) else np.nan
    except Exception: return np.nan

def hours_minutes(v):
    if pd.isna(v): return np.nan
    m = re.search(r"(\d+)\s*h\s*(\d+)\s*m", str(v), re.I)
    return int(m.group(1))*60+int(m.group(2)) if m else np.nan

def pct(a,b): return (a/b*100) if b else 0

def kpi(label, value, hint=""):
    st.markdown(f'<div class="kpi"><div class="label">{label}</div><div class="value">{value}</div><div class="hint">{hint}</div></div>', unsafe_allow_html=True)

def chart_layout(fig, height=360):
    fig.update_layout(height=height, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font_color="#e8efff", margin=dict(l=10,r=10,t=55,b=10), legend_title_text="", title_font_size=16)
    fig.update_xaxes(gridcolor="rgba(160,180,220,.13)", zerolinecolor="rgba(160,180,220,.15)")
    fig.update_yaxes(gridcolor="rgba(160,180,220,.13)", zerolinecolor="rgba(160,180,220,.15)")
    return fig

def show_chart(fig, height=360): st.plotly_chart(chart_layout(fig,height), use_container_width=True, config={"displayModeBar":False,"responsive":True})

if not DATA_FILE.exists():
    st.markdown('<div class="hero"><h1>ACI Premio Plastics | SO Productivity Intelligence</h1><p>Executive read-only dashboard</p></div>', unsafe_allow_html=True)
    st.error("Excel data file পাওয়া যায়নি। GitHub repository-তে `data/SO_Activity_Attendance.xlsx` নামে workbook রাখুন।")
    st.stop()
try:
    summary, activity_all, attendance_all = load_data(str(DATA_FILE), DATA_FILE.stat().st_mtime)
except Exception as e:
    st.error(f"Excel load করতে সমস্যা হয়েছে: {e}")
    st.stop()

st.markdown('<div class="hero"><h1>ACI Premio Plastics</h1><p>SO Work Productivity & Attendance Intelligence Dashboard · Director View</p></div>', unsafe_allow_html=True)
# Filters
st.sidebar.markdown("## Dashboard Filters")
st.sidebar.caption("Choose a reporting month and organizational scope.")
all_dates = pd.concat([activity_all["Visit Date"].dropna(), attendance_all["Attendance Date"].dropna()])
months = sorted(all_dates.dt.to_period("M").astype(str).unique().tolist())
if not months:
    st.error("Workbook-এ valid date পাওয়া যায়নি."); st.stop()
selected_month = st.sidebar.selectbox("Reporting month", months, index=len(months)-1)
sm_col = "SM Name" if "SM Name" in summary else None
rsm_col = "RSM Name" if "RSM Name" in summary else None
group_col = "Group" if "Group" in summary else None
for c in [group_col, sm_col, rsm_col]:
    if c:
        summary[c] = summary[c].astype("string").fillna("Unknown").str.strip()
sel_groups = st.sidebar.multiselect("Group", sorted(summary[group_col].unique()), default=sorted(summary[group_col].unique())) if group_col else []
sm_base = summary[summary[group_col].isin(sel_groups)] if group_col else summary
sel_sms = st.sidebar.multiselect("SM", sorted(sm_base[sm_col].unique()), default=sorted(sm_base[sm_col].unique())) if sm_col else []
rsm_base = sm_base[sm_base[sm_col].isin(sel_sms)] if sm_col else sm_base
sel_rsms = st.sidebar.multiselect("RSM", sorted(rsm_base[rsm_col].unique()), default=sorted(rsm_base[rsm_col].unique())) if rsm_col else []
st.sidebar.caption(f"Data file updated: {datetime.fromtimestamp(DATA_FILE.stat().st_mtime):%d %b %Y, %I:%M %p}")

period = pd.Period(selected_month, freq="M")
act = activity_all[activity_all["Visit Date"].dt.to_period("M").astype(str).eq(selected_month)].copy()
att = attendance_all[attendance_all["Attendance Date"].dt.to_period("M").astype(str).eq(selected_month)].copy()
# Use summary report for authoritative SO targets; map selected team scope to raw records.
mask = pd.Series(True, index=summary.index)
if group_col: mask &= summary[group_col].isin(sel_groups)
if sm_col: mask &= summary[sm_col].isin(sel_sms)
if rsm_col: mask &= summary[rsm_col].isin(sel_rsms)
rep = summary[mask].copy()
keys = set(rep["StaffKey"])
act = act[act["StaffKey"].isin(keys)].copy()
att = att[att["StaffKey"].isin(keys)].copy()

# Normalize targets and compute per-SO metrics.
rep["Visit Achievement %"] = np.where(rep["Total Outlet visitTarget"]>0, rep["Total Number Of Outlet Visit"]/rep["Total Outlet visitTarget"]*100,0)
rep["Memo Achievement %"] = np.where(rep["Total Shop Memo Target"]>0, rep["Number Of Ordered Memo"]/rep["Total Shop Memo Target"]*100,0)
rep["Present %"] = np.where(rep["Attendance ( Count of 21 Days)"]>0, rep["Present Days"]/rep["Attendance ( Count of 21 Days)"]*100,0)
rep["Avg Outlets / Active Day"] = np.where(rep["Attendance ( Count of 21 Days)"]>0, rep["Total Number Of Outlet Visit"]/rep["Attendance ( Count of 21 Days)"],0)
rep["Avg Memos / Active Day"] = np.where(rep["Attendance ( Count of 21 Days)"]>0, rep["Number Of Ordered Memo"]/rep["Attendance ( Count of 21 Days)"],0)
# Raw attendance classification: a record is present if Present or Late; punctuality uses actual in-time <= 08:45.
att["Status"] = att["Attendance Status"].astype(str).str.strip().str.lower()
att["OnTime"] = att["InMinutes"].notna() & (att["InMinutes"] < LATE_CUTOFF_MIN) & att["Status"].isin(["present","late"])
att["Exactly845"] = att["InMinutes"].notna() & (att["InMinutes"] == LATE_CUTOFF_MIN) & att["Status"].isin(["present","late"])
att["AtAfter845"] = att["InMinutes"].notna() & (att["InMinutes"] > LATE_CUTOFF_MIN) & att["Status"].isin(["present","late"])
att["PresentRecord"] = att["Status"].isin(["present","late"])
att["AbsentRecord"] = att["Status"].eq("absent")

# Overview
st.markdown('<div class="section">Executive Summary</div>', unsafe_allow_html=True)
visit_total = int(rep["Total Number Of Outlet Visit"].sum())
memo_total = int(rep["Number Of Ordered Memo"].sum())
visit_target = int(rep["Total Outlet visitTarget"].sum())
memo_target = int(rep["Total Shop Memo Target"].sum())
so_count = rep["Staff ID"].nunique()
att_pct = pct(rep["Present Days"].sum(), rep["Attendance ( Count of 21 Days)"].sum())
cols=st.columns(3)
with cols[0]: kpi("Sales Officers", f"{so_count:,}", f"Reporting month: {selected_month}")
with cols[1]: kpi("Outlet Visits", f"{visit_total:,}", f"{pct(visit_total,visit_target):.1f}% of target · target {visit_target:,}")
with cols[2]: kpi("Memo / Ordered Shops", f"{memo_total:,}", f"{pct(memo_total,memo_target):.1f}% of target · target {memo_target:,}")
cols=st.columns(3)
with cols[0]: kpi("Attendance Present", f"{att_pct:.1f}%", f"{int(rep['Present Days'].sum()):,} present out of {int(rep['Attendance ( Count of 21 Days)'].sum()):,} SO-days")
with cols[1]: kpi("Outlet Target Achievement", f"{pct(visit_total,visit_target):.1f}%", "20 outlets per active day benchmark")
with cols[2]: kpi("Memo Target Achievement", f"{pct(memo_total,memo_target):.1f}%", "8 memos per active day benchmark")

# Month-level daily trend charts
st.markdown('<div class="section">Daily Trend</div>', unsafe_allow_html=True)
c1,c2=st.columns(2)
with c1:
    if not act.empty:
        daily=act.assign(Day=act["Visit Date"].dt.day).groupby("Day",as_index=False).agg(Outlet_Visits=("Number Of Outlet Visit","sum"),Memos=("Number Of Ordered Shop","sum"))
        fig=px.line(daily,x="Day",y="Outlet_Visits",markers=True,title="Daily Outlet Visits",labels={"Day":"Day of month","Outlet_Visits":"Outlet visits"})
        show_chart(fig)
    else: st.info("Selected filters-এ activity data নেই।")
with c2:
    if not act.empty:
        daily=act.assign(Day=act["Visit Date"].dt.day).groupby("Day",as_index=False).agg(Memos=("Number Of Ordered Shop","sum"))
        fig=px.bar(daily,x="Day",y="Memos",title="Daily Memo / Ordered Shop",labels={"Day":"Day of month","Memos":"Memos"})
        show_chart(fig)
    else: st.info("Selected filters-এ activity data নেই।")

# Tabs
tabs=st.tabs(["Attendance & Punctuality","SO Productivity","SM / RSM / Group","Top & Bottom 10","Market & Repeat Visits","AI Improvement Summary","SO Detail"])
with tabs[0]:
    st.subheader("Attendance analysis")
    total_rows=len(att); present_rows=int(att["PresentRecord"].sum()); absent_rows=int(att["AbsentRecord"].sum())
    ontime=int(att["OnTime"].sum()); exactly=int(att["Exactly845"].sum()); after=int(att["AtAfter845"].sum())
    a,b,c,d=st.columns(4)
    with a:kpi("Present records",f"{present_rows:,}",f"{pct(present_rows,total_rows):.1f}% of attendance records")
    with b:kpi("Absent records",f"{absent_rows:,}",f"{pct(absent_rows,total_rows):.1f}% of attendance records")
    with c:kpi("Before 8:45 AM",f"{ontime:,}",f"{pct(ontime,ontime+after):.1f}% of timed present records")
    with d:kpi("At/After 8:45 AM",f"{after:,}",f"{pct(after,ontime+after):.1f}% of timed present records")
    if not att.empty:
        byday=att.assign(Day=att["Attendance Date"].dt.day).groupby("Day",as_index=False).agg(Present=("PresentRecord","sum"),Absent=("AbsentRecord","sum"),Before_845=("OnTime","sum"),Exactly_845=("Exactly845","sum"),After_845=("AtAfter845","sum"))
        timed=att[att["InMinutes"].notna() & att["Status"].isin(["present","late"])].groupby(att.loc[att["InMinutes"].notna() & att["Status"].isin(["present","late"]),"Attendance Date"].dt.day).size()
        byday["Timed_Entries"]=byday["Day"].map(timed).fillna(0)
        for key in ["Before_845","Exactly_845","After_845"]:
            byday[key+"_Pct"]=np.where(byday["Timed_Entries"]>0,byday[key]/byday["Timed_Entries"]*100,0)
        fig=px.bar(byday,x="Day",y=["Present","Absent"],barmode="group",title="Daily Present vs Absent",labels={"Day":"Day of month","value":"Records"})
        show_chart(fig)
        fig=px.bar(byday,x="Day",y=["Before_845","Exactly_845","After_845"],barmode="group",title="Day 1 onward · Entry count by 8:45 AM cutoff",labels={"Day":"Day of month","value":"SO entry count","variable":"Entry time"})
        show_chart(fig)
        pct_daily=byday.melt(id_vars=["Day"],value_vars=["Before_845_Pct","Exactly_845_Pct","After_845_Pct"],var_name="Entry time",value_name="Percentage")
        pct_daily["Entry time"]=pct_daily["Entry time"].map({"Before_845_Pct":"Before 8:45 AM","Exactly_845_Pct":"Exactly 8:45 AM","After_845_Pct":"After 8:45 AM"})
        fig=px.line(pct_daily,x="Day",y="Percentage",color="Entry time",markers=True,title="Day-wise entry percentage (each day = 100%)",labels={"Day":"Day of month","Percentage":"Share of timed present entries (%)"})
        fig.update_yaxes(range=[0,100],ticksuffix="%")
        show_chart(fig)
        st.markdown("#### Day-wise entry count and percentage")
        day_table=byday[["Day","Before_845","Before_845_Pct","Exactly_845","Exactly_845_Pct","After_845","After_845_Pct","Timed_Entries"]].copy()
        day_table=day_table.rename(columns={"Day":"Day of Month","Before_845":"Before 8:45 count","Before_845_Pct":"Before 8:45 %","Exactly_845":"Exactly 8:45 count","Exactly_845_Pct":"Exactly 8:45 %","After_845":"After 8:45 count","After_845_Pct":"After 8:45 %","Timed_Entries":"Timed entries"})
        for col in ["Before 8:45 %","Exactly 8:45 %","After 8:45 %"]: day_table[col]=day_table[col].round(1)
        st.dataframe(day_table,use_container_width=True,hide_index=True)
    # Per-SO punctuality over 21 attendance days; exact 21-day coverage is visible.
    att_so=att.groupby("StaffKey",as_index=False).agg(Attendance_Days=("Attendance Date","nunique"),Present_Days_Raw=("PresentRecord","sum"),Absent_Days_Raw=("AbsentRecord","sum"),Before_845_Days=("OnTime","sum"),At_After_845_Days=("AtAfter845","sum"))
    punctual=rep[["StaffKey","Sr Name","SM Name","RSM Name","Group"]].merge(att_so,on="StaffKey",how="left")
    for c in ["Attendance_Days","Present_Days_Raw","Absent_Days_Raw","Before_845_Days","At_After_845_Days"]: punctual[c]=punctual[c].fillna(0).astype(int)
    punctual["Punctuality % of 21"] = punctual["Before_845_Days"]/21*100
    punctual["At/After 8:45 % of 21"] = punctual["At_After_845_Days"]/21*100
    st.markdown("#### SO-wise 21-day attendance and punctuality")
    st.caption("21-day denominator is used as requested. If the workbook has fewer than 21 dates for an employee, actual recorded days are shown separately; percentages against 21 remain comparable.")
    st.dataframe(punctual.sort_values(["Before_845_Days","Present_Days_Raw"],ascending=False),use_container_width=True,hide_index=True)
    # Monthly attendance status percentages from raw data
    if not att.empty:
        status=att["Attendance Status"].astype(str).value_counts().rename_axis("Status").reset_index(name="Count")
        status["%"]=(status["Count"]/status["Count"].sum()*100).round(1)
        st.markdown("#### Attendance status mix"); st.dataframe(status,use_container_width=True,hide_index=True)

with tabs[1]:
    st.subheader("Outlet, memo and working-day productivity")
    p1,p2=st.columns(2)
    with p1:
        fig=px.bar(rep.sort_values("Visit Achievement %",ascending=False).head(20),x="Visit Achievement %",y="Sr Name",orientation="h",color="Visit Achievement %",title="Top 20 · Outlet target achievement",hover_data=["Total Number Of Outlet Visit","Total Outlet visitTarget"])
        fig.update_layout(yaxis={"categoryorder":"total ascending"}); show_chart(fig,520)
    with p2:
        fig=px.bar(rep.sort_values("Memo Achievement %",ascending=False).head(20),x="Memo Achievement %",y="Sr Name",orientation="h",color="Memo Achievement %",title="Top 20 · Memo target achievement",hover_data=["Number Of Ordered Memo","Total Shop Memo Target"])
        fig.update_layout(yaxis={"categoryorder":"total ascending"}); show_chart(fig,520)
    st.markdown("#### 11-hour workday benchmark")
    st.caption("Benchmark: 20 outlet visits/day, 8 memos/day, 11 working hours/day. Actual recorded working hours are only calculated if a Total Hours field exists in the Attendance sheet; this workbook currently has no Total Hours column, so hour achievement cannot be reliably calculated from this file.")
    work=rep[["Staff ID","Sr Name","Group","SM Name","RSM Name","Attendance ( Count of 21 Days)","Total Outlet visitTarget","Total Number Of Outlet Visit","Visit Achievement %","Total Shop Memo Target","Number Of Ordered Memo","Memo Achievement %"]].copy()
    work=work.rename(columns={"Attendance ( Count of 21 Days)":"Recorded Attendance Days","Total Outlet visitTarget":"Outlet Target","Total Number Of Outlet Visit":"Outlet Visits","Total Shop Memo Target":"Memo Target","Number Of Ordered Memo":"Memos"})
    work["Outlets / Attendance Day"] = np.where(work["Recorded Attendance Days"]>0,work["Outlet Visits"]/work["Recorded Attendance Days"],0)
    work["Memos / Attendance Day"] = np.where(work["Recorded Attendance Days"]>0,work["Memos"]/work["Recorded Attendance Days"],0)
    work["Estimated Target Work Hours"] = work["Recorded Attendance Days"]*WORK_HOURS_TARGET
    work["Target Outlets at 20/day"] = work["Recorded Attendance Days"]*OUTLET_DAILY_TARGET
    work["Target Memos at 8/day"] = work["Recorded Attendance Days"]*MEMO_DAILY_TARGET
    st.dataframe(work.sort_values("Visit Achievement %",ascending=False),use_container_width=True,hide_index=True)

with tabs[2]:
    st.subheader("Leadership comparison")
    for col,label in [("Group","Group-wise"),("SM Name","SM-wise"),("RSM Name","RSM-wise")]:
        if col not in rep: continue
        agg=rep.groupby(col,dropna=False).agg(SO_Count=("Staff ID","nunique"),Outlet_Visits=("Total Number Of Outlet Visit","sum"),Outlet_Target=("Total Outlet visitTarget","sum"),Memos=("Number Of Ordered Memo","sum"),Memo_Target=("Total Shop Memo Target","sum"),Present_Days=("Present Days","sum"),Scheduled_Days=("Attendance ( Count of 21 Days)","sum")).reset_index()
        agg["Outlet Achievement %"]=np.where(agg["Outlet_Target"]>0,agg["Outlet_Visits"]/agg["Outlet_Target"]*100,0)
        agg["Memo Achievement %"]=np.where(agg["Memo_Target"]>0,agg["Memos"]/agg["Memo_Target"]*100,0)
        agg["Present %"]=np.where(agg["Scheduled_Days"]>0,agg["Present_Days"]/agg["Scheduled_Days"]*100,0)
        st.markdown(f"#### {label}")
        c1,c2=st.columns(2)
        with c1: show_chart(px.bar(agg.sort_values("Outlet Achievement %",ascending=False),x=col,y="Outlet Achievement %",color="Outlet Achievement %",title=f"{label} outlet target achievement"))
        with c2: show_chart(px.bar(agg.sort_values("Memo Achievement %",ascending=False),x=col,y="Memo Achievement %",color="Memo Achievement %",title=f"{label} memo target achievement"))
        st.dataframe(agg.sort_values("Outlet Achievement %",ascending=False),use_container_width=True,hide_index=True)
    st.markdown("#### SM → RSM / SO overview")
    st.dataframe(rep.groupby(["Group","SM Name","RSM Name"],dropna=False).agg(SOs=("Staff ID","nunique"),Outlet_Visits=("Total Number Of Outlet Visit","sum"),Memo_Count=("Number Of Ordered Memo","sum"),Attendance_Present=("Present Days","sum")).reset_index(),use_container_width=True,hide_index=True)

with tabs[3]:
    st.subheader("SO leaderboard")
    metric=st.selectbox("Rank by",["Total Number Of Outlet Visit","Number Of Ordered Memo","Present Days","Visit Achievement %","Memo Achievement %","Present %"],index=0)
    display_cols=["Staff ID","Sr Name","Group","SM Name","RSM Name","Total Number Of Outlet Visit","Visit Achievement %","Number Of Ordered Memo","Memo Achievement %","Present Days","Present %"]
    top=rep.sort_values(metric,ascending=False).head(10)[display_cols].copy(); top.insert(0,"Rank",range(1,len(top)+1))
    bottom=rep.sort_values(metric,ascending=True).head(10)[display_cols].copy(); bottom.insert(0,"Rank",range(1,len(bottom)+1))
    c1,c2=st.columns(2)
    with c1:
        st.markdown("#### 🏆 Top 10")
        fig=px.bar(top.sort_values(metric),x=metric,y="Sr Name",orientation="h",text=metric,title=f"Top 10 SO · {metric}",hover_data=["Staff ID","Group","SM Name","RSM Name"])
        fig.update_layout(yaxis={"categoryorder":"total ascending"}); show_chart(fig,520)
        st.dataframe(top,use_container_width=True,hide_index=True)
    with c2:
        st.markdown("#### ⚠️ Bottom 10")
        fig=px.bar(bottom.sort_values(metric,ascending=False),x=metric,y="Sr Name",orientation="h",text=metric,title=f"Bottom 10 SO · {metric}",hover_data=["Staff ID","Group","SM Name","RSM Name"])
        fig.update_layout(yaxis={"categoryorder":"total descending"}); show_chart(fig,520)
        st.dataframe(bottom,use_container_width=True,hide_index=True)
    st.caption("Ranking is based on the selected metric. For bottom rankings, low values are shown first; interpret attendance and achievement together rather than using one metric alone.")

with tabs[4]:
    st.subheader("Market and repeat-visit analysis")
    if act.empty: st.info("Selected filters-এ activity data নেই।")
    else:
        market_col="Market Name" if "Market Name" in act else None
        if market_col:
            market=act.groupby(market_col,dropna=False).agg(Visit_Events=("Number Of Outlet Visit","sum"),Memo_Orders=("Number Of Ordered Shop","sum"),Visit_Days=("Visit Date","nunique"),SO_Count=("StaffKey","nunique")).reset_index().sort_values("Visit_Events",ascending=False)
            c1,c2=st.columns(2)
            with c1: show_chart(px.bar(market.head(15),x="Visit_Events",y=market_col,orientation="h",title="Most visited markets (top 15)"))
            with c2: show_chart(px.bar(market.sort_values("Memo_Orders",ascending=False).head(15),x="Memo_Orders",y=market_col,orientation="h",title="Markets with most memo orders (top 15)"))
            st.dataframe(market,use_container_width=True,hide_index=True)
            repeat=act.groupby(["StaffKey","Sr Name",market_col],dropna=False).agg(Visit_Events=("Number Of Outlet Visit","sum"),Visit_Dates=("Visit Date","nunique"),Memo_Orders=("Number Of Ordered Shop","sum")).reset_index().sort_values(["Visit_Dates","Visit_Events"],ascending=False)
            st.markdown("#### Which SO visited the same market most?")
            st.caption("The workbook has Market Name but no individual shop/outlet name. This chart ranks repeat visits to the same market, not repeat visits to the same individual shop.")
            most_same=repeat.sort_values(["Visit_Events","Visit_Dates"],ascending=False).drop_duplicates("StaffKey").head(10).copy()
            if not most_same.empty:
                most_same["SO / Market"]=most_same["Sr Name"].astype(str)+" · "+most_same[market_col].astype(str)
                fig=px.bar(most_same.sort_values("Visit_Events"),x="Visit_Events",y="SO / Market",orientation="h",text="Visit_Events",title="Top 10 SOs by visits in their most-visited market",hover_data=["Visit_Dates","Memo_Orders"])
                fig.update_layout(yaxis={"categoryorder":"total ascending"}); show_chart(fig,480)
                st.dataframe(most_same[["Sr Name",market_col,"Visit_Events","Visit_Dates","Memo_Orders"]],use_container_width=True,hide_index=True)
            st.markdown("#### SO-wise repeat market activity")
            st.dataframe(repeat[repeat["Visit_Dates"]>1].head(300),use_container_width=True,hide_index=True)
            act["Week"] = act["Visit Date"].dt.to_period("W").astype(str)
            weekly=act.groupby(["Week",market_col],dropna=False).agg(Outlet_Visits=("Number Of Outlet Visit","sum"),Memo_Orders=("Number Of Ordered Shop","sum")).reset_index()
            st.markdown("#### Weekly busiest markets")
            st.dataframe(weekly.sort_values("Outlet_Visits",ascending=False).groupby("Week").head(5),use_container_width=True,hide_index=True)
        else: st.warning("Market Name column পাওয়া যায়নি।")

with tabs[5]:
    st.subheader("AI-style management summary")
    st.caption("These are rule-based observations calculated from the uploaded workbook, not a generative AI API response.")
    insights=[]
    va=pct(visit_total,visit_target); ma=pct(memo_total,memo_target)
    if va>=100: insights.append(f"Outlet visits are at {va:.1f}% of target; the selected team has met or exceeded the benchmark.")
    else: insights.append(f"Outlet visits are at {va:.1f}% of target, leaving approximately {max(0,visit_target-visit_total):,} visits to reach the workbook target.")
    if ma>=100: insights.append(f"Memo achievement is {ma:.1f}%: memo target is met or exceeded.")
    else: insights.append(f"Memo achievement is {ma:.1f}%; approximately {max(0,memo_target-memo_total):,} additional memos are needed to reach target.")
    if not rep.empty:
        lowv=rep.sort_values("Visit Achievement %").head(min(10,len(rep)))
        lowm=rep.sort_values("Memo Achievement %").head(min(10,len(rep)))
        bestg=rep.groupby("Group").apply(lambda x: pd.Series({"Outlet Achievement %":pct(x["Total Number Of Outlet Visit"].sum(),x["Total Outlet visitTarget"].sum()),"Memo Achievement %":pct(x["Number Of Ordered Memo"].sum(),x["Total Shop Memo Target"].sum())}),include_groups=False).sort_values("Outlet Achievement %",ascending=False) if "Group" in rep else pd.DataFrame()
        if not bestg.empty:
            insights.append(f"Highest group outlet achievement: {bestg.index[0]} ({bestg.iloc[0]['Outlet Achievement %']:.1f}%). Compare memo achievement before concluding overall leadership quality.")
        insights.append(f"{len(rep[rep['Visit Achievement %']<60])} SO(s) are below 60% outlet target achievement and should be reviewed for route coverage, attendance and field execution.")
        insights.append(f"{len(rep[rep['Memo Achievement %']<60])} SO(s) are below 60% memo target achievement; check outlet quality, order conversion and stock availability before assigning corrective actions.")
        st.markdown("**Recommended management actions**")
        for x in insights: st.markdown(f"- {x}")
        c1,c2=st.columns(2)
        with c1: st.markdown("#### Outlet visit improvement priority"); st.dataframe(lowv[["Sr Name","Group","SM Name","RSM Name","Total Number Of Outlet Visit","Visit Achievement %"]],use_container_width=True,hide_index=True)
        with c2: st.markdown("#### Memo improvement priority"); st.dataframe(lowm[["Sr Name","Group","SM Name","RSM Name","Number Of Ordered Memo","Memo Achievement %"]],use_container_width=True,hide_index=True)
    st.warning("Data limitation: Attendance sheet has no Total Hours field, and no individual outlet-name field is present. Working-hour actuals and individual-shop repeat counts cannot be calculated accurately unless those fields are added to the source workbook.")

with tabs[6]:
    st.subheader("SO-level consolidated report")
    view=rep.copy()
    view=view.rename(columns={"Total Outlet visitTarget":"Outlet Visit Target","Total Number Of Outlet Visit":"Outlet Visits","percentage of Visit":"Outlet Achievement (source)","Total Shop Memo Target":"Memo Target","Number Of Ordered Memo":"Memo Count","Percentage of Memo":"Memo Achievement (source)","Attendance ( Count of 21 Days)":"Attendance Days","Present Days":"Present Days (source)","Percentage of Present":"Present % (source)"})
    st.dataframe(view,use_container_width=True,hide_index=True)
    csv=view.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download filtered SO report (CSV)",data=csv,file_name=f"SO_Productivity_{selected_month}.csv",mime="text/csv")

st.markdown("---")
st.caption("Internal management analytics · Attendance timing uses actual In Time Only records. Exactly 8:45 AM is shown separately; before and after percentages use timed Present/Late records for each day.")
st.caption("Security note: hiding Streamlit's toolbar is only a UI change. Keep the GitHub repository private and never commit secrets if you need to protect source code.")

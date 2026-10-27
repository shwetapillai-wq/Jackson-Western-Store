import os
import altair as alt
import pandas as pd
import streamlit as st

TITLE = "JW Store Work Dashboard"
st.set_page_config(page_title=TITLE, layout="wide")

def setting(name, default=""):
    try:
        return st.secrets[name]
    except Exception:
        return os.environ.get(name, default)

SHEET_ID = setting("SHEET_ID")
GID = setting("GID", "0")                      # gid of the "Daily Data" tab
GID_UPDATES = setting("GID_UPDATES")           # gid of the "Team Updates" tab (optional)
GID_KEYWORDS = setting("GID_KEYWORDS")         # gid of the "SEO Keywords" tab (optional)
LOCAL_FILE = os.environ.get("LOCAL_FILE", "")  # only for testing
LOCAL_UPDATES = os.environ.get("LOCAL_UPDATES", "")  # only for testing

def parse_dates(series):
    """Reads dates safely whatever format the sheet exports (2026-10-09, 9/10/2026, 10/9/2026, 9-Oct-2026)."""
    s = series.astype(str).str.strip()
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    iso = s.str.match(r"^\d{4}-\d{1,2}-\d{1,2}")
    out[iso] = pd.to_datetime(s[iso].str[:10], format="%Y-%m-%d", errors="coerce")
    sl = s.str.match(r"^\d{1,2}/\d{1,2}/\d{4}")
    if sl.any():
        parts = s[sl].str.extract(r"^(\d{1,2})/(\d{1,2})/(\d{4})").astype(int)
        # day first (India) unless the data proves month first (second number above 12)
        month_first = parts[0].max() <= 12 and parts[1].max() > 12
        out[sl] = pd.to_datetime(s[sl].str[:10].str.extract(r"^(\d{1,2}/\d{1,2}/\d{4})")[0],
                                 format="%m/%d/%Y" if month_first else "%d/%m/%Y", errors="coerce")
    rest = ~(iso | sl)
    if rest.any():
        out[rest] = pd.to_datetime(s[rest], errors="coerce")   # e.g. 9-Oct-2026
    return out

@st.cache_data(ttl=300)  # re-reads the sheet at most every 5 minutes
def load():
    if LOCAL_FILE:
        raw = pd.read_excel(LOCAL_FILE, sheet_name="Daily Data", header=None)
    else:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
        raw = pd.read_csv(url, header=None, dtype=str)
    # Find the header row by name, so extra blank rows or moved columns do not matter
    hits = raw.index[raw.apply(lambda r: r.astype(str).str.strip().eq("SKUs Worked").any(), axis=1)]
    if len(hits) == 0:
        raise ValueError("Could not find a 'SKUs Worked' header. Is GID the Daily Data tab?")
    h = hits[0]
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[h]]
    df = df.loc[:, [c for c in df.columns if c not in ("nan", "")]]
    for c in ["Comments", "Name"]:
        if c not in df.columns:
            df[c] = ""
    df["Date"] = parse_dates(df["Date"])
    for c in ["SKUs Worked", "SKUs Completed", "Hours"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    stats = {"blank_dates": int(df["Date"].isna().sum())}
    df = df.dropna(subset=["Date"])
    before = len(df)
    df = df.drop_duplicates(subset=["Date", "Name", "SKUs Worked", "SKUs Completed", "Hours", "Comments"])
    stats["dups"] = before - len(df)
    stats["same_day_twice"] = int(df.duplicated(subset=["Date", "Name"]).sum())
    df["Name"] = df["Name"].fillna("").astype(str).str.strip()
    df["Comments"] = df["Comments"].fillna("").astype(str).str.strip()
    df["Week Start"] = pd.to_datetime(df["Date"] - pd.to_timedelta(df["Date"].dt.weekday, unit="D"))
    df["Month"] = df["Date"].dt.to_period("M").astype(str)
    df["Worked day"] = (df["SKUs Worked"] > 0) | (df["Hours"] > 0)
    df.attrs["stats"] = stats
    return df

def read_tab(local, sheet, gid, key):
    """Reads a tab and finds its header row by the column name `key`."""
    if local:
        raw = pd.read_excel(local, sheet_name=sheet, header=None)
    else:
        raw = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}", header=None, dtype=str)
    hits = raw.index[raw.apply(lambda r: r.astype(str).str.strip().eq(key).any(), axis=1)]
    if len(hits) == 0:
        raise ValueError(f"Could not find a '{key}' header in {sheet}.")
    df = raw.iloc[hits[0] + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[hits[0]]]
    return df

def monday(s):
    d = parse_dates(s)
    return d - pd.to_timedelta(d.dt.weekday, unit="D")

@st.cache_data(ttl=300)
def load_updates():
    if not (GID_UPDATES or LOCAL_UPDATES):
        return None, None
    u = read_tab(LOCAL_UPDATES, "Team Updates", GID_UPDATES, "Activity")
    u["Week Start"] = monday(u["Week Start"])
    for c in ["Team", "Activity", "Details"]:
        u[c] = u[c].fillna("").astype(str).str.strip()
    for c in ["Count", "Hours"]:
        u[c] = pd.to_numeric(u[c], errors="coerce")
    u = u[(u["Activity"] != "") & u["Week Start"].notna()]
    k = None
    if GID_KEYWORDS or LOCAL_UPDATES:
        k = read_tab(LOCAL_UPDATES, "SEO Keywords", GID_KEYWORDS, "Keyword")
        k["Week Start"] = monday(k["Week Start"])
        k["Rank"] = pd.to_numeric(k["Rank"], errors="coerce")
        k = k[k["Keyword"].notna() & k["Week Start"].notna()]
    return u, k

def summarise(g):
    t = g.agg(Worked=("SKUs Worked", "sum"), Completed=("SKUs Completed", "sum"),
              Hours=("Hours", "sum"), Days=("Worked day", "sum"))
    t["Completion %"] = (t["Completed"] / t["Worked"].where(t["Worked"] > 0) * 100).round(0)
    t["Completed / hour"] = (t["Completed"] / t["Hours"].where(t["Hours"] > 0)).round(1)
    return t

def worked_vs_done(t, x):
    d = t.reset_index()
    enc = (alt.X(f"{x}:O", timeUnit="yearmonthdate", title=None, axis=alt.Axis(format="%d %b", labelAngle=0))
           if x == "Week Start" else alt.X(f"{x}:N", title=None, sort=None))
    tips = [x, "Worked", "Completed"]
    a = alt.Chart(d).mark_bar(opacity=0.35).encode(x=enc, y=alt.Y("Worked:Q", title="SKUs"), tooltip=tips)
    b = alt.Chart(d).mark_bar().encode(x=enc, y="Completed:Q", tooltip=tips)
    st.altair_chart((a + b).properties(width="container", height=300))
    st.caption("Light bar = SKUs worked, dark bar = SKUs completed.")

st.title(TITLE)
if not SHEET_ID and not LOCAL_FILE:
    st.error("SHEET_ID is not set. Add it in Settings > Secrets.")
    st.stop()
full = load()
stats = full.attrs.get("stats", {})
df = full

with st.sidebar:
    st.header("Filters")
    d0, d1 = df["Date"].min().date(), df["Date"].max().date()
    rng = st.date_input("Date range", (d0, d1), min_value=d0, max_value=d1)
    names = sorted([n for n in df["Name"].unique() if n])
    pick = st.multiselect("Team member", names)
    if st.button("Refresh now"):
        st.cache_data.clear()
        st.rerun()
if len(rng) == 2:
    df = df[(df["Date"].dt.date >= rng[0]) & (df["Date"].dt.date <= rng[1])]
if pick:
    df = df[df["Name"].isin(pick)]

tot = summarise(df.assign(k=1).groupby("k")).iloc[0] if len(df) else None
if tot is None:
    st.warning("No data for this selection.")
    st.stop()
c = st.columns(6)
c[0].metric("SKUs worked", f"{tot.Worked:,.0f}")
c[1].metric("SKUs completed", f"{tot.Completed:,.0f}")
c[2].metric("Completion rate", f"{tot['Completion %']:.0f}%")
c[3].metric("Hours", f"{tot.Hours:,.0f}")
c[4].metric("Completed / hour", f"{tot['Completed / hour']:.1f}")
c[5].metric("Days worked", f"{tot.Days:,.0f}")
st.caption(f"Latest entry: {df['Date'].max().date()}. Data refreshes every 5 minutes.")

t1, t2, t3, t4, t5, t6 = st.tabs(["Weekly", "Monthly", "Daily", "Notes", "Dev & SEO", "Data checks"])
with t1:
    w = summarise(df.groupby("Week Start"))
    worked_vs_done(w, "Week Start")
    show = w.reset_index()
    show["Week Start"] = show["Week Start"].dt.date
    st.dataframe(show.sort_values("Week Start", ascending=False), hide_index=True)
    st.caption("Weeks run Monday to Sunday and are named by their Monday date.")
with t2:
    m = summarise(df.groupby("Month"))
    worked_vs_done(m, "Month")
    st.dataframe(m.reset_index())
with t3:
    d = df.sort_values("Date", ascending=False)[["Date", "Name", "SKUs Worked", "SKUs Completed", "Hours", "Comments"]]
    d["Date"] = d["Date"].dt.date
    st.dataframe(d, hide_index=True)
with t4:
    n = df[df["Comments"] != ""].sort_values("Date", ascending=False)
    lv = n["Comments"].str.lower().str.contains("leave|holiday").sum()
    st.write(f"Days with a comment: **{len(n)}**. Of these, leave or holiday: **{int(lv)}**.")
    top = n["Comments"].str.lower().value_counts().head(10).rename("Times used")
    st.dataframe(top)
    nn = n[["Date", "Name", "Comments"]].copy()
    nn["Date"] = nn["Date"].dt.date
    st.dataframe(nn, hide_index=True)
with t5:
    upd, kw = load_updates()
    if upd is None or upd.empty:
        st.info("Dev and SEO updates are not connected yet. Add GID_UPDATES (and GID_KEYWORDS) in Settings > Secrets.")
    else:
        weeks = [pd.Timestamp(x) for x in sorted(upd["Week Start"].unique(), reverse=True)]
        wk = st.selectbox("Week", weeks, format_func=lambda d: f"{d:%d %b %Y} to {d + pd.Timedelta(days=6):%d %b %Y}")
        for team in sorted(upd["Team"].unique()):
            t = upd[(upd["Week Start"] == wk) & (upd["Team"] == team)][["Activity", "Count", "Hours", "Details"]]
            if len(t):
                st.subheader(f"{team} team")
                st.dataframe(t.dropna(axis=1, how="all").replace({"": pd.NA}).dropna(axis=1, how="all"), hide_index=True)
        if kw is not None and len(kw):
            kk = kw[kw["Week Start"] == wk].copy()
            prev = kw[kw["Week Start"] == wk - pd.Timedelta(days=7)].set_index("Keyword")["Rank"]
            if len(kk):
                kk["Last week"] = kk["Keyword"].map(prev)
                st.subheader("Keyword rankings")
                st.dataframe(kk[["Keyword", "Rank", "Last week"]], hide_index=True)
        counted = upd[upd["Count"].notna()]
        if len(counted):
            st.subheader("Weekly trend")
            act = st.selectbox("Activity", sorted(counted["Activity"].unique()))
            tr = counted[counted["Activity"] == act].groupby("Week Start")["Count"].sum()
            st.bar_chart(tr)
        hrs = upd[upd["Hours"].notna()].pivot_table(index="Week Start", columns="Team", values="Hours", aggfunc="sum")
        if len(hrs):
            st.subheader("Hours by team and week")
            st.bar_chart(hrs)

with t6:
    st.write(f"Rows read from the sheet: **{len(full)}**, across **{full['Date'].dt.date.nunique()}** different dates")
    st.write(f"Blank or unreadable date rows ignored: **{stats.get('blank_dates', 0)}**")
    st.write(f"Exact duplicate rows removed: **{stats.get('dups', 0)}**")
    st.write(f"Extra rows for the same person and date: **{stats.get('same_day_twice', 0)}**")
    st.write(f"Date range in sheet: **{full['Date'].min().date()}** to **{full['Date'].max().date()}**. "
             "If the latest date looks wrong, check the Date column format in the sheet (use 24-Oct-2026 style).")
    st.write(f"Days with SKUs but no hours: **{int(((full['SKUs Worked'] > 0) & (full['Hours'] == 0)).sum())}**")

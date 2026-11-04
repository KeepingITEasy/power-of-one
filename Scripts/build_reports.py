"""Build Power of One member trend reports from everything in Source Data.
Run again after dropping a new month's file into Source Data."""
import os, re, json, warnings
import numpy as np, pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter
warnings.filterwarnings("ignore")
import sys; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from extract import load, BASE

OUT = os.environ.get("POWER_OUT", os.path.join(BASE, "Reports"))  # set POWER_OUT to build elsewhere, then copy (avoids OneDrive locks)
CATS = [("Attendance","Att Pts"),("Ref/Wk","Ref Pts"),("Visitor/Wk","Visitor Pts"),("121/Wk","121 Pts"),("CEU/Wk","CEU Pts")]
CATNAME = {"Att Pts":"Attendance","Ref Pts":"Referrals","Visitor Pts":"Visitors","121 Pts":"1-to-1s","CEU Pts":"CEUs"}

df, hist = load()
reports = sorted(df.Report.unique()); latest, first = reports[-1], reports[0]
prev = reports[-2]
lab = {r: r.strftime("%b %Y") for r in reports}
win = {r: f"{(r-6).strftime('%b')}–{(r-1).strftime('%b %Y')}" for r in reports}

def slope(y):
    y = np.array(y, float); x = np.arange(len(y)); m = ~np.isnan(y)
    return round(np.polyfit(x[m], y[m], 1)[0], 6) if m.sum() >= 3 else np.nan

# ---------------- member summary ----------------
wide = df.pivot_table(index="Key", columns="Report", values="Total Score", aggfunc="first").reindex(columns=reports)
info = df.sort_values("Report").groupby("Key").last()
rows = []
BYKEY = {k: g.sort_values("Report") for k, g in df.groupby("Key")}
for k, s in wide.iterrows():
    i = info.loc[k]; vals = s.values.astype(float); have = ~np.isnan(vals)
    fi = int(np.argmax(have)); li = len(vals)-1-int(np.argmax(have[::-1]))
    cur = vals[-1]; pv = vals[-2]
    sl = slope(vals)
    chg_all = vals[li] - vals[fi] if have.sum() > 1 else np.nan
    if np.isnan(cur): status = "Not in latest report"
    elif have.sum() == 1: status = "New this month"
    elif have.sum() < 3: status = "New (≤2 months)"
    elif sl >= 2 or chg_all >= 15: status = "Improving"
    elif sl <= -2 or chg_all <= -15: status = "Declining"
    else: status = "Steady"
    gk = BYKEY[k]
    last = gk[gk.Report == reports[li]].iloc[0]
    firstrow = gk[gk.Report == reports[fi]].iloc[0]
    pts = {c: last[c] for _, c in CATS}
    focus = min(pts, key=lambda c: (pts[c], ["CEU Pts","Visitor Pts","Ref Pts","121 Pts","Att Pts"].index(c)))
    deltas = {c: last[c]-firstrow[c] for _, c in CATS}
    drop = min(deltas, key=deltas.get)
    gain = max(deltas, key=deltas.get)
    chist = gk.Chapter.drop_duplicates().tolist()
    rows.append(dict(Key=k, Region=i.Region, Chapter=i.Chapter, Member=i.Name,
        **{lab[r]: v for r, v in zip(reports, vals)},
        **{"1-Mo Change": cur - pv if not (np.isnan(cur) or np.isnan(pv)) else np.nan,
           "Change Since First": chg_all, "Trend (pts/mo)": round(sl, 1) if not np.isnan(sl) else np.nan,
           "Months Reported": int(have.sum()), "Status": status,
           "Weeks of Data (latest)": last.Weeks,
           "Attendance %": last.Attendance, "Ref/Wk": last["Ref/Wk"], "Visitor/Wk": last["Visitor/Wk"],
           "121/Wk": last["121/Wk"], "CEU/Wk": last["CEU/Wk"],
           "Focus Area (lowest pts)": CATNAME[focus],
           "Biggest Gain": CATNAME[gain] + f" (+{deltas[gain]:.0f})" if deltas[gain] > 0 else "—",
           "Chapter Moves": " → ".join(chist) if len(chist) > 1 else "",
           "Biggest Drop": CATNAME[drop] + f" ({deltas[drop]:.0f})" if deltas[drop] < 0 else "—"}))
summ = pd.DataFrame(rows)
summ = summ[[c for c in summ.columns if c != "Chapter Moves"] + ["Chapter Moves"]]

# ---------------- chapter rosters (Source Data/Rosters/<Chapter>.csv, one name per line) ----------------
def name_key(n):
    n = re.sub(r"\([^)]*\)", " ", str(n).lower()); n = re.sub(r"[^a-z\s'-]", " ", n)
    return re.sub(r"\s+", " ", n).strip()
first_last = lambda k: (lambda t: f"{t[0]} {t[-1]}" if t else k)(k.split())
rosters = {}
rdir = os.path.join(BASE, "Source Data", "Rosters")
if os.path.isdir(rdir):
    for f in sorted(os.listdir(rdir)):
        if f.lower().endswith((".csv", ".txt")):
            lines = [l.split(",")[0].split("\t")[0].strip() for l in open(os.path.join(rdir, f), encoding="utf-8-sig")]
            names = [l for l in lines if l and not re.match(r"^(member names?|name)$", l, re.I)]
            if names: rosters[os.path.splitext(f)[0]] = names
former_keys = set(); stubs = []
for chn, names in rosters.items():
    inch = summ[summ.Chapter == chn]
    if inch.empty: print(f"Roster '{chn}': chapter not found in data – skipped"); continue
    exact = {name_key(n): n for n in names}; fl = {first_last(name_key(n)): n for n in names}
    found = set()
    for idx, r in inch.iterrows():
        k = name_key(r.Member); hit = exact.get(k) or fl.get(first_last(k))
        if hit: found.add(hit)
        else: summ.loc[idx, "Status"] = "Former member"; former_keys.add(r.Key)
    for n in names:
        if n not in found:
            stubs.append(dict(Key=f"roster|{chn}|{n}", Region=inch.Region.iloc[0], Chapter=chn, Member=n, Status="New – no data yet", **{"Months Reported": 0}))
    print(f"Roster '{chn}': {len(names)} current · {len([s for s in stubs if s['Chapter']==chn])} new without data · {inch.Key.isin(former_keys).sum()} former")
if stubs: summ = pd.concat([summ, pd.DataFrame(stubs)], ignore_index=True)
summ = summ.sort_values(["Region","Chapter",lab[latest]], ascending=[True,True,False], na_position="last")
df_stats = df[~df.Key.isin(former_keys)]   # chapter/region averages count current members only

# ---------------- metric trends (long → wide per metric) ----------------
metric_wide = {}
for raw, pts in [("Total Score",None)] + CATS:
    w = df.pivot_table(index="Key", columns="Report", values=raw, aggfunc="first").reindex(columns=reports)
    w.columns = [lab[c] for c in w.columns]
    w = info[["Region","Chapter","Name"]].join(w).reset_index(drop=True).rename(columns={"Name":"Member"})
    metric_wide[raw] = w.sort_values(["Region","Chapter","Member"])

# ---------------- score history (11 windows) ----------------
ends = sorted(hist.WindowEnd.unique())
hw = hist.pivot_table(index="Key", columns="WindowEnd", values="Score", aggfunc="first").reindex(columns=ends)
hw.columns = [f"{(e-5).strftime('%b')}–{e.strftime('%b %y')}" for e in ends]
hinfo = hist.sort_values("WindowEnd").groupby("Key")[["Region","Chapter","Name"]].last()
hinfo.update(info[["Region","Chapter","Name"]])
hw = hinfo.join(hw).reset_index(drop=True).rename(columns={"Name":"Member"}).sort_values(["Region","Chapter","Member"])

# ---------------- chapter & region trends ----------------
def grp(by):
    g = df_stats.groupby(by + ["Report"]).agg(Members=("Name","size"), AvgScore=("Total Score","mean"),
        Attendance=("Attendance","mean"), Ref=("Ref/Wk","mean"), Vis=("Visitor/Wk","mean"),
        OneToOne=("121/Wk","mean"), CEU=("CEU/Wk","mean"),
        At70=("Total Score", lambda s: (s >= 70).mean())).reset_index()
    return g
chap = grp(["Region","Chapter"]); reg = grp(["Region"])
chap_w = chap.pivot_table(index=["Region","Chapter"], columns="Report", values="AvgScore").reindex(columns=reports)
chap_w.columns = [lab[c] for c in chap_w.columns]
chap_m = chap.pivot_table(index=["Region","Chapter"], columns="Report", values="Members").reindex(columns=reports)
chap_w["Members (latest)"] = chap_m[latest].values
chap_w["1-Mo Change"] = chap_w[lab[latest]] - chap_w[lab[prev]]
chap_w["Change Since First"] = chap_w[lab[latest]] - chap_w[lab[first]]
chap_w = chap_w.reset_index().sort_values(lab[latest], ascending=False)
reg_w = reg.pivot_table(index="Region", columns="Report", values="AvgScore").reindex(columns=reports)
reg_w.columns = [lab[c] for c in reg_w.columns]
reg_w["Change Since First"] = reg_w[lab[latest]] - reg_w[lab[first]]
reg_w = reg_w.reset_index()
overall = df_stats.groupby("Report").agg(Members=("Name","size"), AvgScore=("Total Score","mean"),
    Attendance=("Attendance","mean"), Ref=("Ref/Wk","mean"), Vis=("Visitor/Wk","mean"),
    OneToOne=("121/Wk","mean"), CEU=("CEU/Wk","mean"),
    At70=("Total Score", lambda s: (s >= 70).mean()), Under30=("Total Score", lambda s: (s < 30).mean()))
overall.index = [lab[r] for r in overall.index]

# ---------------- workbook ----------------
fname = f"Power of One Member Trends {first.strftime('%Y-%m')} to {latest.strftime('%Y-%m')}.xlsx"
path = os.path.join(OUT, fname)
GONE = ["Not in latest report", "Former member"]
active = summ[~summ.Status.isin(GONE + ["New – no data yet"])]
# ---------------- weighted rank ----------------
W_LATEST, W_AVG, W_TREND = 0.5, 0.3, 0.2          # weights for the blended score
def trend_score(sl): return float(np.clip(50 + 5 * (0 if np.isnan(sl) else sl), 0, 100))
def weighted_table(upto):
    """Score current members using reports[0..upto]; rank overall and within chapter."""
    cols = reports[:upto + 1]; rows_ = []
    for k in active.Key:
        if k not in wide.index: continue
        v = wide.loc[k, cols].values.astype(float)
        if np.isnan(v[-1]): continue
        have = ~np.isnan(v); sl = slope(v) if have.sum() >= 3 else np.nan
        L, A, T = v[-1], np.nanmean(v), trend_score(sl)
        rows_.append(dict(Key=k, Latest=L, Average=round(A, 1), TrendPts=None if np.isnan(sl) else round(sl, 1), TrendScore=round(T, 1),
                          Weighted=round(W_LATEST * L + W_AVG * A + W_TREND * T, 1)))
    t = pd.DataFrame(rows_).merge(active[["Key","Region","Chapter","Member"]], on="Key")
    t["Overall Rank"] = t.Weighted.rank(ascending=False, method="min").astype(int)
    t["Chapter Rank"] = t.groupby("Chapter").Weighted.rank(ascending=False, method="min").astype(int)
    c = t.groupby(["Region","Chapter"]).agg(Members=("Key","size"), Weighted=("Weighted","mean"), Latest=("Latest","mean"),
                                            Average=("Average","mean"), TrendScore=("TrendScore","mean")).reset_index()
    c["Chapter Rank"] = c.Weighted.rank(ascending=False, method="min").astype(int)
    return t, c
wt, wc = weighted_table(len(reports) - 1)
if len(reports) > 1:
    pt, pc = weighted_table(len(reports) - 2)
    wt = wt.merge(pt[["Key","Overall Rank"]].rename(columns={"Overall Rank":"Prev"}), on="Key", how="left")
    wt["Rank Change"] = wt.Prev - wt["Overall Rank"]
    wc = wc.merge(pc[["Chapter","Chapter Rank"]].rename(columns={"Chapter Rank":"Prev"}), on="Chapter", how="left")
    wc["Rank Change"] = wc.Prev - wc["Chapter Rank"]
wt = wt.sort_values("Overall Rank"); wc = wc.sort_values("Chapter Rank")
rank_members = wt[["Overall Rank","Chapter Rank","Rank Change","Member","Chapter","Region","Weighted","Latest","Average","TrendPts","TrendScore"]].rename(columns={
    "Weighted":"Weighted Score", "Latest":f"Latest ({lab[latest]})", "Average":"Avg of All Reports", "TrendPts":"Trend (pts/mo)", "TrendScore":"Trend Score (0-100)", "Rank Change":"Rank Change vs Last Month"})
rank_chapters = wc[["Chapter Rank","Rank Change","Chapter","Region","Members","Weighted","Latest","Average","TrendScore"]].round(1).rename(columns={
    "Weighted":"Avg Weighted Score", "Latest":f"Avg Latest ({lab[latest]})", "Average":"Avg of All Reports", "TrendScore":"Avg Trend Score", "Rank Change":"Rank Change vs Last Month"})
movers_up = active[active.Status.ne("New this month")].nlargest(25, "Change Since First")
movers_dn = active[active.Status.ne("New this month")].nsmallest(25, "Change Since First")
APP_VERSION, APP_CREATED = "1.0", "November 4, 2026"
readme = pd.DataFrame({"Power of One – Member Trend Analysis": [
    f"Created by BNI Lakefront (Highland Park, IL) · built by Michael Federman, Keeping IT Easy · version {APP_VERSION} · created {APP_CREATED}.",
    "Free for any BNI chapter to use – web version: https://keepingiteasy.github.io/power-of-one/ · Not an official BNI Global product.",
    f"Reports included: {', '.join(lab[r] for r in reports)} (each report covers the 6 months of PALMS data before it; e.g. {lab[latest]} = {win[latest]}).",
    f"Members in latest report: {len(active)} across {df[df.Report==latest].Chapter.nunique()} chapters.",
    "Score scale: 5 categories × 20 pts (Attendance, Referrals, Visitors, 1-to-1s, CEUs) = 100 max.",
    "Status rules: Improving = trend ≥ +2 pts/mo or ≥ +15 since first report; Declining = trend ≤ −2 pts/mo or ≤ −15; otherwise Steady. Needs 3+ months.",
    "Focus Area = category with the lowest points in the member's most recent report (where coaching helps most).",
    "Biggest Gain / Drop = category points change from the member's first to most recent report.",
    "Score History (11 mo) comes from the 'Monthly' tab of each source file – a longer look-back for Total Score only.",
    "Members are matched by name across chapters (chapter moves are listed in 'Chapter Moves'); a name-spelling change will show as two rows. Chapter shown = most recent chapter.",
    "Sheets: Overview · Member Summary · Score History · one tab per metric · Chapter Trends · Region Trends · Top Movers · All Data.",
    f"Weighted rank: score = {int(W_LATEST*100)}% latest score + {int(W_AVG*100)}% average of all reports + {int(W_TREND*100)}% trend (pts/month mapped to 0–100: flat = 50, +10/mo = 100, −10/mo = 0; under 3 months = 50). Current members only; chapters ranked by their members' average.",
    "Chapter rosters: put a file named '<Chapter>.csv' (one member name per line) in 'Source Data/Rosters'. Members not on it show as 'Former member' and are left out of chapter averages; roster names with no data show as 'New – no data yet'.",
    "To refresh: drop the new month's file in 'Source Data' (name starting YYYY-MM) and run Scripts/build_reports.py.",
]})
with pd.ExcelWriter(path, engine="openpyxl") as xw:
    readme.to_excel(xw, sheet_name="Read Me", index=False)
    ov = overall.T; ov.index = ["Members","Avg Total Score","Avg Attendance %","Avg Referrals/Wk","Avg Visitors/Wk","Avg 1-to-1s/Wk","Avg CEUs/Wk","% Scoring 70+","% Scoring Under 30"]
    ov.to_excel(xw, sheet_name="Overview")
    pd.DataFrame(active.Status.value_counts()).rename(columns={"count":"Members"}).to_excel(xw, sheet_name="Overview", startrow=13)
    summ.drop(columns="Key").to_excel(xw, sheet_name="Member Summary", index=False)
    hw.to_excel(xw, sheet_name="Score History (11 mo)", index=False)
    for raw, nm in [("Total Score","Total Score"),("Attendance","Attendance %"),("Ref/Wk","Referrals per Wk"),("Visitor/Wk","Visitors per Wk"),("121/Wk","1-to-1s per Wk"),("CEU/Wk","CEUs per Wk")]:
        metric_wide[raw].to_excel(xw, sheet_name=nm, index=False)
    chap_w.to_excel(xw, sheet_name="Chapter Trends", index=False)
    reg_w.to_excel(xw, sheet_name="Region Trends", index=False)
    rank_members.to_excel(xw, sheet_name="Weighted Rank - Members", index=False)
    rank_chapters.to_excel(xw, sheet_name="Weighted Rank - Chapters", index=False)
    movers_up[["Region","Chapter","Member",lab[first],lab[latest],"Change Since First","Biggest Gain"]].to_excel(xw, sheet_name="Top Movers", index=False, startrow=1)
    movers_dn[["Region","Chapter","Member",lab[first],lab[latest],"Change Since First","Biggest Drop"]].to_excel(xw, sheet_name="Top Movers", index=False, startrow=1, startcol=8)
    ad = df.copy(); ad["Report"] = ad.Report.astype(str)
    ad[["Report","Region","Chapter","Name","Weeks","Total Score"] + [c for p in CATS for c in p]].to_excel(xw, sheet_name="All Data", index=False)

HDR = PatternFill("solid", fgColor="1F4E79"); WHITE = Font(bold=True, color="FFFFFF")
def style(ws, pct_cols=(), dec_cols=(), score_cols=(), chg_cols=(), widths=None, freeze="D2"):
    for c in ws[1]:
        c.fill, c.font = HDR, WHITE; c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32
    if freeze: ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions
    hdr = {c.value: c.column_letter for c in ws[1]}
    n = ws.max_row
    for h, L in hdr.items():
        w = 34 if h in ("Chapter",) else 22 if h in ("Member","Region","Status","Focus Area (lowest pts)","Biggest Gain","Biggest Drop") else 11
        ws.column_dimensions[L].width = (widths or {}).get(h, w)
        fmt = "0%" if h in pct_cols else "0.00" if h in dec_cols else "0.0" if h in chg_cols or h in score_cols else None
        if fmt:
            for r in range(2, n+1): ws[f"{L}{r}"].number_format = fmt
    for h in score_cols:
        if h in hdr: ws.conditional_formatting.add(f"{hdr[h]}2:{hdr[h]}{n}", ColorScaleRule(start_type="num", start_value=0, start_color="F8696B", mid_type="num", mid_value=50, mid_color="FFEB84", end_type="num", end_value=100, end_color="63BE7B"))
    for h in chg_cols:
        if h in hdr:
            rng = f"{hdr[h]}2:{hdr[h]}{n}"
            ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="006100", bold=True)))
            ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=Font(color="9C0006", bold=True)))
    return hdr

wb = load_workbook(path)
mlabs = [lab[r] for r in reports]
ws = wb["Read Me"]; ws.column_dimensions["A"].width = 140; ws["A1"].font = Font(bold=True, size=14)
ws = wb["Overview"]; ws.column_dimensions["A"].width = 24
for c in ws[1]: c.fill, c.font = HDR, WHITE
for col in range(2, len(reports)+2):
    for r, f in [(4,"0%"),(5,"0.00"),(6,"0.00"),(7,"0.00"),(8,"0.00"),(9,"0%"),(10,"0%"),(3,"0.0")]:
        ws.cell(r, col).number_format = f
ws.cell(14,1).value = "Trend status (latest members)"; ws.cell(14,1).font = Font(bold=True)
ch = LineChart(); ch.title = "Average Total Score"; ch.height, ch.width = 7, 16; ch.legend = None
ch.add_data(Reference(ws, min_col=2, max_col=len(reports)+1, min_row=3), from_rows=True, titles_from_data=False)
ch.set_categories(Reference(ws, min_col=2, max_col=len(reports)+1, min_row=1)); ws.add_chart(ch, "J2")
ws = wb["Member Summary"]
hdr = style(ws, pct_cols=["Attendance %"], dec_cols=["Ref/Wk","Visitor/Wk","121/Wk","CEU/Wk"], score_cols=mlabs, chg_cols=["1-Mo Change","Change Since First","Trend (pts/mo)"])
for r in range(2, ws.max_row+1):
    c = ws[f"{hdr['Status']}{r}"]
    c.fill = PatternFill("solid", fgColor={"Improving":"C6EFCE","Declining":"FFC7CE","Steady":"F2F2F2"}.get(c.value, "DDEBF7"))
ws = wb["Score History (11 mo)"]; style(ws, score_cols=list(hw.columns[3:]))
for nm, kind in [("Total Score","s"),("Attendance %","p"),("Referrals per Wk","d"),("Visitors per Wk","d"),("1-to-1s per Wk","d"),("CEUs per Wk","d")]:
    ws = wb[nm]
    hdr = style(ws, pct_cols=mlabs if kind=="p" else (), dec_cols=mlabs if kind=="d" else (), score_cols=mlabs if kind=="s" else ())
    if kind != "s":
        for h in mlabs:
            ws.conditional_formatting.add(f"{hdr[h]}2:{hdr[h]}{ws.max_row}", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="percentile", end_value=95, end_color="5B9BD5"))
ws = wb["Chapter Trends"]; style(ws, score_cols=mlabs, chg_cols=["1-Mo Change","Change Since First"], freeze="C2")
ws = wb["Region Trends"]; style(ws, score_cols=mlabs, chg_cols=["Change Since First"], freeze="B2")
ch = LineChart(); ch.title = "Region Average Score"; ch.height, ch.width = 9, 20
for r in range(2, ws.max_row+1):
    ch.add_data(Reference(ws, min_col=1, max_col=len(reports)+1, min_row=r), from_rows=True, titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, max_col=len(reports)+1, min_row=1)); ws.add_chart(ch, "A10")
for nm in ["Weighted Rank - Members", "Weighted Rank - Chapters"]:
    ws = wb[nm]; style(ws, chg_cols=["Rank Change vs Last Month"], score_cols=[c for c in [ "Weighted Score","Avg Weighted Score"]], freeze="E2" if "Members" in nm else "D2")
    ws.column_dimensions["A"].width = 9; ws.column_dimensions["B"].width = 9; ws.column_dimensions["C"].width = 12
ws = wb["Top Movers"]
ws["A1"], ws["I1"] = f"Top 25 Improvers ({lab[first]} → {lab[latest]})", "Top 25 Decliners"
ws["A1"].font = ws["I1"].font = Font(bold=True, size=12)
for c in ws[2]:
    if c.value: c.fill, c.font = HDR, WHITE
for L, w in zip("ABCDEFGHIJKLMNO", [22,30,22,9,9,10,18,3,22,30,22,9,9,10,18]): ws.column_dimensions[L].width = w
ws = wb["All Data"]; style(ws, freeze="E2")
wb.save(path)

# ---------------- per-chapter workbooks ----------------
cd = os.path.join(OUT, "Chapters"); os.makedirs(cd, exist_ok=True)
for chn, g in summ[~summ.Status.isin(GONE)].groupby("Chapter"):
    p = os.path.join(cd, re.sub(r'[\\/:*?"<>|]', "-", chn) + ".xlsx")
    cols = ["Member"] + mlabs + ["Chapter Moves","1-Mo Change","Change Since First","Trend (pts/mo)","Status","Attendance %","Ref/Wk","Visitor/Wk","121/Wk","CEU/Wk","Focus Area (lowest pts)","Biggest Gain","Biggest Drop"]
    cr = chap[chap.Chapter == chn].copy(); cr["Report"] = cr.Report.map(lab)
    cr = cr.rename(columns={"AvgScore":"Avg Score","Ref":"Ref/Wk","Vis":"Visitor/Wk","OneToOne":"121/Wk","CEU":"CEU/Wk","At70":"% at 70+"}).drop(columns=["Region","Chapter"])
    with pd.ExcelWriter(p, engine="openpyxl") as xw:
        g.sort_values(lab[latest], ascending=False)[cols].to_excel(xw, sheet_name="Members", index=False)
        cr.to_excel(xw, sheet_name="Chapter Trend", index=False)
        rm = rank_members[rank_members.Chapter == chn].sort_values("Chapter Rank").drop(columns=["Chapter","Region"])
        if len(rm): rm.to_excel(xw, sheet_name="Weighted Rank", index=False)
        left = summ[(summ.Chapter == chn) & (summ.Status.isin(GONE))]
        if len(left): left[["Member","Status"] + mlabs].to_excel(xw, sheet_name="No Longer Reported", index=False)
    wb2 = load_workbook(p); ws = wb2["Members"]
    hdr = style(ws, pct_cols=["Attendance %"], dec_cols=["Ref/Wk","Visitor/Wk","121/Wk","CEU/Wk"], score_cols=mlabs, chg_cols=["1-Mo Change","Change Since First","Trend (pts/mo)"], freeze="B2")
    for r in range(2, ws.max_row+1):
        c = ws[f"{hdr['Status']}{r}"]; c.fill = PatternFill("solid", fgColor={"Improving":"C6EFCE","Declining":"FFC7CE","Steady":"F2F2F2"}.get(c.value, "DDEBF7"))
    ws = wb2["Chapter Trend"]; style(ws, pct_cols=["Attendance","% at 70+"], dec_cols=["Ref/Wk","Visitor/Wk","121/Wk","CEU/Wk"], chg_cols=["Avg Score"], freeze="B2")
    ch = LineChart(); ch.title = f"{chn} – Average Score"; ch.legend = None; ch.height, ch.width = 7, 16
    ch.add_data(Reference(ws, min_col=3, min_row=2, max_row=ws.max_row)); ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
    ws.add_chart(ch, "A12")
    if "No Longer Reported" in wb2.sheetnames: style(wb2["No Longer Reported"], score_cols=mlabs, freeze="B2")
    if "Weighted Rank" in wb2.sheetnames: style(wb2["Weighted Rank"], chg_cols=["Rank Change vs Last Month"], score_cols=["Weighted Score"], freeze="D2")
    wb2.save(p)

# ---------------- dashboard data ----------------
H = hist.copy()
members = []
SUMM_BY = summ.drop_duplicates("Key").set_index("Key"); H_BY = {k: g.set_index("WindowEnd")["Score"] for k, g in H.groupby("Key")}
for k, g in df.sort_values("Report").groupby("Key"):
    s = SUMM_BY.loc[k]
    hh = H_BY.get(k, pd.Series(dtype=float))
    members.append(dict(n=s.Member, c=s.Chapter, r=s.Region, st=s.Status,
        cur=None if pd.isna(s[lab[latest]]) else float(s[lab[latest]]),
        d1=None if pd.isna(s["1-Mo Change"]) else float(s["1-Mo Change"]),
        dA=None if pd.isna(s["Change Since First"]) else float(s["Change Since First"]),
        f=s["Focus Area (lowest pts)"], mv=s["Chapter Moves"],
        m=[None if (g.Report == r).sum() == 0 else (lambda x: [float(x["Total Score"])] + [round(float(x[a]),3) for a,_ in CATS] + [float(x[p]) for _,p in CATS] + [float(x.Weeks)])(g[g.Report == r].iloc[0]) for r in reports],
        h=[None if e not in hh.index else float(hh[e]) for e in ends]))
data = dict(reports=[lab[r] for r in reports], windows=[win[r] for r in reports],
    hist=[f"{(e-5).strftime('%b')}–{e.strftime('%b %y')}" for e in ends], members=members,
    overall={k: [float(v) for v in overall[k]] for k in overall.columns})
with open(os.path.join(BASE, "Scripts", "dashboard_template.html")) as f: tpl = f.read()
with open(os.path.join(OUT, "Power of One Dashboard.html"), "w") as f:
    f.write(tpl.replace("/*__DATA__*/null", json.dumps(data, separators=(",",":"))))

print("Wrote", path)
print(active.Status.value_counts().to_string())
print(overall.round(2).to_string())
print(chap_w.head(5).to_string()); print(chap_w.tail(5).to_string())

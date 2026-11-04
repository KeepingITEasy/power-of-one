"""Load every 'Power of One' report in Source Data into one long table."""
import re, glob, os
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "Source Data")

# Regions to include. Some monthly files are nationwide (70,000+ members); keeping the same
# regions every month keeps the trends comparable. Set to None to include every region.
REGIONS = ["IL Central", "IL Northern (Northeast Suburbs)", "IL Northern (South Suburbs)",
           "IL Northern (West Suburbs)", "IA 2 (Cedar Rapids)"]

def norm(s):
    return re.sub(r"\s+", " ", str(s)).strip() if s is not None else None

def load():
    rows, hist = [], []
    for f in sorted(glob.glob(os.path.join(SRC, "*.xlsx"))):
        m = re.match(r"(\d{4})-(\d{2})", os.path.basename(f))
        if not m: continue
        period = pd.Period(f"{m[1]}-{m[2]}", "M")
        x = pd.ExcelFile(f)
        p = pd.read_excel(x, "Power of One")
        mo = pd.read_excel(x, "Monthly")
        if REGIONS: mo = mo[mo["Region"].isin(REGIONS)]
        p.columns = [norm(c) for c in p.columns]
        raw_region = p["Region"].map(norm) if "Region" in p.columns else None
        # positional rename (duplicate 'Points' headers)
        cols = list(p.columns)
        start = cols.index("Total Score")
        names = ["Total Score","Attendance","Att Pts","Ref/Wk","Ref Pts","Visitor/Wk","Visitor Pts","121/Wk","121 Pts","CEU/Wk","CEU Pts"]
        p = p.iloc[:, [cols.index("Chapter"), cols.index("Name (weeks of PALMS data)")] + list(range(start, start+11))]
        p.columns = ["Chapter","NameWeeks"] + names
        p = p.dropna(subset=["NameWeeks"])
        p["Weeks"] = p["NameWeeks"].astype(str).str.extract(r"\((\d+)\)\s*$")[0].astype(float)
        p["Name"] = p["NameWeeks"].astype(str).str.replace(r"\s*\(\d+\)\s*$", "", regex=True).map(norm)
        p["Chapter"] = p["Chapter"].map(norm)
        mo["Chapter"] = mo["Chapter"].map(norm); mo["Name"] = mo["Name"].map(norm)
        reg = mo.drop_duplicates(["Chapter","Name"]).set_index(["Chapter","Name"])["Region"]
        p["Region"] = [reg.get((c,n)) for c,n in zip(p.Chapter, p.Name)]
        if raw_region is not None: p["Region"] = p["Region"].fillna(raw_region.loc[p.index])
        if REGIONS: p = p[p.Region.isin(REGIONS) | p.Region.isna()]   # drop other regions before name matching
        p["Report"] = period
        rows.append(p.drop(columns="NameWeeks"))
        # rolling-window score history from the Monthly sheet
        wcols = [c for c in mo.columns if re.match(r"^[A-Z][a-z]{2}-[A-Z][a-z]{2}$", str(c))]
        for i, c in enumerate(wcols):
            end = period - (len(wcols) - i)   # window ending month before report
            h = mo[["Region","Chapter","Name"]].copy()
            h["WindowEnd"] = end; h["Window"] = c; h["Score"] = mo[c]; h["Source"] = period
            hist.append(h.dropna(subset=["Score"]))
    df = pd.concat(rows, ignore_index=True)
    # canonical chapter names: "X" and "BNI X" are the same chapter
    chs = set(df.Chapter) | set(pd.concat(hist).Chapter)
    canon = {c: "BNI " + c for c in chs if not c.startswith("BNI ") and "BNI " + c in chs}
    df["Chapter"] = df.Chapter.replace(canon)
    for h in hist: h["Chapter"] = h.Chapter.replace(canon)
    # person key: name, unless two people share a name in the same report
    clash = set(df.groupby(["Report","Name"]).Chapter.nunique().loc[lambda s: s > 1].index.get_level_values(1))
    df["Key"] = [n + " | " + c if n in clash else n for n, c in zip(df.Name, df.Chapter)]
    # fill missing region from any other month
    rmap = df.dropna(subset=["Region"]).drop_duplicates(["Chapter","Name"]).set_index(["Chapter","Name"])["Region"]
    df["Region"] = df["Region"].fillna(pd.Series([rmap.get((c,n)) for c,n in zip(df.Chapter, df.Name)]))
    if REGIONS: df = df[df.Region.isin(REGIONS)].copy()
    h = pd.concat(hist, ignore_index=True)
    h["Key"] = [n + " | " + c if n in clash else n for n, c in zip(h.Name, h.Chapter)]
    h = h.sort_values("Source").drop_duplicates(["Key","WindowEnd"], keep="last")
    return df, h

if __name__ == "__main__":
    df, h = load()
    print(df.shape, h.shape)
    print(df.groupby("Report").agg(n=("Name","size"), chapters=("Chapter","nunique"), avg=("Total Score","mean"), noregion=("Region", lambda s: s.isna().sum())))
    print(df.Region.value_counts(dropna=False))
    print("dupe keys:", df.duplicated(["Report","Chapter","Name"]).sum())
    print(h.groupby(["WindowEnd"]).agg(n=("Score","size"), w=("Window","first"), avg=("Score","mean")))
    print(df.head(3).T)

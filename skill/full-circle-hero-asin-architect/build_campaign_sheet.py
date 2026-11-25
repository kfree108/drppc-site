#!/usr/bin/env python3
"""
Full Circle Hero ASIN Campaign Architect  (v2.0)
================================================
Reads a seller's real Amazon exports and builds a complete, plug-and-play
Sponsored Products architecture for ONE hero ASIN, priced per keyword.

  python3 build_campaign_sheet.py --sample
  python3 build_campaign_sheet.py --sqp sqp.csv --str str.csv \
      --asin B0XXXXXXXX --sku MY-SKU --brand "My Brand" \
      --price 29.99 --cogs 6.40 --fba-fee 5.12

Inputs (CSV, TSV or XLSX; column headers are matched flexibly)
  --sqp          Brand Analytics > Search Query Performance export
  --str          Sponsored Products > Search Term Report export
  --competitors  optional competitor list (Scout, Helium 10 Xray, Keepa ...)

Outputs (all four, in one run, written to --out-dir)
  <ASIN>-hero-asin-blueprint.html       interactive report + live calculator
  <ASIN>-hero-asin-architecture.xlsx    6 tabs, live formulas on Tab 1 inputs
  <ASIN>-amazon-bulk-2.0.csv            Bulk 2.0 upload, everything paused
  <ASIN>-drppc-strategy-brief.md        copy-paste brief for Dr. PPC in Orbit

The math, per keyword (never one flat CPC for the whole ASIN):
  Break-Even ACoS = (Price - COGS - Referral fee - FBA fee - Returns/storage) / Price
  Keyword CVR     = live PPC CVR when clicks >= 15, otherwise a Bayesian blend
                    (orders + 15 x prior CVR) / (clicks + 15); prior = the
                    keyword's own SQP CVR when it has one, else the ASIN CVR
  Max CPC         = Price x Break-Even ACoS x Keyword CVR      (break-even click)
  Target CPC      = Price x Target ACoS  x Keyword CVR         (capped vs Max CPC)
  Starting bid    = Target CPC / (1 + placement boost), never above the market
                    ceiling (1.25 x the CPC you already pay on that keyword)

Requires Python 3.9+ and openpyxl (pip3 install openpyxl).
"""

import argparse
import csv
import io
import json
import math
import os
import re
import sys
from datetime import date, datetime

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is missing. Install it with:  pip3 install openpyxl")

VERSION = "2.0.0"
PRIOR = 15                      # Bayesian prior, in clicks
SEG_FACTOR = {"gen": 1.0, "skc": 1.15, "def": 0.65, "br": 0.50}   # x non-branded Target ACoS
STAGE_FACTOR = {"profit": 0.42, "balanced": 0.56, "scale": 0.75}  # x Break-Even ACoS
CEIL_MULT = 1.25                # market ceiling = 1.25 x the CPC already paid on the keyword
SIGNUP_URL = ("https://fullcircleorbit.com/signup?from=drppc&utm_source=creator"
              "&utm_campaign=hormozi_skill_funnel&code=FOUNDER10")
CALL_URL = "https://calendly.com/nick-fullcircle/full-circle-chat-website"
ASIN_RE = re.compile(r"^b0[a-z0-9]{8}$", re.I)

BULK_HDR = ["Product", "Entity", "Operation", "Campaign ID", "Ad Group ID", "Portfolio ID", "Ad ID", "Keyword ID",
            "Product Targeting ID", "Campaign Name", "Ad Group Name", "Start Date", "End Date", "Targeting Type",
            "State", "Daily Budget", "SKU", "ASIN", "Ad Group Default Bid", "Bid", "Keyword Text", "Match Type",
            "Bidding Strategy", "Placement", "Percentage", "Product Targeting Expression"]

# key, lane, kind, tail, label, target-ACoS key, match, bidding, TOS %, PP %, budget weight, cap x Max CPC, proof
CAMPAIGN_DEFS = [
    ("auto",   "NB", "AUTO",      "Auto",        "Auto Discovery",              "gen", "Auto",       "Dynamic bids - down only", 0,  0,  7,  1.00, 1.00),
    ("phrase", "NB", "KW-PHRASE", "Discovery",   "Phrase Discovery",            "gen", "Phrase",     "Dynamic bids - down only", 10, 0,  8,  1.00, 1.00),
    ("broad",  "NB", "KW-BROAD",  "Discovery",   "Broad Discovery",             "gen", "Broad",      "Dynamic bids - down only", 0,  0,  4,  1.00, 1.00),
    ("grow",   "NB", "KW-EXACT",  "Growth",      "Growth Exact (SQP gaps)",     "gen", "Exact",      "Dynamic bids - down only", 20, 5,  12, 1.25, 1.00),
    ("domin",  "NB", "KW-EXACT",  "Domination",  "Domination Exact",            "gen", "Exact",      "Dynamic bids - down only", 35, 10, 18, 1.25, 1.00),
    ("skc1",   "NB", "KW-EXACT",  "SKC 1",       "Single-Keyword #1",           "skc", "Exact",      "Fixed bid",                60, 0,  7,  1.25, 1.00),
    ("skc2",   "NB", "KW-EXACT",  "SKC 2",       "Single-Keyword #2",           "skc", "Exact",      "Fixed bid",                60, 0,  7,  1.25, 1.00),
    ("skc3",   "NB", "KW-EXACT",  "SKC 3",       "Single-Keyword #3",           "skc", "Exact",      "Fixed bid",                60, 0,  7,  1.25, 1.00),
    ("br",     "BR", "KW-EXACT",  "Branded",     "Branded Protection",          "br",  "Exact",      "Dynamic bids - down only", 25, 10, 6,  1.00, 1.00),
    ("def",    "BR", "PAT",       "Defensive",   "Defensive PAT (own catalog)", "def", "ASIN exact", "Dynamic bids - down only", 0,  30, 5,  1.00, 1.00),
    ("offd",   "NB", "PAT",       "Off Domin",   "Offensive PAT: Top 25",       "gen", "ASIN exact", "Dynamic bids - down only", 0,  40, 14, 1.25, 1.00),
    ("off2",   "NB", "PAT",       "Off Disc T2", "Offensive PAT: Tier 2",       "gen", "ASIN exact", "Dynamic bids - down only", 0,  20, 7,  1.00, 0.85),
    ("off3",   "NB", "PAT",       "Off Disc T3", "Offensive PAT: Tier 3",       "gen", "ASIN exact", "Dynamic bids - down only", 0,  10, 5,  1.00, 0.70),
]
SEG_LABEL = {"auto": "Discovery", "phrase": "Discovery", "broad": "Discovery", "grow": "SQP Growth", "domin": "Domination",
             "skc1": "SKC", "skc2": "SKC", "skc3": "SKC", "br": "Branded", "def": "Defensive", "offd": "Offensive PAT",
             "off2": "Offensive PAT", "off3": "Offensive PAT"}
MT_CSV = {"Exact": "exact", "Phrase": "phrase", "Broad": "broad", "Negative exact": "negativeExact",
          "Negative phrase": "negativePhrase"}
EXACT_KEYS = ("skc1", "skc2", "skc3", "domin")


# ---------------------------------------------------------------------------- helpers
def r2(x):
    """Round half up to cents, the way a spreadsheet ROUND does."""
    return math.floor(x * 100 + 0.5 + 1e-9) / 100


def rint(x):
    return int(math.floor(x + 0.5))


def norm(h):
    return re.sub(r"[^a-z0-9]+", " ", str(h if h is not None else "").lower()).strip()


def num(v):
    """'$1,234.50', '12.5%', '(3)', '-', 12 -> float. Blank or unreadable -> 0."""
    if v is None:
        return 0.0
    if isinstance(v, (bool, int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return 0.0
    negative = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^0-9.\-eE]", "", s)
    try:
        x = float(s)
    except ValueError:
        return 0.0
    return -x if negative else x


def warn(msg):
    print("  ! " + msg, file=sys.stderr)


def parse_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip().split("T")[0]
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%Y/%m/%d",
                "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def read_table(path):
    """Read a CSV / TSV / TXT / XLSX export into a list of rows (lists of cells)."""
    if not os.path.isfile(path):
        sys.exit("File not found: %s" % path)
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        wb = load_workbook(path, read_only=True, data_only=True)
        return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    raw = open(path, "rb").read()
    text = None
    for enc in ("utf-8-sig", "utf-16", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    head = "\n".join(text.splitlines()[:5])
    delim = "\t" if head.count("\t") > head.count(",") else (";" if head.count(";") > head.count(",") else ",")
    return [row for row in csv.reader(io.StringIO(text), delimiter=delim)]


def pick(hdrs, *specs):
    """First column whose normalised header matches a spec. A spec is a string (exact match)
    or (must-contain tokens, must-not-contain tokens)."""
    for spec in specs:
        if isinstance(spec, str):
            for i, h in enumerate(hdrs):
                if h == spec:
                    return i
            continue
        must, must_not = spec
        for i, h in enumerate(hdrs):
            if h and all(t in h for t in must) and not any(t in h for t in must_not):
                return i
    return None


def find_header(rows, term_specs, what):
    for idx, row in enumerate(rows[:25]):
        hdrs = [norm(c) for c in row]
        if sum(1 for h in hdrs if h) >= 3 and pick(hdrs, *term_specs) is not None:
            return idx, hdrs
    sys.exit("Could not find the header row in the %s file. It needs a 'Search Query' / 'Customer Search Term' "
             "column. Export it again from Amazon without editing the columns." % what)


def cell(row, i):
    return row[i] if i is not None and i < len(row) else None


# ---------------------------------------------------------------------------- parsers
TERM_SPECS_SQP = ["search query", (["search query"], ["score", "volume"]), (["search term"], []),
                  (["query"], ["score", "volume"]), "keyword", (["keyword"], ["volume", "id"])]
TERM_SPECS_STR = ["customer search term", (["customer search term"], []), "search term", (["search term"], []),
                  (["query"], [])]


def parse_sqp(path):
    """-> ({query: {...numbers scaled to 30 days...}}, notes)"""
    rows = read_table(path)
    hi, hdrs = find_header(rows, TERM_SPECS_SQP, "SQP")
    meta = " ".join(str(c) for r in rows[:hi] for c in r if c)
    period = 30 if re.search(r"month", meta, re.I) else 90 if re.search(r"quarter", meta, re.I) else 7
    c = dict(
        q=pick(hdrs, *TERM_SPECS_SQP),
        vol=pick(hdrs, "search query volume", (["search query volume"], []), (["search volume"], []), (["volume"], [])),
        imp_t=pick(hdrs, (["impressions", "total"], [])),
        imp_b=pick(hdrs, (["impressions", "asin count"], []), (["impressions", "brand count"], [])),
        imp_s=pick(hdrs, (["impressions", "share"], []), (["impression share"], [])),
        clk_t=pick(hdrs, (["clicks", "total"], [])),
        clk_b=pick(hdrs, (["clicks", "asin count"], []), (["clicks", "brand count"], [])),
        clk_s=pick(hdrs, (["clicks", "share"], []), (["click share"], [])),
        pur_t=pick(hdrs, (["purchases", "total"], [])),
        pur_b=pick(hdrs, (["purchases", "asin count"], []), (["purchases", "brand count"], [])),
        pur_s=pick(hdrs, (["purchases", "share"], []), (["purchase share"], [])),
        dt=pick(hdrs, "reporting date", (["reporting date"], []), "date", "week"),
    )
    body = [r for r in rows[hi + 1:] if any(x not in (None, "") for x in r)]

    def share_scale(col):
        if col is None:
            return 1.0
        vals = [cell(r, col) for r in body]
        if any("%" in str(v) for v in vals if v is not None) or any(num(v) > 1 for v in vals):
            return 0.01
        return 1.0

    sc = {k: share_scale(c[k]) for k in ("imp_s", "clk_s", "pur_s")}
    out, dates, modeled = {}, set(), False
    for r in body:
        q = str(cell(r, c["q"]) or "").strip().lower()
        if not q:
            continue
        if c["dt"] is not None and cell(r, c["dt"]) not in (None, ""):
            dates.add(str(cell(r, c["dt"])))
        vol = num(cell(r, c["vol"]))
        clk_t = num(cell(r, c["clk_t"])) if c["clk_t"] is not None else 0.0
        if c["clk_t"] is None and vol:
            clk_t, modeled = vol * 0.6, True
        pur_t = num(cell(r, c["pur_t"])) if c["pur_t"] is not None else 0.0
        if c["pur_t"] is None and clk_t:
            pur_t, modeled = clk_t * 0.10, True
        imp_t = num(cell(r, c["imp_t"])) if c["imp_t"] is not None else 0.0
        clk_b = num(cell(r, c["clk_b"])) if c["clk_b"] is not None else num(cell(r, c["clk_s"])) * sc["clk_s"] * clk_t
        pur_b = num(cell(r, c["pur_b"])) if c["pur_b"] is not None else num(cell(r, c["pur_s"])) * sc["pur_s"] * pur_t
        imp_b = num(cell(r, c["imp_b"])) if c["imp_b"] is not None else num(cell(r, c["imp_s"])) * sc["imp_s"] * imp_t
        d = out.setdefault(q, dict(vol=0.0, imp_t=0.0, imp_b=0.0, clk_t=0.0, clk_b=0.0, pur_t=0.0, pur_b=0.0))
        for k, v in (("vol", vol), ("imp_t", imp_t), ("imp_b", imp_b), ("clk_t", clk_t), ("clk_b", clk_b),
                     ("pur_t", pur_t), ("pur_b", pur_b)):
            d[k] += v
    if not out:
        sys.exit("The SQP file has a header row but no search queries under it.")
    days = period * max(1, len(dates))
    k = 30.0 / days
    for d in out.values():
        for key in d:
            d[key] *= k
    notes = ["SQP: %d search queries, %d-day window (%s report), scaled to 30 days." %
             (len(out), days, "weekly" if period == 7 else "monthly" if period == 30 else "quarterly")]
    if modeled:
        notes.append("SQP: some click / purchase counts were missing in the file and are modeled from search volume.")
    return out, notes


def parse_str(path, campaign_filter=None, days_override=None):
    """-> ({term: {clicks, spend, orders, sales, units}}, window_days, notes)"""
    rows = read_table(path)
    hi, hdrs = find_header(rows, TERM_SPECS_STR, "Search Term Report")
    c = dict(
        term=pick(hdrs, *TERM_SPECS_STR),
        clicks=pick(hdrs, "clicks", (["clicks"], ["thru", "rate", "share", "cost"])),
        spend=pick(hdrs, "spend", (["spend"], []), "cost", (["cost"], ["per", "acos", "of sales", "cpc"])),
        sales=pick(hdrs, (["7 day total sales"], []), (["14 day total sales"], []), (["total sales"], ["acos", "cost", "roas"]),
                   (["sales"], ["acos", "cost", "roas", "return", "sku"])),
        orders=pick(hdrs, (["total orders"], []), (["orders"], ["sku"]), "purchases"),
        units=pick(hdrs, (["total units"], []), (["units"], ["sku"])),
        camp=pick(hdrs, "campaign name", (["campaign name"], []), (["campaign"], ["id", "budget", "status"])),
        d0=pick(hdrs, "start date", "date", (["start date"], [])),
        d1=pick(hdrs, "end date", (["end date"], [])),
    )
    missing = [k for k in ("clicks", "spend", "sales", "orders") if c[k] is None]
    if missing:
        sys.exit("The Search Term Report is missing these columns: %s. Found headers: %s" %
                 (", ".join(missing), ", ".join(h for h in hdrs if h)))
    out, dmin, dmax, skipped = {}, None, None, 0
    flt = campaign_filter.lower() if campaign_filter else None
    for r in rows[hi + 1:]:
        term = str(cell(r, c["term"]) or "").strip().lower()
        if not term or term == "*":
            continue
        camp = str(cell(r, c["camp"]) or "").strip()
        if flt and flt not in camp.lower():
            skipped += 1
            continue
        for col in (c["d0"], c["d1"]):
            dt = parse_date(cell(r, col))
            if dt:
                dmin = dt if dmin is None or dt < dmin else dmin
                dmax = dt if dmax is None or dt > dmax else dmax
        d = out.setdefault(term, dict(clicks=0.0, spend=0.0, orders=0.0, sales=0.0, units=0.0))
        d["clicks"] += num(cell(r, c["clicks"]))
        d["spend"] += num(cell(r, c["spend"]))
        d["orders"] += num(cell(r, c["orders"]))
        d["sales"] += num(cell(r, c["sales"]))
        d["units"] += num(cell(r, c["units"]))
    if not out:
        sys.exit("No search terms were read from the Search Term Report%s." %
                 (" (the --campaign-filter matched nothing)" if flt else ""))
    notes = []
    if days_override:
        days = int(days_override)
    elif dmin and dmax:
        days = max(1, (dmax - dmin).days + 1)
    else:
        days = 30
        notes.append("Search Term Report: no date columns found, assumed a 30-day window (pass --days to correct it).")
    notes.insert(0, "Search Term Report: %d search terms, %d-day window." % (len(out), days))
    if flt:
        notes.append("Search Term Report: kept campaigns containing '%s' (%d rows from other campaigns skipped)." %
                     (campaign_filter, skipped))
    return out, days, notes


def parse_competitors(path):
    rows = read_table(path)
    asin_specs = ["asin", (["asin"], ["parent", "revenue", "sales"])]
    hi, hdrs = None, None
    for idx, row in enumerate(rows[:25]):
        h = [norm(x) for x in row]
        if pick(h, *asin_specs) is not None:
            hi, hdrs = idx, h
            break
    if hi is None:
        sys.exit("The competitors file needs an ASIN column.")
    c = dict(
        asin=pick(hdrs, *asin_specs),
        brand=pick(hdrs, "brand", (["brand"], [])),
        title=pick(hdrs, "title", "product details", "product name", (["title"], []), (["product"], ["type", "group"])),
        price=pick(hdrs, "price", (["price"], ["trend", "change", "list"]), (["buy box", "current"], []), (["price"], [])),
        rating=pick(hdrs, "rating", "ratings", (["rating"], ["count", "format"])),
        reviews=pick(hdrs, "review count", "reviews", (["review count"], []), (["rating", "count"], []),
                     (["reviews"], ["rating", "velocity"])),
        rev=pick(hdrs, (["revenue"], ["trend", "change", "parent"]), (["revenue"], []), (["sales"], ["trend", "rank"])),
    )
    out, seen = [], set()
    for r in rows[hi + 1:]:
        m = re.search(r"B0[A-Z0-9]{8}", str(cell(r, c["asin"]) or "").upper())
        if not m or m.group(0) in seen:
            continue
        seen.add(m.group(0))
        out.append(dict(asin=m.group(0), brand=str(cell(r, c["brand"]) or "").strip(),
                        title=str(cell(r, c["title"]) or "").strip()[:110], price=num(cell(r, c["price"])),
                        rating=num(cell(r, c["rating"])), reviews=num(cell(r, c["reviews"])), rev=num(cell(r, c["rev"]))))
    return out, ["Competitors: %d ASINs read." % len(out)]


# ---------------------------------------------------------------------------- built-in sample
def sample_inputs():
    """A demo hero ASIN (fictional brand) so the skill can be shown without any files."""
    prod = dict(brand="Apex Recovery Labs", asin="B09X4K8M2P", sku="ARL-MGPRO-BLK-01", price=69.99, cogs=15.40,
                fba=6.85, ret=2.10, monthly_sales=201000.0, target_tacos=12.5,
                own_asins=["B09X4K9R1L", "B0B2M7N4Q8", "B0C1T5P9W3"])
    # query, SQP volume, click share, purchase share, market CVR, 30d clicks, orders, spend, sales
    core = [
        ("deep tissue massage gun", 148200, .112, .162, .115, 412, 68, 1087.68, 4759.32),
        ("percussion massager for athletes", 64500, .095, .149, .108, 248, 41, 582.80, 2869.59),
        ("brushless massage gun quiet", 39800, .108, .175, .124, 195, 36, 421.20, 2519.64),
        ("back and shoulder massage gun", 52100, .082, .118, .102, 214, 31, 496.48, 2169.69),
        ("muscle recovery massager", 44900, .068, .104, .096, 178, 27, 377.36, 1889.73),
        ("handheld deep tissue massager", 38400, .079, .112, .105, 162, 25, 353.16, 1749.75),
        ("sports massage gun 6 heads", 19200, .135, .194, .131, 134, 24, 261.30, 1679.76),
        ("percussion therapy gun", 31600, .059, .088, .094, 146, 21, 319.74, 1469.79),
        ("cordless massage gun heavy duty", 27400, .081, .121, .111, 129, 20, 278.64, 1399.80),
        ("massage gun for runners", 22900, .074, .109, .104, 118, 19, 245.44, 1329.81),
        ("physical therapy massage gun", 28700, .062, .092, .098, 132, 18, 297.00, 1259.82),
        ("massage gun for sore muscles", 41200, .051, .075, .089, 141, 9, 515.84, 629.91),
        ("high torque percussion massager", 14800, .119, .168, .128, 92, 16, 180.32, 1119.84),
        ("neck and back percussion gun", 24300, .048, .069, .091, 108, 14, 232.20, 979.86),
        ("massage gun with carrying case", 18600, .064, .091, .103, 12, 2, 24.12, 139.98),
        ("gym recovery massage tool", 15900, .012, .008, .095, 0, 0, 0, 0),
        ("quiet massage gun for home", 17400, .009, .006, .093, 0, 0, 0, 0),
        ("deep muscle stimulator gun", 12100, .011, .007, .088, 6, 0, 12.24, 0),
        ("professional massage gun for back pain", 36200, .007, .004, .097, 0, 0, 0, 0),
        ("massage gun deep tissue for legs", 29500, .010, .006, .101, 0, 0, 0, 0),
        ("apex recovery massage gun", 16400, .645, .785, .242, 268, 74, 343.04, 5179.26),
        ("apex pro percussion massager", 8900, .680, .812, .260, 154, 45, 181.72, 3149.55),
        ("cheap massage gun under 25", 33200, .018, .000, .042, 47, 0, 106.22, 0),
        ("mini pocket travel massager", 46800, .015, .002, .051, 39, 0, 88.92, 0),
        ("massage chair pad heated", 58900, .011, .000, .039, 34, 0, 81.94, 0),
        ("foot massager machine shiatsu", 74100, .009, .000, .044, 29, 0, 71.05, 0),
    ]
    sqp, st = {}, {}
    for q, vol, cs, ps, mcvr, clk, od, sp, sa in core:
        clk_t = vol * 0.62
        pur_t = clk_t * mcvr
        sqp[q] = dict(vol=float(vol), imp_t=vol * 2.4, imp_b=vol * 2.4 * cs / 1.15, clk_t=clk_t, clk_b=clk_t * cs,
                      pur_t=pur_t, pur_b=pur_t * ps)
        if clk:
            st[q] = dict(clicks=float(clk), spend=float(sp), orders=float(od), sales=float(sa), units=float(od))
    extra = [  # search terms that only show up in the Search Term Report
        ("tens unit muscle stimulator", 31, 0, 74.40, 0), ("neck pillow massager for car", 26, 0, 58.76, 0),
        ("scalp massager electric waterproof", 24, 0, 51.12, 0), ("cellulite massager handheld", 23, 0, 49.45, 0),
        ("massage oil lavender organic", 16, 0, 31.20, 0), ("foam roller trigger point", 14, 0, 29.82, 0),
        ("heated massage gun under 40", 52, 1, 118.56, 69.99), ("massage gun for kids", 33, 1, 74.91, 69.99),
        ("best budget massage gun", 41, 1, 95.12, 69.99), ("massage gun attachments set", 22, 3, 44.10, 209.97),
        ("percussion gun for shoulders", 19, 3, 39.71, 209.97), ("travel case massage gun", 9, 1, 17.55, 69.99),
        ("b07z9p1k8m", 18, 0, 41.40, 0), ("b08c3v7r2n", 15, 0, 34.50, 0), ("b09m4t8v1x", 44, 6, 96.80, 419.94),
    ]
    for t, clk, od, sp, sa in extra:
        st[t] = dict(clicks=float(clk), spend=float(sp), orders=float(od), sales=float(sa), units=float(od))
    comps = [dict(asin=a, brand=b, title=t, price=p, rating=rt, reviews=rv, rev=mr) for a, b, t, p, rt, rv, mr in [
        ("B08K9L2P4Q", "Competitor A", "Mid-tier percussion massager", 89.99, 4.3, 1420, 165000),
        ("B09M4T8V1X", "Competitor B", "Legacy brand pro massage gun", 119.99, 4.4, 3810, 290000),
        ("B0B7W3N6K2", "Competitor C", "Generic 4-head massage gun", 74.99, 4.1, 680, 92000),
        ("B09R2H5C8M", "Competitor D", "Athlete massage gun", 94.99, 4.4, 2150, 198000),
        ("B0C4J9Q1Z7", "Competitor E", "Quiet pro massage gun V2", 79.99, 4.2, 510, 74000),
        ("B08V6X1D3P", "Competitor F", "Heated percussion massager", 99.99, 4.3, 1890, 152000),
        ("B0B1F8S4Y9", "Competitor G", "Deep muscle XL massager", 84.99, 4.2, 940, 110000),
        ("B09P5G7L2W", "Competitor H", "Physio relief massage gun", 109.99, 4.5, 2640, 215000),
        ("B0C8N2B5R4", "Competitor I", "Cordless elite massager", 76.99, 4.1, 390, 61000),
        ("B08T3K6M9J", "Competitor J", "Gym recovery pro gun", 89.95, 4.3, 1280, 134000),
        ("B0B9Q4V8H1", "Competitor K", "Dual-speed massage gun", 69.99, 4.0, 440, 55000),
        ("B09W1D7F6N", "Competitor L", "Sports clinic massage gun", 129.99, 4.4, 1750, 142000),
        ("B09X4K9R1L", "Apex Recovery Labs", "Apex Mini Travel Gun", 49.99, 4.7, 1920, 88000),
        ("B0B2M7N4Q8", "Apex Recovery Labs", "Apex Replacement Heads", 24.99, 4.8, 860, 34000),
        ("B0C1T5P9W3", "Apex Recovery Labs", "Apex Recovery Roller", 39.99, 4.7, 1140, 52000),
    ]]
    return prod, sqp, st, comps


# ---------------------------------------------------------------------------- the build
def build(args):
    notes, assumed = [], []
    if args.sample:
        sp, sqp, st, comps = sample_inputs()
        days = 30
        for k, v in (("brand", sp["brand"]), ("asin", sp["asin"]), ("sku", sp["sku"]), ("price", sp["price"]),
                     ("cogs", sp["cogs"]), ("fba_fee", sp["fba"]), ("returns_storage", sp["ret"]),
                     ("monthly_sales", sp["monthly_sales"]), ("target_tacos", sp["target_tacos"])):
            if getattr(args, k) is None:
                setattr(args, k, v)
        if not args.own_asins:
            args.own_asins = ",".join(sp["own_asins"])
        notes.append("SAMPLE MODE: demo data for a fictional brand. Run with --sqp and --str for your own account.")
    else:
        if not args.asin:
            sys.exit("--asin is required (the hero ASIN this build is for).")
        sqp, st, comps, days = {}, {}, [], 30
        if args.sqp:
            sqp, n = parse_sqp(args.sqp)
            notes += n
        if args.str:
            st, days, n = parse_str(args.str, args.campaign_filter, args.days)
            notes += n
        if args.competitors:
            comps, n = parse_competitors(args.competitors)
            notes += n
        if not args.sqp:
            notes.append("No SQP file: search volume and share columns are blank, and no SQP gap keywords can be found.")
        if not args.str:
            notes.append("No Search Term Report: every keyword CVR is an SQP estimate and no bleeders can be found.")

    asin = args.asin.strip().upper()
    brand = (args.brand or "").strip() or "My Brand"
    if not args.brand:
        assumed.append("Brand name not given (--brand): branded search terms could not be separated.")
    brand_terms = [t.strip().lower() for t in (args.brand_terms or "").split(",") if t.strip()]
    if not brand_terms and args.brand:
        generic = ("labs", "brand", "brands", "company", "store", "official", "home", "products", "group", "the")
        words = [w for w in re.findall(r"[a-z0-9]+", brand.lower()) if len(w) >= 4 and w not in generic]
        brand_terms = words[:1] if words else [brand.lower()]

    kw_live = {t: d for t, d in st.items() if not ASIN_RE.match(t.replace(" ", ""))}
    pat_live = {t.replace(" ", "").upper(): d for t, d in st.items() if ASIN_RE.match(t.replace(" ", ""))}
    tot = {k: sum(d[k] for d in st.values()) for k in ("clicks", "spend", "orders", "sales", "units")}

    # ---- unit economics
    price = args.price
    if price is None:
        if tot["units"] > 0 or tot["orders"] > 0:
            price = r2(tot["sales"] / (tot["units"] or tot["orders"]))
            assumed.append("Price not given (--price): used $%.2f, the average ad sale in the Search Term Report." % price)
        else:
            sys.exit("--price is required (no ad sales in the files to infer it from).")
    if price <= 0:
        sys.exit("--price must be above 0.")
    cogs = args.cogs
    if cogs is None:
        cogs = r2(price * 0.30)
        assumed.append("COGS not given (--cogs): ASSUMED 30%% of price = $%.2f. Replace it in the calculator." % cogs)
    fba = args.fba_fee
    if fba is None:
        fba = r2(price * 0.12 + 2.50)
        assumed.append("FBA fee not given (--fba-fee): ASSUMED $%.2f. Replace it in the calculator." % fba)
    ret = args.returns_storage if args.returns_storage is not None else 0.0
    ref_pct = args.referral_fee_pct / 100.0
    ref_fee = r2(price * ref_pct)
    cm = price - cogs - ref_fee - fba - ret
    be = max(0.0, cm / price)
    if be <= 0:
        warn("Break-even ACoS is 0%: at these costs the product loses money before ads. Check price, COGS and fees.")
    if args.target_acos is not None:
        t = args.target_acos / 100.0
    else:
        t = math.floor(be * STAGE_FACTOR[args.stage] * 1000 + 0.5) / 1000
    t = max(t, 0.01)

    # ---- ASIN conversion rate
    sqp_clk = sum(d["clk_b"] for d in sqp.values())
    sqp_pur = sum(d["pur_b"] for d in sqp.values())
    if tot["clicks"] >= 100 and tot["orders"] > 0:
        asin_cvr, cvr_src = tot["orders"] / tot["clicks"], "Search Term Report, all clicks"
    elif sqp_clk >= 50 and sqp_pur > 0:
        asin_cvr, cvr_src = sqp_pur / sqp_clk, "SQP, all ASIN clicks"
    elif tot["clicks"] > 0 and tot["orders"] > 0:
        asin_cvr, cvr_src = tot["orders"] / tot["clicks"], "Search Term Report (thin data)"
    else:
        asin_cvr, cvr_src = 0.10, "ASSUMED 10%"
        assumed.append("Not enough clicks to measure the ASIN conversion rate: ASSUMED 10%.")
    asin_cvr = min(asin_cvr, 0.95)
    cfs = int(math.ceil(1.0 / asin_cvr))
    neg_clicks = 3 * cfs
    avg_cpc = tot["spend"] / tot["clicks"] if tot["clicks"] else 0.0
    k30 = 30.0 / days

    # ---- sales, TACoS and budget goals
    mo_ad_spend, mo_ad_sales = tot["spend"] * k30, tot["sales"] * k30
    monthly_sales = args.monthly_sales
    if monthly_sales is None:
        if mo_ad_sales > 0:
            monthly_sales = float(rint(mo_ad_sales / 0.45))
            assumed.append("Total monthly sales not given (--monthly-sales): ASSUMED $%s (ad sales are 45%% of total)."
                           % format(int(monthly_sales), ","))
        else:
            monthly_sales = float(rint(max(sqp_pur * price, 3000.0)))
            assumed.append("Total monthly sales not given (--monthly-sales): ASSUMED $%s from SQP purchases."
                           % format(int(monthly_sales), ","))
    tacos = args.target_tacos / 100.0 if args.target_tacos is not None else math.floor(t * 0.45 * 1000 + 0.5) / 1000
    budget_manual = args.daily_budget is not None
    daily_budget = float(args.daily_budget) if budget_manual else float(max(10, rint(monthly_sales * tacos / 30.0)))

    # ---- the keyword universe
    terms = {}
    for q, d in sqp.items():
        terms[q] = dict(kw=q, vol=d["vol"], clk_t=d["clk_t"], clk_b=d["clk_b"], pur_t=d["pur_t"], pur_b=d["pur_b"],
                        imp_share=(d["imp_b"] / d["imp_t"] if d["imp_t"] else None),
                        click_share=(d["clk_b"] / d["clk_t"] if d["clk_t"] else None),
                        purch_share=(d["pur_b"] / d["pur_t"] if d["pur_t"] else None), in_sqp=True)
    for q, d in kw_live.items():
        e = terms.setdefault(q, dict(kw=q, vol=None, clk_t=0, clk_b=0, pur_t=0, pur_b=0, imp_share=None, click_share=None,
                                     purch_share=None, in_sqp=False))
        e.update(clicks=d["clicks"], spend=d["spend"], orders=d["orders"], sales=d["sales"])
    for e in terms.values():
        for k in ("clicks", "spend", "orders", "sales"):
            e.setdefault(k, 0.0)
        e["branded"] = any(b in e["kw"] for b in brand_terms)
        e["acos"] = e["spend"] / e["sales"] if e["sales"] > 0 else None
        e["cur_cpc"] = e["spend"] / e["clicks"] if e["clicks"] > 0 else None
        # keyword CVR: live when there is enough of it, otherwise blended with the best prior
        sqp_cvr = None
        if e["clk_b"] >= PRIOR and e["pur_b"] >= 1:
            sqp_cvr = e["pur_b"] / e["clk_b"]
        elif e["clk_b"] > 0:
            sqp_cvr = (e["pur_b"] + PRIOR * asin_cvr) / (e["clk_b"] + PRIOR)
        prior = sqp_cvr if sqp_cvr is not None else asin_cvr
        if e["clicks"] >= PRIOR and e["orders"] > 0:
            e["cvr"], e["cvr_src"] = e["orders"] / e["clicks"], "Live PPC"
        elif e["clicks"] > 0:
            e["cvr"], e["cvr_src"] = (e["orders"] + PRIOR * prior) / (e["clicks"] + PRIOR), "Blended"
        elif sqp_cvr is not None:
            e["cvr"], e["cvr_src"] = sqp_cvr, "SQP est."
        else:
            mkt = e["pur_t"] / e["clk_t"] if e["clk_t"] else asin_cvr
            e["cvr"], e["cvr_src"] = 0.5 * mkt + 0.5 * asin_cvr, "SQP est."
        e["cvr"] = round(min(e["cvr"], 0.95), 4)
        e["ceil"] = r2(e["cur_cpc"] * CEIL_MULT) if e["cur_cpc"] else None

    negs, watch, cand = [], [], []
    for e in terms.values():
        if e["branded"]:
            continue
        if e["clicks"] >= neg_clicks and e["orders"] == 0:
            e["why"] = "3x rule: %d clicks, 0 orders, $%.2f spent" % (e["clicks"], e["spend"])
            negs.append(e)
        elif (e["acos"] is not None and e["acos"] >= 3 * t and e["spend"] >= 1.25 * cfs * avg_cpc
              and e["cvr"] < 0.6 * asin_cvr):
            # a term that converts like the ASIN is repriced through its own Max CPC instead of negated
            e["why"] = "3x ACoS: %.0f%% ACoS on $%.2f spend (limit %.0f%%)" % (e["acos"] * 100, e["spend"], 3 * t * 100)
            negs.append(e)
        elif e["orders"] == 0 and e["clicks"] >= cfs:
            e["why"] = "%d clicks, 0 orders: not negated yet (the 3x line is %d clicks)" % (e["clicks"], neg_clicks)
            watch.append(e)
        else:
            cand.append(e)
    negs.sort(key=lambda e: -e["spend"])
    watch.sort(key=lambda e: -e["spend"])

    def proven(e):
        # a winner converts; its ACoS today does not matter, because its bid is repriced from its own CVR
        return e["orders"] >= 2 and e["cvr"] >= 0.6 * asin_cvr

    winners = sorted([e for e in cand if proven(e)], key=lambda e: (-(e["vol"] or 0), -e["orders"]))
    n_skc = min(3, len(winners) // 2)
    assign = {}
    for i, e in enumerate(winners[:n_skc]):
        assign[e["kw"]] = "skc%d" % (i + 1)
    for e in sorted(winners[n_skc:], key=lambda e: -e["orders"])[:40]:
        assign[e["kw"]] = "domin"
    rest = [e for e in cand if e["kw"] not in assign]
    # SQP gaps: real search volume, not proven in ads yet, and a CVR that is not a relevance miss
    for e in sorted([e for e in rest if e["in_sqp"] and (e["vol"] or 0) > 0 and e["orders"] < 2
                     and e["cvr"] >= 0.35 * asin_cvr], key=lambda e: -(e["vol"] or 0))[:25]:
        assign[e["kw"]] = "grow"
    rest = [e for e in rest if e["kw"] not in assign]
    for e in sorted([e for e in rest if e["orders"] >= 1], key=lambda e: -e["orders"])[:20]:
        assign[e["kw"]] = "phrase"
    for e in sorted([e for e in terms.values() if e["branded"]],
                    key=lambda e: -(e["clicks"] + (e["vol"] or 0) / 1000.0))[:15]:
        assign[e["kw"]] = "br"

    # broad discovery stems: the two-word roots the winners share
    stems, stop = {}, {"for", "and", "the", "with", "of", "to", "in", "a", "an", "on"}
    for e in terms.values():
        if assign.get(e["kw"]) in EXACT_KEYS + ("grow",):
            w = [x for x in e["kw"].split() if x not in stop]
            for a, b in zip(w, w[1:]):
                stems[a + " " + b] = stems.get(a + " " + b, 0) + 1
    broad = [s for s, n in sorted(stems.items(), key=lambda x: (-x[1], x[0])) if n >= 2 and s not in assign][:6]

    order = [d[0] for d in CAMPAIGN_DEFS]
    match_of = {d[0]: d[6] for d in CAMPAIGN_DEFS}
    by_camp = {k: [] for k in order}
    for kw, ck in assign.items():
        by_camp[ck].append(terms[kw])
    for f, expr in ((1.00, "close-match"), (0.60, "loose-match"), (0.75, "substitutes"), (0.45, "complements")):
        by_camp["auto"].append(dict(kw=expr, _auto=round(asin_cvr * f, 4)))
    for s in broad:
        by_camp["broad"].append(dict(kw=s, _broad=round(asin_cvr * 0.6, 4)))

    # ---- product targets (PAT)
    own = set(a.strip().upper() for a in (args.own_asins or "").split(",") if a.strip())
    own.discard(asin)
    neg_asins, pat = [], []
    comp_by = {c["asin"]: c for c in comps if c["asin"] != asin}
    pat_neg_clicks = min(10, max(2, int(math.ceil(cfs * 2.0))))
    for a, d in pat_live.items():
        if a == asin:
            continue
        if d["clicks"] >= pat_neg_clicks and d["orders"] == 0:
            neg_asins.append(dict(asin=a, clicks=rint(d["clicks"]), spend=round(d["spend"], 2),
                                  why="PAT rule: %d clicks, 0 orders, $%.2f spent" % (d["clicks"], d["spend"])))
            comp_by.pop(a, None)
        elif a not in comp_by and d["orders"] >= 1:
            comp_by[a] = dict(asin=a, brand="", title="Converting ASIN from your Search Term Report", price=0, rating=0,
                              reviews=0, rev=0)
    max_rev = max([c["rev"] for c in comp_by.values()] + [1.0])
    for c in comp_by.values():
        is_own = c["asin"] in own or (bool(brand_terms) and any(b in (c["brand"] + " " + c["title"]).lower() for b in brand_terms))
        price_adv = max(-1.0, min(1.0, (c["price"] - price) / price)) if c["price"] else 0.0
        weak = max(-1.0, min(1.0, (4.5 - c["rating"]) / 1.0)) if c["rating"] else 0.0
        c["score"] = round(50 * (c["rev"] / max_rev) + 30 * max(0.0, price_adv) + 20 * max(0.0, weak), 1)
        c["own"] = is_own
        est = asin_cvr * (1.0 if is_own else 0.5 * max(0.5, min(1.4, 1 + 0.3 * price_adv + 0.2 * weak)))
        live = pat_live.get(c["asin"])
        if live and live["clicks"] >= PRIOR and live["orders"] > 0:
            c["cvr"], c["cvr_src"] = live["orders"] / live["clicks"], "Live PPC"
        elif live and live["clicks"] > 0:
            c["cvr"], c["cvr_src"] = (live["orders"] + PRIOR * est) / (live["clicks"] + PRIOR), "Blended"
        else:
            c["cvr"], c["cvr_src"] = est, "Modeled"
        c["cvr"] = round(min(c["cvr"], 0.95), 4)
        c["live"] = live or {}
        pat.append(c)
    offs = sorted([c for c in pat if not c["own"]], key=lambda c: -c["score"])
    for i, c in enumerate(offs[:150]):
        by_camp["offd" if i < 25 else "off2" if i < 75 else "off3"].append(c)
    for c in [c for c in pat if c["own"]][:25]:
        by_camp["def"].append(c)
    if not args.sample and not comps:
        notes.append("No competitors file: product-targeting campaigns only use ASINs found in your Search Term Report.")

    # ---- campaigns that actually have targets, and one row per target
    camps, rows = [], []
    for (key, lane, kind, tail, label, tk, match, bidding, tos, pp, weight, cap, proof) in CAMPAIGN_DEFS:
        if not by_camp[key]:
            continue
        camps.append(dict(key=key, name="%s | %s | %s | SP | %s | %s" % (brand, lane, asin, kind, tail),
                          adgroup="%s | %s" % (asin, tail), label=label, seg=SEG_LABEL[key], tk=tk, match=match,
                          bidding=bidding, tos=tos, pp=pp, mult=max(tos, pp) / 100.0, weight=weight, cap=cap, proof=proof,
                          targeting="Auto" if key == "auto" else "Manual"))
    if len(camps) <= 1:
        sys.exit("Nothing to build: no keywords or product targets survived. Check the input files.")

    def add_row(kw, rtype, ck, match, e=None, cvr=None, cvr_src="", ceil=None, extra=None):
        e = e or {}
        row = dict(kw=kw, type=rtype, ck=ck, match=match, vol=(rint(e["vol"]) if e.get("vol") else None),
                   imp_share=e.get("imp_share"), click_share=e.get("click_share"), purch_share=e.get("purch_share"),
                   clicks=rint(e.get("clicks", 0) or 0), spend=round(e.get("spend", 0) or 0, 2),
                   orders=rint(e.get("orders", 0) or 0), sales=round(e.get("sales", 0) or 0, 2),
                   cvr=(cvr if cvr is not None else e.get("cvr")), cvr_src=(cvr_src or e.get("cvr_src", "")),
                   ceil=(ceil if ceil is not None else e.get("ceil")), why=e.get("why", ""))
        if extra:
            row.update(extra)
        rows.append(row)

    for c in camps:
        for e in by_camp[c["key"]]:
            if "_auto" in e:
                add_row(e["kw"], "auto", "auto", "Auto", cvr=e["_auto"], cvr_src="Modeled")
            elif "_broad" in e:
                add_row(e["kw"], "keyword", "broad", "Broad", cvr=e["_broad"], cvr_src="Modeled",
                        ceil=(r2(avg_cpc * 0.8) if avg_cpc else None))
            elif "asin" in e:
                lv = e["live"]
                add_row(e["asin"], "pat", c["key"], "ASIN exact",
                        e=dict(clicks=lv.get("clicks", 0), spend=lv.get("spend", 0), orders=lv.get("orders", 0),
                               sales=lv.get("sales", 0)),
                        cvr=e["cvr"], cvr_src=e["cvr_src"],
                        ceil=(r2(lv["spend"] / lv["clicks"] * CEIL_MULT) if lv.get("clicks") else None),
                        extra=dict(brand=e["brand"], title=e["title"], price=e["price"], rating=e["rating"],
                                   reviews=rint(e["reviews"]), rev=rint(e["rev"]), score=e["score"]))
            else:
                add_row(e["kw"], "keyword", c["key"], match_of[c["key"]], e=e)
    n_targets = len(rows)
    for e in negs:
        add_row(e["kw"], "negative", None, "Negative exact", e=e)
    for e in watch[:40]:
        add_row(e["kw"], "watch", None, "", e=e)
    for i, r in enumerate(rows):
        r["i"] = i

    # ---- negative keyword plan per campaign
    nb_kw = [c["key"] for c in camps if c["key"] in ("auto", "phrase", "broad", "grow", "domin", "skc1", "skc2", "skc3")]
    disc = [k for k in nb_kw if k in ("auto", "phrase", "broad")]
    winners_iso = [r["kw"] for r in rows if r["ck"] in EXACT_KEYS]
    neg_plan = {c["key"]: [] for c in camps}
    for k in nb_kw:
        for b in brand_terms:
            neg_plan[k].append((b, "Negative phrase"))
    for k in disc:
        for e in negs[:200]:
            neg_plan[k].append((e["kw"], "Negative exact"))
        for w in winners_iso[:60]:
            neg_plan[k].append((w, "Negative exact"))

    # ---- bulk plan: row templates. fill_plan() (and the browser) drop the live bids and budgets in.
    start = date.today().strftime("%Y%m%d")
    plan = []

    def brow(c, entity, **kw):
        d = {"Entity": entity, "Campaign ID": c["name"]}
        d.update(kw)
        plan.append(d)

    for c in camps:
        brow(c, "Campaign", **{"Campaign Name": c["name"], "Start Date": start, "Targeting Type": c["targeting"],
                               "State": "paused", "Bidding Strategy": c["bidding"], "_budget": c["key"]})
        for plc, pct in (("Placement Top", c["tos"]), ("Placement Product Page", c["pp"])):
            if pct:
                brow(c, "Bidding Adjustment", **{"Bidding Strategy": c["bidding"], "Placement": plc, "Percentage": str(pct)})
        ag = {"Ad Group ID": c["adgroup"]}
        brow(c, "Ad Group", **dict(ag, **{"Ad Group Name": c["adgroup"], "State": "paused", "_dbid": c["key"]}))
        brow(c, "Product Ad", **dict(ag, **({"State": "paused", "SKU": args.sku} if args.sku else
                                          {"State": "paused", "ASIN": asin})))
        for r in rows:
            if r["ck"] != c["key"]:
                continue
            if r["type"] == "keyword":
                brow(c, "Keyword", **dict(ag, **{"State": "paused", "Keyword Text": r["kw"],
                                                 "Match Type": MT_CSV[r["match"]], "_bid": r["i"]}))
            elif r["type"] == "auto":
                brow(c, "Product Targeting", **dict(ag, **{"State": "paused", "Product Targeting Expression": r["kw"],
                                                           "_bid": r["i"]}))
            else:
                brow(c, "Product Targeting", **dict(ag, **{"State": "paused", "_bid": r["i"],
                                                           "Product Targeting Expression": 'asin="%s"' % r["kw"]}))
        seen = set()
        for term, mt in neg_plan[c["key"]]:
            if (term, mt) in seen:
                continue
            seen.add((term, mt))
            brow(c, "Campaign Negative Keyword", **{"State": "enabled", "Keyword Text": term, "Match Type": MT_CSV[mt]})
        if c["key"] == "auto":
            for na in neg_asins:
                brow(c, "Negative Product Targeting", **dict(ag, **{"State": "enabled",
                                                                    "Product Targeting Expression": 'asin="%s"' % na["asin"]}))
    if not args.sku:
        assumed.append("SKU not given (--sku): the Product Ad rows carry the ASIN. Sellers must add their SKU before upload.")

    return dict(
        version=VERSION, sample=bool(args.sample), brand=brand, asin=asin, sku=args.sku or "", built=date.today().isoformat(),
        days=days, prior=PRIOR, segf=SEG_FACTOR, stage=args.stage, stage_factor=STAGE_FACTOR,
        inp=dict(price=price, cogs=cogs, ref_pct=round(ref_pct * 100, 4), fba=fba, ret=ret, t=round(t * 100, 2),
                 tacos=round(tacos * 100, 2), monthly_sales=monthly_sales, daily_budget=daily_budget,
                 budget_manual=budget_manual),
        asin_cvr=round(asin_cvr, 4), cvr_src=cvr_src, cfs=cfs, neg_clicks=neg_clicks, avg_cpc=round(avg_cpc, 4),
        acct=dict(clicks=rint(tot["clicks"]), spend=round(tot["spend"], 2), orders=rint(tot["orders"]),
                  sales=round(tot["sales"], 2), mo_spend=round(mo_ad_spend, 2), mo_sales=round(mo_ad_sales, 2)),
        camps=camps, rows=rows, n_targets=n_targets, neg_asins=neg_asins, brand_terms=brand_terms, plan=plan,
        hdr=BULK_HDR, notes=notes, assumed=assumed, urls=dict(signup=SIGNUP_URL, call=CALL_URL),
        files=dict(html="%s-hero-asin-blueprint.html" % asin, xlsx="%s-hero-asin-architecture.xlsx" % asin,
                   csv="%s-amazon-bulk-2.0.csv" % asin, md="%s-drppc-strategy-brief.md" % asin),
    )


# ---------------------------------------------------------------------------- the calculator (mirrored in the HTML)
def compute(D, inp=None):
    """Same arithmetic, in the same order, as compute() in the HTML report."""
    inp = inp or D["inp"]
    price = inp["price"]
    ref = r2(price * inp["ref_pct"] / 100.0)
    cm = price - inp["cogs"] - ref - inp["fba"] - inp["ret"]
    be = max(0.0, cm / price) if price > 0 else 0.0
    t = inp["t"] / 100.0
    seg_t = {"gen": t, "skc": min(be, t * D["segf"]["skc"]), "def": t * D["segf"]["def"], "br": t * D["segf"]["br"]}
    camp = {c["key"]: c for c in D["camps"]}
    out_rows, neg_spend, over = [], 0.0, 0.0
    for r in D["rows"]:
        o = dict(i=r["i"])
        cur = r["spend"] / r["clicks"] if r["clicks"] > 0 else None
        o["cur"] = cur
        o["acos"] = r["spend"] / r["sales"] if r["sales"] > 0 else None
        if r["type"] == "negative":
            neg_spend += r["spend"]
            o.update(mx=r2(price * be * r["cvr"]), tg=None, sb=None, tone="red", action="Negate: 3x rule tripped")
        elif r["type"] == "watch":
            o.update(mx=r2(price * be * r["cvr"]), tg=None, sb=None, tone="amber", action="Watch: no orders yet")
        else:
            c = camp[r["ck"]]
            mx = r2(price * be * r["cvr"])
            tg = r2(min(price * seg_t[c["tk"]] * r["cvr"], mx * c["cap"]))
            raw = tg / (1 + c["mult"]) * c["proof"]
            sb = max(0.02, r2(min(raw, r["ceil"]) if r["ceil"] is not None else raw))
            o.update(mx=mx, tg=tg, sb=sb)
            over += max(0.0, r["spend"] - r["sales"] * be)
            if cur is not None:
                d = cur - tg
                if cur > mx and d > 0:
                    o.update(tone="red", action="Overpaying $%.2f a click" % d)
                elif d > 0.005:
                    o.update(tone="amber", action="Trim $%.2f a click to target" % d)
                else:
                    o.update(tone="green", action="Headroom $%.2f a click" % (-d))
            elif r["type"] == "keyword" and (r["vol"] or 0) >= 1000 and r["ck"] in EXACT_KEYS + ("grow",):
                o.update(tone="indigo", action="Untapped SQP gap")
            else:
                o.update(tone="slate", action="Launch at starting bid")
        out_rows.append(o)
    total = inp["daily_budget"]
    wsum = float(sum(c["weight"] for c in D["camps"]))
    camps = []
    for c in D["camps"]:
        mine = [o for o, r in zip(out_rows, D["rows"]) if r["ck"] == c["key"]]
        bids = sorted(o["sb"] for o in mine)
        n = len(bids)
        med = bids[n // 2] if n % 2 else r2((bids[n // 2 - 1] + bids[n // 2]) / 2.0)
        mxs = [o["mx"] for o in mine]
        camps.append(dict(key=c["key"], budget=max(5, rint(total * c["weight"] / wsum)), dbid=med, n=n, bid_min=bids[0],
                          bid_max=bids[-1], max_min=min(mxs), max_max=max(mxs), tacos=seg_t[c["tk"]]))
    k30 = 30.0 / D["days"]
    kw_max = [o["mx"] for o, r in zip(out_rows, D["rows"]) if r["type"] == "keyword" and r["ck"]]
    return dict(ref=ref, cm=cm, be=be, t=t, seg_t=seg_t, rows=out_rows, camps=camps,
                budget_total=sum(c["budget"] for c in camps), bleed_mo=(neg_spend + over) * k30,
                neg_spend_mo=neg_spend * k30, over_mo=over * k30,
                max_lo=(min(kw_max) if kw_max else 0.0), max_hi=(max(kw_max) if kw_max else 0.0),
                implied_tacos=(total * 30.0 / inp["monthly_sales"] if inp["monthly_sales"] else 0.0))


def fill_plan(D, S):
    bid = {o["i"]: o["sb"] for o in S["rows"]}
    cc = {c["key"]: c for c in S["camps"]}
    out = []
    for p in D["plan"]:
        d = dict.fromkeys(BULK_HDR, "")
        d["Product"], d["Operation"] = "Sponsored Products", "Create"
        for k, v in p.items():
            if k == "_budget":
                d["Daily Budget"] = "%.2f" % cc[v]["budget"]
            elif k == "_dbid":
                d["Ad Group Default Bid"] = "%.2f" % cc[v]["dbid"]
            elif k == "_bid":
                d["Bid"] = "%.2f" % bid[v]
            else:
                d[k] = v
        out.append(d)
    return out


def bulk_csv(D, S):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=BULK_HDR, lineterminator="\r\n")
    w.writeheader()
    w.writerows(fill_plan(D, S))
    return buf.getvalue()


# ---------------------------------------------------------------------------- strategy brief (markdown)
def money(v, d=2):
    return "$" + format(v, ",.%df" % d)


def brief_md(D, S):
    i, camp = D["inp"], {c["key"]: c for c in D["camps"]}
    L = []
    a = L.append
    a("# Strategy, Goal & Budget Brief for Dr. PPC")
    a("")
    a("**Hero ASIN:** %s  |  **Brand:** %s  |  **SKU:** %s  |  **Built:** %s" %
      (D["asin"], D["brand"], D["sku"] or "add yours", D["built"]))
    a("")
    a("> Paste everything below the line into the Dr. PPC chat inside Orbit. "
      "Start for $1 with code FOUNDER10: %s" % D["urls"]["signup"])
    a("")
    a("---")
    a("")
    a("Dr. PPC, please take over the Sponsored Products architecture below for my hero ASIN **%s**. Upload it paused, "
      "check it against my live account, then launch and manage it daily inside the limits I set here." % D["asin"])
    a("")
    a("## 1. Unit economics (confirmed by me)")
    a("")
    a("| Item | Value |")
    a("|---|---|")
    a("| Selling price | %s |" % money(i["price"]))
    a("| Landed COGS | %s |" % money(i["cogs"]))
    a("| Amazon referral fee (%s%%) | %s |" % ("%g" % i["ref_pct"], money(S["ref"])))
    a("| FBA fulfillment fee | %s |" % money(i["fba"]))
    a("| Returns and storage per unit | %s |" % money(i["ret"]))
    a("| Profit per unit before ads | %s |" % money(S["cm"]))
    a("| **Break-even ACoS** | **%.1f%%** |" % (S["be"] * 100))
    a("| ASIN conversion rate (%s) | %.1f%% |" % (D["cvr_src"], D["asin_cvr"] * 100))
    a("")
    a("## 2. Goal")
    a("")
    a("- **Target ACoS, non-branded:** %.1f%%  (single-keyword rank campaigns %.1f%%, defensive %.1f%%, branded %.1f%%)" %
      (S["t"] * 100, S["seg_t"]["skc"] * 100, S["seg_t"]["def"] * 100, S["seg_t"]["br"] * 100))
    a("- **Target TACoS:** %g%% of %s total monthly sales" % (i["tacos"], money(i["monthly_sales"], 0)))
    a("- **Hard ceiling:** never bid a keyword above its own Max CPC (the break-even click) unless I approve it.")
    if D["acct"]["spend"]:
        a("- **Where I am today:** %s ad spend and %s ad sales a month (%.1f%% ACoS), average CPC %s." %
          (money(D["acct"]["mo_spend"], 0), money(D["acct"]["mo_sales"], 0),
           (D["acct"]["spend"] / D["acct"]["sales"] * 100 if D["acct"]["sales"] else 0), money(D["avg_cpc"])))
        a("- **Bleed found:** about %s a month (%s on search terms that trip the 3x rule, %s spent above break-even)." %
          (money(S["bleed_mo"], 0), money(S["neg_spend_mo"], 0), money(S["over_mo"], 0)))
    a("")
    a("## 3. Budget")
    a("")
    a("**Total daily budget: %s** (%s a month, %.1f%% TACoS)." %
      (money(S["budget_total"], 0), money(S["budget_total"] * 30, 0), S["implied_tacos"] * 100))
    a("")
    a("| Campaign | Role | Targets | Daily budget | Target ACoS | Starting bids | Max CPC range |")
    a("|---|---|---|---|---|---|---|")
    for c in S["camps"]:
        cd = camp[c["key"]]
        a("| %s | %s | %d | %s | %.1f%% | %s to %s | %s to %s |" %
          (cd["name"].replace("|", "/"), cd["label"], c["n"], money(c["budget"], 0), c["tacos"] * 100,
           money(c["bid_min"]), money(c["bid_max"]), money(c["max_min"]), money(c["max_max"])))
    a("")
    a("## 4. Bid rules, per keyword (not one flat CPC)")
    a("")
    a("- Keyword CVR = live PPC CVR once a keyword has %d clicks, otherwise blended with its SQP CVR or the ASIN CVR." % D["prior"])
    a("- **Max CPC** = price x break-even ACoS x keyword CVR. **Target CPC** = price x target ACoS x keyword CVR.")
    a("- Starting bid = Target CPC / (1 + placement boost), never above 1.25 x the CPC I already pay on that keyword.")
    a("- Per-keyword Max CPC runs from **%s to %s** across this build." % (money(S["max_lo"]), money(S["max_hi"])))
    a("")
    top = sorted([(r, o) for r, o in zip(D["rows"], S["rows"]) if r["type"] == "keyword" and r["ck"]],
                 key=lambda x: -(x[0]["orders"] * 1000 + (x[0]["vol"] or 0) / 1000.0))[:15]
    a("| Keyword | Campaign | Match | CVR used | Max CPC | Target CPC | Starting bid |")
    a("|---|---|---|---|---|---|---|")
    for r, o in top:
        a("| %s | %s | %s | %.1f%% (%s) | %s | %s | %s |" %
          (r["kw"], camp[r["ck"]]["label"], r["match"], r["cvr"] * 100, r["cvr_src"], money(o["mx"]), money(o["tg"]),
           money(o["sb"])))
    a("")
    a("The full list (%d targets) is in `%s` and `%s`." % (D["n_targets"], D["files"]["csv"], D["files"]["xlsx"]))
    a("")
    a("## 5. 3x Negation Shield")
    a("")
    a("- Negate a search term at **%d clicks with 0 orders** (3 x %d clicks per sale), or at **%.0f%% ACoS** (3 x target)." %
      (D["neg_clicks"], D["cfs"], S["t"] * 300))
    a("- Keep branded searches out of every non-branded campaign%s." %
      ((": negative phrase `" + "`, `".join(D["brand_terms"]) + "`") if D["brand_terms"] else ""))
    negs = [r for r in D["rows"] if r["type"] == "negative"]
    if negs:
        a("- Negate these %d search terms on day 1:" % len(negs))
        for r in negs[:25]:
            a("  - `%s`: %s" % (r["kw"], r["why"]))
        if len(negs) > 25:
            a("  - and %d more in the workbook, tab 6" % (len(negs) - 25))
    if D["neg_asins"]:
        a("- Negative product targets: %s" % ", ".join("`%s`" % n["asin"] for n in D["neg_asins"]))
    a("")
    a("## 6. What I want you to do every day")
    a("")
    a("1. Move each bid toward its Target CPC as real clicks come in, and never past its Max CPC.")
    a("2. Harvest converting search terms from Auto, Phrase and Broad into the exact campaigns.")
    a("3. Apply the 3x Negation Shield and tell me what you negated and why.")
    a("4. Reset a top-of-search boost to 0% when that placement runs above target.")
    a("5. Keep total spend inside the daily budget and the TACoS goal above.")
    a("6. Write every change in the Reason Log in plain English so I can read it in the morning.")
    a("")
    if D["assumed"]:
        a("## 7. Numbers I still need to confirm")
        a("")
        for x in D["assumed"]:
            a("- " + x)
        a("")
    a("---")
    a("")
    a("**Two ways to use this build**")
    a("")
    a("1. **Do it myself:** upload `%s` in Amazon Ads > Bulk operations. Everything arrives paused; I enable it and "
      "manage bids by hand." % D["files"]["csv"])
    a("2. **Let Dr. PPC run it:** paste this brief into Dr. PPC inside Orbit. $1 for the first 30 days with code "
      "FOUNDER10, then $299 a month + 2%% of ad spend, month to month: %s" % D["urls"]["signup"])
    a("")
    a("Selling $1M+ a year or spending $50K+ a month on ads? Book a 1-on-1 PPC architecture audit with Nick and the "
      "Full Circle team: %s" % D["urls"]["call"])
    a("")
    return "\n".join(L)


# ---------------------------------------------------------------------------- workbook
INK, LINE = "0F172A", "CBD5E1"
THIN = Side(border_style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_H = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
F_B = Font(name="Calibri", size=10, color=INK)
FILL_H = PatternFill("solid", fgColor=INK)
FILL_IN = PatternFill("solid", fgColor="DBEAFE")
FILL_OUT = PatternFill("solid", fgColor="ECFDF5")
TONE = {"red": "DC2626", "amber": "B45309", "green": "047857", "indigo": "4338CA", "slate": "475569"}
S1, S2, S3, S4 = "Inputs & Summary", "Per-Keyword CPC", "Competitor Targets", "Campaign Blueprint"
S5, S6 = "Bulk 2.0 Upload", "3x Negation Shield"
E = "'%s'!" % S1
USD, USD0, PCT = '"$"#,##0.00', '"$"#,##0', "0.0%"


def _sheet(wb, title, heading, sub, widths, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.sheet_view.showGridLines = False
    ws["A1"] = heading
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=INK)
    ws["A2"] = sub
    ws["A2"].font = Font(name="Calibri", italic=True, size=10, color="475569")
    for n, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(n)].width = w
    return ws


def _header(ws, row, labels):
    ws.row_dimensions[row].height = 30
    for n, text in enumerate(labels, start=1):
        c = ws.cell(row=row, column=n, value=text)
        c.font, c.fill, c.border = F_H, FILL_H, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _put(ws, row, col, v, fmt=None, bold=False, fill=None, color=None, align=None):
    c = ws.cell(row=row, column=col, value=v)
    c.font = Font(name="Calibri", size=10, bold=bold, color=color or INK)
    c.border = BOX
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    c.alignment = Alignment(horizontal=align, vertical="center")
    return c


def write_xlsx(D, S, path):
    wb = Workbook()
    i = D["inp"]
    negs = [r for r in D["rows"] if r["type"] == "negative"]
    kw_rows = [r for r in D["rows"] if r["type"] != "pat"]
    pats = [r for r in D["rows"] if r["type"] == "pat"]
    kw_last = 4 + len(kw_rows)
    neg_total_row = 5 + max(1, len(negs))

    # ---- Tab 1: inputs and summary. Blue cells are the seller's to edit: every other tab recalculates.
    ws = _sheet(wb, S1, "Hero ASIN PPC Architecture: %s (%s)" % (D["brand"], D["asin"]),
                "Blue cells are inputs. Change COGS, fees, Target ACoS or TACoS and every keyword bid and campaign "
                "budget in this workbook recalculates. Built %s." % D["built"], [46, 16, 70], first=True)
    _header(ws, 4, ["Your numbers (edit the blue cells)", "Value", "What it is"])
    inputs = [
        (5, "Selling price", i["price"], USD, "What a shopper pays for one unit."),
        (6, "Landed COGS per unit", i["cogs"], USD, "Product cost plus freight and duty to the Amazon warehouse."),
        (7, "Amazon referral fee %", i["ref_pct"] / 100.0, PCT, "15% in most categories."),
        (8, "FBA fulfillment fee", i["fba"], USD, "Per unit, from your fee preview."),
        (9, "Returns and storage per unit", i["ret"], USD, "Your average cost per unit sold."),
        (10, "Target ACoS, non-branded", i["t"] / 100.0, PCT, "The ad cost you want per ad sale."),
        (11, "Target TACoS", i["tacos"] / 100.0, PCT, "Ad spend as a share of ALL sales, ads plus organic."),
        (12, "Total monthly sales (ads + organic)", i["monthly_sales"], USD0, "For this ASIN, last 30 days."),
        (13, "Total daily ad budget", i["daily_budget"] if i["budget_manual"] else "=MAX(10,ROUND(B12*B11/30,0))", USD0,
         "Your own number." if i["budget_manual"] else "Follows TACoS: monthly sales x Target TACoS / 30. Type a number to override."),
    ]
    for row, label, val, fmt, note in inputs:
        _put(ws, row, 1, label, bold=True)
        _put(ws, row, 2, val, fmt, bold=True, fill=FILL_IN, color="1D4ED8", align="right")
        _put(ws, row, 3, note)
    _header(ws, 15, ["Calculated for you", "Value", "How"])
    calc = [
        (16, "Referral fee per unit", "=ROUND(B5*B7+1E-11,2)", USD, "Price x referral %"),
        (17, "Profit per unit before ads", "=B5-B6-B16-B8-B9", USD, "Price - COGS - referral fee - FBA - returns/storage"),
        (18, "Break-even ACoS", "=IF(B5>0,MAX(0,B17/B5),0)", PCT, "Above this, an ad sale loses money."),
        (19, "Target ACoS: single-keyword rank campaigns", "=MIN(B18,B10*%s)" % SEG_FACTOR["skc"], PCT, "Target x 1.15, capped at break-even."),
        (20, "Target ACoS: defensive product targeting", "=B10*%s" % SEG_FACTOR["def"], PCT, "Target x 0.65"),
        (21, "Target ACoS: branded", "=B10*%s" % SEG_FACTOR["br"], PCT, "Target x 0.50. Branded never blends with non-branded."),
        (22, "ASIN conversion rate", D["asin_cvr"], PCT, D["cvr_src"]),
        (23, "Clicks per sale (rounded up)", "=ROUNDUP(1/B22,0)", "0", "1 / conversion rate"),
        (24, "3x rule: negate at this many clicks, 0 orders", "=3*B23", "0", "3 x clicks per sale"),
        (25, "3x rule: negate at this ACoS", "=3*B10", PCT, "3 x Target ACoS"),
        (26, "TACoS this budget implies", "=IF(B12>0,B13*30/B12,0)", PCT, "Daily budget x 30 / total monthly sales"),
        (27, "Monthly bleed found", "=('%s'!F%d+SUM('%s'!Z5:Z%d)+SUM('%s'!X5:X%d))*30/%d"
         % (S6, neg_total_row, S2, kw_last, S3, max(5, 4 + len(pats)), D["days"]),
         USD0, "Spend on 3x-rule search terms + spend above break-even, scaled to 30 days."),
        (28, "Lowest / highest keyword Max CPC", "=MIN('%s'!V5:V%d)" % (S2, kw_last), USD, "Highest is in the next cell."),
    ]
    for row, label, val, fmt, note in calc:
        _put(ws, row, 1, label, bold=True)
        _put(ws, row, 2, val, fmt, bold=True, fill=FILL_OUT, align="right")
        _put(ws, row, 3, note)
    _put(ws, 28, 3, "=MAX('%s'!V5:V%d)" % (S2, kw_last), USD, bold=True, fill=FILL_OUT, align="left")
    _header(ws, 30, ["Read this first", "", ""])
    lines = D["notes"] + ["CHECK: " + x for x in D["assumed"]] + [
        "Path 1, do it yourself: upload %s in Amazon Ads > Bulk operations. Everything arrives paused." % D["files"]["csv"],
        "Path 2, let Dr. PPC run it: paste %s into Dr. PPC inside Orbit. $1 for the first 30 days, code FOUNDER10." % D["files"]["md"],
        D["urls"]["signup"],
    ]
    for n, text in enumerate(lines, start=31):
        ws.merge_cells(start_row=n, start_column=1, end_row=n, end_column=3)
        c = ws.cell(row=n, column=1, value=text)
        c.font = Font(name="Calibri", size=10, color=("B45309" if text.startswith("CHECK") else INK))
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A5"

    seg_cell = {"gen": E + "$B$10", "skc": E + "$B$19", "def": E + "$B$20", "br": E + "$B$21"}
    camp = {c["key"]: c for c in D["camps"]}
    srow = {o["i"]: o for o in S["rows"]}
    xl_row = {}

    # ---- Tab 2: every keyword, priced on its own
    ws = _sheet(wb, S2, "Per-keyword Max CPC, Target CPC and starting bid",
                "Max CPC = price x break-even ACoS x keyword CVR. Target CPC = price x Target ACoS x keyword CVR. "
                "The green bid columns are formulas tied to tab 1.",
                [38, 10, 14, 26, 14, 13, 11, 11, 9, 11, 9, 11, 11, 9, 9, 11, 10, 8, 9, 8, 11, 11, 11, 12, 34, 13])
    _header(ws, 4, ["Keyword / target", "Type", "Segment", "Campaign", "Match", "SQP volume / mo", "Click share",
                    "Purchase share", "Clicks", "Spend", "Orders", "Sales", "Current CPC", "ACoS", "CVR used", "CVR source",
                    "Target ACoS", "Cap x Max", "Placement boost", "Proof", "Market ceiling", "Max CPC", "Target CPC",
                    "Starting bid", "Action", "Spend above break-even"])
    for n, r in enumerate(kw_rows, start=5):
        xl_row[r["i"]] = (S2, n)
        o = srow[r["i"]]
        is_t = r["ck"] is not None
        c = camp.get(r["ck"])
        tone = TONE[o["tone"]]
        _put(ws, n, 1, r["kw"], bold=True)
        _put(ws, n, 2, "Target" if is_t else ("Negate" if r["type"] == "negative" else "Watch"), align="center",
             color=(INK if is_t else tone))
        _put(ws, n, 3, c["seg"] if c else "3x Shield", align="center")
        _put(ws, n, 4, c["label"] if c else "")
        _put(ws, n, 5, r["match"], align="center")
        _put(ws, n, 6, r["vol"], "#,##0", align="right")
        _put(ws, n, 7, r["click_share"], PCT, align="right")
        _put(ws, n, 8, r["purch_share"], PCT, align="right")
        _put(ws, n, 9, r["clicks"], "#,##0", align="right")
        _put(ws, n, 10, r["spend"], USD, align="right")
        _put(ws, n, 11, r["orders"], "#,##0", align="right")
        _put(ws, n, 12, r["sales"], USD, align="right")
        _put(ws, n, 13, '=IF(I%d>0,J%d/I%d,"")' % (n, n, n), USD, align="right")
        _put(ws, n, 14, '=IF(L%d>0,J%d/L%d,"")' % (n, n, n), PCT, align="right")
        _put(ws, n, 15, r["cvr"], PCT, fill=FILL_IN, color="1D4ED8", align="right")
        _put(ws, n, 16, r["cvr_src"], align="center")
        _put(ws, n, 22, "=ROUND(%s$B$5*%s$B$18*O%d+1E-11,2)" % (E, E, n), USD, bold=True, fill=FILL_OUT, align="right")
        if is_t:
            _put(ws, n, 17, "=" + seg_cell[c["tk"]], PCT, align="right")
            _put(ws, n, 18, c["cap"], "0.00", align="right")
            _put(ws, n, 19, c["mult"], "0%", align="right")
            _put(ws, n, 20, c["proof"], "0.00", align="right")
            _put(ws, n, 21, r["ceil"], USD, align="right")
            _put(ws, n, 23, "=ROUND(MIN(%s$B$5*Q%d*O%d,V%d*R%d)+1E-11,2)" % (E, n, n, n, n), USD, bold=True, fill=FILL_OUT,
                 align="right")
            _put(ws, n, 24, '=MAX(0.02,ROUND(MIN(W%d/(1+S%d)*T%d,IF(U%d="",W%d/(1+S%d)*T%d,U%d))+1E-11,2))' % ((n,) * 8), USD,
                 bold=True, fill=FILL_OUT, align="right")
            _put(ws, n, 26, "=MAX(0,J%d-L%d*%s$B$18)" % (n, n, E), USD, align="right")
        else:
            for col in (17, 18, 19, 20, 21, 23, 24):
                _put(ws, n, col, None)
            _put(ws, n, 26, 0, USD, align="right")
        _put(ws, n, 25, o["action"] if is_t else (r["why"] or o["action"]), color=tone, bold=True)
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = "A4:Z%d" % kw_last

    # ---- Tab 3: competitor product targets
    ws = _sheet(wb, S3, "Competitor and own-catalog product targets (%d)" % len(pats),
                "Offensive = competitor listings where you can win the click. Defensive = your own listings. "
                "The green bid columns are formulas tied to tab 1.",
                [14, 20, 44, 26, 10, 9, 10, 14, 10, 9, 11, 9, 11, 11, 10, 11, 8, 9, 8, 11, 11, 11, 12])
    _header(ws, 4, ["ASIN", "Brand", "Title", "Campaign", "Price", "Rating", "Reviews", "Revenue / mo", "Conquest score",
                    "Clicks", "Spend", "Orders", "Sales", "CVR used", "CVR source", "Target ACoS", "Cap x Max",
                    "Placement boost", "Proof", "Market ceiling", "Max CPC", "Target CPC", "Starting bid",
                    "Spend above break-even"])
    ws.column_dimensions["X"].width = 13
    for n, r in enumerate(pats, start=5):
        xl_row[r["i"]] = (S3, n)
        c = camp[r["ck"]]
        _put(ws, n, 1, r["kw"], bold=True)
        _put(ws, n, 2, r.get("brand", ""))
        _put(ws, n, 3, r.get("title", ""))
        _put(ws, n, 4, c["label"])
        _put(ws, n, 5, r.get("price") or None, USD, align="right")
        _put(ws, n, 6, r.get("rating") or None, "0.0", align="right")
        _put(ws, n, 7, r.get("reviews") or None, "#,##0", align="right")
        _put(ws, n, 8, r.get("rev") or None, USD0, align="right")
        _put(ws, n, 9, r.get("score"), "0.0", align="right")
        _put(ws, n, 10, r["clicks"], "#,##0", align="right")
        _put(ws, n, 11, r["spend"], USD, align="right")
        _put(ws, n, 12, r["orders"], "#,##0", align="right")
        _put(ws, n, 13, r["sales"], USD, align="right")
        _put(ws, n, 14, r["cvr"], PCT, fill=FILL_IN, color="1D4ED8", align="right")
        _put(ws, n, 15, r["cvr_src"], align="center")
        _put(ws, n, 16, "=" + seg_cell[c["tk"]], PCT, align="right")
        _put(ws, n, 17, c["cap"], "0.00", align="right")
        _put(ws, n, 18, c["mult"], "0%", align="right")
        _put(ws, n, 19, c["proof"], "0.00", align="right")
        _put(ws, n, 20, r["ceil"], USD, align="right")
        _put(ws, n, 21, "=ROUND(%s$B$5*%s$B$18*N%d+1E-11,2)" % (E, E, n), USD, bold=True, fill=FILL_OUT, align="right")
        _put(ws, n, 22, "=ROUND(MIN(%s$B$5*P%d*N%d,U%d*Q%d)+1E-11,2)" % (E, n, n, n, n), USD, bold=True, fill=FILL_OUT, align="right")
        _put(ws, n, 23, '=MAX(0.02,ROUND(MIN(V%d/(1+R%d)*S%d,IF(T%d="",V%d/(1+R%d)*S%d,T%d))+1E-11,2))' % ((n,) * 8), USD,
             bold=True, fill=FILL_OUT, align="right")
        _put(ws, n, 24, "=MAX(0,K%d-M%d*%s$B$18)" % (n, n, E), USD, align="right")
    if not pats:
        ws["A5"] = "No product targets in this build. Add a competitor list with --competitors to fill this tab."
        ws["X5"] = 0
    ws.freeze_panes = "B5"

    def bid_ref(idx):
        sh, n = xl_row[idx]
        return "'%s'!$%s$%d" % (sh, "X" if sh == S2 else "W", n)

    # ---- Tab 4: campaign blueprint
    ws = _sheet(wb, S4, "Campaign blueprint: %d campaigns, all created paused" % len(D["camps"]),
                "Budgets are formulas: each campaign's share of the total daily budget on tab 1.",
                [5, 54, 28, 14, 12, 24, 9, 9, 9, 13, 13, 12, 9, 13, 13, 13])
    _header(ws, 4, ["No.", "Campaign name", "Role", "Segment", "Match", "Bidding", "TOS %", "PP %", "Budget weight",
                    "Daily budget", "Monthly budget", "Target ACoS", "Targets", "Default bid", "Lowest bid", "Highest bid"])
    camp_row = {}
    last = 4 + len(D["camps"])
    for n, c in enumerate(D["camps"], start=5):
        camp_row[c["key"]] = n
        mine = [r["i"] for r in D["rows"] if r["ck"] == c["key"]]
        sh, first_n = xl_row[mine[0]]
        col = "X" if sh == S2 else "W"
        rng = "'%s'!$%s$%d:$%s$%d" % (sh, col, first_n, col, first_n + len(mine) - 1)
        _put(ws, n, 1, n - 4, align="center")
        _put(ws, n, 2, c["name"], bold=True)
        _put(ws, n, 3, c["label"])
        _put(ws, n, 4, c["seg"], align="center")
        _put(ws, n, 5, c["match"], align="center")
        _put(ws, n, 6, c["bidding"])
        _put(ws, n, 7, c["tos"] / 100.0, "0%", align="right")
        _put(ws, n, 8, c["pp"] / 100.0, "0%", align="right")
        _put(ws, n, 9, c["weight"], "0", fill=FILL_IN, color="1D4ED8", align="right")
        _put(ws, n, 10, "=MAX(5,ROUND(%s$B$13*I%d/SUM($I$5:$I$%d),0))" % (E, n, last), USD0, bold=True, fill=FILL_OUT,
             align="right")
        _put(ws, n, 11, "=J%d*30" % n, USD0, align="right")
        _put(ws, n, 12, "=" + seg_cell[c["tk"]], PCT, align="right")
        _put(ws, n, 13, len(mine), "0", align="right")
        _put(ws, n, 14, "=ROUND(MEDIAN(%s)+1E-11,2)" % rng, USD, fill=FILL_OUT, align="right")
        _put(ws, n, 15, "=MIN(%s)" % rng, USD, align="right")
        _put(ws, n, 16, "=MAX(%s)" % rng, USD, align="right")
    tr = last + 1
    _put(ws, tr, 2, "TOTAL", bold=True)
    _put(ws, tr, 9, "=SUM(I5:I%d)" % last, "0", bold=True, align="right")
    _put(ws, tr, 10, "=SUM(J5:J%d)" % last, USD0, bold=True, align="right")
    _put(ws, tr, 11, "=SUM(K5:K%d)" % last, USD0, bold=True, align="right")
    _put(ws, tr, 13, "=SUM(M5:M%d)" % last, "0", bold=True, align="right")
    ws.freeze_panes = "C5"

    # ---- Tab 5: bulk upload (a live view of the CSV)
    ws = _sheet(wb, S5, "Amazon Sponsored Products Bulk 2.0 upload (live)",
                "Budgets and bids here follow tab 1. To upload after changing tab 1: copy rows 4 and below, paste as "
                "values into a new sheet, save as CSV. Or use %s." % D["files"]["csv"],
                [18, 26, 10, 46, 22, 9, 7, 9, 12, 46, 22, 11, 9, 12, 9, 12, 20, 13, 12, 9, 34, 16, 24, 22, 11, 26])
    _header(ws, 4, BULK_HDR)
    for n, (p, d) in enumerate(zip(D["plan"], fill_plan(D, S)), start=5):
        for col, h in enumerate(BULK_HDR, start=1):
            v, fmt = d[h], None
            if h == "Daily Budget" and "_budget" in p:
                v, fmt = "='%s'!$J$%d" % (S4, camp_row[p["_budget"]]), "0.00"
            elif h == "Ad Group Default Bid" and "_dbid" in p:
                v, fmt = "='%s'!$N$%d" % (S4, camp_row[p["_dbid"]]), "0.00"
            elif h == "Bid" and "_bid" in p:
                v, fmt = "=" + bid_ref(p["_bid"]), "0.00"
            c = ws.cell(row=n, column=col, value=(v if v != "" else None))
            c.font = Font(name="Calibri", size=10, bold=(d["Entity"] == "Campaign"), color=INK)
            if d["Entity"] == "Campaign":
                c.fill = FILL_OUT
            if fmt:
                c.number_format = fmt
    ws.freeze_panes = "C5"

    # ---- Tab 6: 3x Negation Shield
    ws = _sheet(wb, S6, "3x Negation Shield: %d search terms, %d product targets" % (len(negs), len(D["neg_asins"])),
                "Negate at 3 x clicks-per-sale with 0 orders, or at 3 x Target ACoS. These are already in the bulk "
                "file as campaign negatives.", [40, 16, 11, 11, 11, 13, 13, 13, 60])
    _header(ws, 4, ["Search term / ASIN", "Negative type", "Clicks", "Orders", "ACoS", "Spend", "Sales", "Wasted / mo", "Rule"])
    n = 5
    for r in negs:
        _put(ws, n, 1, r["kw"], bold=True)
        _put(ws, n, 2, "Negative exact", align="center")
        _put(ws, n, 3, r["clicks"], "#,##0", align="right")
        _put(ws, n, 4, r["orders"], "#,##0", align="right")
        _put(ws, n, 5, '=IF(G%d>0,F%d/G%d,"no sales")' % (n, n, n), PCT, align="right")
        _put(ws, n, 6, r["spend"], USD, align="right", color=TONE["red"], bold=True)
        _put(ws, n, 7, r["sales"], USD, align="right")
        _put(ws, n, 8, "=F%d*30/%d" % (n, D["days"]), USD, align="right")
        _put(ws, n, 9, r["why"])
        n += 1
    if not negs:
        _put(ws, n, 1, "No search term trips the 3x rule yet.")
        _put(ws, n, 6, 0, USD)
        n += 1
    assert n == neg_total_row
    _put(ws, n, 1, "TOTAL", bold=True)
    _put(ws, n, 6, "=SUM(F5:F%d)" % (n - 1), USD, bold=True, align="right")
    _put(ws, n, 8, "=F%d*30/%d" % (n, D["days"]), USD, bold=True, align="right")
    n += 2
    if D["neg_asins"]:
        _header(ws, n, ["Negative product target", "Type", "Clicks", "", "", "Spend", "", "", "Rule"])
        for na in D["neg_asins"]:
            n += 1
            _put(ws, n, 1, na["asin"], bold=True)
            _put(ws, n, 2, "Negative ASIN", align="center")
            _put(ws, n, 3, na["clicks"], "#,##0", align="right")
            _put(ws, n, 6, na["spend"], USD, align="right")
            _put(ws, n, 9, na["why"])
        n += 2
    watch = [r for r in D["rows"] if r["type"] == "watch"]
    if watch:
        _header(ws, n, ["On watch (not negated yet)", "", "Clicks", "Orders", "", "Spend", "", "", "Why"])
        for r in watch:
            n += 1
            _put(ws, n, 1, r["kw"])
            _put(ws, n, 3, r["clicks"], "#,##0", align="right")
            _put(ws, n, 4, r["orders"], "#,##0", align="right")
            _put(ws, n, 6, r["spend"], USD, align="right")
            _put(ws, n, 9, r["why"])
    ws.freeze_panes = "B5"
    wb.save(path)


# ---------------------------------------------------------------------------- HTML report
HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>Hero ASIN Blueprint</title>
<style>
:root{--white:#fff;--canvas:#F8FAFC;--line:#E2E8F0;--line2:#CBD5E1;--ink:#0F172A;--ink2:#334155;--mute:#64748B;
--green:#047857;--green-bg:#ECFDF5;--green-ln:#A7F3D0;--amber:#B45309;--amber-bg:#FFFBEB;--amber-ln:#FDE68A;
--red:#DC2626;--red-bg:#FEF2F2;--red-ln:#FECACA;--indigo:#4338CA;--indigo-bg:#EEF2FF;--indigo-ln:#C7D2FE;
--sans:ui-sans-serif,system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
--mono:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace}
*{box-sizing:border-box}
html,body{margin:0;background:var(--canvas);color:var(--ink);font-family:var(--sans);line-height:1.5;-webkit-font-smoothing:antialiased}
h1,h2,h3,p{margin:0}
.wrap{max-width:1400px;margin:0 auto;padding:0 20px}
td.n,th.n{font-family:var(--mono);font-variant-numeric:tabular-nums}
.bar{position:sticky;top:0;z-index:40;background:rgba(255,255,255,.94);backdrop-filter:blur(10px);border-bottom:1px solid var(--line)}
.bar .wrap{display:flex;align-items:center;gap:12px;min-height:60px;flex-wrap:wrap;padding-top:8px;padding-bottom:8px}
.mark{font-weight:700;font-size:14px}.mark small{display:block;font-weight:500;font-size:11px;color:var(--mute)}
.pill{border:1px solid var(--indigo-ln);background:var(--indigo-bg);color:var(--indigo);border-radius:999px;padding:4px 12px;font:600 12.5px var(--mono)}
.sp{flex:1}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;border-radius:10px;padding:9px 14px;font:600 13px var(--sans);
text-decoration:none;border:1px solid var(--line2);background:var(--white);color:var(--ink);cursor:pointer}
.btn:hover{border-color:var(--ink2)}
.btn.go{background:var(--green);border-color:var(--green);color:#fff}
.btn.go:hover{background:#065F46}
.btn.big{padding:13px 18px;font-size:15px;border-radius:12px}
.hero{padding:36px 0 18px}
.eyebrow{font:600 12px var(--sans);letter-spacing:.09em;text-transform:uppercase;color:var(--indigo)}
h1{font-size:34px;line-height:1.12;letter-spacing:-.02em;font-weight:800;margin-top:8px;max-width:900px}
.lede{color:var(--ink2);font-size:16px;max-width:860px;margin-top:12px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-top:22px}
.kpi{background:var(--white);border:1px solid var(--line);border-radius:14px;padding:14px 16px}
.kpi b{display:block;font:700 24px var(--mono);letter-spacing:-.02em;margin-top:2px}
.kpi span{font-size:12px;color:var(--mute);font-weight:600;text-transform:uppercase;letter-spacing:.05em}
.kpi small{display:block;color:var(--mute);font-size:12px;margin-top:2px}
.kpi.red b{color:var(--red)}.kpi.green b{color:var(--green)}.kpi.indigo b{color:var(--indigo)}
section{padding:26px 0}
.snum{font:700 12px var(--mono);color:var(--indigo);letter-spacing:.06em}
h2{font-size:23px;letter-spacing:-.02em;font-weight:700;margin-top:2px}
.sub{color:var(--ink2);font-size:14.5px;max-width:900px;margin-top:6px}
.card{background:var(--white);border:1px solid var(--line);border-radius:14px;padding:18px;margin-top:16px}
.calc{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}
.calc label{display:block;font-size:12px;font-weight:600;color:var(--ink2)}
.calc input{width:100%;margin-top:4px;border:1px solid var(--indigo-ln);background:var(--indigo-bg);color:#1D4ED8;border-radius:9px;
padding:9px 10px;font:600 15px var(--mono)}
.calc input:focus{outline:2px solid var(--indigo);outline-offset:1px}
.calc small{color:var(--mute);font-size:11.5px}
.outs{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-top:16px;padding-top:16px;border-top:1px dashed var(--line2)}
.out span{font-size:12px;font-weight:600;color:var(--mute)}.out b{display:block;font:700 19px var(--mono)}
.row{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:14px}
.chip{border:1px solid var(--line2);background:var(--white);border-radius:999px;padding:6px 12px;font:600 12.5px var(--sans);cursor:pointer;color:var(--ink2)}
.chip.on{background:var(--ink);border-color:var(--ink);color:#fff}
.search{border:1px solid var(--line2);border-radius:9px;padding:8px 10px;font:500 13.5px var(--sans);min-width:220px}
.tw{overflow-x:auto;border:1px solid var(--line);border-radius:14px;background:var(--white);margin-top:14px}
table{border-collapse:collapse;width:100%;font-size:13px}
th{background:var(--canvas);text-align:left;font:600 11px var(--sans);letter-spacing:.04em;text-transform:uppercase;
color:var(--mute);padding:10px;border-bottom:1px solid var(--line);white-space:nowrap}
td{padding:9px 10px;border-bottom:1px solid var(--line);vertical-align:middle}
td.n,th.n{text-align:right;white-space:nowrap}
tr:last-child td{border-bottom:0}
td.k{font-weight:600;min-width:210px}
td.k small{display:block;font-weight:500;color:var(--mute);font-size:11.5px}
.tag{display:inline-block;border-radius:6px;padding:2px 7px;font:600 11px var(--sans);border:1px solid var(--line);background:var(--canvas);color:var(--ink2);white-space:nowrap}
.tag.red{color:var(--red);background:var(--red-bg);border-color:var(--red-ln)}
.tag.amber{color:var(--amber);background:var(--amber-bg);border-color:var(--amber-ln)}
.tag.green{color:var(--green);background:var(--green-bg);border-color:var(--green-ln)}
.tag.indigo{color:var(--indigo);background:var(--indigo-bg);border-color:var(--indigo-ln)}
.hl{background:var(--green-bg);font-weight:700}
.two{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px;margin-top:16px}
.path{background:var(--white);border:1px solid var(--line);border-radius:14px;padding:20px}
.path.best{border:2px solid var(--green)}
.path h3{font-size:18px}.path ol{margin:10px 0 14px;padding-left:20px;color:var(--ink2);font-size:14.5px}
.path li{margin:4px 0}
.note{background:var(--amber-bg);border:1px solid var(--amber-ln);color:#78350F;border-radius:12px;padding:12px 14px;font-size:13.5px;margin-top:14px}
.note ul{margin:6px 0 0;padding-left:18px}
.muted{color:var(--mute);font-size:12.5px}
.toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:var(--ink);color:#fff;padding:10px 16px;border-radius:10px;
font:600 13.5px var(--sans);opacity:0;pointer-events:none;transition:opacity .2s;z-index:60}
.toast.on{opacity:1}
textarea#briefBox{width:100%;height:220px;margin-top:12px;border:1px solid var(--line2);border-radius:10px;padding:10px;font:12.5px var(--mono);display:none}
footer{padding:26px 0 50px;color:var(--mute);font-size:12.5px}
@media (max-width:640px){h1{font-size:26px}.wrap{padding:0 16px}.kpi b{font-size:21px}}
</style>
</head>
<body>
<div class="bar"><div class="wrap">
  <div class="mark">Hero ASIN Blueprint<small>Full Circle &times; Dr. PPC</small></div>
  <span class="pill" id="asinPill"></span>
  <span class="sp"></span>
  <button class="btn" id="dlCsvTop">Download Bulk 2.0 CSV</button>
  <button class="btn go" id="copyTop">Copy Strategy, Goal &amp; Budget Brief for Dr. PPC</button>
</div></div>

<div class="wrap">
<div class="hero">
  <div class="eyebrow" id="eyebrow"></div>
  <h1 id="h1"></h1>
  <p class="lede" id="lede"></p>
  <div class="kpis" id="kpis"></div>
  <div id="assumed"></div>
</div>

<section>
  <div class="snum">01 &middot; LIVE CALCULATOR</div>
  <h2>Confirm your COGS, Amazon fees, Target ACoS and TACoS</h2>
  <p class="sub">Change any blue box. Every keyword's Max CPC, Target CPC and starting bid, every campaign budget, the bulk file and the Dr. PPC brief recalculate right here in your browser.</p>
  <div class="card">
    <div class="calc">
      <div><label>Selling price ($)<input id="i_price" type="number" step="0.01" min="0"></label></div>
      <div><label>Landed COGS ($)<input id="i_cogs" type="number" step="0.01" min="0"></label></div>
      <div><label>Referral fee (%)<input id="i_ref_pct" type="number" step="0.1" min="0"></label></div>
      <div><label>FBA fee ($)<input id="i_fba" type="number" step="0.01" min="0"></label></div>
      <div><label>Returns + storage ($)<input id="i_ret" type="number" step="0.01" min="0"></label></div>
      <div><label>Target ACoS (%)<input id="i_t" type="number" step="0.5" min="1"></label></div>
      <div><label>Target TACoS (%)<input id="i_tacos" type="number" step="0.5" min="0"></label><small>sets the daily budget</small></div>
      <div><label>Total monthly sales ($)<input id="i_monthly_sales" type="number" step="100" min="0"></label><small>ads + organic</small></div>
      <div><label>Daily ad budget ($)<input id="i_daily_budget" type="number" step="1" min="1"></label><small id="budgetNote"></small></div>
    </div>
    <div class="row">
      <span class="muted">Target ACoS presets:</span>
      <button class="chip" data-stage="profit">Profit (42% of break-even)</button>
      <button class="chip" data-stage="balanced">Balanced (56%)</button>
      <button class="chip" data-stage="scale">Scale (75%)</button>
      <button class="chip" id="reset">Reset to my numbers</button>
    </div>
    <div class="outs" id="outs"></div>
  </div>
</section>

<section>
  <div class="snum">02 &middot; PER-KEYWORD BIDS</div>
  <h2>Every keyword priced on its own, not one flat CPC</h2>
  <p class="sub">Max CPC is the most a click can cost before that keyword loses money. Target CPC hits your Target ACoS. The starting bid is what goes in the bulk file.</p>
  <div class="row">
    <button class="chip on" data-f="all">All targets</button>
    <button class="chip" data-f="red">Overpaying</button>
    <button class="chip" data-f="amber">Trim</button>
    <button class="chip" data-f="green">Headroom</button>
    <button class="chip" data-f="indigo">Untapped SQP gaps</button>
    <button class="chip" data-f="neg">Negate (3x rule)</button>
    <input class="search" id="q" placeholder="Search keywords">
    <span class="muted" id="count"></span>
  </div>
  <div class="tw"><table>
    <thead><tr><th>Keyword / target</th><th>Campaign</th><th class="n">SQP vol / mo</th><th class="n">Click share</th><th class="n">Clicks</th>
    <th class="n">Orders</th><th class="n">CPC now</th><th class="n">ACoS</th><th class="n">CVR used</th><th class="n">Max CPC</th>
    <th class="n">Target CPC</th><th class="n">Starting bid</th><th>Action</th></tr></thead>
    <tbody id="kwBody"></tbody>
  </table></div>
</section>

<section>
  <div class="snum">03 &middot; CAMPAIGN BLUEPRINT</div>
  <h2 id="campH2"></h2>
  <p class="sub">Branded and non-branded never share a campaign. Every campaign, ad group and target is created paused.</p>
  <div class="tw"><table>
    <thead><tr><th>Campaign</th><th>Role</th><th>Match</th><th class="n">TOS / PP boost</th><th class="n">Targets</th><th class="n">Target ACoS</th>
    <th class="n">Starting bids</th><th class="n">Max CPC range</th><th class="n">Daily budget</th></tr></thead>
    <tbody id="campBody"></tbody>
  </table></div>
</section>

<section>
  <div class="snum">04 &middot; 3X NEGATION SHIELD</div>
  <h2 id="negH2"></h2>
  <p class="sub" id="negSub"></p>
  <div class="tw"><table>
    <thead><tr><th>Search term / ASIN</th><th class="n">Clicks</th><th class="n">Orders</th><th class="n">Spend</th><th>Rule</th></tr></thead>
    <tbody id="negBody"></tbody>
  </table></div>
</section>

<section>
  <div class="snum">05 &middot; LAUNCH IT</div>
  <h2>Two ways to put this live</h2>
  <div class="two">
    <div class="path">
      <h3>Path 1: Do it yourself</h3>
      <ol>
        <li>Download the Bulk 2.0 CSV (it uses the numbers in the calculator above).</li>
        <li>Amazon Ads &rarr; Bulk operations &rarr; Upload.</li>
        <li>Everything arrives <b>paused</b>. Review, then enable.</li>
        <li>Check bids, harvest and negate every morning.</li>
      </ol>
      <button class="btn big" id="dlCsv">Download Bulk 2.0 CSV</button>
    </div>
    <div class="path best">
      <h3>Path 2: Let Dr. PPC upload and manage it daily for $1</h3>
      <ol>
        <li>Copy the Strategy, Goal &amp; Budget Brief (one click).</li>
        <li>Start Dr. PPC inside Orbit: <b>$1 for the first 30 days</b> with code <b>FOUNDER10</b>.</li>
        <li>Paste the brief into the Dr. PPC chat.</li>
        <li>Read the plain-English Reason Log each morning.</li>
      </ol>
      <button class="btn go big" id="copyBrief">Copy Strategy, Goal &amp; Budget Brief for Dr. PPC</button>
      <a class="btn big" id="signup" target="_blank" rel="noopener" style="margin-top:8px">Start Dr. PPC for $1 &rarr;</a>
      <p class="muted" style="margin-top:10px">Then $299 a month + 2% of ad spend, month to month. FOUNDER10 = 10% off for life.</p>
      <textarea id="briefBox" readonly></textarea>
    </div>
  </div>
  <p class="sub" style="margin-top:16px">Selling $1M+ a year or spending $50K+ a month on ads? <a id="call" target="_blank" rel="noopener">Book a 1-on-1 PPC architecture audit with Nick and the Full Circle team &rarr;</a></p>
</section>

<footer id="foot"></footer>
</div>
<div class="toast" id="toast"></div>

<script id="data" type="application/json">__DATA__</script>
<script id="calc">
/* Pure calculator: it never touches the page, so it can be run and checked outside a browser.
   Mirrors compute(), bulk_csv() and the brief in build_campaign_sheet.py. */
function r2(x){return Math.floor(x*100+0.5+1e-9)/100}
function rint(x){return Math.floor(x+0.5)}
function compute(D,inp){
  var price=inp.price,ref=r2(price*inp.ref_pct/100),cm=price-inp.cogs-ref-inp.fba-inp.ret;
  var be=price>0?Math.max(0,cm/price):0,t=inp.t/100;
  var segT={gen:t,skc:Math.min(be,t*D.segf.skc),def:t*D.segf.def,br:t*D.segf.br};
  var camp={};D.camps.forEach(function(c){camp[c.key]=c});
  var negSpend=0,over=0,EX=['grow','skc1','skc2','skc3','domin'];
  var rows=D.rows.map(function(r){
    var o={i:r.i},cur=r.clicks>0?r.spend/r.clicks:null;
    o.cur=cur;o.acos=r.sales>0?r.spend/r.sales:null;
    if(r.type==='negative'){negSpend+=r.spend;o.mx=r2(price*be*r.cvr);o.tg=null;o.sb=null;o.tone='red';o.action='Negate: 3x rule tripped'}
    else if(r.type==='watch'){o.mx=r2(price*be*r.cvr);o.tg=null;o.sb=null;o.tone='amber';o.action='Watch: no orders yet'}
    else{
      var c=camp[r.ck],mx=r2(price*be*r.cvr),tg=r2(Math.min(price*segT[c.tk]*r.cvr,mx*c.cap));
      var raw=tg/(1+c.mult)*c.proof,hasCeil=r.ceil!==null&&r.ceil!==undefined;
      var sb=Math.max(0.02,r2(hasCeil?Math.min(raw,r.ceil):raw));
      o.mx=mx;o.tg=tg;o.sb=sb;over+=Math.max(0,r.spend-r.sales*be);
      if(cur!==null){var d=cur-tg;
        if(cur>mx&&d>0){o.tone='red';o.action='Overpaying $'+d.toFixed(2)+' a click'}
        else if(d>0.005){o.tone='amber';o.action='Trim $'+d.toFixed(2)+' a click to target'}
        else{o.tone='green';o.action='Headroom $'+(-d).toFixed(2)+' a click'}}
      else if(r.type==='keyword'&&(r.vol||0)>=1000&&EX.indexOf(r.ck)>=0){o.tone='indigo';o.action='Untapped SQP gap'}
      else{o.tone='slate';o.action='Launch at starting bid'}
    }
    return o});
  var total=inp.daily_budget,wsum=0;D.camps.forEach(function(c){wsum+=c.weight});
  var camps=D.camps.map(function(c){
    var bids=[],mxs=[];
    D.rows.forEach(function(r,n){if(r.ck===c.key){bids.push(rows[n].sb);mxs.push(rows[n].mx)}});
    bids.sort(function(a,b){return a-b});
    var n=bids.length,med=n%2?bids[(n-1)/2]:r2((bids[n/2-1]+bids[n/2])/2);
    return {key:c.key,budget:Math.max(5,rint(total*c.weight/wsum)),dbid:med,n:n,bid_min:bids[0],bid_max:bids[n-1],
            max_min:Math.min.apply(null,mxs),max_max:Math.max.apply(null,mxs),tacos:segT[c.tk]}});
  var k30=30/D.days,kwMax=[];
  D.rows.forEach(function(r,n){if(r.type==='keyword'&&r.ck)kwMax.push(rows[n].mx)});
  var bt=0;camps.forEach(function(c){bt+=c.budget});
  return {ref:ref,cm:cm,be:be,t:t,seg_t:segT,rows:rows,camps:camps,budget_total:bt,bleed_mo:(negSpend+over)*k30,
          neg_spend_mo:negSpend*k30,over_mo:over*k30,max_lo:kwMax.length?Math.min.apply(null,kwMax):0,
          max_hi:kwMax.length?Math.max.apply(null,kwMax):0,implied_tacos:inp.monthly_sales?total*30/inp.monthly_sales:0}}
function csvCell(v){v=String(v);return /[",\r\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v}
function bulkCsv(D,S){
  var bid={},cc={};S.rows.forEach(function(o){bid[o.i]=o.sb});S.camps.forEach(function(c){cc[c.key]=c});
  var lines=[D.hdr.map(csvCell).join(',')];
  D.plan.forEach(function(p){
    var d={};D.hdr.forEach(function(h){d[h]=''});d.Product='Sponsored Products';d.Operation='Create';
    Object.keys(p).forEach(function(k){
      if(k==='_budget')d['Daily Budget']=cc[p[k]].budget.toFixed(2);
      else if(k==='_dbid')d['Ad Group Default Bid']=cc[p[k]].dbid.toFixed(2);
      else if(k==='_bid')d['Bid']=bid[p[k]].toFixed(2);
      else d[k]=p[k]});
    lines.push(D.hdr.map(function(h){return csvCell(d[h])}).join(','))});
  return lines.join('\r\n')+'\r\n'}
function usd(v,d){d=d===undefined?2:d;return '$'+Number(v).toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d})}
function pct(v,d){return (v*100).toFixed(d===undefined?1:d)+'%'}
function briefText(D,S,inp){
  var camp={};D.camps.forEach(function(c){camp[c.key]=c});var L=[];
  L.push('STRATEGY, GOAL & BUDGET BRIEF FOR DR. PPC');
  L.push('Hero ASIN: '+D.asin+' | Brand: '+D.brand+' | SKU: '+(D.sku||'add yours'));
  L.push('');
  L.push('Dr. PPC, please take over the Sponsored Products architecture below for my hero ASIN '+D.asin+'. Upload it paused, check it against my live account, then launch and manage it daily inside the limits I set here.');
  L.push('');
  L.push('1. UNIT ECONOMICS (confirmed by me)');
  L.push('- Selling price: '+usd(inp.price));
  L.push('- Landed COGS: '+usd(inp.cogs));
  L.push('- Amazon referral fee ('+inp.ref_pct+'%): '+usd(S.ref));
  L.push('- FBA fulfillment fee: '+usd(inp.fba));
  L.push('- Returns and storage per unit: '+usd(inp.ret));
  L.push('- Profit per unit before ads: '+usd(S.cm));
  L.push('- Break-even ACoS: '+pct(S.be));
  L.push('- ASIN conversion rate: '+pct(D.asin_cvr)+' ('+D.cvr_src+')');
  L.push('');
  L.push('2. GOAL');
  L.push('- Target ACoS, non-branded: '+pct(S.t)+' (single-keyword rank campaigns '+pct(S.seg_t.skc)+', defensive '+pct(S.seg_t.def)+', branded '+pct(S.seg_t.br)+')');
  L.push('- Target TACoS: '+inp.tacos+'% of '+usd(inp.monthly_sales,0)+' total monthly sales');
  L.push('- Hard ceiling: never bid a keyword above its own Max CPC (the break-even click) unless I approve it.');
  if(D.acct.spend){
    L.push('- Where I am today: '+usd(D.acct.mo_spend,0)+' ad spend and '+usd(D.acct.mo_sales,0)+' ad sales a month ('+pct(D.acct.sales?D.acct.spend/D.acct.sales:0)+' ACoS), average CPC '+usd(D.avg_cpc)+'.');
    L.push('- Bleed found: about '+usd(S.bleed_mo,0)+' a month ('+usd(S.neg_spend_mo,0)+' on search terms that trip the 3x rule, '+usd(S.over_mo,0)+' spent above break-even).')}
  L.push('');
  L.push('3. BUDGET');
  L.push('Total daily budget: '+usd(S.budget_total,0)+' ('+usd(S.budget_total*30,0)+' a month, '+pct(S.implied_tacos)+' TACoS).');
  S.camps.forEach(function(c){var cd=camp[c.key];
    L.push('- '+cd.name+' | '+cd.label+' | '+c.n+' targets | '+usd(c.budget,0)+'/day | target ACoS '+pct(c.tacos)+' | starting bids '+usd(c.bid_min)+' to '+usd(c.bid_max)+' | Max CPC '+usd(c.max_min)+' to '+usd(c.max_max))});
  L.push('');
  L.push('4. BID RULES, PER KEYWORD (not one flat CPC)');
  L.push('- Keyword CVR = live PPC CVR once a keyword has '+D.prior+' clicks, otherwise blended with its SQP CVR or the ASIN CVR.');
  L.push('- Max CPC = price x break-even ACoS x keyword CVR. Target CPC = price x target ACoS x keyword CVR.');
  L.push('- Starting bid = Target CPC / (1 + placement boost), never above 1.25 x the CPC I already pay on that keyword.');
  L.push('- Per-keyword Max CPC runs from '+usd(S.max_lo)+' to '+usd(S.max_hi)+' across this build.');
  var top=[];D.rows.forEach(function(r,n){if(r.type==='keyword'&&r.ck)top.push([r,S.rows[n]])});
  top.sort(function(a,b){return (b[0].orders*1000+(b[0].vol||0)/1000)-(a[0].orders*1000+(a[0].vol||0)/1000)});
  top.slice(0,15).forEach(function(x){var r=x[0],o=x[1];
    L.push('- '+r.kw+' | '+camp[r.ck].label+' | '+r.match+' | CVR '+pct(r.cvr)+' ('+r.cvr_src+') | Max CPC '+usd(o.mx)+' | Target CPC '+usd(o.tg)+' | Starting bid '+usd(o.sb))});
  L.push('The full list ('+D.n_targets+' targets) is in my Bulk 2.0 file.');
  L.push('');
  L.push('5. 3X NEGATION SHIELD');
  L.push('- Negate a search term at '+D.neg_clicks+' clicks with 0 orders (3 x '+D.cfs+' clicks per sale), or at '+(S.t*300).toFixed(0)+'% ACoS (3 x target).');
  L.push('- Keep branded searches out of every non-branded campaign'+(D.brand_terms.length?': negative phrase "'+D.brand_terms.join('", "')+'"':'')+'.');
  var negs=D.rows.filter(function(r){return r.type==='negative'});
  if(negs.length){L.push('- Negate these '+negs.length+' search terms on day 1:');
    negs.slice(0,25).forEach(function(r){L.push('  - '+r.kw+': '+r.why)});
    if(negs.length>25)L.push('  - and '+(negs.length-25)+' more in my workbook')}
  if(D.neg_asins.length)L.push('- Negative product targets: '+D.neg_asins.map(function(n){return n.asin}).join(', '));
  L.push('');
  L.push('6. WHAT I WANT YOU TO DO EVERY DAY');
  L.push('1. Move each bid toward its Target CPC as real clicks come in, and never past its Max CPC.');
  L.push('2. Harvest converting search terms from Auto, Phrase and Broad into the exact campaigns.');
  L.push('3. Apply the 3x Negation Shield and tell me what you negated and why.');
  L.push('4. Reset a top-of-search boost to 0% when that placement runs above target.');
  L.push('5. Keep total spend inside the daily budget and the TACoS goal above.');
  L.push('6. Write every change in the Reason Log in plain English so I can read it in the morning.');
  return L.join('\n')}
</script>
<script>
(function(){
  var D=JSON.parse(document.getElementById('data').textContent);
  var KEYS=['price','cogs','ref_pct','fba','ret','t','tacos','monthly_sales','daily_budget'];
  var inp=JSON.parse(JSON.stringify(D.inp)),manual=!!D.inp.budget_manual,S,filter='all';
  var $=function(id){return document.getElementById(id)};
  function esc(s){return String(s===null||s===undefined?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
  function toast(t){var e=$('toast');e.textContent=t;e.classList.add('on');setTimeout(function(){e.classList.remove('on')},2200)}
  function setInputs(){KEYS.forEach(function(k){$('i_'+k).value=inp[k]})}
  function read(changed){
    KEYS.forEach(function(k){var v=parseFloat($('i_'+k).value);inp[k]=isFinite(v)&&v>=0?v:0});
    if(changed==='daily_budget')manual=true;
    if(changed==='tacos'||changed==='monthly_sales')manual=false;
    if(!manual){inp.daily_budget=Math.max(10,rint(inp.monthly_sales*inp.tacos/100/30));$('i_daily_budget').value=inp.daily_budget}
    render()}
  function render(){
    S=compute(D,inp);
    var camp={};D.camps.forEach(function(c){camp[c.key]=c});
    $('budgetNote').textContent=manual?'your own number ('+pct(S.implied_tacos)+' TACoS)':'follows TACoS';
    $('outs').innerHTML=[
      ['Referral fee / unit',usd(S.ref)],['Profit / unit before ads',usd(S.cm)],['Break-even ACoS',pct(S.be)],
      ['Profit left at Target ACoS',pct(S.be-S.t)+' of price'],['Max CPC range',usd(S.max_lo)+' to '+usd(S.max_hi)],
      ['Total daily budget',usd(S.budget_total,0)]].map(function(x){return '<div class="out"><span>'+x[0]+'</span><b>'+x[1]+'</b></div>'}).join('')
      +(S.be<=0?'<div class="note" style="grid-column:1/-1">At these costs the product loses money before ads. Check price, COGS and fees.</div>':'')
      +(S.t>S.be&&S.be>0?'<div class="note" style="grid-column:1/-1">Your Target ACoS is above break-even: every ad sale at target loses money. That can be a deliberate rank push, but know it.</div>':'');
    $('kpis').innerHTML=[
      ['red','Monthly bleed found',usd(S.bleed_mo,0),'3x-rule terms + spend above break-even'],
      ['indigo','Per-keyword Max CPC',usd(S.max_lo)+' to '+usd(S.max_hi),'one number per keyword, not one flat CPC'],
      ['','Break-even ACoS',pct(S.be),'target '+pct(S.t)],
      ['green','Ready to launch',D.camps.length+' campaigns','all paused · '+D.n_targets+' targets · '+usd(S.budget_total,0)+' a day']
    ].map(function(k){return '<div class="kpi '+k[0]+'"><span>'+k[1]+'</span><b>'+k[2]+'</b><small>'+k[3]+'</small></div>'}).join('');
    var q=$('q').value.trim().toLowerCase(),shown=0,html='';
    D.rows.forEach(function(r,n){
      var o=S.rows[n],isT=!!r.ck;
      if(filter==='neg'){if(isT)return}
      else if(filter==='all'){if(!isT)return}
      else if(!isT||o.tone!==filter)return;
      if(q&&r.kw.toLowerCase().indexOf(q)<0)return;
      shown++;
      var sub=r.type==='pat'?esc((r.brand?r.brand+' · ':'')+(r.title||'')):esc(r.match);
      html+='<tr><td class="k">'+esc(r.kw)+'<small>'+sub+'</small></td><td>'+(isT?esc(camp[r.ck].label):'<span class="tag red">3x Shield</span>')+'</td>'
        +'<td class="n">'+(r.vol?r.vol.toLocaleString('en-US'):'')+'</td><td class="n">'+(r.click_share!==null&&r.click_share!==undefined?pct(r.click_share):'')+'</td>'
        +'<td class="n">'+(r.clicks||'')+'</td><td class="n">'+(r.clicks?r.orders:'')+'</td><td class="n">'+(o.cur!==null?usd(o.cur):'')+'</td>'
        +'<td class="n">'+(o.acos!==null?pct(o.acos,0):'')+'</td><td class="n">'+pct(r.cvr)+' <span class="tag">'+esc(r.cvr_src)+'</span></td>'
        +'<td class="n hl">'+usd(o.mx)+'</td><td class="n">'+(o.tg!==null?usd(o.tg):'')+'</td><td class="n hl">'+(o.sb!==null?usd(o.sb):'')+'</td>'
        +'<td><span class="tag '+o.tone+'">'+esc(isT?o.action:(r.why||o.action))+'</span></td></tr>'});
    $('kwBody').innerHTML=html||'<tr><td colspan="13" class="muted">Nothing matches this filter.</td></tr>';
    $('count').textContent=shown+' shown';
    $('campBody').innerHTML=S.camps.map(function(c){var cd=camp[c.key];
      return '<tr><td class="k">'+esc(cd.name)+'</td><td>'+esc(cd.label)+'</td><td>'+esc(cd.match)+'</td><td class="n">'+cd.tos+'% / '+cd.pp+'%</td>'
        +'<td class="n">'+c.n+'</td><td class="n">'+pct(c.tacos)+'</td><td class="n">'+usd(c.bid_min)+' to '+usd(c.bid_max)+'</td>'
        +'<td class="n">'+usd(c.max_min)+' to '+usd(c.max_max)+'</td><td class="n hl">'+usd(c.budget,0)+'</td></tr>'}).join('')
      +'<tr><td class="k">Total</td><td></td><td></td><td></td><td class="n">'+D.n_targets+'</td><td></td><td></td><td></td><td class="n hl">'+usd(S.budget_total,0)+'</td></tr>';
    $('campH2').textContent=D.camps.length+' campaigns, '+D.n_targets+' targets, '+usd(S.budget_total,0)+' a day';
    if($('briefBox').style.display==='block')$('briefBox').value=briefText(D,S,inp)}
  function download(){
    var blob=new Blob([bulkCsv(D,S)],{type:'text/csv;charset=utf-8'}),a=document.createElement('a');
    a.href=URL.createObjectURL(blob);a.download=D.files.csv;document.body.appendChild(a);a.click();
    setTimeout(function(){URL.revokeObjectURL(a.href);a.remove()},500);toast('Bulk 2.0 CSV downloaded: everything is paused')}
  function copy(){
    var text=briefText(D,S,inp),box=$('briefBox');box.style.display='block';box.value=text;
    function fallback(){box.focus();box.select();try{document.execCommand('copy');toast('Brief copied. Paste it into Dr. PPC.')}catch(e){toast('Select the text below and copy it')}}
    if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(text).then(function(){toast('Brief copied. Paste it into Dr. PPC.')},fallback)}else fallback()}
  document.title='Hero ASIN Blueprint · '+D.brand+' '+D.asin;
  $('asinPill').textContent=D.asin;
  $('eyebrow').textContent=(D.sample?'Sample build · ':'')+D.brand+' · built '+D.built;
  $('h1').textContent='Your hero ASIN, rebuilt keyword by keyword.';
  $('lede').textContent='Built from your Search Query Performance and Search Term data'+(D.acct.clicks?' ('+D.acct.clicks.toLocaleString('en-US')+' ad clicks over '+D.days+' days)':'')
    +'. ASIN conversion rate '+pct(D.asin_cvr)+': '+D.cfs+' clicks per sale, so a search term is negated at '+D.neg_clicks+' clicks with no order.';
  $('assumed').innerHTML=D.assumed.length?'<div class="note"><b>Check these numbers before you upload:</b><ul>'+D.assumed.map(function(x){return '<li>'+esc(x)+'</li>'}).join('')+'</ul></div>':'';
  var negs=D.rows.filter(function(r){return r.type==='negative'}),watch=D.rows.filter(function(r){return r.type==='watch'});
  $('negH2').textContent=negs.length+' search terms and '+D.neg_asins.length+' product targets to negate on day 1';
  $('negSub').textContent='Rule: '+D.neg_clicks+' clicks with 0 orders (3 x '+D.cfs+' clicks per sale), or 3 x your Target ACoS. Already in the bulk file as campaign negatives.'+(watch.length?' '+watch.length+' more terms are on watch.':'');
  var negHtml=negs.map(function(r){return '<tr><td class="k">'+esc(r.kw)+'</td><td class="n">'+r.clicks+'</td><td class="n">'+r.orders+'</td><td class="n">'+usd(r.spend)+'</td><td>'+esc(r.why)+'</td></tr>'}).join('')
    +D.neg_asins.map(function(n){return '<tr><td class="k">'+esc(n.asin)+' <span class="tag">ASIN</span></td><td class="n">'+n.clicks+'</td><td class="n">0</td><td class="n">'+usd(n.spend)+'</td><td>'+esc(n.why)+'</td></tr>'}).join('');
  $('negBody').innerHTML=negHtml||'<tr><td colspan="5" class="muted">No search term trips the 3x rule yet.</td></tr>';
  $('signup').href=D.urls.signup;$('call').href=D.urls.call;
  $('foot').innerHTML=D.notes.map(esc).join('<br>')+'<br>Full Circle Hero ASIN Architect v'+esc(D.version)+'. Starting bids are a plan, not a promise: review before you enable.';
  KEYS.forEach(function(k){$('i_'+k).addEventListener('input',function(){read(k)})});
  Array.prototype.forEach.call(document.querySelectorAll('[data-stage]'),function(b){b.addEventListener('click',function(){
    $('i_t').value=Math.round(S.be*D.stage_factor[b.getAttribute('data-stage')]*1000)/10;read('t')})});
  $('reset').addEventListener('click',function(){inp=JSON.parse(JSON.stringify(D.inp));manual=!!D.inp.budget_manual;setInputs();render()});
  Array.prototype.forEach.call(document.querySelectorAll('[data-f]'),function(b){b.addEventListener('click',function(){
    filter=b.getAttribute('data-f');Array.prototype.forEach.call(document.querySelectorAll('[data-f]'),function(x){x.classList.toggle('on',x===b)});render()})});
  $('q').addEventListener('input',render);
  ['dlCsv','dlCsvTop'].forEach(function(id){$(id).addEventListener('click',download)});
  ['copyBrief','copyTop'].forEach(function(id){$(id).addEventListener('click',copy)});
  setInputs();render();
  window.__HAA={D:D,state:function(){return S},inputs:function(){return inp}};
})();
</script>
</body>
</html>
"""


def write_html(D, path):
    payload = json.dumps(D, separators=(",", ":")).replace("</", "<\\/")
    open(path, "w", encoding="utf-8").write(HTML.replace("__DATA__", payload))


# ---------------------------------------------------------------------------- CLI
def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Full Circle Hero ASIN Campaign Architect: per-keyword Max CPC, Target CPC "
                                            "and starting bids from your real SQP + Search Term Report exports.")
    p.add_argument("--sample", action="store_true", help="build the demo (fictional brand), no files needed")
    p.add_argument("--sqp", help="Brand Analytics Search Query Performance export (CSV/XLSX)")
    p.add_argument("--str", help="Sponsored Products Search Term Report export (CSV/XLSX)")
    p.add_argument("--competitors", help="optional competitor list (Scout, Helium 10, Keepa): needs an ASIN column")
    p.add_argument("--asin", help="hero ASIN")
    p.add_argument("--sku", help="your seller SKU for the hero ASIN (needed for the Product Ad rows)")
    p.add_argument("--brand", help="brand name (used in campaign names and to separate branded search terms)")
    p.add_argument("--brand-terms", help="comma-separated words that mark a search as branded (default: from --brand)")
    p.add_argument("--own-asins", help="comma-separated ASINs from your own catalog (defensive product targets)")
    p.add_argument("--price", type=float, help="selling price")
    p.add_argument("--cogs", type=float, help="landed cost of goods per unit")
    p.add_argument("--referral-fee-pct", type=float, default=15.0, help="Amazon referral fee percent (default 15)")
    p.add_argument("--fba-fee", type=float, help="FBA fulfillment fee per unit")
    p.add_argument("--returns-storage", type=float, help="returns + storage cost per unit (default 0)")
    p.add_argument("--target-acos", type=float, help="target ACoS percent (default: break-even ACoS x the stage factor)")
    p.add_argument("--stage", choices=sorted(STAGE_FACTOR), default="balanced",
                   help="profit = 42%% of break-even, balanced = 56%% (default), scale = 75%%")
    p.add_argument("--target-tacos", type=float, help="target TACoS percent (default: 45%% of the target ACoS)")
    p.add_argument("--monthly-sales", type=float, help="total monthly sales for this ASIN, ads + organic")
    p.add_argument("--daily-budget", type=float, help="total daily ad budget (default: monthly sales x TACoS / 30)")
    p.add_argument("--campaign-filter", help="only use Search Term Report rows whose campaign name contains this text")
    p.add_argument("--days", type=int, help="days covered by the Search Term Report (default: read from its dates)")
    p.add_argument("--out-dir", default=".", help="where to write the 4 files (default: current folder)")
    args = p.parse_args(argv)
    if not (args.sample or args.sqp or args.str):
        args.sample = True
        print("No input files given: building the sample. Use --sqp and --str for your own account (see --help).")
    return args


def main(argv=None):
    args = parse_args(argv)
    D = build(args)
    S = compute(D)
    os.makedirs(args.out_dir, exist_ok=True)
    out = {k: os.path.join(args.out_dir, v) for k, v in D["files"].items()}
    write_html(D, out["html"])
    write_xlsx(D, S, out["xlsx"])
    open(out["csv"], "w", newline="", encoding="utf-8").write(bulk_csv(D, S))
    open(out["md"], "w", encoding="utf-8").write(brief_md(D, S))

    i = D["inp"]
    print("")
    print("Hero ASIN %s (%s)" % (D["asin"], D["brand"]))
    for n in D["notes"]:
        print("  " + n)
    print("  Break-even ACoS %.1f%% | Target ACoS %.1f%% | Target TACoS %.1f%% | ASIN CVR %.1f%% (%d clicks per sale)"
          % (S["be"] * 100, i["t"], i["tacos"], D["asin_cvr"] * 100, D["cfs"]))
    print("  Per-keyword Max CPC %s to %s | %d campaigns | %d targets | %s a day | all paused"
          % (money(S["max_lo"]), money(S["max_hi"]), len(D["camps"]), D["n_targets"], money(S["budget_total"], 0)))
    print("  Monthly bleed found: %s (%d search terms + %d ASINs to negate)"
          % (money(S["bleed_mo"], 0), len([r for r in D["rows"] if r["type"] == "negative"]), len(D["neg_asins"])))
    for x in D["assumed"]:
        warn("CHECK: " + x)
    print("")
    print("4 files written:")
    for k in ("html", "xlsx", "csv", "md"):
        print("  " + os.path.abspath(out[k]))
    print("")
    print("Next: open the .html, confirm COGS / fees / Target ACoS / TACoS in the calculator, then either")
    print("  1) upload the Bulk 2.0 CSV yourself (Amazon Ads > Bulk operations), or")
    print("  2) click 'Copy Strategy, Goal & Budget Brief for Dr. PPC' and paste it into Dr. PPC in Orbit")
    print("     ($1 for the first 30 days, code FOUNDER10): " + SIGNUP_URL)


if __name__ == "__main__":
    main()

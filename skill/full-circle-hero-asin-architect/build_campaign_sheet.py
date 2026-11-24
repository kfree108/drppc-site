#!/usr/bin/env python3
"""
Full Circle Hero ASIN Campaign Architecture Builder
--------------------------------------------------
Generates a complete, formatted 6-tab Excel workbook (.xlsx) implementing
Full Circle's production Amazon PPC system ($500M+ managed Amazon revenue):
  - Tab 1: Executive Summary & Perfect CPC (MIN(Profitability Anchor, Market P90+15%))
  - Tab 2: SQP + STR Opportunity Matrix
  - Tab 3: Competitor Scout Targets (Top 15 ASINs, Offensive vs Defensive PAT)
  - Tab 4: Full Circle 8-Segment Structure (auto/disc/grow/domin/skc/br/def/off)
  - Tab 5: Amazon Bulk Upload Ready (Sponsored Products Bulk Operations schema)
  - Tab 6: 3x Negation Shield (Zero-Sale 3x CFS, 3x Target ACoS, Branded Isolation)

Output: sample-hero-asin-campaign-architecture.xlsx
"""

import math
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# Brand Palette (Full Circle Orbit: Dark Slate + Emerald + Full Circle Gold)
# ==============================================================================
DARK_NAVY   = "0B0E16"
SLATE_HDR   = "131A2B"
SLATE_SUB   = "1E293B"
EMERALD     = "10B981"
EMERALD_LT  = "D1FAE5"
EMERALD_DK  = "065F46"
GOLD        = "F59E0B"
GOLD_LT     = "FEF3C7"
GOLD_DK     = "92400E"
RED_LT      = "FEE2E2"
RED_DK      = "991B1B"
BLUE_LT     = "DBEAFE"
BLUE_DK     = "1E40AF"
ROW_ALT     = "F8FAFC"
BORDER_CLR  = "CBD5E1"
WHITE       = "FFFFFF"
MUTED_TXT   = "475569"

THIN_SIDE = Side(border_style="thin", color=BORDER_CLR)
MED_EMERALD = Side(border_style="medium", color=EMERALD)
CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

FONT_FAMILY = "Calibri"

def make_font(size=10, bold=False, color="0F172A", italic=False):
    return Font(name=FONT_FAMILY, size=size, bold=bold, color=color, italic=italic)

def make_fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


# ==============================================================================
# Sample $2.4M/yr Hero ASIN Dataset (Apex Recovery Labs — B09X4K8M2P)
# ==============================================================================
HERO = {
    "brand": "Apex Recovery Labs",
    "asin": "B09X4K8M2P",
    "sku": "ARL-MGPRO-BLK-01",
    "title": "Apex Pro Deep-Tissue Percussion Massage Gun (6-Head Brushless Motor)",
    "category": "Sports & Outdoors > Massage & Relaxation",
    "stage": "Balanced",
    "annual_rev": 2412000.00,
    "monthly_rev": 201000.00,
    "monthly_units": 2872,
    "aov": 69.99,
    "cogs": 15.40,
    "referral_fee_pct": 0.15,
    "fba_fee": 6.85,
    "returns_storage": 2.10,
    "cvr_30d": 0.148,
    "avg_cpc_30d": 2.18,
    "p90_cpc_30d": 2.75,
    "cpc_buffer_pct": 0.15,
    "target_acos": 0.28,
    "brand_target_acos": 0.14,
    "monthly_ad_spend": 38400.00,
    "current_acos": 0.364,
    "current_tacos": 0.191,
    "target_tacos": 0.125,
    "own_defense_asins": ["B09X4K9R1L", "B0B2M7N4Q8", "B0C1T5P9W3"],
}

# Derived Full Circle / Orbit cpc-resolver.ts metrics
REFERRAL_FEE = round(HERO["aov"] * HERO["referral_fee_pct"], 2)
CONTRIB_MARGIN = round(HERO["aov"] - HERO["cogs"] - REFERRAL_FEE - HERO["fba_fee"] - HERO["returns_storage"], 2)
BREAKEVEN_ACOS = CONTRIB_MARGIN / HERO["aov"]
CFS = math.ceil(1.0 / HERO["cvr_30d"])  # Always ROUNDUP (L99 Spec Warning #5)
ANCHOR_CPC = round((HERO["aov"] * HERO["target_acos"]) / CFS, 2)
MARKET_CPC = round(HERO["p90_cpc_30d"] * (1.0 + HERO["cpc_buffer_pct"]), 2)
PERFECT_CPC = min(ANCHOR_CPC, MARKET_CPC)

# Balanced Stage 3x Negation Shield thresholds (negation-evaluator.ts)
NEG_AG_CLICKS = math.ceil(CFS * 1.0)
NEG_ASIN_CLICKS = math.ceil(CFS * 3.0)
NEG_ACOS_LIMIT = HERO["target_acos"] * 3.0
NEG_SPEND_THRESHOLD = round(CFS * HERO["avg_cpc_30d"] * 1.25, 2)
NEG_PAT_AG_CLICKS = min(4, max(1, math.ceil(CFS * 0.67)))
NEG_PAT_ASIN_CLICKS = min(10, math.ceil(CFS * 2.0))


def style_banner(ws, max_col, title, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    c1 = ws.cell(row=1, column=1, value=title)
    c1.font = make_font(size=15, bold=True, color=WHITE)
    c1.fill = make_fill(DARK_NAVY)
    c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    c2 = ws.cell(row=2, column=1, value=subtitle)
    c2.font = make_font(size=10, italic=True, color="A7F3D0")
    c2.fill = make_fill(SLATE_HDR)
    c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 22

    # Paint full banner background across merged cells
    for col in range(1, max_col + 1):
        ws.cell(row=1, column=col).fill = make_fill(DARK_NAVY)
        ws.cell(row=2, column=col).fill = make_fill(SLATE_HDR)


def style_table_header(ws, row_idx, headers):
    ws.row_dimensions[row_idx].height = 26
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=text)
        cell.font = make_font(size=10, bold=True, color=WHITE)
        cell.fill = make_fill(SLATE_SUB)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=MED_EMERALD)


def set_col_widths(ws, widths):
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w


# ==============================================================================
# TAB 1: Executive Summary & Perfect CPC
# ==============================================================================
def build_tab1(wb):
    ws = wb.active
    ws.title = "Executive Summary & Perfect CPC"
    ws.views.sheetView[0].showGridLines = True

    style_banner(
        ws, 6,
        f"FULL CIRCLE HERO ASIN ARCHITECTURE — {HERO['brand']} ({HERO['asin']})",
        "Generated by Full Circle Hero ASIN Architect (Claude Code Skill) · Powered by Full Circle Orbit (Dr. PPC + Ava) · $500M+ Managed Playbook"
    )

    # Section A: Hero ASIN Snapshot & Unit Economics
    ws.cell(row=4, column=1, value="1. HERO ASIN UNIT ECONOMICS & MARGIN PROFILE").font = make_font(11, True, EMERALD_DK)
    style_table_header(ws, 5, ["Metric", "Value", "Formula / Source", "Full Circle Benchmark", "Status", "Notes"])

    econ_rows = [
        ("Hero ASIN / SKU", f"{HERO['asin']} / {HERO['sku']}", "Auto-locked #1 Revenue ASIN", "1 Free Hero ASIN", "LOCKED", HERO["title"]),
        ("Annualized Revenue", HERO["annual_rev"], "SP-API 30d × 12", "$1M+ Hero Tier", "QUALIFIED", f"${HERO['monthly_rev']:,.0f}/mo ({HERO['monthly_units']:,} units/mo)"),
        ("Selling Price (AOV)", HERO["aov"], "30d Avg Order Value", ">$35 Ideal", "HEALTHY", "Strong CPC absorption headroom"),
        ("Landed COGS", HERO["cogs"], "22.0% of AOV", "<25% Target", "HEALTHY", "Confirmed in Orbit Product Config"),
        ("Amazon Referral Fee (15%)", "=B8*0.15", "=AOV * 15%", "Category Standard", "FIXED", "Sports & Outdoors referral tier"),
        ("FBA Fulfillment + Storage/Returns", HERO["fba_fee"] + HERO["returns_storage"], "$6.85 FBA + $2.10 Buf", "<15% of AOV", "HEALTHY", "Includes 3% return/storage buffer"),
        ("Pre-Ad Contribution Margin ($)", "=B8-B9-B10-B11", "=AOV - COGS - Fees", ">$20.00 / unit", "STRONG", "Gross profit available before ad spend"),
        ("Break-Even ACoS (%)", "=B12/B8", "=Margin / AOV", ">=40% Ideal", "STRONG", "Absolute ceiling before losing money on ad sale"),
        ("Target ACoS — Non-Branded (Balanced)", HERO["target_acos"], "Break-Even × 56%", "25%–32%", "TARGET", "Leaves 22.2% net profit margin on ad sales"),
        ("Target ACoS — Branded Lane (br)", HERO["brand_target_acos"], "Full Circle Brand Rule", "10%–16%", "ISOLATED", "Strictly separated from Non-Branded campaigns"),
        ("Measured 30d Listing CVR", HERO["cvr_30d"], "SP-API Unit Session %", ">=12% Hero Bar", "STRONG", "Excludes last 2–3 attribution lag days"),
        ("Clicks For Sale (CFS — ROUNDUP)", "=ROUNDUP(1/B16,0)", "=ROUNDUP(1 / CVR, 0)", "Always Round Up", "VERIFIED", "L99 Spec Warning #5: Never round CFS down"),
    ]

    for i, r in enumerate(econ_rows, start=6):
        ws.row_dimensions[i].height = 21
        fill = make_fill(ROW_ALT) if i % 2 == 1 else make_fill(WHITE)
        for c_idx, val in enumerate(r, start=1):
            cell = ws.cell(row=i, column=c_idx, value=val)
            cell.font = make_font(10, bold=(c_idx in (1, 2)))
            cell.fill = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="center", horizontal="right" if c_idx == 2 else "left")

        # Number formats for Column B
        if i in (7, 8, 9, 10, 11, 12):
            ws.cell(row=i, column=2).number_format = "$#,##0.00"
        elif i in (13, 14, 15, 16):
            ws.cell(row=i, column=2).number_format = "0.0%"
        elif i == 17:
            ws.cell(row=i, column=2).number_format = "0"

        # Status badge
        st_cell = ws.cell(row=i, column=5)
        st_cell.alignment = Alignment(horizontal="center", vertical="center")
        st_cell.font = make_font(9, True, EMERALD_DK)
        st_cell.fill = make_fill(EMERALD_LT)

    # Section B: Full Circle Perfect CPC Formula (cpc-resolver.ts)
    ws.cell(row=19, column=1, value="2. FULL CIRCLE PERFECT CPC FORMULA (ORBIT cpc-resolver.ts)").font = make_font(11, True, EMERALD_DK)
    style_table_header(ws, 20, ["CPC Component", "Value", "Exact Orbit Formula", "Resolution Logic", "Winning Branch", "Interpretation"])

    cpc_rows = [
        ("Measured 30d Avg CPC", HERO["avg_cpc_30d"], "30d Spend / 30d Clicks", "Cold-start fallback", "INPUT", "Current blended average click cost"),
        ("Measured 30d P90 CPC", HERO["p90_cpc_30d"], "90th Percentile CPC", "Auction reality check", "INPUT", "Top-decile winning click price in category"),
        ("Profitability Anchor CPC", "=ROUND((B8*B14)/B17,2)", "=AOV × Target_ACoS / CFS", "Max profitable click", "WINNER (MIN)", "Max CPC that still hits 28.0% Target ACoS at 7 CFS"),
        ("Market Rate Ceiling (P90 + 15%)", "=ROUND(B22*(1+0.15),2)", "=P90_CPC × (1 + 15%)", "Max market asks", "CEILING", "15% headroom above P90 auction clearing rate"),
        ("PERFECT CPC (CPC-for-Rule)", "=MIN(B23,B24)", "=MIN(Anchor, Market)", "Tighter ceiling wins", "ACTIVE RULE CPC", "Don't pay more than profitable OR more than market asks"),
    ]
    for i, r in enumerate(cpc_rows, start=21):
        ws.row_dimensions[i].height = 23
        is_hero_row = (i == 25)
        fill = make_fill(GOLD_LT) if is_hero_row else (make_fill(ROW_ALT) if i % 2 == 1 else make_fill(WHITE))
        for c_idx, val in enumerate(r, start=1):
            cell = ws.cell(row=i, column=c_idx, value=val)
            cell.font = make_font(11 if is_hero_row else 10, bold=(is_hero_row or c_idx in (1, 2)), color=GOLD_DK if is_hero_row else "0F172A")
            cell.fill = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="center", horizontal="right" if c_idx == 2 else "left")
        ws.cell(row=i, column=2).number_format = "$#,##0.00"
        ws.cell(row=i, column=5).alignment = Alignment(horizontal="center", vertical="center")
        if is_hero_row:
            ws.cell(row=i, column=5).fill = make_fill(EMERALD)
            ws.cell(row=i, column=5).font = make_font(10, True, WHITE)

    # Section C: Monthly Bleed Found & 30-Day Recovery Impact
    ws.cell(row=27, column=1, value="3. MONTHLY AD BLEED FOUND & PROFIT RECOVERY SUMMARY").font = make_font(11, True, RED_DK)
    style_table_header(ws, 28, ["Leak Vector Detected", "Monthly Bleed ($)", "Annualized Impact ($)", "% of Ad Spend", "Full Circle Fix Applied", "Target Tab"])

    bleed_rows = [
        ("3x Negation Shield Bleed (0-Sale >=21 Clicks & >84% ACoS)", 2914.00, "=B29*12", "=B29/38400", "28 Negative Exact KWs + 6 Negative PAT ASINs", "Tab 6: 3x Negation Shield"),
        ("Branded / Non-Branded Lane Contamination & Over-Bidding", 1890.00, "=B30*12", "=B30/38400", "Isolate 'br' lane (14% ACoS) + Cap bids at $2.80 Perfect CPC", "Tab 4: Full Circle 8-Segment"),
        ("Unharvested Exact Winners Buried in Auto/Broad Discovery", 1268.00, "=B31*12", "=B31/38400", "Graduate 14 proven terms to DOMIN & 4 roots to SKC", "Tab 2 & Tab 5: Bulk Upload"),
        ("Uncontrolled Placement Bleed (TOS >36% ACoS / PP Leak)", 770.00, "=B32*12", "=B32/38400", "Reset high-ACoS TOS to 0% + Defensive PAT lock", "Tab 3 & Tab 4"),
        ("TOTAL RECOVERABLE MONTHLY BLEED", "=SUM(B29:B32)", "=SUM(C29:C32)", "=SUM(D29:D32)", "Immediate Bulk Upload + Daily Orbit Guardrails", "All 6 Tabs Ready"),
    ]
    for i, r in enumerate(bleed_rows, start=29):
        ws.row_dimensions[i].height = 23
        is_total = (i == 33)
        fill = make_fill(RED_LT) if is_total else (make_fill(ROW_ALT) if i % 2 == 1 else make_fill(WHITE))
        for c_idx, val in enumerate(r, start=1):
            cell = ws.cell(row=i, column=c_idx, value=val)
            cell.font = make_font(11 if is_total else 10, bold=(is_total or c_idx in (2, 3)), color=RED_DK if is_total else "0F172A")
            cell.fill = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="center", horizontal="right" if c_idx in (2, 3, 4) else "left")
        ws.cell(row=i, column=2).number_format = "$#,##0"
        ws.cell(row=i, column=3).number_format = "$#,##0"
        ws.cell(row=i, column=4).number_format = "0.0%"

    # Footer CTA box
    ws.merge_cells("A35:F36")
    cta = ws.cell(
        row=35, column=1,
        value="NEXT STEP — AUTOMATE THIS EVERY MORNING: Dr. PPC + Ava (inside Full Circle Orbit) runs Full Circle's 8-segment rule engine, 3x Negation Shield, placement resets & plain-English Reason Log automatically. First 30 days for $1, then $299/mo + 2% of ad spend (month-to-month · code FOUNDER10 = 10% off for life) → fullcircleorbit.com/signup?from=drppc&utm_source=creator&utm_campaign=hormozi_skill_workbook&code=FOUNDER10"
    )
    cta.font = make_font(10, True, EMERALD_DK)
    cta.fill = make_fill(EMERALD_LT)
    cta.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for r in (35, 36):
        ws.row_dimensions[r].height = 20
        for c in range(1, 7):
            ws.cell(row=r, column=c).fill = make_fill(EMERALD_LT)
            ws.cell(row=r, column=c).border = CELL_BORDER

    set_col_widths(ws, [42, 18, 28, 24, 18, 44])
    ws.freeze_panes = "A6"


# ==============================================================================
# TAB 2: SQP + STR Opportunity Matrix
# ==============================================================================
SQP_STR_DATA = [
    # query, sq_vol, imp_sh, clk_sh, cart_sh, purch_sh, cat_cvr, clicks30, orders30, spend30, sales30, cpc30, segment, action, bid
    ("deep tissue massage gun", 148200, 0.084, 0.112, 0.138, 0.162, 0.115, 412, 68, 1087.68, 4759.32, 2.64, "skc", "LAUNCH SKC (+60% TOS)", 2.80),
    ("percussion massager for athletes", 64500, 0.062, 0.095, 0.121, 0.149, 0.108, 248, 41, 582.80, 2869.59, 2.35, "skc", "LAUNCH SKC (+50% TOS)", 2.75),
    ("brushless massage gun quiet", 39800, 0.071, 0.108, 0.142, 0.175, 0.124, 195, 36, 421.20, 2519.64, 2.16, "skc", "LAUNCH SKC (+50% TOS)", 2.65),
    ("back and shoulder massage gun", 52100, 0.055, 0.082, 0.098, 0.118, 0.102, 214, 31, 496.48, 2169.69, 2.32, "skc", "LAUNCH SKC (+40% TOS)", 2.60),
    ("muscle recovery massager", 44900, 0.049, 0.068, 0.089, 0.104, 0.096, 178, 27, 377.36, 1889.73, 2.12, "domin", "HARVEST TO DOMIN (EXACT)", 2.45),
    ("handheld deep tissue massager", 38400, 0.064, 0.079, 0.094, 0.112, 0.105, 162, 25, 353.16, 1749.75, 2.18, "domin", "HARVEST TO DOMIN (EXACT)", 2.40),
    ("sports massage gun 6 heads", 19200, 0.092, 0.135, 0.168, 0.194, 0.131, 134, 24, 261.30, 1679.76, 1.95, "domin", "HARVEST TO DOMIN (EXACT)", 2.35),
    ("percussion therapy gun", 31600, 0.041, 0.059, 0.076, 0.088, 0.094, 146, 21, 319.74, 1469.79, 2.19, "domin", "HARVEST TO DOMIN (EXACT)", 2.30),
    ("cordless massage gun heavy duty", 27400, 0.058, 0.081, 0.102, 0.121, 0.111, 129, 20, 278.64, 1399.80, 2.16, "domin", "HARVEST TO DOMIN (EXACT)", 2.35),
    ("massage gun for runners", 22900, 0.053, 0.074, 0.091, 0.109, 0.104, 118, 19, 245.44, 1329.81, 2.08, "domin", "HARVEST TO DOMIN (EXACT)", 2.25),
    ("physical therapy massage gun", 28700, 0.044, 0.062, 0.079, 0.092, 0.098, 132, 18, 297.00, 1259.82, 2.25, "domin", "HARVEST TO DOMIN (EXACT)", 2.20),
    ("massage gun for sore muscles", 41200, 0.038, 0.051, 0.064, 0.075, 0.089, 141, 17, 315.84, 1189.83, 2.24, "domin", "HARVEST TO DOMIN (EXACT)", 2.10),
    ("high torque percussion massager", 14800, 0.081, 0.119, 0.145, 0.168, 0.128, 92, 16, 180.32, 1119.84, 1.96, "domin", "HARVEST TO DOMIN (EXACT)", 2.35),
    ("neck and back percussion gun", 24300, 0.036, 0.048, 0.058, 0.069, 0.091, 108, 14, 232.20, 979.86, 2.15, "grow", "SCALE IN GROW (PHRASE)", 1.95),
    ("massage gun with carrying case", 18600, 0.049, 0.064, 0.078, 0.091, 0.103, 94, 13, 188.94, 909.87, 2.01, "grow", "SCALE IN GROW (PHRASE)", 1.90),
    ("gym recovery massage tool", 15900, 0.042, 0.055, 0.069, 0.081, 0.095, 86, 11, 171.14, 769.89, 1.99, "grow", "SCALE IN GROW (PHRASE)", 1.85),
    ("quiet massage gun for home", 17400, 0.039, 0.052, 0.063, 0.074, 0.093, 79, 10, 155.63, 699.90, 1.97, "grow", "SCALE IN GROW (PHRASE)", 1.85),
    ("deep muscle stimulator gun", 12100, 0.034, 0.045, 0.056, 0.066, 0.088, 68, 8, 138.72, 559.92, 2.04, "grow", "SCALE IN GROW (PHRASE)", 1.75),
    ("apex recovery massage gun", 16400, 0.482, 0.645, 0.720, 0.785, 0.242, 268, 74, 343.04, 5179.26, 1.28, "br", "ISOLATE TO BRAND LANE (br)", 1.45),
    ("apex pro percussion massager", 8900, 0.515, 0.680, 0.748, 0.812, 0.260, 154, 45, 181.72, 3149.55, 1.18, "br", "ISOLATE TO BRAND LANE (br)", 1.35),
    ("cheap massage gun under 25", 33200, 0.021, 0.018, 0.004, 0.000, 0.042, 47, 0, 106.22, 0.00, 2.26, "neg", "NEGATE 3x SHIELD (0 ORDERS)", 0.00),
    ("mini pocket travel massager", 46800, 0.019, 0.015, 0.005, 0.002, 0.051, 39, 0, 88.92, 0.00, 2.28, "neg", "NEGATE 3x SHIELD (0 ORDERS)", 0.00),
    ("massage chair pad heated", 58900, 0.014, 0.011, 0.002, 0.000, 0.039, 34, 0, 81.94, 0.00, 2.41, "neg", "NEGATE 3x SHIELD (0 ORDERS)", 0.00),
    ("foot massager machine shiatsu", 74100, 0.012, 0.009, 0.001, 0.000, 0.044, 29, 0, 71.05, 0.00, 2.45, "neg", "NEGATE 3x SHIELD (0 ORDERS)", 0.00),
]

def build_tab2(wb):
    ws = wb.create_sheet("SQP + STR Opportunity Matrix")
    style_banner(
        ws, 16,
        "SQP (SEARCH QUERY PERFORMANCE) + STR OPPORTUNITY MATRIX",
        "Cross-joining Brand Analytics Share Gaps with 30-Day Search Term Report (2-day attribution lag excluded) · Anchored to $2.80 Perfect CPC"
    )
    headers = [
        "Search Query", "SQP Monthly Vol", "Brand Impr %", "Brand Click %", "Brand Cart %",
        "Brand Purch %", "Purch vs Impr Gap", "30d Clicks", "30d Orders", "30d Spend",
        "30d Sales", "30d ACoS", "30d CVR", "Target Segment", "Full Circle Action", "Perfect CPC Bid"
    ]
    style_table_header(ws, 4, headers)

    for idx, row in enumerate(SQP_STR_DATA, start=5):
        q, vol, imp, clk, cart, pur, cat_cvr, clicks, orders, spend, sales, cpc, seg, action, bid = row
        ws.row_dimensions[idx].height = 21
        fill = make_fill(ROW_ALT) if idx % 2 == 1 else make_fill(WHITE)

        vals = [
            q, vol, imp, clk, cart, pur,
            f"=F{idx}-C{idx}",  # Purchase share vs Impression share edge
            clicks, orders, spend, sales,
            f"=IF(K{idx}>0,J{idx}/K{idx},0)",
            f"=IF(H{idx}>0,I{idx}/H{idx},0)",
            seg.upper(), action, bid
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=idx, column=c_idx, value=val)
            cell.font = make_font(10, bold=(c_idx in (1, 14, 16)))
            cell.fill = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="center", horizontal="left" if c_idx in (1, 15) else ("center" if c_idx == 14 else "right"))

        ws.cell(row=idx, column=2).number_format = "#,##0"
        for c in (3, 4, 5, 6, 7, 12, 13):
            ws.cell(row=idx, column=c).number_format = "0.0%"
        ws.cell(row=idx, column=8).number_format = "#,##0"
        ws.cell(row=idx, column=9).number_format = "#,##0"
        ws.cell(row=idx, column=10).number_format = "$#,##0.00"
        ws.cell(row=idx, column=11).number_format = "$#,##0.00"
        ws.cell(row=idx, column=16).number_format = "$#,##0.00"

        # Action badge color coding
        act_cell = ws.cell(row=idx, column=15)
        seg_cell = ws.cell(row=idx, column=14)
        if seg == "neg":
            act_cell.fill = make_fill(RED_LT); act_cell.font = make_font(9, True, RED_DK)
            seg_cell.fill = make_fill(RED_LT); seg_cell.font = make_font(9, True, RED_DK)
        elif seg == "skc":
            act_cell.fill = make_fill(GOLD_LT); act_cell.font = make_font(9, True, GOLD_DK)
            seg_cell.fill = make_fill(GOLD_LT); seg_cell.font = make_font(9, True, GOLD_DK)
        elif seg == "br":
            act_cell.fill = make_fill(BLUE_LT); act_cell.font = make_font(9, True, BLUE_DK)
            seg_cell.fill = make_fill(BLUE_LT); seg_cell.font = make_font(9, True, BLUE_DK)
        else:
            act_cell.fill = make_fill(EMERALD_LT); act_cell.font = make_font(9, True, EMERALD_DK)
            seg_cell.fill = make_fill(EMERALD_LT); seg_cell.font = make_font(9, True, EMERALD_DK)

    # Summary row
    tot = 5 + len(SQP_STR_DATA)
    ws.row_dimensions[tot].height = 24
    ws.cell(row=tot, column=1, value="PORTFOLIO OPPORTUNITY TOTALS (24 CORE QUERIES)").font = make_font(10, True, WHITE)
    ws.cell(row=tot, column=2, value=f"=SUM(B5:B{tot-1})").number_format = "#,##0"
    ws.cell(row=tot, column=8, value=f"=SUM(H5:H{tot-1})").number_format = "#,##0"
    ws.cell(row=tot, column=9, value=f"=SUM(I5:I{tot-1})").number_format = "#,##0"
    ws.cell(row=tot, column=10, value=f"=SUM(J5:J{tot-1})").number_format = "$#,##0.00"
    ws.cell(row=tot, column=11, value=f"=SUM(K5:K{tot-1})").number_format = "$#,##0.00"
    ws.cell(row=tot, column=12, value=f"=J{tot}/K{tot}").number_format = "0.0%"
    ws.cell(row=tot, column=13, value=f"=I{tot}/H{tot}").number_format = "0.0%"
    for c in range(1, 17):
        cell = ws.cell(row=tot, column=c)
        cell.fill = make_fill(SLATE_HDR)
        cell.font = make_font(10, True, WHITE)
        cell.border = CELL_BORDER

    set_col_widths(ws, [34, 15, 13, 13, 13, 13, 15, 11, 11, 13, 13, 11, 11, 13, 27, 15])
    ws.freeze_panes = "B5"


# ==============================================================================
# TAB 3: Competitor Scout Targets (Top 15 ASINs)
# ==============================================================================
SCOUT_TARGETS = [
    # asin, brand_label, price, rating, reviews, est_mo_rev, bsr, lane, edge_reason, pat_bid, pp_mod
    ("B08K9L2P4Q", "Competitor A — Mid-Tier Massager", 89.99, 4.3, 1420, 165000, 1840, "off", "We are $20 cheaper + 0.4★ higher rating", 1.95, "+25% PP"),
    ("B09M4T8V1X", "Competitor B — Legacy Brand Pro", 119.99, 4.4, 3810, 290000, 920, "off", "We are $50 cheaper + brushless motor parity", 2.15, "+25% PP"),
    ("B0B7W3N6K2", "Competitor C — Generic 4-Head", 74.99, 4.1, 680, 92000, 3420, "off", "We beat on price, rating (+0.6★), & 6 heads", 1.85, "+20% PP"),
    ("B09R2H5C8M", "Competitor D — Athlete Gun", 94.99, 4.4, 2150, 198000, 1410, "off", "We are $25 cheaper + stronger review velocity", 2.05, "+25% PP"),
    ("B0C4J9Q1Z7", "Competitor E — Quiet Pro V2", 79.99, 4.2, 510, 74000, 4190, "off", "Weak social proof + $10 higher price", 1.75, "+20% PP"),
    ("B08V6X1D3P", "Competitor F — Heated Percussion", 99.99, 4.3, 1890, 152000, 2080, "off", "$30 cheaper + 0.4★ rating advantage", 1.90, "+20% PP"),
    ("B0B1F8S4Y9", "Competitor G — Deep Muscle XL", 84.99, 4.2, 940, 110000, 2910, "off", "Conquest high-intent PDP traffic ($15 edge)", 1.80, "+20% PP"),
    ("B09P5G7L2W", "Competitor H — Physio Relief", 109.99, 4.5, 2640, 215000, 1250, "off", "$40 price undercut with matching 4.7★ quality", 2.10, "+25% PP"),
    ("B0C8N2B5R4", "Competitor I — Cordless Elite", 76.99, 4.1, 390, 61000, 5120, "off", "High conquest CVR — 0.6★ rating gap", 1.65, "+15% PP"),
    ("B08T3K6M9J", "Competitor J — Gym Recovery Pro", 89.95, 4.3, 1280, 134000, 2410, "off", "$20 price advantage on overlapping SQP roots", 1.85, "+20% PP"),
    ("B0B9Q4V8H1", "Competitor K — Dual-Speed Gun", 69.99, 4.0, 440, 55000, 5890, "off", "Same price, massive +0.7★ quality edge", 1.60, "+15% PP"),
    ("B09W1D7F6N", "Competitor L — Sports Clinic Gun", 129.99, 4.4, 1750, 142000, 2240, "off", "$60 cheaper alternative for comparison shoppers", 2.00, "+25% PP"),
    ("B09X4K9R1L", "OWN CATALOG — Apex Mini Travel Gun", 49.99, 4.7, 1920, 88000, 3150, "def", "Defensive cross-sell: protect Hero PDP carousel", 1.25, "+30% PP"),
    ("B0B2M7N4Q8", "OWN CATALOG — Apex Replacement Heads", 24.99, 4.8, 860, 34000, 6840, "def", "Defensive accessory lock: raise basket AOV", 0.95, "+30% PP"),
    ("B0C1T5P9W3", "OWN CATALOG — Apex Recovery Roller", 39.99, 4.7, 1140, 52000, 4720, "def", "Defensive bundle defense: block competitors", 1.10, "+30% PP"),
]

def build_tab3(wb):
    ws = wb.create_sheet("Competitor Scout Targets")
    style_banner(
        ws, 13,
        "COMPETITOR SCOUT TARGETS — TOP 15 PAT OPPORTUNITIES (OFFENSIVE vs DEFENSIVE)",
        f"Hero ASIN Benchmark: {HERO['asin']} · Price: ${HERO['aov']:.2f} · Rating: 4.7★ · Reviews: 4,280 · PAT Strategy: Full Circle 'off' & 'def' Segments"
    )
    headers = [
        "Target ASIN", "Competitor / Own Product", "Price", "Price Edge vs Hero", "Rating",
        "Rating Edge", "Reviews", "Est. Mo Rev", "BSR", "Segment",
        "Scout Conquest Rationale", "PAT Starting Bid", "Placement Focus"
    ]
    style_table_header(ws, 4, headers)

    for idx, row in enumerate(SCOUT_TARGETS, start=5):
        asin, name, price, rating, rev, mo_rev, bsr, lane, reason, bid, pp = row
        ws.row_dimensions[idx].height = 22
        fill = make_fill(BLUE_LT) if lane == "def" else (make_fill(ROW_ALT) if idx % 2 == 1 else make_fill(WHITE))

        vals = [
            asin, name, price, f"=C{idx}-{HERO['aov']}", rating,
            f"=4.7-E{idx}", rev, mo_rev, bsr, lane.upper(), reason, bid, pp
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=idx, column=c_idx, value=val)
            cell.font = make_font(10, bold=(c_idx in (1, 10, 12)))
            cell.fill = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="center", horizontal="left" if c_idx in (1, 2, 11) else ("center" if c_idx in (10, 13) else "right"))

        ws.cell(row=idx, column=3).number_format = "$#,##0.00"
        ws.cell(row=idx, column=4).number_format = "+$#,##0.00;-$#,##0.00;$0.00"
        ws.cell(row=idx, column=5).number_format = "0.0"
        ws.cell(row=idx, column=6).number_format = "+0.0;-0.0;0.0"
        ws.cell(row=idx, column=7).number_format = "#,##0"
        ws.cell(row=idx, column=8).number_format = "$#,##0"
        ws.cell(row=idx, column=9).number_format = "#,##0"
        ws.cell(row=idx, column=12).number_format = "$#,##0.00"

        seg_cell = ws.cell(row=idx, column=10)
        if lane == "off":
            seg_cell.fill = make_fill(GOLD_LT); seg_cell.font = make_font(9, True, GOLD_DK)
        else:
            seg_cell.fill = make_fill(EMERALD_LT); seg_cell.font = make_font(9, True, EMERALD_DK)

    set_col_widths(ws, [15, 34, 11, 16, 10, 12, 11, 14, 10, 11, 44, 15, 14])
    ws.freeze_panes = "C5"


# ==============================================================================
# TAB 4: Full Circle 8-Segment Structure
# ==============================================================================
FC_8_SEGMENTS = [
    ("1", "auto", "SP_NB_B09X4K8M2P_AUTO_7Day", "7day_iterative", "Auto (Close/Loose/Sub/Comp)", 90, 2700, "Dynamic bids - down only", "0%", "0%", "0%", "28%", "Automated discovery & low-bid long-tail mining; all harvested & branded terms negated."),
    ("2", "disc", "SP_NB_B09X4K8M2P_DISC_Broad", "lean", "Broad / Phrase Discovery", 140, 4200, "Dynamic bids - down only", "10%", "0%", "0%", "28%", "Seeded with top SQP stems; feeds GROW/DOMIN via import rules (>=4-8 orders, <=1.25x ACoS)."),
    ("3", "grow", "SP_NB_B09X4K8M2P_GROW_Phrase", "core", "Phrase Mid-Tail Scaling", 160, 4800, "Dynamic bids - down only", "15%", "0%", "0%", "28%", "Scales mid-tail converting phrases with Core graduated bid rules & 3x negation shield."),
    ("4", "domin", "SP_NB_B09X4K8M2P_DOMIN_Exact", "core", "Exact Core Winners (9 KWs)", 320, 9600, "Dynamic bids - down only", "25%", "0%", "0%", "28%", "Primary profit engine — exact-match high-converting non-branded winners anchored to $2.80 Perfect CPC."),
    ("5", "skc", "SP_NB_B09X4K8M2P_SKC_DeepTissue", "kw_boost", "Single Keyword Campaigns (4 Roots)", 260, 7800, "Dynamic bids - down only", "50%", "0%", "0%", "32%", "1 hero SQP root per campaign for organic rank velocity; TOS auto-resets to 0% if ACoS >=36%."),
    ("6", "br", "SP_BR_B09X4K8M2P_BR_ExactPhrase", "brand", "Branded Search Defense", 85, 2550, "Dynamic bids - down only", "20%", "0%", "0%", "14%", "Strict Two-Lane isolation: Branded queries protected at 10–16% ACoS, negated in all NB lanes."),
    ("7", "def", "SP_DEF_B09X4K8M2P_DEF_PAT", "pat", "Defensive PAT (3 Own ASINs)", 65, 1950, "Dynamic bids - down only", "0%", "0%", "30%", "18%", "Locks own catalog variations & accessories onto Hero PDP carousels so competitors can't poach."),
    ("8", "off", "SP_OFF_B09X4K8M2P_OFF_PAT", "pat", "Offensive PAT (12 Scout ASINs)", 155, 4650, "Dynamic bids - down only", "0%", "0%", "25%", "28%", "Conquests Scout-verified competitor ASINs where Hero holds >=$15 price or >=0.4★ rating edge."),
]

def build_tab4(wb):
    ws = wb.create_sheet("Full Circle 8-Segment")
    style_banner(
        ws, 13,
        "FULL CIRCLE CAMPAIGN ARCHITECTURE — 8 STRUCTURAL SEGMENTS",
        "Exact Orbit segment-to-strategy mapping (strategy-derivation.ts) · Strict Branded vs Non-Branded Lane Isolation · Graduated Placement Control"
    )
    headers = [
        "#", "Segment", "Campaign Naming Convention", "Orbit Strategy Slug", "Targeting & Role",
        "Daily Budget", "Monthly Cap", "Bidding Strategy", "TOS Mod", "ROS Mod",
        "PP Mod", "Target ACoS", "Full Circle Operating Rules & Guardrails"
    ]
    style_table_header(ws, 4, headers)

    for idx, row in enumerate(FC_8_SEGMENTS, start=5):
        num, seg, cname, slug, role, dbud, mbud, bstrat, tos, ros, pp, tacos, notes = row
        ws.row_dimensions[idx].height = 30
        fill = make_fill(ROW_ALT) if idx % 2 == 1 else make_fill(WHITE)
        vals = [num, seg, cname, slug, role, dbud, f"=F{idx}*30", bstrat, tos, ros, pp, tacos, notes]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=idx, column=c_idx, value=val)
            cell.font = make_font(10, bold=(c_idx in (2, 3, 4)))
            cell.fill = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=(c_idx in (5, 13)),
                horizontal="center" if c_idx in (1, 2, 4, 9, 10, 11, 12) else ("right" if c_idx in (6, 7) else "left")
            )
        ws.cell(row=idx, column=6).number_format = "$#,##0"
        ws.cell(row=idx, column=7).number_format = "$#,##0"

        seg_cell = ws.cell(row=idx, column=2)
        if seg in ("domin", "grow"):
            seg_cell.fill = make_fill(EMERALD_LT); seg_cell.font = make_font(10, True, EMERALD_DK)
        elif seg == "skc":
            seg_cell.fill = make_fill(GOLD_LT); seg_cell.font = make_font(10, True, GOLD_DK)
        elif seg == "br":
            seg_cell.fill = make_fill(BLUE_LT); seg_cell.font = make_font(10, True, BLUE_DK)
        else:
            seg_cell.fill = make_fill("F1F5F9"); seg_cell.font = make_font(10, True, SLATE_SUB)

    # Totals Row
    tot = 13
    ws.row_dimensions[tot].height = 25
    ws.cell(row=tot, column=1, value="")
    ws.cell(row=tot, column=2, value="ALL 8")
    ws.cell(row=tot, column=3, value="COMPLETE HERO ASIN FULL CIRCLE ARCHITECTURE")
    ws.cell(row=tot, column=5, value="11 Live Campaigns (4 SKC + 7 Core Lanes)")
    ws.cell(row=tot, column=6, value="=SUM(F5:F12)").number_format = "$#,##0"
    ws.cell(row=tot, column=7, value="=SUM(G5:G12)").number_format = "$#,##0"
    ws.cell(row=tot, column=12, value="24.8% Blend")
    ws.cell(row=tot, column=13, value="Projected TACoS drops from 19.1% → 12.5% within 45 days")
    for c in range(1, 14):
        cell = ws.cell(row=tot, column=c)
        cell.fill = make_fill(SLATE_HDR)
        cell.font = make_font(10, True, WHITE)
        cell.border = CELL_BORDER

    set_col_widths(ws, [5, 10, 34, 16, 28, 13, 13, 24, 10, 10, 10, 13, 52])
    ws.freeze_panes = "D5"


# ==============================================================================
# TAB 5: Amazon Bulk Upload Ready (Sponsored Products Bulk Operations Schema)
# ==============================================================================
def build_tab5(wb):
    ws = wb.create_sheet("Amazon Bulk Upload Ready")
    style_banner(
        ws, 18,
        "AMAZON SPONSORED PRODUCTS BULK UPLOAD SHEET — READY FOR SELLER CENTRAL",
        "Official Amazon Bulk Operations 2.0 Columns · Copy rows 5+ into Amazon Ads > Sponsored Ads > Bulk Operations > Upload Spreadsheet"
    )
    headers = [
        "Product", "Entity", "Operation", "Campaign Id", "Ad Group Id",
        "Campaign Name", "Ad Group Name", "Start Date", "Targeting Type",
        "State", "Daily Budget", "SKU", "Ad Group Default Bid", "Bid",
        "Keyword Text", "Match Type", "Bidding Strategy", "Placement / Targeting Expr"
    ]
    style_table_header(ws, 4, headers)

    start_date = datetime.now().strftime("%Y%m%d")
    sku = HERO["sku"]
    rows = []

    def add_campaign_block(cid, cname, ttype, budget, bstrat, ag_name, def_bid, placement_note, targets):
        # Campaign row
        rows.append(["Sponsored Products", "Campaign", "Create", cid, "", cname, "", start_date, ttype, "enabled", budget, "", "", "", "", "", bstrat, placement_note])
        # Ad Group row
        rows.append(["Sponsored Products", "Ad Group", "Create", cid, f"{cid}_AG1", cname, ag_name, "", "", "enabled", "", "", def_bid, "", "", "", "", ""])
        # Product Ad row
        rows.append(["Sponsored Products", "Product Ad", "Create", cid, f"{cid}_AG1", cname, ag_name, "", "", "enabled", "", sku, "", "", "", "", "", ""])
        # Targets
        for t_entity, kw_text, mtype, bid, expr in targets:
            rows.append(["Sponsored Products", t_entity, "Create", cid, f"{cid}_AG1", cname, ag_name, "", "", "enabled", "", "", "", bid, kw_text, mtype, "", expr])

    # 1. AUTO
    add_campaign_block(
        "FC_AUTO_01", "SP_NB_B09X4K8M2P_AUTO_7Day", "AUTO", 90, "Dynamic bids - down only",
        "Auto_Discovery", 1.35, "Placement: TOS 0% | PP 0%",
        [
            ("Product Targeting", "", "", 1.55, "close-match"),
            ("Product Targeting", "", "", 1.10, "loose-match"),
            ("Product Targeting", "", "", 1.25, "substitutes"),
            ("Product Targeting", "", "", 0.85, "complements"),
            ("Campaign Negative Keyword", "apex recovery", "negativePhrase", "", ""),
            ("Campaign Negative Keyword", "apex pro", "negativePhrase", "", ""),
        ]
    )

    # 2. DISC
    add_campaign_block(
        "FC_DISC_01", "SP_NB_B09X4K8M2P_DISC_Broad", "MANUAL", 140, "Dynamic bids - down only",
        "Broad_Mining", 1.65, "Placement: TOS +10%",
        [
            ("Keyword", "deep tissue massage gun", "broad", 1.65, ""),
            ("Keyword", "percussion massager", "broad", 1.55, ""),
            ("Keyword", "muscle recovery massager", "broad", 1.50, ""),
            ("Keyword", "sports massage gun", "broad", 1.60, ""),
            ("Campaign Negative Keyword", "apex", "negativePhrase", "", ""),
        ]
    )

    # 3. GROW
    add_campaign_block(
        "FC_GROW_01", "SP_NB_B09X4K8M2P_GROW_Phrase", "MANUAL", 160, "Dynamic bids - down only",
        "Phrase_Core", 1.90, "Placement: TOS +15%",
        [
            ("Keyword", "neck and back percussion gun", "phrase", 1.95, ""),
            ("Keyword", "massage gun with carrying case", "phrase", 1.90, ""),
            ("Keyword", "gym recovery massage tool", "phrase", 1.85, ""),
            ("Keyword", "quiet massage gun for home", "phrase", 1.85, ""),
            ("Keyword", "deep muscle stimulator gun", "phrase", 1.75, ""),
            ("Campaign Negative Keyword", "apex", "negativePhrase", "", ""),
        ]
    )

    # 4. DOMIN (Exact Winners)
    domin_kws = [
        ("Keyword", r[0], "exact", r[14], "")
        for r in SQP_STR_DATA if r[12] == "domin"
    ]
    domin_kws.append(("Campaign Negative Keyword", "apex", "negativePhrase", "", ""))
    add_campaign_block(
        "FC_DOMIN_01", "SP_NB_B09X4K8M2P_DOMIN_Exact", "MANUAL", 320, "Dynamic bids - down only",
        "Exact_Winners", 2.30, "Placement: TOS +25%",
        domin_kws
    )

    # 5. SKC (Top 4 SQP Roots)
    skc_items = [r for r in SQP_STR_DATA if r[12] == "skc"]
    for i, r in enumerate(skc_items, start=1):
        slug = "".join(w.capitalize() for w in r[0].split()[:3])
        add_campaign_block(
            f"FC_SKC_0{i}", f"SP_NB_B09X4K8M2P_SKC_{slug}", "MANUAL", 65, "Dynamic bids - down only",
            f"SKC_{slug}", r[14], "Placement: TOS +50% (Reset 0% if ACoS>=36%)",
            [("Keyword", r[0], "exact", r[14], "")]
        )

    # 6. BR (Branded Lane)
    add_campaign_block(
        "FC_BR_01", "SP_BR_B09X4K8M2P_BR_ExactPhrase", "MANUAL", 85, "Dynamic bids - down only",
        "Brand_Defense", 1.40, "Placement: TOS +20% (14% Target ACoS)",
        [
            ("Keyword", "apex recovery massage gun", "exact", 1.45, ""),
            ("Keyword", "apex pro percussion massager", "exact", 1.35, ""),
            ("Keyword", "apex recovery labs", "phrase", 1.25, ""),
            ("Keyword", "apex massage gun", "phrase", 1.30, ""),
        ]
    )

    # 7. DEF (Defensive PAT)
    def_targets = [
        ("Product Targeting", "", "", r[9], f'asin="{r[0]}"')
        for r in SCOUT_TARGETS if r[7] == "def"
    ]
    add_campaign_block(
        "FC_DEF_01", "SP_DEF_B09X4K8M2P_DEF_PAT", "MANUAL", 65, "Dynamic bids - down only",
        "Own_Catalog_Lock", 1.10, "Placement: Product Pages +30%",
        def_targets
    )

    # 8. OFF (Offensive PAT — top 12 Scout ASINs)
    off_targets = [
        ("Product Targeting", "", "", r[9], f'asin="{r[0]}"')
        for r in SCOUT_TARGETS if r[7] == "off"
    ]
    add_campaign_block(
        "FC_OFF_01", "SP_OFF_B09X4K8M2P_OFF_PAT", "MANUAL", 155, "Dynamic bids - down only",
        "Scout_Conquest_PAT", 1.90, "Placement: Product Pages +25%",
        off_targets
    )

    for idx, r in enumerate(rows, start=5):
        ws.row_dimensions[idx].height = 20
        entity = r[1]
        if entity == "Campaign":
            fill = make_fill(EMERALD_LT)
            f_bold = True
        elif entity == "Campaign Negative Keyword":
            fill = make_fill(RED_LT)
            f_bold = False
        else:
            fill = make_fill(ROW_ALT) if idx % 2 == 1 else make_fill(WHITE)
            f_bold = False

        for c_idx, val in enumerate(r, start=1):
            cell = ws.cell(row=idx, column=c_idx, value=val)
            cell.font = make_font(9.5, bold=f_bold)
            cell.fill = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="center", horizontal="right" if c_idx in (11, 13, 14) else "left")

        if r[10] != "":
            ws.cell(row=idx, column=11).number_format = "$#,##0.00"
        if r[12] != "":
            ws.cell(row=idx, column=13).number_format = "$#,##0.00"
        if r[13] != "":
            ws.cell(row=idx, column=14).number_format = "$#,##0.00"

    set_col_widths(ws, [19, 25, 11, 14, 17, 33, 21, 11, 13, 10, 13, 19, 15, 10, 32, 16, 25, 36])
    ws.freeze_panes = "C5"


# ==============================================================================
# TAB 6: 3x Negation Shield
# ==============================================================================
NEGATION_ROWS = [
    # term, type, source_camp, clicks, orders, spend, sales, rule, neg_level, neg_match
    ("cheap massage gun under 25", "Keyword", "SP_NB_AUTO + DISC", 47, 0, 106.22, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("mini pocket travel massager", "Keyword", "SP_NB_DISC_Broad", 39, 0, 88.92, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("massage chair pad heated", "Keyword", "SP_NB_AUTO_7Day", 34, 0, 81.94, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("foot massager machine shiatsu", "Keyword", "SP_NB_AUTO_7Day", 29, 0, 71.05, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("tens unit muscle stimulator", "Keyword", "SP_NB_DISC_Broad", 31, 0, 74.40, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("neck pillow massager for car", "Keyword", "SP_NB_AUTO_7Day", 26, 0, 58.76, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("scalp massager electric waterproof", "Keyword", "SP_NB_AUTO_7Day", 24, 0, 51.12, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("cellulite massager handheld", "Keyword", "SP_NB_DISC_Broad", 23, 0, 49.45, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("theragun prime replacement charger", "Keyword", "SP_NB_DISC_Broad", 22, 0, 52.80, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("hyperice hypervolt battery pack", "Keyword", "SP_NB_AUTO_7Day", 21, 0, 48.30, 0.00, "Rule 2: 3x CFS Zero-Sale (>=21 clicks)", "Campaign", "Negative Exact"),
    ("massage oil lavender organic", "Keyword", "SP_NB_AUTO_7Day", 16, 0, 31.20, 0.00, "Rule 1: AG 1x CFS Zero-Sale (>=7 clicks)", "Ad Group", "Negative Exact"),
    ("foam roller trigger point", "Keyword", "SP_NB_DISC_Broad", 14, 0, 29.82, 0.00, "Rule 1: AG 1x CFS Zero-Sale (>=7 clicks)", "Ad Group", "Negative Exact"),
    ("vibration plate exercise machine", "Keyword", "SP_NB_AUTO_7Day", 19, 0, 44.65, 0.00, "Rule 1: AG 1x CFS Zero-Sale (>=7 clicks)", "Ad Group", "Negative Exact"),
    ("back stretcher lumbar support", "Keyword", "SP_NB_DISC_Broad", 15, 0, 32.70, 0.00, "Rule 1: AG 1x CFS Zero-Sale (>=7 clicks)", "Ad Group", "Negative Exact"),
    ("manual massage cane hook", "Keyword", "SP_NB_DISC_Broad", 12, 0, 25.08, 0.00, "Rule 1: AG 1x CFS Zero-Sale (>=7 clicks)", "Ad Group", "Negative Exact"),
    ("chiropractic adjusting tool", "Keyword", "SP_NB_AUTO_7Day", 11, 0, 26.95, 0.00, "Rule 1: AG 1x CFS Zero-Sale (>=7 clicks)", "Ad Group", "Negative Exact"),
    ("massage gun bulk wholesale", "Keyword", "SP_NB_DISC_Broad", 10, 0, 24.50, 0.00, "Rule 1: AG 1x CFS Zero-Sale (>=7 clicks)", "Ad Group", "Negative Exact"),
    ("heated massage gun under 40", "Keyword", "SP_NB_DISC_Broad", 52, 1, 118.56, 69.99, "Rule 5: 3x Target ACoS Catastrophic (>=84%)", "Campaign", "Negative Exact"),
    ("massage gun for kids", "Keyword", "SP_NB_AUTO_7Day", 33, 1, 74.91, 69.99, "Rule 5: 3x Target ACoS Catastrophic (>=84%)", "Campaign", "Negative Exact"),
    ("best budget massage gun", "Keyword", "SP_NB_DISC_Broad", 41, 1, 95.12, 69.99, "Rule 5: 3x Target ACoS Catastrophic (>=84%)", "Campaign", "Negative Exact"),
    ("full body massage mat", "Keyword", "SP_NB_AUTO_7Day", 38, 1, 87.78, 69.99, "Rule 5: 3x Target ACoS Catastrophic (>=84%)", "Campaign", "Negative Exact"),
    ("massage gun replacement tips only", "Keyword", "SP_NB_DISC_Broad", 29, 1, 63.80, 69.99, "Rule 5: 3x Target ACoS Catastrophic (>=84%)", "Campaign", "Negative Exact"),
    ("asin=\"B07Z9P1K8M\"", "PAT ASIN", "SP_OFF_PAT", 18, 0, 41.40, 0.00, "Rule 7: PAT 2x CFS Zero-Sale (>=10 clicks)", "Ad Group", "Negative Product"),
    ("asin=\"B08C3V7R2N\"", "PAT ASIN", "SP_OFF_PAT", 15, 0, 34.50, 0.00, "Rule 7: PAT 2x CFS Zero-Sale (>=10 clicks)", "Ad Group", "Negative Product"),
    ("asin=\"B09H6W2T5L\"", "PAT ASIN", "SP_AUTO_Substitutes", 14, 0, 31.08, 0.00, "Rule 7: PAT 2x CFS Zero-Sale (>=10 clicks)", "Ad Group", "Negative Product"),
    ("asin=\"B0B4K8M1Q9\"", "PAT ASIN", "SP_OFF_PAT", 12, 0, 27.96, 0.00, "Rule 7: PAT 2x CFS Zero-Sale (>=10 clicks)", "Ad Group", "Negative Product"),
    ("asin=\"B0C2N9X4P6\"", "PAT ASIN", "SP_OFF_PAT", 11, 0, 24.86, 0.00, "Rule 7: PAT 2x CFS Zero-Sale (>=10 clicks)", "Ad Group", "Negative Product"),
    ("asin=\"B07Y5Q3J8W\"", "PAT ASIN", "SP_AUTO_Substitutes", 10, 0, 22.10, 0.00, "Rule 7: PAT 2x CFS Zero-Sale (>=10 clicks)", "Ad Group", "Negative Product"),
]

def build_tab6(wb):
    ws = wb.create_sheet("3x Negation Shield")
    style_banner(
        ws, 11,
        "FULL CIRCLE 3x NEGATION SHIELD — ZERO-SALE & CATASTROPHIC ACOS CUTS",
        f"Balanced Stage Thresholds (negation-evaluator.ts): CFS={CFS} · AG Zero-Sale>={NEG_AG_CLICKS} clk · ASIN 3x Zero-Sale>={NEG_ASIN_CLICKS} clk · 3x ACoS>={NEG_ACOS_LIMIT:.0%} · 2-3d lag ignored"
    )
    headers = [
        "Search Term / Target ASIN", "Target Type", "Source Campaign(s)", "30d Clicks",
        "30d Orders", "30d Wasted Spend", "30d Sales", "Actual ACoS",
        "Negation Rule Triggered", "Apply Level", "Negative Match Type"
    ]
    style_table_header(ws, 4, headers)

    for idx, row in enumerate(NEGATION_ROWS, start=5):
        term, ttype, camp, clicks, orders, spend, sales, rule, lvl, nmatch = row
        ws.row_dimensions[idx].height = 21
        fill = make_fill(ROW_ALT) if idx % 2 == 1 else make_fill(WHITE)
        vals = [
            term, ttype, camp, clicks, orders, spend, sales,
            f'=IF(G{idx}>0,F{idx}/G{idx},"∞ (0 Sales)")',
            rule, lvl, nmatch
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=idx, column=c_idx, value=val)
            cell.font = make_font(10, bold=(c_idx in (1, 6)))
            cell.fill = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="center", horizontal="left" if c_idx in (1, 3, 9) else ("center" if c_idx in (2, 10, 11) else "right"))

        ws.cell(row=idx, column=4).number_format = "#,##0"
        ws.cell(row=idx, column=5).number_format = "#,##0"
        ws.cell(row=idx, column=6).number_format = "$#,##0.00"
        ws.cell(row=idx, column=7).number_format = "$#,##0.00"
        if sales > 0:
            ws.cell(row=idx, column=8).number_format = "0.0%"
        ws.cell(row=idx, column=8).font = make_font(9.5, True, RED_DK)
        ws.cell(row=idx, column=11).fill = make_fill(RED_LT)
        ws.cell(row=idx, column=11).font = make_font(9, True, RED_DK)

    # Totals Row
    tot = 5 + len(NEGATION_ROWS)
    ws.row_dimensions[tot].height = 25
    ws.cell(row=tot, column=1, value=f"3x NEGATION SHIELD TOTAL ({len(NEGATION_ROWS)} TARGETS CUT)")
    ws.cell(row=tot, column=4, value=f"=SUM(D5:D{tot-1})").number_format = "#,##0"
    ws.cell(row=tot, column=5, value=f"=SUM(E5:E{tot-1})").number_format = "#,##0"
    ws.cell(row=tot, column=6, value=f"=SUM(F5:F{tot-1})").number_format = "$#,##0.00"
    ws.cell(row=tot, column=7, value=f"=SUM(G5:G{tot-1})").number_format = "$#,##0.00"
    ws.cell(row=tot, column=9, value="Direct 30d bleed cut ($19,466/yr) + 90d tail shield = $34,968/yr saved")
    for c in range(1, 12):
        cell = ws.cell(row=tot, column=c)
        cell.fill = make_fill(SLATE_HDR)
        cell.font = make_font(10, True, WHITE)
        cell.border = CELL_BORDER

    set_col_widths(ws, [34, 13, 22, 11, 11, 17, 13, 14, 38, 13, 18])
    ws.freeze_panes = "B5"


def main():
    wb = Workbook()
    build_tab1(wb)
    build_tab2(wb)
    build_tab3(wb)
    build_tab4(wb)
    build_tab5(wb)
    build_tab6(wb)

    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "sample-hero-asin-campaign-architecture.xlsx")
    wb.save(out_path)
    print(f"✓ Generated 6-tab workbook: {out_path}")
    print(f"  Hero ASIN: {HERO['asin']} ({HERO['brand']} · ${HERO['annual_rev']:,.0f}/yr)")
    print(f"  Break-Even ACoS: {BREAKEVEN_ACOS:.1%} | Target ACoS: {HERO['target_acos']:.1%} | CFS: {CFS}")
    print(f"  Perfect CPC: ${PERFECT_CPC:.2f} (MIN(Anchor=${ANCHOR_CPC:.2f}, Market=${MARKET_CPC:.2f}))")
    print(f"  Tabs: {', '.join(wb.sheetnames)}")


if __name__ == "__main__":
    main()

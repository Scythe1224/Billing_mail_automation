"""
Billing Report Generator + Auto Email (Outlook / Office 365)
=============================================================
Reads an Excel rental report containing:
  - an "Old" sheet (previous month, case-insensitive)
  - a current-month sheet (name auto-detected from any common
    month format: Jul, July, Jul-26, July-2026, Aug-26, ...)

Produces a professional HTML dashboard summarising:
  - Total Rental Amount (current vs old, delta, %)
  - Final Payable          (current vs old, delta, %)
  - Expected Collection breakdown (gross / wallet-credit / net)
  - Product-wise client counts
  - New clients / discontinued clients
  - Changed rental amounts (top 25 inline, full list attached as CSV)
  - Data-issue report (missing IDs, duplicates, invalid values)

SENDER_EMAIL and SENDER_PASSWORD are read from environment variables.
RECIPIENTS env var (comma separated) sets the receiver; default list used if empty.

Usage:
  python billing_report_generator.py <excel_file> [sheet_name]

  Examples:
    python billing_report_generator.py Rental_Report_for_Jul-26.xlsx
    python billing_report_generator.py Rental_Report_for_Aug-26.xlsx Aug-26
"""

import io
import os
import sys
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime


# ─────────────────────────────────────────────────────────────
#  EMAIL CONFIG
# ─────────────────────────────────────────────────────────────

SMTP_SERVER     = "smtp.office365.com"
SMTP_PORT       = 587

SENDER_EMAIL    = os.environ.get("SENDER_EMAIL", "")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "")

DEFAULT_RECIPIENTS = [
    "rajesh.narula@reliablesoft.co.in",
    "admin@reliablesoft.co.in",
]
# Receiver is now a variable: set via RECIPIENTS env var (comma/semicolon
# separated) - the dashboard passes it in. Falls back to the default list.
_env_rcpt = os.environ.get("RECIPIENTS", "").replace(";", ",")
RECIPIENTS = [r.strip() for r in _env_rcpt.split(",") if r.strip()] or DEFAULT_RECIPIENTS

CC_RECIPIENTS = [
    # "manager@yourdomain.com",
]

# ─────────────────────────────────────────────────────────────
#  DATA CONTRACT
# ─────────────────────────────────────────────────────────────

REQUIRED_COLUMNS = [
    "Cable Operator ID",
    "R-Soft/On-Net",
    "MINIMUM BILLING",
    "Total Rental Amount",
    "Final Payable",
]

PRODUCT_ORDER = [
    "CNMS ON-Net Professional",
    "CNMS ON-Net Standard",
    "R-Soft SMS",
    "NetManazer",
    "R-Soft IPTV",
]

MAX_INLINE_DECREASES = 5     # top-N biggest drops shown in email body
CHANGE_THRESHOLD     = 1.0   # ignore rental deltas under ₹1 (rounding noise)

# ─────────────────────────────────────────────────────────────
#  COLOUR PALETTE
# ─────────────────────────────────────────────────────────────

COLOR_PRIMARY = "#1a3c5e"
COLOR_ACCENT  = "#0078d4"
COLOR_NEW     = "#107c41"
COLOR_DISC    = "#c50f1f"
COLOR_WARN    = "#b7791f"
COLOR_ROW_ALT = "#f9f9f9"


# ─────────────────────────────────────────────────────────────
#  CONFIG VALIDATION
# ─────────────────────────────────────────────────────────────

def validate_config():
    missing = []
    if not SENDER_EMAIL:
        missing.append("SENDER_EMAIL  (environment variable)")
    if not SENDER_PASSWORD:
        missing.append("SENDER_PASSWORD  (environment variable)")
    if not RECIPIENTS:
        missing.append("RECIPIENTS  (add at least one email in the script)")
    if missing:
        print("\nERROR: Missing required config:")
        for v in missing:
            print(f"   - {v}")
        sys.exit(1)


# ─────────────────────────────────────────────────────────────
#  SHEET DETECTION
# ─────────────────────────────────────────────────────────────

# Order matters: try longest / most specific formats first.
_MONTH_FORMATS = [
    "%B-%Y", "%B %Y",           # July-2026, July 2026
    "%B-%y", "%B %y",           # July-26,   July 26
    "%b-%Y", "%b %Y",           # Jul-2026,  Jul 2026
    "%b-%y", "%b %y",           # Jul-26,    Jul 26
    "%B", "%b",                 # July,      Jul
]


def _parse_month(name):
    """Return a datetime for a month-like sheet name, or None."""
    n = str(name).strip()
    for fmt in _MONTH_FORMATS:
        try:
            dt = datetime.strptime(n, fmt)
            # Bare month name (no year) — assume current year
            if fmt in ("%B", "%b"):
                dt = dt.replace(year=datetime.now().year)
            return dt
        except ValueError:
            continue
    return None


def detect_sheets(filepath, forced_current=None):
    """Locate the 'Old' sheet and the current-month sheet."""
    xl = pd.ExcelFile(filepath)
    sheets = xl.sheet_names

    old_sheet = next((s for s in sheets if s.strip().lower() == "old"), None)
    if not old_sheet:
        raise ValueError(
            "Could not find an 'Old' sheet (case-insensitive) in the workbook.\n"
            f"Sheets present: {sheets}"
        )

    if forced_current:
        if forced_current not in sheets:
            raise ValueError(
                f"Sheet '{forced_current}' not found. Available: {sheets}"
            )
        return old_sheet, forced_current

    candidates = []
    for s in sheets:
        if s == old_sheet:
            continue
        dt = _parse_month(s)
        if dt:
            candidates.append((s, dt))

    if not candidates:
        raise ValueError(
            "Could not auto-detect the current-month sheet. Expected a sheet "
            "named like 'Jul-26', 'July-2026', 'Aug', etc.\n"
            f"Sheets present: {sheets}"
        )

    candidates.sort(key=lambda x: x[1], reverse=True)
    return old_sheet, candidates[0][0]


def month_label(sheet_name):
    """Pretty label like 'July 2026' for headers/subjects."""
    dt = _parse_month(sheet_name)
    return dt.strftime("%B %Y") if dt else sheet_name


# ─────────────────────────────────────────────────────────────
#  DATA VALIDATION
# ─────────────────────────────────────────────────────────────

class ValidationError(Exception):
    pass


def validate_sheet_columns(df, sheet_name):
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        return [
            f"Sheet '{sheet_name}' is missing required column: '{c}'"
            for c in missing
        ]
    return []


def validate_workbook(old_df, cur_df, old_name, cur_name):
    errors = []
    errors.extend(validate_sheet_columns(old_df, old_name))
    errors.extend(validate_sheet_columns(cur_df, cur_name))
    if errors:
        raise ValidationError("\n".join(errors))


# ─────────────────────────────────────────────────────────────
#  DATA PREPARATION
# ─────────────────────────────────────────────────────────────

def prepare(df):
    """Normalise IDs, coerce numerics, fill NaN with 0."""
    df = df.copy()
    df["_id"] = df["Cable Operator ID"].astype(str).str.strip().str.upper()
    for c in ("Total Rental Amount", "Final Payable", "MINIMUM BILLING"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["ADJ"] = pd.to_numeric(df["ADJ"], errors="coerce").fillna(0) if "ADJ" in df.columns else 0.0
    df["R-Soft/On-Net"] = df["R-Soft/On-Net"].fillna("(Unspecified)").astype(str).str.strip()
    return df


def find_data_issues(df, sheet_name):
    """Return a dict of non-fatal data issues per sheet."""
    issues = {"sheet": sheet_name, "missing_id": [], "duplicates": [], "missing_product": []}

    empty_id = df["Cable Operator ID"].isna() | (
        df["Cable Operator ID"].astype(str).str.strip().isin(["", "nan", "NaN"])
    )
    if empty_id.any():
        if "S. No." in df.columns:
            issues["missing_id"] = [
                int(x) if pd.notna(x) else "?"
                for x in df.loc[empty_id, "S. No."].tolist()
            ]
        else:
            issues["missing_id"] = [int(i) + 2 for i in df.index[empty_id].tolist()]

    with_id = df.loc[~empty_id]
    dup = with_id["_id"][with_id["_id"].duplicated(keep=False)]
    issues["duplicates"] = sorted(dup.unique().tolist())

    prod_missing = df["R-Soft/On-Net"].astype(str).str.strip().eq("(Unspecified)")
    if prod_missing.any():
        issues["missing_product"] = df.loc[prod_missing & ~empty_id, "_id"].tolist()

    return issues


def has_issues(issues):
    return bool(issues["missing_id"] or issues["duplicates"] or issues["missing_product"])


# ─────────────────────────────────────────────────────────────
#  METRICS
# ─────────────────────────────────────────────────────────────

def _split_positive_negative(series):
    positive = series[series > 0].sum()
    negative = series[series < 0].sum()
    return float(positive), float(negative)


def collection_metrics(df):
    """
    Rental columns rarely go negative (pure rental owed).
    Final Payable = Rental +/- adjustments; negatives here are
    the customers whose wallet/credit exceeds their rental, so we
    treat those negatives as wallet/credit for the expected-collection
    breakdown.
    """
    rental_pos, rental_neg = _split_positive_negative(df["Total Rental Amount"])
    payable_pos, payable_neg = _split_positive_negative(df["Final Payable"])
    return {
        "rental_gross":   rental_pos,
        "rental_wallet":  abs(rental_neg),
        "rental_net":     rental_pos + rental_neg,
        "payable_gross":  payable_pos,
        "payable_wallet": abs(payable_neg),
        "payable_net":    payable_pos + payable_neg,
        # Wallet count is driven by Final Payable (that's where wallets
        # actually manifest as negatives in this workbook).
        "wallet_count":   int((df["Final Payable"] < 0).sum()),
        # ADJ column: + = dues, - = wallet/adjustments
        "invoice_amt":    rental_pos + rental_neg,
        "dues":           float(df["ADJ"][df["ADJ"] > 0].sum()),
        "adj_wallet":     abs(float(df["ADJ"][df["ADJ"] < 0].sum())),
        "adj_wallet_cnt": int((df["ADJ"] < 0).sum()),
        "dues_cnt":       int((df["ADJ"] > 0).sum()),
        "total_clients":  len(df),
    }


def product_counts(df):
    return df.groupby("R-Soft/On-Net").size().to_dict()


def ordered_products(*count_dicts):
    all_p = set()
    for d in count_dicts:
        all_p.update(d.keys())
    tail = sorted(p for p in all_p if p not in PRODUCT_ORDER)
    return [p for p in PRODUCT_ORDER if p in all_p] + tail


# ─────────────────────────────────────────────────────────────
#  DIFF
# ─────────────────────────────────────────────────────────────

def new_clients(old_df, cur_df):
    ids = set(cur_df["_id"]) - set(old_df["_id"])
    return cur_df[cur_df["_id"].isin(ids)][
        ["Cable Operator ID", "Company Name", "R-Soft/On-Net",
         "MINIMUM BILLING", "Total Rental Amount"]
    ].copy()


def discontinued_clients(old_df, cur_df):
    ids = set(old_df["_id"]) - set(cur_df["_id"])
    return old_df[old_df["_id"].isin(ids)][
        ["Cable Operator ID", "Company Name", "R-Soft/On-Net",
         "MINIMUM BILLING", "Total Rental Amount"]
    ].copy()


def changed_rentals(old_df, cur_df):
    old_slim = old_df[["_id", "Total Rental Amount"]].rename(
        columns={"Total Rental Amount": "Rental_Old"}
    )
    cur_slim = cur_df[["_id", "Cable Operator ID", "Company Name",
                       "R-Soft/On-Net", "Total Rental Amount"]].rename(
        columns={"Total Rental Amount": "Rental_Cur"}
    )
    merged = cur_slim.merge(old_slim, on="_id", how="inner")
    merged["Delta"] = merged["Rental_Cur"] - merged["Rental_Old"]
    merged = merged[merged["Delta"].abs() >= CHANGE_THRESHOLD]
    merged["Abs_Delta"] = merged["Delta"].abs()
    return merged.sort_values("Abs_Delta", ascending=False).drop(columns=["Abs_Delta"])


# ─────────────────────────────────────────────────────────────
#  FORMATTERS
# ─────────────────────────────────────────────────────────────

def _inr(n):
    """Indian-format numeric string with rupee sign, e.g. ₹1,23,456.78."""
    try:
        n = float(n)
    except (TypeError, ValueError):
        return "-"
    sign = "-" if n < 0 else ""
    n = abs(n)
    whole, dec = f"{n:.2f}".split(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        head = ",".join(
            [head[max(i - 2, 0):i] for i in range(len(head), 0, -2)][::-1]
        )
        whole = f"{head},{tail}"
    return f"{sign}₹{whole}.{dec}"


def _pct(delta, base):
    if not base:
        return "n/a"
    p = (delta / base) * 100
    sign = "+" if p >= 0 else ""
    return f"{sign}{p:.2f}%"


def _signed(delta):
    return f"+{_inr(delta)}" if delta >= 0 else _inr(delta)


# ─────────────────────────────────────────────────────────────
#  PLAIN TEXT REPORT (email fallback + log artifact)
# ─────────────────────────────────────────────────────────────

def format_plain_text(context):
    old_m   = context["old_metrics"]
    cur_m   = context["cur_metrics"]
    new_df  = context["new_df"]
    disc_df = context["disc_df"]
    chg_df  = context["chg_df"]
    old_iss = context["old_issues"]
    cur_iss = context["cur_issues"]
    cur_sheet = context["current_sheet"]
    label = month_label(cur_sheet)

    L = ["Billing Details Update", "=" * 40, ""]
    L.append(f"Month:     {label}   (sheet: {cur_sheet})")
    L.append(f"Generated: {datetime.now().strftime('%d %b %Y, %I:%M %p')}")
    L.append("")
    L.append("-- Collection Summary --")
    L.append(f"Generated Invoice Amount (Without Dues)   Old: {_inr(old_m['rental_net'])}"
             f"   Current: {_inr(cur_m['rental_net'])}"
             f"   Delta: {_signed(cur_m['rental_net'] - old_m['rental_net'])}"
             f"   ({_pct(cur_m['rental_net'] - old_m['rental_net'], old_m['rental_net'])})")
    L.append(f"Final Payable (With Dues)   Old: {_inr(old_m['payable_net'])}"
             f"   Current: {_inr(cur_m['payable_net'])}"
             f"   Delta: {_signed(cur_m['payable_net'] - old_m['payable_net'])}"
             f"   ({_pct(cur_m['payable_net'] - old_m['payable_net'], old_m['payable_net'])})")
    L.append("")
    L.append("-- Final Expected Collection (Current Month) --")
    L.append(f"Generated Invoice Amount:    {_inr(cur_m['invoice_amt'])}")
    L.append(f"Dues (+ ADJ):               +{_inr(cur_m['dues'])}   [{cur_m['dues_cnt']} client(s)]")
    L.append(f"Wallet/Adjustments (- ADJ): -{_inr(cur_m['adj_wallet'])}   [{cur_m['adj_wallet_cnt']} client(s)]")
    L.append(f"Final Receivable:            {_inr(cur_m['payable_net'])}")
    L.append("")
    L.append(f"Total Clients   Old: {old_m['total_clients']}   "
             f"Current: {cur_m['total_clients']}")
    L.append(f"New: {len(new_df)}   Discontinued: {len(disc_df)}   "
             f"Rental changes: {len(chg_df)}")
    L.append("")

    if not new_df.empty:
        L.append("-- New Clients This Month --")
        for _, r in new_df.iterrows():
            L.append(f"  {r['Cable Operator ID']:<20} {r['R-Soft/On-Net']:<28} "
                     f"Rental: {_inr(r['Total Rental Amount'])}")
        L.append("")

    if not disc_df.empty:
        L.append("-- Discontinued Clients --")
        for _, r in disc_df.iterrows():
            L.append(f"  {r['Cable Operator ID']:<20} {r['R-Soft/On-Net']:<28} "
                     f"Rental: {_inr(r['Total Rental Amount'])}")
        L.append("")

    if not chg_df.empty:
        decreases = chg_df[chg_df["Delta"] < 0]
        increases = chg_df[chg_df["Delta"] > 0]
        L.append(f"-- Rental Changes vs Old Sheet --")
        L.append(f"   Total changed: {len(chg_df)}   "
                 f"Decreased: {len(decreases)}   Increased: {len(increases)}")
        L.append(f"   Total decrease: -{_inr(abs(decreases['Delta'].sum()))}   "
                 f"Total increase: +{_inr(increases['Delta'].sum())}")
        L.append("")
        if not decreases.empty:
            L.append(f"-- Top {min(MAX_INLINE_DECREASES, len(decreases))} "
                     f"Rental Decreases --")
            for _, r in decreases.head(MAX_INLINE_DECREASES).iterrows():
                L.append(f"  {r['Cable Operator ID']:<20} "
                         f"Old: {_inr(r['Rental_Old'])}  ->  "
                         f"New: {_inr(r['Rental_Cur'])}  "
                         f"({_signed(r['Delta'])})")
            if len(decreases) > MAX_INLINE_DECREASES:
                L.append(f"  ... and {len(decreases) - MAX_INLINE_DECREASES} "
                         f"more decrease(s) in the attached CSV")
        L.append("  (Full change list — all decreases and increases — attached as CSV)")
        L.append("")

    for iss in (old_iss, cur_iss):
        if has_issues(iss):
            L.append(f"-- Data Issues in '{iss['sheet']}' --")
            if iss["missing_id"]:
                L.append(f"  Missing Cable Operator ID at rows: {iss['missing_id']}")
            if iss["duplicates"]:
                L.append(f"  Duplicate IDs: {iss['duplicates']}")
            if iss["missing_product"]:
                L.append(f"  Missing product tag for: {iss['missing_product']}")
            L.append("")

    return "\n".join(L)


# ─────────────────────────────────────────────────────────────
#  HTML BUILDERS
# ─────────────────────────────────────────────────────────────

def _card_collection(title, old_val, cur_val, accent):
    delta = cur_val - old_val
    pct   = _pct(delta, old_val)
    up    = delta >= 0
    arrow = "&#9650;" if up else "&#9660;"
    delta_color = COLOR_NEW if up else COLOR_DISC
    return f"""
    <td width="48%" valign="top" style="background:#ffffff;border:1px solid #e5e7eb;
      border-radius:10px;padding:0;overflow:hidden;">
      <div style="background:{accent};padding:10px 16px;color:#fff;font-size:13px;
        font-weight:700;letter-spacing:.2px;">{title}</div>
      <div style="padding:14px 16px 16px;">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td style="font-size:11px;color:#6b7280;padding-bottom:2px;">Previous Month</td>
            <td align="right" style="font-size:11px;color:#6b7280;padding-bottom:2px;">Current Month</td>
          </tr>
          <tr>
            <td style="font-size:16px;font-weight:700;color:#374151;">{_inr(old_val)}</td>
            <td align="right" style="font-size:18px;font-weight:800;color:{COLOR_PRIMARY};">{_inr(cur_val)}</td>
          </tr>
          <tr><td colspan="2" style="border-top:1px dashed #e5e7eb;padding-top:10px;padding-bottom:2px;">
            <span style="font-size:12px;color:#6b7280;">Change</span>
            <span style="float:right;font-size:13px;font-weight:700;color:{delta_color};">
              {arrow} {_signed(delta)} &nbsp;({pct})
            </span>
          </td></tr>
        </table>
      </div>
    </td>"""


def _card_expected(cur_m):
    inv, dues = cur_m["invoice_amt"], cur_m["dues"]
    wallet, net = cur_m["adj_wallet"], cur_m["payable_net"]
    gross = inv + dues
    bar_wallet = min((wallet / gross * 100) if gross else 0, 100)
    bar_net = 100 - bar_wallet
    def cell(w, label, val, color, pad):
        return (f'<td width="{w}%" valign="top" style="{pad}"><div style="font-size:11px;color:#6b7280;">{label}</div>'
                f'<div style="font-size:16px;font-weight:800;color:{color};">{val}</div></td>')
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;
      border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;margin-top:12px;">
      <tr><td style="background:{COLOR_PRIMARY};padding:10px 16px;">
        <span style="color:#fff;font-size:13px;font-weight:700;">Final Expected Collection (Current Month)</span>
      </td></tr>
      <tr><td style="padding:16px;">
        <table width="100%" cellpadding="0" cellspacing="0"><tr>
          {cell(25, "Generated Invoice Amount", _inr(inv), "#111827", "padding:0 6px 0 0;")}
          {cell(25, f"Dues <span style='color:#9ca3af;'>({cur_m['dues_cnt']})</span>", "+" + _inr(dues), "#b45309", "padding:0 6px;")}
          {cell(25, f"Wallet / Adjustments <span style='color:#9ca3af;'>({cur_m['adj_wallet_cnt']})</span>", "-" + _inr(wallet), COLOR_DISC, "padding:0 6px;")}
          {cell(25, "Final Receivable", _inr(net), COLOR_NEW, "padding:0 0 0 6px;")}
        </tr></table>
        <div style="margin-top:14px;height:10px;background:#f3f4f6;border-radius:6px;overflow:hidden;">
          <div style="width:{bar_net:.1f}%;height:100%;background:{COLOR_NEW};float:left;"></div>
          <div style="width:{bar_wallet:.1f}%;height:100%;background:{COLOR_DISC};float:left;"></div>
        </div>
        <div style="margin-top:6px;font-size:11px;color:#6b7280;">
          <span style="display:inline-block;width:8px;height:8px;background:{COLOR_NEW};border-radius:2px;vertical-align:middle;"></span>&nbsp;Final Receivable&nbsp;&nbsp;
          <span style="display:inline-block;width:8px;height:8px;background:{COLOR_DISC};border-radius:2px;vertical-align:middle;"></span>&nbsp;Wallet/Adjustments
        </div>
      </td></tr>
    </table>"""


def _product_rows(counts, ordered):
    rows = ""
    for i, p in enumerate(ordered):
        if p not in counts:
            continue
        bg = "#ffffff" if i % 2 == 0 else COLOR_ROW_ALT
        rows += (
            f'<tr style="background:{bg};">'
            f'<td style="padding:8px 14px;color:#333;font-size:13px;">{p}</td>'
            f'<td style="padding:8px 14px;text-align:center;font-weight:700;'
            f'color:{COLOR_ACCENT};font-size:14px;">{counts[p]}</td>'
            f'</tr>'
        )
    return rows


def _client_group_table(df, accent):
    if df.empty:
        return '<p style="color:#888;font-style:italic;margin:6px 0 0;">None this month</p>'
    html = ""
    for product, grp in df.groupby("R-Soft/On-Net"):
        html += (
            f'<p style="margin:12px 0 4px;font-weight:700;color:#333;font-size:13px;">'
            f'{product}<span style="font-weight:400;color:#666;"> &nbsp;'
            f'({len(grp)} client{"s" if len(grp) > 1 else ""})</span></p>'
            f'<table width="100%" cellpadding="0" cellspacing="0" '
            f'style="border-collapse:collapse;border:1px solid #e0e0e0;'
            f'border-radius:6px;overflow:hidden;">'
            f'<thead><tr style="background:#f5f5f5;">'
            f'<th style="padding:7px 14px;text-align:left;font-size:12px;color:#555;'
            f'font-weight:600;border-bottom:1px solid #e0e0e0;">Client ID</th>'
            f'<th style="padding:7px 14px;text-align:left;font-size:12px;color:#555;'
            f'font-weight:600;border-bottom:1px solid #e0e0e0;">Company</th>'
            f'<th style="padding:7px 14px;text-align:right;font-size:12px;color:#555;'
            f'font-weight:600;border-bottom:1px solid #e0e0e0;">Total Rental</th>'
            f'</tr></thead><tbody>'
        )
        for i, (_, row) in enumerate(grp.iterrows()):
            bg = "#ffffff" if i % 2 == 0 else COLOR_ROW_ALT
            company = row.get("Company Name") or "-"
            html += (
                f'<tr style="background:{bg};">'
                f'<td style="padding:7px 14px;font-size:13px;color:#333;'
                f'font-family:Consolas,monospace;">{row["Cable Operator ID"]}</td>'
                f'<td style="padding:7px 14px;font-size:13px;color:#333;">{company}</td>'
                f'<td style="padding:7px 14px;text-align:right;font-size:13px;'
                f'font-weight:700;color:{accent};">{_inr(row["Total Rental Amount"])}</td>'
                f'</tr>'
            )
        html += "</tbody></table>"
    return html


def _changes_table(chg_df):
    if chg_df.empty:
        return '<p style="color:#888;font-style:italic;margin:6px 0 0;">No rental amount changes</p>'

    decreases = chg_df[chg_df["Delta"] < 0]
    increases = chg_df[chg_df["Delta"] > 0]
    total_dec = abs(decreases["Delta"].sum()) if not decreases.empty else 0.0
    total_inc = increases["Delta"].sum() if not increases.empty else 0.0

    # Summary strip
    summary = (
        f'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:10px;"><tr>'
        f'<td width="33%" style="padding:8px 12px;background:#fff5f5;border:1px solid #ffcdd2;border-radius:6px;">'
        f'<div style="font-size:11px;color:#6b7280;">Decreased</div>'
        f'<div style="font-size:14px;font-weight:800;color:{COLOR_DISC};">{len(decreases)} client(s)</div>'
        f'<div style="font-size:11px;color:{COLOR_DISC};">-{_inr(total_dec)}</div>'
        f'</td><td width="2%"></td>'
        f'<td width="33%" style="padding:8px 12px;background:#f0fff4;border:1px solid #b2dfdb;border-radius:6px;">'
        f'<div style="font-size:11px;color:#6b7280;">Increased</div>'
        f'<div style="font-size:14px;font-weight:800;color:{COLOR_NEW};">{len(increases)} client(s)</div>'
        f'<div style="font-size:11px;color:{COLOR_NEW};">+{_inr(total_inc)}</div>'
        f'</td><td width="2%"></td>'
        f'<td width="30%" style="padding:8px 12px;background:#f5f9ff;border:1px solid #cfe0f5;border-radius:6px;">'
        f'<div style="font-size:11px;color:#6b7280;">Net Change</div>'
        f'<div style="font-size:14px;font-weight:800;color:{COLOR_PRIMARY};">'
        f'{_signed(total_inc - total_dec)}</div>'
        f'</td></tr></table>'
    )

    if decreases.empty:
        body = (
            '<p style="color:#666;font-size:12px;margin:6px 0 0;">'
            'No rental decreases this month. Full change list is attached as CSV.</p>'
        )
        return summary + body

    shown = decreases.head(MAX_INLINE_DECREASES)
    rest  = len(decreases) - len(shown)
    rows = ""
    for i, (_, r) in enumerate(shown.iterrows()):
        bg = "#ffffff" if i % 2 == 0 else COLOR_ROW_ALT
        company = r.get("Company Name") or "-"
        rows += (
            f'<tr style="background:{bg};">'
            f'<td style="padding:7px 12px;font-size:12px;font-family:Consolas,monospace;color:#333;">'
            f'{r["Cable Operator ID"]}</td>'
            f'<td style="padding:7px 12px;font-size:12px;color:#333;">{company}</td>'
            f'<td style="padding:7px 12px;font-size:12px;text-align:right;color:#6b7280;">'
            f'{_inr(r["Rental_Old"])}</td>'
            f'<td style="padding:7px 12px;font-size:12px;text-align:right;color:#111827;font-weight:700;">'
            f'{_inr(r["Rental_Cur"])}</td>'
            f'<td style="padding:7px 12px;font-size:12px;text-align:right;font-weight:700;color:{COLOR_DISC};">'
            f'&#9660; {_signed(r["Delta"])}</td>'
            f'</tr>'
        )
    footer = (
        f'<tr><td colspan="5" style="padding:8px 12px;font-size:11px;color:#6b7280;'
        f'background:#f9fafb;font-style:italic;">'
        f'Top {len(shown)} biggest decrease(s) shown. '
        f'{rest} more decrease(s) plus all increases in the attached CSV.</td></tr>'
        if rest > 0 else
        f'<tr><td colspan="5" style="padding:8px 12px;font-size:11px;color:#6b7280;'
        f'background:#f9fafb;font-style:italic;">'
        f'Full change list (decreases + increases) attached as CSV.</td></tr>'
    )
    heading = (
        f'<div style="font-size:12px;font-weight:700;color:{COLOR_DISC};margin:0 0 6px;">'
        f'Top {len(shown)} Rental Decreases</div>'
    )
    table = (
        f'<table width="100%" cellpadding="0" cellspacing="0" '
        f'style="border-collapse:collapse;border:1px solid #e0e0e0;'
        f'border-radius:6px;overflow:hidden;">'
        f'<thead><tr style="background:#f5f5f5;">'
        f'<th style="padding:7px 12px;text-align:left;font-size:12px;color:#555;font-weight:600;">Client ID</th>'
        f'<th style="padding:7px 12px;text-align:left;font-size:12px;color:#555;font-weight:600;">Company</th>'
        f'<th style="padding:7px 12px;text-align:right;font-size:12px;color:#555;font-weight:600;">Old Rental</th>'
        f'<th style="padding:7px 12px;text-align:right;font-size:12px;color:#555;font-weight:600;">New Rental</th>'
        f'<th style="padding:7px 12px;text-align:right;font-size:12px;color:#555;font-weight:600;">Decrease</th>'
        f'</tr></thead><tbody>{rows}{footer}</tbody></table>'
    )
    return summary + heading + table


def _issues_block(old_iss, cur_iss):
    if not (has_issues(old_iss) or has_issues(cur_iss)):
        return ""
    parts = []
    for iss in (old_iss, cur_iss):
        if not has_issues(iss):
            continue
        rows = ""
        if iss["missing_id"]:
            rows += (
                f'<li><b>Missing Cable Operator ID</b> at row(s): '
                f'{", ".join(str(x) for x in iss["missing_id"][:20])}'
                f'{" ..." if len(iss["missing_id"]) > 20 else ""}</li>'
            )
        if iss["duplicates"]:
            rows += (
                f'<li><b>Duplicate ID(s):</b> '
                f'{", ".join(iss["duplicates"][:20])}'
                f'{" ..." if len(iss["duplicates"]) > 20 else ""}</li>'
            )
        if iss["missing_product"]:
            rows += (
                f'<li><b>Missing product tag for:</b> '
                f'{", ".join(iss["missing_product"][:20])}'
                f'{" ..." if len(iss["missing_product"]) > 20 else ""}</li>'
            )
        parts.append(
            f'<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;'
            f'padding:12px 16px;margin-top:8px;">'
            f'<div style="font-weight:700;color:{COLOR_WARN};font-size:13px;margin-bottom:6px;">'
            f'&#9888; Data issues in sheet "{iss["sheet"]}"</div>'
            f'<ul style="margin:0 0 0 18px;padding:0;color:#374151;font-size:12px;line-height:1.7;">'
            f'{rows}</ul></div>'
        )
    return f"""
    <tr><td style="padding:8px 32px 20px;">
      <div style="border-left:4px solid {COLOR_WARN};padding-left:12px;margin-bottom:8px;">
        <span style="font-size:15px;font-weight:700;color:{COLOR_WARN};">
          &#9888; Exception Report
        </span>
      </div>
      {''.join(parts)}
    </td></tr>"""


def format_html_email(context):
    old_df   = context["old_df"]
    cur_df   = context["cur_df"]
    old_m    = context["old_metrics"]
    cur_m    = context["cur_metrics"]
    new_df   = context["new_df"]
    disc_df  = context["disc_df"]
    chg_df   = context["chg_df"]
    old_iss  = context["old_issues"]
    cur_iss  = context["cur_issues"]
    cur_sheet = context["current_sheet"]

    old_counts = product_counts(old_df)
    cur_counts = product_counts(cur_df)
    ordered    = ordered_products(old_counts, cur_counts)
    prev_total = sum(old_counts.values())
    cur_total  = sum(cur_counts.values())
    diff       = cur_total - prev_total
    diff_str   = f"+{diff}" if diff > 0 else str(diff)

    label = month_label(cur_sheet)
    generated = datetime.now().strftime("%d %b %Y, %I:%M %p")

    new_badge_bg  = COLOR_NEW  if not new_df.empty  else "#aaaaaa"
    disc_badge_bg = COLOR_DISC if not disc_df.empty else "#aaaaaa"
    chg_badge_bg  = COLOR_ACCENT if not chg_df.empty else "#aaaaaa"

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f0f2f5;font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f2f5;padding:24px 0;">
<tr><td align="center">
<table width="720" cellpadding="0" cellspacing="0"
  style="background:#ffffff;border-radius:10px;overflow:hidden;
         box-shadow:0 2px 16px rgba(0,0,0,0.10);max-width:720px;">

  <tr><td style="background:linear-gradient(135deg,{COLOR_PRIMARY} 0%,{COLOR_ACCENT} 100%);padding:28px 32px;">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td>
        <div style="font-size:11px;color:rgba(255,255,255,0.65);letter-spacing:1.5px;text-transform:uppercase;margin-bottom:5px;">Monthly Billing Report</div>
        <div style="font-size:22px;font-weight:700;color:#ffffff;line-height:1.2;">&#128226; Billing Details Update</div>
        <div style="font-size:13px;color:rgba(255,255,255,0.8);margin-top:5px;">&#128197; {label}</div>
      </td>
      <td align="right" valign="middle">
        <div style="background:rgba(255,255,255,0.15);border-radius:10px;padding:12px 20px;text-align:center;display:inline-block;">
          <div style="font-size:11px;color:rgba(255,255,255,0.7);margin-bottom:2px;">Net Client Change</div>
          <div style="font-size:28px;font-weight:800;color:#ffffff;">{diff_str}</div>
          <div style="font-size:11px;color:rgba(255,255,255,0.7);margin-top:2px;">clients</div>
        </div>
      </td>
    </tr></table>
  </td></tr>

  <tr><td style="background:#fafafa;padding:14px 32px;border-bottom:1px solid #eeeeee;">
    <table cellpadding="0" cellspacing="0"><tr>
      <td style="padding-right:8px;"><span style="background:{new_badge_bg};color:#fff;border-radius:20px;padding:5px 14px;font-size:12px;font-weight:600;">&#127379; {len(new_df)} New</span></td>
      <td style="padding-right:8px;"><span style="background:{disc_badge_bg};color:#fff;border-radius:20px;padding:5px 14px;font-size:12px;font-weight:600;">&#10060; {len(disc_df)} Discontinued</span></td>
      <td style="padding-right:8px;"><span style="background:{chg_badge_bg};color:#fff;border-radius:20px;padding:5px 14px;font-size:12px;font-weight:600;">&#8646; {len(chg_df)} Rental changes</span></td>
      <td><span style="background:{COLOR_PRIMARY};color:#fff;border-radius:20px;padding:5px 14px;font-size:12px;font-weight:600;">&#128100; {cur_m['total_clients']} Total clients</span></td>
    </tr></table>
  </td></tr>

  <tr><td style="padding:20px 32px 6px;">
    <div style="font-size:11px;letter-spacing:1px;color:#6b7280;text-transform:uppercase;margin-bottom:10px;">Collection Comparison</div>
    <table width="100%" cellpadding="0" cellspacing="0"><tr valign="top">
      {_card_collection("Generated Invoice Amount (Without Dues)", old_m["rental_net"], cur_m["rental_net"], COLOR_ACCENT)}
      <td width="4%"></td>
      {_card_collection("Final Payable (With Dues)", old_m["payable_net"], cur_m["payable_net"], COLOR_NEW)}
    </tr></table>
    {_card_expected(cur_m)}
  </td></tr>

  <tr><td style="padding:24px 32px 8px;">
    <div style="font-size:11px;letter-spacing:1px;color:#6b7280;text-transform:uppercase;margin-bottom:10px;">Product-wise Client Count</div>
    <table width="100%" cellpadding="0" cellspacing="0"><tr valign="top">
      <td width="48%" style="background:#eaf2fb;border-radius:8px;overflow:hidden;border:1px solid #b8d6f0;">
        <div style="background:#1a6fad;padding:10px 14px;">
          <span style="color:#fff;font-size:13px;font-weight:700;">&#128197; Previous Month</span>
        </div>
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr style="background:#c5def5;">
            <td style="padding:9px 14px;font-size:12px;color:#1a3c5e;font-weight:700;">Total Clients</td>
            <td style="padding:9px 14px;text-align:center;font-size:16px;font-weight:800;color:#1a3c5e;">{prev_total}</td>
          </tr>
          {_product_rows(old_counts, ordered)}
        </table>
      </td>
      <td width="4%"></td>
      <td width="48%" style="background:#e8f5e9;border-radius:8px;overflow:hidden;border:1px solid #a5d6a7;">
        <div style="background:#107c41;padding:10px 14px;">
          <span style="color:#fff;font-size:13px;font-weight:700;">&#128197; {label}</span>
        </div>
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr style="background:#c8ecd8;">
            <td style="padding:9px 14px;font-size:12px;color:#0a4e28;font-weight:700;">Total Clients</td>
            <td style="padding:9px 14px;text-align:center;font-size:16px;font-weight:800;color:#0a4e28;">{cur_total}</td>
          </tr>
          {_product_rows(cur_counts, ordered)}
        </table>
      </td>
    </tr></table>
  </td></tr>

  <tr><td style="padding:20px 32px 8px;">
    <div style="border-left:4px solid {COLOR_NEW};padding-left:12px;margin-bottom:12px;">
      <span style="font-size:15px;font-weight:700;color:{COLOR_NEW};">&#127379; New Billing This Month</span>
    </div>
    <div style="background:#f0fff4;border:1px solid #b2dfdb;border-radius:8px;padding:14px 16px;">
      {_client_group_table(new_df, COLOR_NEW)}
    </div>
  </td></tr>

  <tr><td style="padding:12px 32px 8px;">
    <div style="border-left:4px solid {COLOR_DISC};padding-left:12px;margin-bottom:12px;">
      <span style="font-size:15px;font-weight:700;color:{COLOR_DISC};">&#10060; Clients Discontinued This Month</span>
    </div>
    <div style="background:#fff5f5;border:1px solid #ffcdd2;border-radius:8px;padding:14px 16px;">
      {_client_group_table(disc_df, COLOR_DISC)}
    </div>
  </td></tr>

  <tr><td style="padding:12px 32px 8px;">
    <div style="border-left:4px solid {COLOR_ACCENT};padding-left:12px;margin-bottom:12px;">
      <span style="font-size:15px;font-weight:700;color:{COLOR_ACCENT};">&#8646; Rental Changes vs Old Sheet</span>
    </div>
    <div style="background:#f5f9ff;border:1px solid #cfe0f5;border-radius:8px;padding:14px 16px;">
      {_changes_table(chg_df)}
    </div>
  </td></tr>

  {_issues_block(old_iss, cur_iss)}

  <tr><td style="background:#f5f5f5;padding:14px 32px;border-top:1px solid #e0e0e0;border-radius:0 0 10px 10px;">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td style="font-size:11px;color:#999;">Auto-generated on {generated} &nbsp;|&nbsp; Source sheet: <b>{cur_sheet}</b></td>
      <td align="right" style="font-size:11px;color:#999;">Rental Billing System</td>
    </tr></table>
  </td></tr>

</table>
</td></tr>
</table>
</body></html>"""


# ─────────────────────────────────────────────────────────────
#  SUBJECT
# ─────────────────────────────────────────────────────────────

def build_subject(cur_sheet, new_df, disc_df, chg_df, cur_metrics):
    label = month_label(cur_sheet)
    parts = []
    if not new_df.empty:
        parts.append(f"{len(new_df)} New")
    if not disc_df.empty:
        parts.append(f"{len(disc_df)} Disc")
    if not chg_df.empty:
        parts.append(f"{len(chg_df)} Δ")
    detail = f" | {', '.join(parts)}" if parts else ""
    net = _inr(cur_metrics["rental_net"])
    return (f"Billing Update – {label} | "
            f"Clients: {cur_metrics['total_clients']} | Net Rental: {net}{detail}")


# ─────────────────────────────────────────────────────────────
#  CSV ATTACHMENT
# ─────────────────────────────────────────────────────────────

def changes_csv_bytes(chg_df):
    buf = io.StringIO()
    out = chg_df[["Cable Operator ID", "Company Name", "R-Soft/On-Net",
                  "Rental_Old", "Rental_Cur", "Delta"]].rename(columns={
        "Rental_Old": "Total Rental Amount (Old)",
        "Rental_Cur": "Total Rental Amount (Current)",
        "Delta":      "Change",
    })
    out.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")


# ─────────────────────────────────────────────────────────────
#  SEND EMAIL
# ─────────────────────────────────────────────────────────────

def send_email(subject, html_body, plain_body, attachments):
    outer = MIMEMultipart("mixed")
    outer["Subject"] = subject
    outer["From"]    = SENDER_EMAIL
    outer["To"]      = ", ".join(RECIPIENTS)
    if CC_RECIPIENTS:
        outer["Cc"]  = ", ".join(CC_RECIPIENTS)

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(plain_body, "plain", "utf-8"))
    alt.attach(MIMEText(html_body,  "html",  "utf-8"))
    outer.attach(alt)

    for filename, data in attachments:
        part = MIMEApplication(data, _subtype="octet-stream")
        part.add_header("Content-Disposition", "attachment", filename=filename)
        outer.attach(part)

    all_recipients = RECIPIENTS + CC_RECIPIENTS

    print(f"\nConnecting to {SMTP_SERVER}:{SMTP_PORT} ...")
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, all_recipients, outer.as_string())
        print(f"Email sent to: {', '.join(all_recipients)}")
    except smtplib.SMTPAuthenticationError:
        print("Authentication failed. If MFA is on, use an App Password.")
        sys.exit(1)
    except smtplib.SMTPException as e:
        print(f"SMTP error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


# ─────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python billing_report_generator.py <excel_file> [sheet_name]")
        sys.exit(1)

    filepath      = sys.argv[1]
    forced_sheet  = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        sys.exit(1)

    validate_config()

    try:
        old_name, cur_name = detect_sheets(filepath, forced_sheet)
    except ValueError as e:
        print(f"SHEET DETECTION FAILED:\n{e}")
        sys.exit(2)

    print(f"Old sheet:            {old_name}")
    print(f"Current-month sheet:  {cur_name}  ({month_label(cur_name)})")

    old_raw = pd.read_excel(filepath, sheet_name=old_name)
    cur_raw = pd.read_excel(filepath, sheet_name=cur_name)

    try:
        validate_workbook(old_raw, cur_raw, old_name, cur_name)
    except ValidationError as e:
        print("VALIDATION FAILED:")
        print(e)
        sys.exit(3)

    old_df = prepare(old_raw)
    cur_df = prepare(cur_raw)

    context = {
        "old_df":        old_df,
        "cur_df":        cur_df,
        "old_metrics":   collection_metrics(old_df),
        "cur_metrics":   collection_metrics(cur_df),
        "new_df":        new_clients(old_df, cur_df),
        "disc_df":       discontinued_clients(old_df, cur_df),
        "chg_df":        changed_rentals(old_df, cur_df),
        "old_issues":    find_data_issues(old_df, old_name),
        "cur_issues":    find_data_issues(cur_df, cur_name),
        "current_sheet": cur_name,
        "old_sheet":     old_name,
    }

    plain_body = format_plain_text(context)
    html_body  = format_html_email(context)
    subject    = build_subject(cur_name, context["new_df"],
                               context["disc_df"], context["chg_df"],
                               context["cur_metrics"])

    print("\n" + "=" * 60)
    print(plain_body)
    print("=" * 60)
    print(f"\nSubject: {subject}")

    txt_file = f"Billing_Message_{cur_name}.txt"
    with open(txt_file, "w", encoding="utf-8") as f:
        f.write(plain_body)
    print(f"Plain text saved: {txt_file}")

    html_file = f"Billing_Report_{cur_name}.html"
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_body)
    print(f"HTML preview saved: {html_file}")

    attachments = []
    if not context["chg_df"].empty:
        attachments.append(
            (f"Rental_Changes_{cur_name}.csv",
             changes_csv_bytes(context["chg_df"]))
        )

    send_email(subject, html_body, plain_body, attachments)


if __name__ == "__main__":
    main()

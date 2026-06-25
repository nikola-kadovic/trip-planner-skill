"""
trip_xlsx.py — reusable workbook builder for the trip-planner skill (Phase 5: Assembly).

Why this exists
---------------
Phase 5 turns the researched data (hotels, flights, activities, transport, dates)
into a polished, formula-driven Excel workbook. The *styling and plumbing* of that
workbook — fonts, fills, borders, currency formats, hyperlink mechanics, the link
audit, and the recalc step — are the same on every trip and are fiddly to get right
from scratch (the classic bug: coloring a cell blue but forgetting to set
`cell.hyperlink`, so it isn't actually clickable). Rewriting ~400 lines of that
boilerplate per run is slow and error-prone, so it lives here once.

What this does NOT do
---------------------
It does not decide which tabs a trip needs or how the cost formulas read — that's
trip-specific and stays in the orchestrator's build script. A ski trip wants a
ski-pass tab; a beach trip won't. This module gives you primitives and a couple of
table helpers; you compose the tabs.

Typical usage (orchestrator writes a short script like this)
------------------------------------------------------------
    import sys
    sys.path.insert(0, "<skill-dir>/scripts")
    import trip_xlsx as tx

    wb = tx.new_workbook()
    ws = tx.sheet(wb, "Hotels", first=True,
                  widths={1: 26, 2: 12, 3: 10, 4: 12, 5: 12})

    tx.title(ws, 1, "HOTELS — Paris", ncols=5)
    tx.header(ws, 3, ["Property", "Platform", "Score", "Total stay", "Total (home)"])
    r = 4
    tx.link(ws, r, 1, "Hotel Lutetia", "https://www.booking.com/hotel/...")
    tx.lbl(ws, r, 2, "Booking.com")
    tx.lbl(ws, r, 3, "9.1", align="center")
    tx.num(ws, r, 4, 1840, fmt=tx.money("EUR"))
    tx.formula(ws, r, 5, f"=D{r}*'Trip Summary'!$C$11")
    tx.zebra(ws, 4, r, 1, 5)
    tx.borders(ws, 3, 1, r, 5)

    # Save + audit links + recalc, all in one call:
    report = tx.finalize(wb, "/mnt/user-data/outputs/Paris_Trip.xlsx",
                         link_checks=[("Hotels", "D")])   # (sheet, price column)
    print(report)

Every cell helper returns the cell, so you can tweak it further if needed.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

# ── Palette ──────────────────────────────────────────────────────────────────
# Override any of these after import if a trip wants a different look:
#   import trip_xlsx as tx; tx.DARK = "1F4E79"
DARK = "14505C"        # title band fill
MID = "2E7D8A"         # section / header band fill
LIGHT = "EAF3F5"       # zebra stripe (alternate rows)
ACCENT = "C9A227"      # emphasis (e.g. per-person row)
WHITE = "FFFFFF"
CHOSEN = "FBF3D5"      # highlight fill for the recommended/chosen row
NOTE_GRAY = "F2F2F2"   # subtle fill for note/estimate rows

# ── Font colors (text, not fills) ────────────────────────────────────────────
C_TEXT = "000000"      # default text / formulas
C_INPUT = "0000FF"     # hardcoded numbers the user may edit (blue)
C_XREF = "375623"      # cross-sheet reference formulas (green)
C_LINK = "1155CC"      # hyperlinks (blue, underlined)
C_NOTE = "808080"      # gray italic notes / estimates
C_WHITE = "FFFFFF"
C_GOOD = "375623"      # affirmative inline text (e.g. "included")
C_WARN = "C00000"      # cautionary inline text (e.g. "extra fee")

# ── Number formats ───────────────────────────────────────────────────────────
FMT_INT = "0"
FMT_TEXT = "@"

# Currency symbols for the formats commonly needed. Extend as required.
_CURRENCY_SYMBOL = {
    "CAD": "C$", "USD": "$", "EUR": "€", "GBP": "£", "AUD": "A$",
    "NZD": "NZ$", "JPY": "¥", "CHF": "CHF ", "MXN": "$", "INR": "₹",
}


def money(currency: str = "USD") -> str:
    """Return an Excel number format for a currency code, e.g. money("EUR")."""
    sym = _CURRENCY_SYMBOL.get((currency or "").upper(), (currency or "") + " ")
    # positive ; negative-in-parens ; zero-as-dash
    return f'"{sym}"#,##0_);("{sym}"#,##0);-'


# Backwards-friendly default used widely in build scripts:
FMT_MONEY = money("USD")

_THIN = Side(style="thin", color="D0D9DB")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


# ── Core cell writer ─────────────────────────────────────────────────────────
def cell(ws, r, c, value=None, *, bold=False, size=10, color=C_TEXT, fill=None,
         align="left", wrap=False, fmt=None, italic=False, underline=None,
         border=True):
    """Write a value to (r, c) with Arial styling. Returns the cell.

    This is the single styling primitive; the named helpers below are thin
    wrappers around it for readability.
    """
    cl = ws.cell(row=r, column=c, value=value)
    cl.font = Font(name="Arial", bold=bold, size=size, color=color,
                   italic=italic, underline=underline)
    cl.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fill:
        cl.fill = PatternFill("solid", start_color=fill)
    if fmt:
        cl.number_format = fmt
    if border:
        cl.border = BORDER
    return cl


# ── Named cell helpers ───────────────────────────────────────────────────────
def lbl(ws, r, c, text, *, bold=False, align="left", fill=None, wrap=False,
        color=C_TEXT, size=10):
    """Plain text label."""
    return cell(ws, r, c, text, bold=bold, align=align, fill=fill, wrap=wrap,
                color=color, size=size, fmt=FMT_TEXT if isinstance(text, str) else None)


def num(ws, r, c, value, *, fmt=FMT_MONEY, color=C_INPUT, fill=None, align="right"):
    """Hardcoded number the user may edit — blue by default."""
    return cell(ws, r, c, value, color=color, fill=fill, align=align, fmt=fmt)


def formula(ws, r, c, expr, *, fmt=FMT_MONEY, color=C_TEXT, fill=None, align="right"):
    """A computed formula — black by default. `expr` must start with '='."""
    if not (isinstance(expr, str) and expr.startswith("=")):
        raise ValueError(f"formula() expects a string starting with '=', got: {expr!r}")
    return cell(ws, r, c, expr, color=color, fill=fill, align=align, fmt=fmt)


def xref(ws, r, c, expr, *, fmt=FMT_MONEY, fill=None, align="right"):
    """A cross-sheet reference formula — green, so the user can see it's a link
    to a chosen value elsewhere (e.g. ='Hotels'!J7)."""
    return formula(ws, r, c, expr, fmt=fmt, color=C_XREF, fill=fill, align=align)


def link(ws, r, c, text, url, *, bold=False, fill=None, wrap=True, align="left"):
    """A real, clickable hyperlink. Sets BOTH cell.hyperlink and link styling —
    coloring text blue alone does not make it clickable.

    If `url` is falsy this raises, because a priced row must never ship without a
    working link (see the skill's link-coverage gate). Pass a search URL for the
    exact item rather than leaving it blank.
    """
    if not url:
        raise ValueError(
            f"link() called with empty url for {text!r} at {get_column_letter(c)}{r}. "
            "Every booking/source cell needs a real URL — use a search URL for the "
            "exact item if no deep link exists."
        )
    cl = cell(ws, r, c, text, bold=bold, color=C_LINK, underline="single",
              fill=fill, wrap=wrap, align=align, fmt=FMT_TEXT)
    cl.hyperlink = url
    return cl


def note(ws, r, c, text, *, colspan=1, fill=None):
    """Gray italic note / source / estimate flag. Optionally merges colspan cells."""
    cl = cell(ws, r, c, text, color=C_NOTE, size=9, italic=True, wrap=True,
              align="left", fill=fill, border=False)
    cl.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    if colspan > 1:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + colspan - 1)
    return cl


# ── Layout helpers ───────────────────────────────────────────────────────────
_INVALID_SHEET = re.compile(r'[:\\/?*\[\]]')


def _safe_title(title: str) -> str:
    """Excel sheet titles: <=31 chars, no : \\ / ? * [ ]. Emoji are allowed."""
    t = _INVALID_SHEET.sub(" ", title).strip()
    return t[:31] if len(t) > 31 else t


def new_workbook() -> Workbook:
    """Fresh workbook. Use sheet(..., first=True) for the first tab so you don't
    leave a stray empty 'Sheet' behind."""
    return Workbook()


def sheet(wb, title, *, widths=None, first=False, hide_gridlines=True):
    """Create (or, with first=True, repurpose the initial blank sheet) a tab.
    `widths` is an optional {col_index_or_letter: width} dict."""
    title = _safe_title(title)
    if first:
        ws = wb.active
        ws.title = title
    else:
        ws = wb.create_sheet(title)
    if hide_gridlines:
        ws.sheet_view.showGridLines = False
    if widths:
        set_col_widths(ws, widths)
    return ws


def set_col_widths(ws, widths):
    """widths: {1: 26, 'B': 12, ...} — keys may be 1-based indexes or letters."""
    for key, w in widths.items():
        col = key if isinstance(key, str) else get_column_letter(key)
        ws.column_dimensions[col].width = w


def _band(ws, r, c, text, ncols, *, fill, fg, size, height):
    cl = cell(ws, r, c, text, bold=True, color=fg, fill=fill, size=size,
              align="center", wrap=True)
    if ncols > 1:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + ncols - 1)
        # carry the fill across the merged span so the whole band is colored
        for cc in range(c + 1, c + ncols):
            ws.cell(r, cc).fill = PatternFill("solid", start_color=fill)
    if height:
        ws.row_dimensions[r].height = height
    return cl


def title(ws, r, text, *, ncols=1, start_col=1, fill=DARK, fg=C_WHITE, size=13,
          height=28):
    """Big title band across `ncols` columns."""
    return _band(ws, r, start_col, text, ncols, fill=fill, fg=fg, size=size, height=height)


def band(ws, r, text, *, ncols=1, start_col=1, fill=MID, fg=C_WHITE, size=10,
         height=22):
    """Section band (sub-header)."""
    return _band(ws, r, start_col, text, ncols, fill=fill, fg=fg, size=size, height=height)


def header(ws, r, headers, *, start_col=1, fill=MID, fg=C_WHITE, height=26):
    """Write a row of column headers (bold, centered, wrapped, banded)."""
    for i, h in enumerate(headers):
        cell(ws, r, start_col + i, h, bold=True, color=fg, fill=fill,
             align="center", wrap=True)
    if height:
        ws.row_dimensions[r].height = height
    return r


def fill_row(ws, r, c1, c2, color):
    f = PatternFill("solid", start_color=color)
    for c in range(c1, c2 + 1):
        ws.cell(r, c).fill = f


def zebra(ws, r1, r2, c1, c2, colors=(LIGHT, WHITE)):
    """Apply alternating row fills across [r1, r2] x [c1, c2].
    Skips any cell that already has a non-default fill (so a CHOSEN highlight or a
    manual fill_row is preserved)."""
    for i, r in enumerate(range(r1, r2 + 1)):
        color = colors[i % len(colors)]
        for c in range(c1, c2 + 1):
            cl = ws.cell(r, c)
            existing = cl.fill
            already = (existing is not None and existing.patternType == "solid"
                       and (existing.fgColor.rgb not in (None, "00000000")))
            if not already:
                cl.fill = PatternFill("solid", start_color=color)


def borders(ws, r1, c1, r2, c2):
    for row in ws.iter_rows(min_row=r1, min_col=c1, max_row=r2, max_col=c2):
        for cl in row:
            cl.border = BORDER


def highlight_row(ws, r, c1, c2, color=CHOSEN):
    """Mark the chosen/recommended option row."""
    fill_row(ws, r, c1, c2, color)


# ── Link-coverage audit ──────────────────────────────────────────────────────
def audit_links(wb_or_path, checks):
    """Verify every PRICED row on the given sheets carries a clickable hyperlink.

    checks: list of (sheet_name, price_col) where price_col is the column letter
            (or 1-based index) whose presence of a number or formula marks a row
            as "priced". A row counts as covered if ANY cell in that row has a
            hyperlink, so it doesn't matter whether the link sits on the name
            column or a dedicated "Book" column.

    Pass only the bookable tabs (Hotels, Flights, Activities, Transport...) — not
    the Trip Summary, whose per-person formula cells are priced but shouldn't link.

    Returns a list of human-readable strings for offending rows (empty == all good).
    """
    wb = load_workbook(wb_or_path) if isinstance(wb_or_path, (str, os.PathLike)) else wb_or_path
    missing = []
    for sheet_name, price_col in checks:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        pc = price_col if isinstance(price_col, int) else column_index_from_string(price_col)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            if len(row) < pc:
                continue
            price = row[pc - 1].value
            is_number = isinstance(price, (int, float))
            is_formula = isinstance(price, str) and price.startswith("=")
            if not (is_number or is_formula):
                continue  # not a priced row
            if not any(c.hyperlink for c in row):
                missing.append(f"{sheet_name}!row {row[0].row} (priced, no clickable link)")
    return missing


# ── Recalc + finalize ────────────────────────────────────────────────────────
def _find_recalc():
    for p in (
        "/mnt/skills/public/xlsx/scripts/recalc.py",
        "/mnt/skills/private/xlsx/scripts/recalc.py",
    ):
        if os.path.exists(p):
            return p
    return None


def recalc(path):
    """Run the xlsx skill's recalc.py on the file if available.
    Returns (ok: bool, output: str). ok is None if recalc.py wasn't found."""
    script = _find_recalc()
    if not script:
        return None, "recalc.py not found — run the xlsx skill's recalc manually."
    proc = subprocess.run([sys.executable, script, path],
                          capture_output=True, text=True)
    out = ((proc.stdout or "") + (proc.stderr or "")).strip()
    # recalc.py prints JSON like {"status": "success", "total_errors": 0, ...}.
    # Parse it rather than substring-matching "error" (which appears in the keys).
    ok = proc.returncode == 0
    try:
        data = json.loads(out)
        ok = (data.get("status") == "success") and (data.get("total_errors", 0) == 0)
    except (json.JSONDecodeError, AttributeError):
        # Couldn't parse — fall back to: success only if no obvious failure text.
        ok = proc.returncode == 0 and "traceback" not in out.lower()
    return ok, out


def finalize(wb, out_path, *, link_checks=None, run_recalc=True):
    """Save the workbook, audit links, and run recalc — the standard close-out.

    out_path:    where to write (parent dirs are created if missing).
    link_checks: list of (sheet, price_col) for audit_links; pass the bookable
                 tabs. If None, the audit is skipped (do it manually).
    run_recalc:  run the xlsx skill's recalc.py and report formula errors.

    Returns a report dict. Inspect it: if report["missing_links"] is non-empty or
    report["recalc_ok"] is False, fix and call finalize() again before presenting.
    Never present a workbook with missing links or formula errors.
    """
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)

    report = {"path": out_path, "missing_links": [], "recalc_ok": None, "recalc_output": ""}

    if link_checks:
        report["missing_links"] = audit_links(out_path, link_checks)

    if run_recalc:
        ok, out = recalc(out_path)
        report["recalc_ok"] = ok
        report["recalc_output"] = out

    return report

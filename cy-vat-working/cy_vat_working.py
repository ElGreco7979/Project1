#!/usr/bin/env python3
"""
Cyprus VAT working-paper generator (VAT 4 return).

Builds and maintains an Excel workbook that turns a full general ledger (GL)
into a Cyprus VAT return, with intercompany / EU / third-country analysis,
a VIES summary, reconciliations to the VAT control accounts and quarterly
average exchange rates from the European Central Bank (the euro reference
rates that the Central Bank of Cyprus republishes).

Commands
--------
  new        Create a blank working (with clearly marked example data).
  import-gl  Load your GL export (.xlsx / .csv) into the working, keeping the
             account mapping and counterparty list from earlier quarters.
  fx         Download the ECB daily reference rates for the quarter and
             compute the quarterly average (or load them from a CSV file).
  roll       Set the working to a new quarter (dates in Settings).

Examples
--------
  python cy_vat_working.py new CY_VAT_Q3_2026.xlsx --year 2026 --quarter 3
  python cy_vat_working.py import-gl CY_VAT_Q3_2026.xlsx GL_export.xlsx
  python cy_vat_working.py fx CY_VAT_Q3_2026.xlsx --year 2026 --quarter 3

Every amount in the return is an Excel formula, so after the script has run
you can open the file in Excel, change a mapping and watch the return update.

Requires: Python 3.9+, openpyxl, pandas (pip install openpyxl pandas)
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import statistics
import sys
import urllib.error
import urllib.request
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# --------------------------------------------------------------------------
# Styling
# --------------------------------------------------------------------------
FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
F_H2 = Font(name=FONT, size=11, bold=True, color="1F3864")
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_INPUT = Font(name=FONT, size=10, color="0000FF")          # hardcoded inputs
F_LINK = Font(name=FONT, size=10, color="008000")           # links to other sheets
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
F_BOX = Font(name=FONT, size=11, bold=True)
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_INPUT = PatternFill("solid", fgColor="FFFF99")
FILL_CALC = PatternFill("solid", fgColor="F2F2F2")
FILL_KEY = PatternFill("solid", fgColor="DDEBF7")
FILL_EXAMPLE = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
EUR = '#,##0.00;(#,##0.00);"-"'
RATE = "0.0000"
PCT = "0.0%"
DATE = "dd/mm/yyyy"

EXAMPLE_TAG = "EXAMPLE"
SIGN_OFF = ["Prepared by / date", "Reviewed by / date", "Submitted on TFA (date / ref.)", "Paid (date / amount)"]

# --------------------------------------------------------------------------
# Sheet layout constants (formulas depend on these - change with care)
# --------------------------------------------------------------------------
GL = "GL_Input"
GL_HDR_ROW = 4
GL_FIRST = 5
GL_TEMPLATE_ROWS = 2000            # rows pre-filled with formulas in a new file

# Input columns A-L (you paste / the script imports into these)
GL_INPUT_COLS = [
    ("A", "date", "Date", 12),
    ("B", "doc", "Journal / Doc no.", 14),
    ("C", "account", "Account code", 12),
    ("D", "account_name", "Account name", 28),
    ("E", "description", "Description", 34),
    ("F", "cp_code", "Counterparty code", 14),
    ("G", "cp_name", "Counterparty name", 24),
    ("H", "currency", "Currency", 9),
    ("I", "amount_fc", "Amount in currency (Dr + / Cr -)", 16),
    ("J", "amount_eur", "Amount EUR if in GL (Dr + / Cr -)", 16),
    ("K", "vat_code", "VAT code override (optional)", 12),
    ("L", "vat_fc", "VAT amount in currency (optional)", 14),
]
# Calculated columns M-Z
GL_CALC_COLS = [
    ("M", "Account VAT code", 10),
    ("N", "VAT code used", 10),
    ("O", "Country", 8),
    ("P", "Region (CY/EU/3RD)", 9),
    ("Q", "Inter-company (Y/N)", 9),
    ("R", "FX rate (ccy per EUR)", 10),
    ("S", "Amount EUR (Dr + / Cr -)", 15),
    ("T", "Nature", 8),
    ("U", "Return value EUR", 15),
    ("V", "VAT rate", 8),
    ("W", "VAT EUR", 13),
    ("X", "VIES type", 7),
    ("Y", "Month in quarter", 8),
    ("Z", "Check", 26),
]

SET = "Settings"
# Settings cells
S_COMPANY, S_VATNO, S_TIC, S_QLABEL = "C4", "C5", "C6", "C7"
S_START, S_END, S_PRORATA, S_RC_IN_GL = "C8", "C9", "C10", "C11"
S_TOL, S_DUE, S_PRORATA_NOTE = "C12", "C13", "C14"
OUT_VAT_ROWS = range(18, 23)      # output VAT control accounts C18:C22
IN_VAT_ROWS = range(26, 31)       # input VAT control accounts C26:C30
MAP_FIRST = 35                    # GL column map rows

CODES = "VAT_Codes"
CODES_FIRST = 5
CODES_LAST = 80                   # blank rows up to here for your own codes
BOX_COLS = {                      # VAT_Codes column for each box flag
    "1": "I", "2": "J", "4": "K", "6": "L", "7": "M", "8A": "N",
    "8B": "O", "9": "P", "10": "Q", "11A": "R", "11B": "S",
}

MAPS = "Account_Mapping"
CPS = "Counterparties"
FX = "FX_Rates"
FXD = "FX_Daily"
LISTS = "Lists"

EU_MEMBERS = [  # ISO 3166 codes of the other 26 EU member states (+ EL, XI)
    ("AT", "Austria"), ("BE", "Belgium"), ("BG", "Bulgaria"), ("HR", "Croatia"),
    ("CZ", "Czechia"), ("DK", "Denmark"), ("EE", "Estonia"), ("FI", "Finland"),
    ("FR", "France"), ("DE", "Germany"), ("GR", "Greece"), ("EL", "Greece (VIES prefix)"),
    ("HU", "Hungary"), ("IE", "Ireland"), ("IT", "Italy"), ("LV", "Latvia"),
    ("LT", "Lithuania"), ("LU", "Luxembourg"), ("MT", "Malta"), ("NL", "Netherlands"),
    ("PL", "Poland"), ("PT", "Portugal"), ("RO", "Romania"), ("SK", "Slovakia"),
    ("SI", "Slovenia"), ("ES", "Spain"), ("SE", "Sweden"),
    ("XI", "Northern Ireland (goods only)"),
]

# --------------------------------------------------------------------------
# VAT treatment codes
# (code, description, nature, rate, expected region, line-is-VAT, recovery,
#  VIES type, boxes, guidance)
# recovery: 1 = full, 0 = blocked, "PR" = Settings pro-rata %
# --------------------------------------------------------------------------
VAT_CODES = [
    # ---------------- OUTPUTS ----------------
    ("S19", "Domestic sales - standard rate 19%", "Output", 0.19, "CY", "N", 1, "",
     ["1", "6"], "Supplies taxable in Cyprus at 19%. Includes inter-company supplies to Cyprus group companies (unless in a VAT group)."),
    ("S9", "Domestic sales - reduced rate 9%", "Output", 0.09, "CY", "N", 1, "",
     ["1", "6"], "e.g. hotel accommodation, restaurant/catering services - check the reduced-rate schedule."),
    ("S5", "Domestic sales - reduced rate 5%", "Output", 0.05, "CY", "N", 1, "",
     ["1", "6"], "e.g. foodstuffs, pharmaceuticals, qualifying residential property - check the reduced-rate schedule."),
    ("S3", "Domestic sales - super-reduced rate 3%", "Output", 0.03, "CY", "N", 1, "",
     ["1", "6"], "e.g. books, newspapers, periodicals - check the reduced-rate schedule."),
    ("Z0", "Domestic zero-rated supplies (not exports)", "Output", 0, "ANY", "N", 1, "",
     ["6", "9"], "0% supplies in Cyprus (e.g. items temporarily zero-rated, qualifying ships/aircraft supplies)."),
    ("EXPG", "Exports of goods to third countries", "Output", 0, "3RD", "N", 1, "",
     ["6", "9"], "Goods dispatched outside the EU. Keep customs export evidence (SAD / ECS exit message)."),
    ("EUG", "Intra-EU supplies of goods (B2B)", "Output", 0, "EU", "N", 1, "G",
     ["6", "8A"], "Goods dispatched to a VAT-registered customer in another member state. Customer VAT no. must be valid on VIES. Reported in VIES."),
    ("EUS", "Services to EU taxable persons (general rule - reverse charge by customer)", "Output", 0, "EU", "N", 1, "S",
     ["6", "8B"], "B2B services where the place of supply is the customer's member state (s.14 / Art. 44). Includes inter-company management fees, recharges, IT, consultancy to EU group companies. Invoice must say 'Reverse charge'. Reported in VIES."),
    ("TRI", "Triangular transactions (Cyprus as intermediary)", "Output", 0, "EU", "N", 1, "T",
     ["6", "10"], "Simplified triangulation - goods go directly from the first to the third member state. Reported in VIES with the triangulation indicator."),
    ("OS3", "Services to customers outside the EU (outside scope, with right to deduct)", "Output", 0, "3RD", "N", 1, "",
     ["6", "10"], "B2B services to third-country customers, including inter-company services to non-EU group companies. Not in VIES."),
    ("OSS", "Distance sales / B2C services declared through OSS", "Output", 0, "EU", "N", 1, "",
     ["6", "10"], "Sales taxed in another member state and declared in the OSS return. The value only goes in Boxes 6 and 10."),
    ("FIN3", "Financial / insurance services to recipients outside the EU", "Output", 0, "3RD", "N", 1, "",
     ["6", "10"], "Financial and insurance services (e.g. inter-company loan interest) to non-EU recipients give a right to deduct input VAT."),
    ("FINEU", "Financial / insurance services to EU recipients (exempt, no right to deduct)", "Output", 0, "EU", "N", 1, "",
     ["6"], "e.g. interest on loans to EU group companies. Not in VIES and gives no right to deduct - affects the pro-rata."),
    ("EXS", "Exempt supplies (domestic)", "Output", 0, "ANY", "N", 1, "",
     ["6"], "Exempt supplies (e.g. financial services, insurance, letting of immovable property, medical, education). No right to deduct related input VAT."),
    ("DRCS", "Domestic reverse-charge sales (customer accounts for VAT)", "Output", 0, "CY", "N", 1, "",
     ["6"], "Supplies where the Cyprus customer accounts for the VAT (construction services, scrap metals, mobile phones / laptops / consoles, precious metals, etc. - s.11B/11C). Invoice without VAT, marked 'reverse charge'."),
    # ---------------- INPUTS ----------------
    ("P19", "Domestic purchases 19% - fully recoverable", "Input", 0.19, "CY", "N", 1, "",
     ["4", "7"], "Purchases from Cyprus suppliers with a valid VAT invoice, used for taxable (or outside-scope-with-right) supplies. Includes capital assets."),
    ("P9", "Domestic purchases 9% - fully recoverable", "Input", 0.09, "CY", "N", 1, "",
     ["4", "7"], "As P19 at 9%."),
    ("P5", "Domestic purchases 5% - fully recoverable", "Input", 0.05, "CY", "N", 1, "",
     ["4", "7"], "As P19 at 5%."),
    ("P3", "Domestic purchases 3% - fully recoverable", "Input", 0.03, "CY", "N", 1, "",
     ["4", "7"], "As P19 at 3%."),
    ("PM19", "Domestic purchases 19% - mixed use (pro-rata applies)", "Input", 0.19, "CY", "N", "PR", "",
     ["4", "7"], "Overheads used for both taxable and exempt activities. Recovery = pro-rata % in Settings."),
    ("PB19", "Domestic purchases 19% - VAT not recoverable (blocked)", "Input", 0.19, "ANY", "N", 0, "",
     ["7"], "Business entertainment, private use, saloon cars (subject to exceptions), costs for exempt supplies. Value in Box 7, no Box 4 claim."),
    ("PZ", "Purchases with no VAT (zero-rated, exempt, non-registered supplier)", "Input", 0, "ANY", "N", 1, "",
     ["7"], "e.g. insurance, bank charges, rent without VAT, purchases from non-registered suppliers."),
    ("AEG19", "Acquisitions of goods from EU suppliers - 19%", "Input", 0.19, "EU", "N", 1, "",
     ["2", "4", "7", "11A"], "Goods received from a VAT-registered EU supplier invoicing without VAT. Self-account VAT in Box 2, recover in Box 4."),
    ("AEG9", "Acquisitions of goods from EU suppliers - 9%", "Input", 0.09, "EU", "N", 1, "",
     ["2", "4", "7", "11A"], "As AEG19 at 9%."),
    ("AEG5", "Acquisitions of goods from EU suppliers - 5%", "Input", 0.05, "EU", "N", 1, "",
     ["2", "4", "7", "11A"], "As AEG19 at 5%."),
    ("AEG0", "Acquisitions of goods from EU suppliers - zero-rated goods", "Input", 0, "EU", "N", 1, "",
     ["7", "11A"], "Zero-rated goods acquired from the EU (value only)."),
    ("AES19", "Services received from EU suppliers - reverse charge 19%", "Input", 0.19, "EU", "N", 1, "",
     ["1", "4", "6", "7", "11B"], "B2B services from EU suppliers (including inter-company recharges, management fees, licences, IT from EU group companies). VAT in Box 1 and Box 4; value in Boxes 6, 7 and 11B."),
    ("AESM", "Services received from EU suppliers - reverse charge, mixed use", "Input", 0.19, "EU", "N", "PR", "",
     ["1", "4", "6", "7", "11B"], "As AES19 but input recovery limited to the pro-rata % in Settings."),
    ("A3S19", "Services received from non-EU suppliers - reverse charge 19%", "Input", 0.19, "3RD", "N", 1, "",
     ["1", "4", "6", "7"], "B2B services from third-country suppliers (e.g. US software, UK consultancy, inter-company charges from non-EU group companies). Not in Box 11B."),
    ("A3SM", "Services received from non-EU suppliers - reverse charge, mixed use", "Input", 0.19, "3RD", "N", "PR", "",
     ["1", "4", "6", "7"], "As A3S19 but input recovery limited to the pro-rata % in Settings."),
    ("DRCP", "Domestic reverse-charge purchases 19% (we account for the VAT)", "Input", 0.19, "CY", "N", 1, "",
     ["1", "4", "7"], "Purchases under s.11B/11C (construction, scrap, phones/laptops, etc.) invoiced without VAT by a Cyprus supplier."),
    ("IMP", "Imports of goods from third countries (value)", "Input", 0, "3RD", "N", 1, "",
     ["7"], "Value of imported goods. The import VAT paid to Customs is posted separately with code IVAT."),
    ("IVAT", "Import VAT paid to / deferred with Customs (line is the VAT)", "Input", 0, "ANY", "Y", 1, "",
     ["4"], "Map the 'Import VAT' account here. The GL line amount itself is the VAT claimed in Box 4 (needs the customs import declaration)."),
    ("FOR", "Foreign VAT incurred abroad (not on the Cyprus return)", "None", 0, "ANY", "N", 0, "",
     [], "e.g. hotels or fuel abroad. Reclaim through the EU VAT refund procedure (Directive 2008/9) or the foreign country's procedure."),
    ("NR", "Not reportable on the VAT return", "None", 0, "ANY", "N", 0, "",
     [], "Salaries, social insurance, depreciation, dividends, bank transfers, loans principal, balance-sheet movements, VAT control accounts, corporation tax."),
]

# Official box titles and what each should include
BOXES = [
    ("1", "VAT due in this period on sales and other outputs",
     "Output VAT on domestic taxable supplies (19%, 9%, 5%, 3%) PLUS the VAT you self-account under the reverse charge on services received from abroad (EU and non-EU) and on domestic reverse-charge purchases. Credit notes issued reduce this box.",
     "S19, S9, S5, S3, AES19, AESM, A3S19, A3SM, DRCP"),
    ("2", "VAT due in this period on acquisitions from other EU Member States",
     "VAT self-accounted on goods (and related services) acquired from VAT-registered suppliers in other EU member states.",
     "AEG19, AEG9, AEG5"),
    ("3", "Total VAT due (sum of Boxes 1 and 2)", "Calculated: Box 1 + Box 2.", "Box 1 + Box 2"),
    ("4", "VAT reclaimed in this period on purchases and other inputs (including acquisitions from the EU)",
     "Deductible input VAT on domestic purchases and capital goods, import VAT paid to Customs, and the reverse-charge VAT of Boxes 1 and 2 to the extent it is deductible. Exclude blocked VAT (entertainment, private use, saloon cars, VAT on exempt activities) and foreign VAT. Mixed-use costs at the pro-rata %.",
     "P19, P9, P5, P3, PM19 (pro-rata), AEG*, AES19, AESM (pro-rata), A3S19, A3SM (pro-rata), DRCP, IVAT"),
    ("5", "Net VAT to be paid or reclaimed (difference between Boxes 3 and 4)",
     "Calculated: Box 3 - Box 4. Positive = payable by the due date; negative = credit carried forward / refund claim.",
     "Box 3 - Box 4"),
    ("6", "Total value of sales and other outputs excluding any VAT",
     "All outputs excluding VAT: domestic taxable, zero-rated, exempt, exports, EU supplies, outside-scope supplies, domestic reverse-charge sales. Must include the values in Boxes 8A, 8B, 9, 10 and 11B, and the value of services received from outside the EU under the reverse charge.",
     "S*, Z0, EXPG, EUG, EUS, TRI, OS3, OSS, FIN3, FINEU, EXS, DRCS, AES*, A3S*"),
    ("7", "Total value of purchases and other inputs excluding any VAT",
     "All purchases and expenses (including capital goods) excluding VAT: domestic, imports, EU acquisitions of goods (11A), services received from abroad (11B and non-EU). Exclude salaries, depreciation, dividends, loans and other non-supplies.",
     "P*, PM19, PB19, PZ, AEG*, AES*, A3S*, DRCP, IMP"),
    ("8A", "Total value of supplies of goods to other EU Member States",
     "Intra-community supplies of goods to VAT-registered customers in other member states (0%). Must agree with the goods lines of the VIES statements for the quarter.",
     "EUG"),
    ("8B", "Total value of services supplied to other EU Member States (general rule)",
     "B2B services to EU taxable persons where the customer accounts for the VAT (reverse charge). Must agree with the services lines of the VIES statements.",
     "EUS"),
    ("9", "Total value of outputs at zero rate (other than those in Boxes 8A and 8B)",
     "Exports of goods outside the EU and other zero-rated supplies.",
     "EXPG, Z0"),
    ("10", "Total value of out-of-scope sales with the right to deduct input VAT",
     "Services to non-EU customers whose place of supply is outside Cyprus, financial/insurance services to non-EU recipients, triangular transactions, OSS sales.",
     "OS3, FIN3, TRI, OSS"),
    ("11A", "Total value of acquisitions of goods and related services from other EU Member States",
     "Value (excluding VAT) of goods acquired from EU suppliers. Also included in Box 7.",
     "AEG19, AEG9, AEG5, AEG0"),
    ("11B", "Total value of services received from other EU Member States (reverse charge)",
     "Value (excluding VAT) of B2B services received from EU suppliers. Also included in Boxes 6 and 7.",
     "AES19, AESM"),
]


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def quarter_dates(year: int, quarter: int) -> tuple[dt.date, dt.date]:
    if quarter not in (1, 2, 3, 4):
        raise SystemExit("--quarter must be 1-4")
    start = dt.date(year, 3 * quarter - 2, 1)
    end_month = 3 * quarter
    end = (dt.date(year + (end_month == 12), end_month % 12 + 1, 1) - dt.timedelta(days=1))
    return start, end


def hdr(ws, row, values, widths=None, height=30):
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[row].height = height
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def put(ws, ref, value, font=F_BASE, fill=None, fmt=None, border=True, wrap=False):
    c = ws[ref]
    c.value = value
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if border:
        c.border = BORDER
    if wrap:
        c.alignment = WRAP
    return c


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    if sub:
        ws["A2"] = sub
        ws["A2"].font = F_NOTE


def text_rows(ws, start_row, rows, col_width=150):
    """rows: list of (style, text). style in {'h','b','t','n',''}"""
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = col_width
    r = start_row
    for style, text in rows:
        if style == "h" and r != start_row:
            r += 1
        c = ws.cell(row=r, column=2, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = {"h": F_H2, "b": F_BOLD, "n": F_NOTE}.get(style, F_BASE)
        lines = max(1, len(text) // 140 + 1) if text else 1
        ws.row_dimensions[r].height = 14 * lines if lines > 1 else None
        r += 1
    return r


def set_ref(cell):
    return f"{SET}!${cell[0]}${cell[1:]}"


# --------------------------------------------------------------------------
# GL formulas
# --------------------------------------------------------------------------
def code_lookup(col, r, default='""'):
    return f'IFERROR(INDEX({CODES}!${col}:${col},MATCH($N{r},{CODES}!$A:$A,0)),{default})'


def gl_formulas(r: int) -> dict[str, str]:
    start, end = set_ref(S_START), set_ref(S_END)
    expected = code_lookup("E", r, '"ANY"')
    return {
        "M": f'=IF($C{r}="","",IFERROR(INDEX({MAPS}!$C:$C,MATCH($C{r}&"",{MAPS}!$A:$A,0))&"",""))',
        "N": f'=IF($K{r}<>"",$K{r},$M{r})',
        "O": f'=IF($F{r}="","",IFERROR(INDEX({CPS}!$C:$C,MATCH($F{r}&"",{CPS}!$A:$A,0))&"",""))',
        "P": f'=IF($O{r}="","",IF($O{r}="CY","CY",IF(ISNUMBER(MATCH($O{r},{LISTS}!$A:$A,0)),"EU","3RD")))',
        "Q": f'=IF($F{r}="","N",IFERROR(IF(INDEX({CPS}!$E:$E,MATCH($F{r}&"",{CPS}!$A:$A,0))="Y","Y","N"),"N"))',
        "R": f'=IF(OR($H{r}="",$H{r}="EUR"),1,IFERROR(INDEX({FX}!$D:$D,MATCH($H{r},{FX}!$A:$A,0)),""))',
        "S": f'=IF(AND($I{r}="",$J{r}=""),"",IF($J{r}<>"",$J{r},IF(OR($R{r}="",$R{r}=0),"",ROUND($I{r}/$R{r},2))))',
        "T": f'=IF($N{r}="","",{code_lookup("C", r)}&"")',
        "U": f'=IF(OR($S{r}="",$T{r}="",$T{r}="None"),0,IF($T{r}="Output",-$S{r},$S{r}))',
        "V": f'=IF($N{r}="",0,{code_lookup("D", r, "0")})',
        "W": (f'=IF($U{r}=0,0,IF({code_lookup("F", r, chr(34) + "N" + chr(34))}="Y",$U{r},'
              f'IF($L{r}<>"",ROUND(ABS($L{r})/IF(OR($R{r}="",$R{r}=0),1,$R{r}),2)*SIGN($U{r}),ROUND($U{r}*$V{r},2))))'),
        "X": f'=IF($N{r}="","",{code_lookup("H", r)}&"")',
        "Y": f'=IF($A{r}="","",(YEAR($A{r})-YEAR({start}))*12+MONTH($A{r})-MONTH({start})+1)',
        "Z": (f'=IF(AND($C{r}="",$I{r}="",$J{r}=""),"",'
              f'IF($N{r}="","NO VAT CODE",'
              f'IF($T{r}="","UNKNOWN VAT CODE",'
              f'IF($S{r}="","MISSING FX RATE",'
              f'IF(OR($A{r}<{start},$A{r}>{end}),"OUTSIDE PERIOD",'
              f'IF(AND({expected}<>"ANY",$P{r}=""),"MISSING COUNTERPARTY COUNTRY",'
              f'IF(AND({expected}<>"ANY",$P{r}<>{expected}),"REGION MISMATCH","OK")))))))'),
    }


def write_gl_row(ws, r, values: dict | None = None, example=False):
    if values:
        for col, key, _, _ in GL_INPUT_COLS:
            v = values.get(key)
            if v is None or v == "":
                continue
            c = ws[f"{col}{r}"]
            c.value = v
            c.font = F_INPUT
            if key == "date":
                c.number_format = DATE
            elif key in ("amount_fc", "amount_eur", "vat_fc"):
                c.number_format = EUR
            if example:
                c.fill = FILL_EXAMPLE
    for col, f in gl_formulas(r).items():
        c = ws[f"{col}{r}"]
        c.value = f
        c.font = F_BASE
        if col in ("S", "U", "W"):
            c.number_format = EUR
        elif col == "R":
            c.number_format = RATE
        elif col == "V":
            c.number_format = PCT


# --------------------------------------------------------------------------
# Sheet builders
# --------------------------------------------------------------------------
def build_instructions(wb):
    ws = wb.create_sheet("Instructions", 0)
    ws.sheet_view.showGridLines = False
    title(ws, "Cyprus VAT working - instructions",
          "How the working is built, what you do each quarter, and how each part feeds the VAT 4 return.")
    rows = [
        ("h", "1. What this working does"),
        ("", "It takes the full general ledger (GL) for the VAT quarter, gives every line a VAT treatment code (through the account mapping, "
             "or an override on the line), classifies the counterparty as Cyprus (CY), other EU member state (EU) or third country (3RD) and as "
             "inter-company (Y) or third party (N), converts foreign-currency amounts to EUR at the quarterly average ECB rate, and builds the "
             "VAT 4 return (Boxes 1-11B), the VIES summary, an inter-company analysis and reconciliations to the VAT control accounts."),
        ("", "All results are Excel formulas. Change a mapping, a country or a rate and every sheet updates."),
        ("h", "2. Sheets"),
        ("b", "Instructions - this page."),
        ("b", "Settings - company details, quarter dates, pro-rata %, VAT control accounts, GL column names for the import script. (Yellow cells = your input.)"),
        ("b", "GL_Input - the full GL for the quarter. Columns A-L are inputs (paste or import). Columns M-Z are formulas - do not overwrite them."),
        ("b", "Account_Mapping - one row per GL account with its default VAT code. Carried forward from quarter to quarter."),
        ("b", "Counterparties - one row per customer/supplier: country (ISO code), VAT number and inter-company flag. Drives the EU / 3rd-country / inter-company split."),
        ("b", "VAT_Codes - the VAT treatments and which return boxes each one feeds. Totals per code are calculated here."),
        ("b", "FX_Rates / FX_Daily - daily ECB euro reference rates for the quarter and the quarterly average used in GL_Input."),
        ("b", "VAT_Return - the VAT 4 form with every box, the amount, what the box must include and which codes feed it."),
        ("b", "IC_Analysis - every VAT code split into inter-company / third party and CY / EU / third country."),
        ("b", "VIES - the monthly recapitulative statement amounts per EU customer (goods, triangular, services)."),
        ("b", "Reconciliation - Box 1 and Box 4 against the GL VAT control accounts."),
        ("b", "Checks - the exception report. Every check must show 0 / OK before you file."),
        ("b", "Legal_Notes - rates, deadlines, penalties, reverse-charge rules, place of supply, exchange-rate rules and sources."),
        ("h", "3. Colour code"),
        ("", "Blue text on yellow = input you enter or the script imports.  Black = formula.  Green = link to another sheet.  "
             "Orange fill = EXAMPLE data that shows the expected format (the import script removes it; delete it if you paste manually)."),
        ("h", "4. One-off set-up (first quarter only)"),
        ("", "a) Install Python 3.9+ and run:  pip install openpyxl pandas"),
        ("", "b) Create the working:  python cy_vat_working.py new CY_VAT_Q3_2026.xlsx --year 2026 --quarter 3"),
        ("", "c) Settings: enter company name, VAT number (e.g. 10123456X), TIC, the output and input VAT control account codes, and the pro-rata % "
             "(100% if the company makes no exempt supplies). Under 'GL column map' type the column headers exactly as they appear in your GL export."),
        ("", "d) After the first import (step 5b), complete Account_Mapping: give every account a VAT code from VAT_Codes. Balance-sheet, payroll, "
             "depreciation, dividend, loan, bank and VAT control accounts = NR. Where one account mixes treatments (e.g. a 'Consultancy' account with "
             "Cyprus, EU and US suppliers), map it to the most common code and override the other lines in GL_Input column K."),
        ("", "e) Complete Counterparties: ISO country code (CY, GR, DE, GB, US, AE ...), VAT number (with country prefix, e.g. DE123456789) and "
             "Inter-company = Y for every group company. Use EL for Greek VAT numbers in VIES if you wish - EL and GR are both treated as EU."),
        ("h", "5. Every quarter"),
        ("", "a) Copy last quarter's working (it keeps the mapping and counterparties) and run:  python cy_vat_working.py roll CY_VAT_Q4_2026.xlsx --year 2026 --quarter 4"),
        ("", "b) Import the GL:  python cy_vat_working.py import-gl CY_VAT_Q4_2026.xlsx GL_Q4_2026.xlsx   (add --sheet NAME if needed). "
             "The script clears the old GL lines, writes the new ones with formulas, and adds any NEW accounts / counterparties to the bottom of "
             "Account_Mapping and Counterparties with a blank code - complete them. Alternatively paste the GL into GL_Input A-L yourself "
             "and copy the formulas in M-Z down to the last row."),
        ("", "c) Exchange rates:  python cy_vat_working.py fx CY_VAT_Q4_2026.xlsx --year 2026 --quarter 4   - downloads the ECB daily euro reference "
             "rates (also published by the Central Bank of Cyprus) for every currency found in the GL, writes them to FX_Daily and sets FX_Rates to the "
             "simple average of the quarter's business days. No internet? Download the rates as CSV from the CBC or ECB site and run with --csv FILE. "
             "For currencies the ECB does not publish (e.g. AED, RUB), type a rate in FX_Rates column C (manual override) and note the source."),
        ("", "d) Open the file in Excel (formulas calculate on opening; press F9 if calculation is manual)."),
        ("", "e) Checks sheet: clear every exception - unmapped accounts, unknown codes, counterparties without a country, region mismatches "
             "(e.g. code EUS on a US customer), missing FX rates, lines outside the period."),
        ("", "f) Review IC_Analysis: inter-company services to EU group companies must be in EUS (Box 8B + VIES), to non-EU group companies in OS3 (Box 10); "
             "inter-company charges received from EU group companies in AES19 (Boxes 1, 4, 6, 7, 11B) and from non-EU group companies in A3S19 (Boxes 1, 4, 6, 7)."),
        ("", "g) Reconciliation: Box 1 and Box 4 against the VAT control accounts. Investigate any difference above the tolerance in Settings "
             "(typical causes: VAT posted to the wrong account, rounding, blocked VAT, credit notes, invoices dated outside the quarter)."),
        ("", "h) VIES: check that the VIES totals agree with Boxes 8A, 8B and the triangular part of Box 10, and that every EU customer has a valid VAT "
             "number (check on the EU VIES site). VIES is filed MONTHLY, by the 15th of the following month."),
        ("", "i) Transfer the VAT_Return amounts (Boxes 1-11B) to the VAT 4 return on the Tax For All (TFA) portal. Submit and pay by the 10th day of the "
             "second month after the quarter end (shown in Settings). Save a PDF of the submitted return with this working."),
        ("h", "6. How the calculation works (GL_Input columns M-Z)"),
        ("", "M  Account VAT code - looked up from Account_Mapping using the account code (C)."),
        ("", "N  VAT code used - the override in K if entered, otherwise M."),
        ("", "O-Q  Country, region and inter-company flag - looked up from Counterparties using the counterparty code (F). Region: CY = Cyprus, "
             "EU = code listed on the Lists sheet, 3RD = any other country."),
        ("", "R-S  FX rate and EUR amount. If the GL already has an EUR amount (J) it is used as it is; otherwise amount in currency (I) / quarterly "
             "average rate (ECB rates are units of currency per 1 EUR). EUR lines use rate 1."),
        ("", "T-U  Nature and return value. GL sign convention is Debit = + and Credit = -. Outputs (sales) are reversed so sales show positive; credit "
             "notes therefore reduce the boxes automatically. Code NR / FOR lines give 0."),
        ("", "V-W  VAT rate and VAT. VAT = return value x rate, unless you give the actual VAT amount on the line (L) - then the actual amount is used. "
             "For IVAT (import VAT) the line amount itself is the VAT."),
        ("", "X-Y  VIES type (G = goods, T = triangular, S = services) and month 1-3 of the quarter, used by the VIES sheet."),
        ("", "Z  Check - OK or the first exception found on the line. Filter column Z for anything other than OK."),
        ("", "VAT_Codes totals each code (SUMIFS on columns N, U, W). The recoverable VAT = VAT x recovery % (100%, 0% or the pro-rata). "
             "Each return box adds up the codes flagged Y in that box's column."),
        ("h", "7. Foreign exchange - important"),
        ("", "Art. 91(2) of the VAT Directive (2006/112/EC), which the Cyprus VAT Law follows, requires a foreign-currency taxable amount to be converted "
             "at the latest rate published by the ECB (or the most representative market rate) at the time the VAT becomes chargeable (the tax point). "
             "A quarterly average is a practical simplification that many groups use for reporting. Before relying on it, confirm with your "
             "VAT adviser / the Tax Department that it is acceptable for your circumstances, and keep using the invoice-date rate where the invoice "
             "itself shows the VAT in EUR. To use transaction-date rates, enter the EUR amounts in GL_Input column J - they take priority over the average."),
        ("h", "8. Limits of this working"),
        ("", "The VAT codes and box mapping follow the Tax Department's VAT 4 guidance and Big Four / practitioner guides as at October 2026. "
             "VAT rules (rates, temporary zero rates, reverse-charge lists) change often - review VAT_Codes and Legal_Notes every quarter. "
             "This working is a preparation tool, not tax advice; complex items (property, construction, capital goods adjustments, "
             "partial exemption annual adjustments, VAT groups, fixed establishments) need specific review."),
    ]
    text_rows(ws, 4, rows)
    ws.freeze_panes = "A4"


def build_settings(wb, start: dt.date, end: dt.date, quarter_label: str):
    ws = wb.create_sheet(SET)
    title(ws, "Settings", "Yellow cells are inputs. Dates drive the period checks, VIES months and due date.")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 70
    rows = [
        (4, "Company name", "Example Holdings Ltd", None, "Legal name as registered with the Tax Department."),
        (5, "VAT registration number", "10000000X", None, "Cyprus VAT number (8 digits + letter)."),
        (6, "Tax Identification Code (TIC)", "", None, "Optional."),
        (7, "Quarter", quarter_label, None, "Label only."),
        (8, "Period start", start, DATE, "First day of the VAT quarter."),
        (9, "Period end", end, DATE, "Last day of the VAT quarter."),
        (10, "Input VAT recovery % (pro-rata) for mixed-use costs", 1.0, PCT,
         "Provisional % = taxable (incl. outside-scope with right) supplies / total supplies of the previous year. 100% if no exempt supplies. Annual adjustment at year end."),
        (11, "Reverse-charge VAT is booked in the GL VAT control accounts? (Y/N)", "N", None,
         "Y if your bookkeeping posts the self-assessed reverse-charge VAT to the output and input VAT accounts. Used by the Reconciliation sheet."),
        (12, "Reconciliation tolerance (EUR)", 5, EUR, "Differences above this amount are flagged."),
    ]
    for r, label, val, fmt, note in rows:
        put(ws, f"B{r}", label, F_BOLD)
        put(ws, f"C{r}", val, F_INPUT, FILL_INPUT, fmt)
        put(ws, f"D{r}", note, F_NOTE, wrap=True)
    put(ws, "B13", "Return and payment due date", F_BOLD)
    put(ws, S_DUE, f"=DATE(YEAR({S_END}),MONTH({S_END})+2,10)", F_BASE, FILL_CALC, DATE)
    put(ws, "D13", "10th day of the second month after the end of the quarter (next working day if a weekend / holiday).", F_NOTE, wrap=True)
    put(ws, "B14", "VIES statements due", F_BOLD)
    put(ws, S_PRORATA_NOTE, "15th of each following month", F_BASE, FILL_CALC)
    put(ws, "D14", "Monthly recapitulative statement for EU supplies of goods and services (TFA portal).", F_NOTE, wrap=True)

    ws["B17"] = "Output VAT control account codes (credit balances)"
    ws["B17"].font = F_H2
    for i, r in enumerate(OUT_VAT_ROWS):
        put(ws, f"B{r}", f"Output VAT account {i + 1}")
        put(ws, f"C{r}", "2201" if i == 0 else None, F_INPUT, FILL_INPUT)
    ws["B25"] = "Input VAT control account codes (debit balances) - excl. import VAT account"
    ws["B25"].font = F_H2
    for i, r in enumerate(IN_VAT_ROWS):
        put(ws, f"B{r}", f"Input VAT account {i + 1}")
        put(ws, f"C{r}", "2101" if i == 0 else None, F_INPUT, FILL_INPUT)

    ws[f"B{MAP_FIRST - 2}"] = "GL column map (for the import-gl command)"
    ws[f"B{MAP_FIRST - 2}"].font = F_H2
    ws[f"D{MAP_FIRST - 2}"] = ("Type the header text exactly as it appears in your GL export. Leave blank if the column does not exist. "
                               "Give either 'Amount' (Dr + / Cr -) or both 'Debit' and 'Credit'.")
    ws[f"D{MAP_FIRST - 2}"].font = F_NOTE
    ws[f"D{MAP_FIRST - 2}"].alignment = WRAP
    hdr_cells = ["Working field", "Header in your GL export", "Notes"]
    for i, t in enumerate(hdr_cells):
        c = ws.cell(row=MAP_FIRST - 1, column=2 + i, value=t)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
    fields = [
        ("date", "Date", "Posting or document date."),
        ("doc", "Doc No", ""),
        ("account", "Account", "Account code."),
        ("account_name", "Account Name", ""),
        ("description", "Description", ""),
        ("cp_code", "Counterparty", "Customer / supplier code (or name) - the key to Counterparties."),
        ("cp_name", "Counterparty Name", ""),
        ("currency", "Currency", "ISO currency code. Blank = EUR."),
        ("amount_fc", "Amount", "Signed amount in transaction currency (Dr + / Cr -)."),
        ("debit_fc", "", "Use instead of 'Amount' when the export has separate debit and credit columns."),
        ("credit_fc", "", ""),
        ("amount_eur", "", "Optional: amount already in EUR (Dr + / Cr -). Takes priority over the FX conversion."),
        ("vat_code", "", "Optional: VAT/tax code from your accounting system, if it uses the same codes as VAT_Codes."),
        ("vat_fc", "", "Optional: VAT amount of the line in transaction currency."),
    ]
    for i, (key, default, note) in enumerate(fields):
        r = MAP_FIRST + i
        put(ws, f"B{r}", key)
        put(ws, f"C{r}", default or None, F_INPUT, FILL_INPUT)
        put(ws, f"D{r}", note, F_NOTE, wrap=True)


def build_gl(wb, rows=GL_TEMPLATE_ROWS, examples=None):
    ws = wb.create_sheet(GL)
    title(ws, "GL_Input - full general ledger for the VAT quarter",
          "Columns A-L: paste/import your GL (Debit = +, Credit = -).  Columns M-Z: formulas, do not overwrite.  Filter column Z for exceptions.")
    ws["A3"] = "INPUT  >>>"
    ws["A3"].font = F_BOLD
    ws["M3"] = "CALCULATED  >>>"
    ws["M3"].font = F_BOLD
    headers = [h for _, _, h, _ in GL_INPUT_COLS] + [h for _, h, _ in GL_CALC_COLS]
    widths = [w for _, _, _, w in GL_INPUT_COLS] + [w for _, _, w in GL_CALC_COLS]
    hdr(ws, GL_HDR_ROW, headers, widths, height=45)
    for i in range(12, 26):
        ws.cell(row=GL_HDR_ROW, column=i + 1).fill = PatternFill("solid", fgColor="595959")
    examples = examples or []
    n = max(rows, len(examples))
    for i in range(n):
        r = GL_FIRST + i
        write_gl_row(ws, r, examples[i] if i < len(examples) else None, example=i < len(examples))
    ws.freeze_panes = ws[f"D{GL_FIRST}"]
    ws.auto_filter.ref = f"A{GL_HDR_ROW}:Z{GL_FIRST + n - 1}"
    dv = DataValidation(type="list", formula1=f"={CODES}!$A${CODES_FIRST}:$A${CODES_LAST}", allow_blank=True)
    dv.error = "Choose a code from VAT_Codes"
    ws.add_data_validation(dv)
    dv.add(f"K{GL_FIRST}:K{GL_FIRST + max(n, 50000)}")
    ws.conditional_formatting.add(
        f"Z{GL_FIRST}:Z{GL_FIRST + max(n, 50000)}",
        FormulaRule(formula=[f'AND($Z{GL_FIRST}<>"",$Z{GL_FIRST}<>"OK")'],
                    fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONT, color="9C0006", bold=True)))
    return ws


def build_mapping(wb, accounts):
    ws = wb.create_sheet(MAPS)
    title(ws, "Account_Mapping - default VAT code per GL account",
          "Give every account a code from VAT_Codes. Non-VAT accounts (balance sheet, payroll, depreciation, VAT control) = NR. Override single lines in GL_Input column K.")
    hdr(ws, 4, ["Account code", "Account name", "VAT code", "Notes / reasoning",
                "GL movement EUR (Dr+/Cr-)", "No. of GL lines", "Lines with exceptions"],
        [14, 36, 11, 50, 18, 12, 14])
    for i, acc in enumerate(accounts):
        write_mapping_row(ws, 5 + i, *acc)
    ws.freeze_panes = "C5"
    dv = DataValidation(type="list", formula1=f"={CODES}!$A${CODES_FIRST}:$A${CODES_LAST}", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("C5:C5000")
    ws.conditional_formatting.add(
        "C5:C5000", FormulaRule(formula=['AND($A5<>"",$C5="")'], fill=PatternFill("solid", fgColor="F8CBAD")))
    return ws


def write_mapping_row(ws, r, code, name, vat_code, note, example=False):
    fill = FILL_EXAMPLE if example else FILL_INPUT
    put(ws, f"A{r}", str(code), F_INPUT, fill)
    put(ws, f"B{r}", name, F_INPUT, fill)
    put(ws, f"C{r}", vat_code or None, F_INPUT, fill)
    put(ws, f"D{r}", note, F_INPUT, fill)
    put(ws, f"E{r}", f'=IF($A{r}="","",SUMIFS({GL}!$S:$S,{GL}!$C:$C,$A{r}))', F_LINK, fmt=EUR)
    put(ws, f"F{r}", f'=IF($A{r}="","",COUNTIFS({GL}!$C:$C,$A{r}))', F_LINK)
    put(ws, f"G{r}", f'=IF($A{r}="","",COUNTIFS({GL}!$C:$C,$A{r},{GL}!$Z:$Z,"<>OK"))', F_LINK)


def build_counterparties(wb, cps):
    ws = wb.create_sheet(CPS)
    title(ws, "Counterparties - country, VAT number and inter-company flag",
          "Country = ISO 2-letter code (CY, GR, DE, GB, US, AE...). Inter-company = Y for all group companies. The code must match GL_Input column F.")
    hdr(ws, 4, ["Counterparty code", "Name", "Country (ISO)", "VAT number (with prefix)", "Inter-company (Y/N)",
                "Relationship / notes", "Region", "Lines in GL", "Sales value EUR", "Purchase value EUR"],
        [16, 34, 10, 20, 12, 36, 9, 10, 16, 16])
    for i, cp in enumerate(cps):
        write_cp_row(ws, 5 + i, *cp)
    ws.freeze_panes = "C5"
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E5:E5000")
    ws.conditional_formatting.add(
        "C5:C5000", FormulaRule(formula=['AND($A5<>"",$C5="")'], fill=PatternFill("solid", fgColor="F8CBAD")))
    return ws


def write_cp_row(ws, r, code, name, country, vatno, ic, note, example=False):
    fill = FILL_EXAMPLE if example else FILL_INPUT
    for col, v in zip("ABCDEF", (str(code), name, country, vatno, ic, note)):
        put(ws, f"{col}{r}", v or None, F_INPUT, fill)
    put(ws, f"G{r}", f'=IF($C{r}="","",IF($C{r}="CY","CY",IF(ISNUMBER(MATCH($C{r},{LISTS}!$A:$A,0)),"EU","3RD")))', F_BASE)
    put(ws, f"H{r}", f'=IF($A{r}="","",COUNTIFS({GL}!$F:$F,$A{r}))', F_LINK)
    put(ws, f"I{r}", f'=IF($A{r}="","",SUMIFS({GL}!$U:$U,{GL}!$F:$F,$A{r},{GL}!$T:$T,"Output"))', F_LINK, fmt=EUR)
    put(ws, f"J{r}", f'=IF($A{r}="","",SUMIFS({GL}!$U:$U,{GL}!$F:$F,$A{r},{GL}!$T:$T,"Input"))', F_LINK, fmt=EUR)


def build_codes(wb):
    ws = wb.create_sheet(CODES)
    title(ws, "VAT_Codes - VAT treatments and the return boxes they feed",
          "Y = the code feeds that box. Add your own codes in the blank rows (keep the columns). Recovery: 100% / 0% / pro-rata from Settings.")
    heads = ["Code", "Description", "Nature", "VAT rate", "Expected region", "Line is VAT (Y/N)",
             "Input recovery %", "VIES type"] + [f"Box {b}" for b in BOX_COLS] + \
            ["Net value EUR", "VAT EUR", "VAT reclaimable EUR", "No. of lines", "Guidance / legal basis"]
    widths = [8, 46, 8, 8, 9, 8, 10, 7] + [6] * len(BOX_COLS) + [15, 13, 13, 8, 80]
    hdr(ws, 4, heads, widths, height=45)
    for i, (code, desc, nature, rate, region, isvat, rec, vies, boxes, note) in enumerate(VAT_CODES):
        r = CODES_FIRST + i
        write_code_row(ws, r, code, desc, nature, rate, region, isvat, rec, vies, boxes, note)
    for r in range(CODES_FIRST + len(VAT_CODES), CODES_LAST + 1):
        write_code_row(ws, r, None, None, None, None, None, None, None, None, [], None)
    put(ws, f"B{CODES_LAST + 1}", "TOTAL", F_BOLD)
    for col in "TUVW":
        put(ws, f"{col}{CODES_LAST + 1}", f"=SUM({col}{CODES_FIRST}:{col}{CODES_LAST})", F_BOLD, FILL_KEY, EUR if col != "W" else None)
    ws.freeze_panes = "C5"
    dv = DataValidation(type="list", formula1='"Output,Input,None"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C{CODES_FIRST}:C{CODES_LAST}")
    dv2 = DataValidation(type="list", formula1='"CY,EU,3RD,ANY"', allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add(f"E{CODES_FIRST}:E{CODES_LAST}")


def write_code_row(ws, r, code, desc, nature, rate, region, isvat, rec, vies, boxes, note):
    put(ws, f"A{r}", code, F_INPUT if code else F_BASE, FILL_INPUT)
    put(ws, f"B{r}", desc, F_INPUT, FILL_INPUT)
    put(ws, f"C{r}", nature, F_INPUT, FILL_INPUT)
    put(ws, f"D{r}", rate, F_INPUT, FILL_INPUT, PCT)
    put(ws, f"E{r}", region, F_INPUT, FILL_INPUT)
    put(ws, f"F{r}", isvat, F_INPUT, FILL_INPUT)
    if rec == "PR":
        put(ws, f"G{r}", f"={set_ref(S_PRORATA)}", F_LINK, FILL_CALC, PCT)
    else:
        put(ws, f"G{r}", rec, F_INPUT, FILL_INPUT, PCT)
    put(ws, f"H{r}", vies or None, F_INPUT, FILL_INPUT)
    for b, col in BOX_COLS.items():
        c = put(ws, f"{col}{r}", "Y" if b in boxes else None, F_INPUT, FILL_INPUT)
        c.alignment = Alignment(horizontal="center")
    put(ws, f"T{r}", f'=IF($A{r}="",0,SUMIFS({GL}!$U:$U,{GL}!$N:$N,$A{r}))', F_LINK, FILL_CALC, EUR)
    put(ws, f"U{r}", f'=IF($A{r}="",0,SUMIFS({GL}!$W:$W,{GL}!$N:$N,$A{r}))', F_LINK, FILL_CALC, EUR)
    put(ws, f"V{r}", f"=ROUND(U{r}*N(G{r}),2)", F_BASE, FILL_CALC, EUR)
    put(ws, f"W{r}", f'=IF($A{r}="",0,COUNTIFS({GL}!$N:$N,$A{r}))', F_LINK, FILL_CALC)
    put(ws, f"X{r}", note, F_BASE, wrap=False)


def build_fx(wb, ccys=("USD", "GBP")):
    ws = wb.create_sheet(FX)
    title(ws, "FX_Rates - quarterly average euro reference rates (units of currency per 1 EUR)",
          "Average of the ECB daily reference rates in FX_Daily (the same rates the Central Bank of Cyprus publishes). Column C overrides the average.")
    hdr(ws, 4, ["Currency", "Quarter average (FX_Daily)", "Manual override", "Rate used", "No. of daily rates",
                "First date", "Last date", "Source / date retrieved"], [10, 16, 14, 12, 12, 12, 12, 60])
    for i, ccy in enumerate(ccys):
        write_fx_row(ws, 5 + i, ccy, i, f"{EXAMPLE_TAG} - run the fx command to download the real rates")
    ws.freeze_panes = "B5"

    wd = wb.create_sheet(FXD)
    title(wd, "FX_Daily - ECB euro foreign exchange reference rates (units per 1 EUR)",
          "Written by the fx command. Business days only (no rates on weekends / TARGET holidays).")
    hdr(wd, 4, ["Date"] + list(ccys), [12] + [10] * len(ccys))
    sample = [(dt.date(2026, 7, 1), 1.1702, 0.8601), (dt.date(2026, 8, 3), 1.1655, 0.8662),
              (dt.date(2026, 9, 1), 1.1590, 0.8687)]
    for j, row in enumerate(sample):
        put(wd, f"A{5 + j}", row[0], F_INPUT, FILL_EXAMPLE, DATE)
        for k in range(len(ccys)):
            put(wd, f"{get_column_letter(2 + k)}{5 + j}", row[1 + k] if k < 2 else None, F_INPUT, FILL_EXAMPLE, RATE)
    wd.freeze_panes = "B5"


def write_fx_row(ws, r, ccy, idx, source):
    col = get_column_letter(2 + idx)
    rng = f"{FXD}!${col}$5:${col}$2000"
    put(ws, f"A{r}", ccy, F_BOLD)
    put(ws, f"B{r}", f"=IFERROR(ROUND(AVERAGE({rng}),4),\"\")", F_LINK, FILL_CALC, RATE)
    put(ws, f"C{r}", None, F_INPUT, FILL_INPUT, RATE)
    put(ws, f"D{r}", f'=IF(C{r}<>"",C{r},B{r})', F_BOLD, FILL_KEY, RATE)
    put(ws, f"E{r}", f"=COUNT({rng})", F_LINK)
    put(ws, f"F{r}", f'=IF(E{r}=0,"",MIN({FXD}!$A$5:$A$2000))', F_LINK, fmt=DATE)
    put(ws, f"G{r}", f'=IF(E{r}=0,"",MAX({FXD}!$A$5:$A$2000))', F_LINK, fmt=DATE)
    put(ws, f"H{r}", source, F_NOTE)


def box_sum(box):
    col = BOX_COLS[box]
    val = {"1": "U", "2": "U", "4": "V"}.get(box, "T")
    rng = f"{CODES}!${{c}}${CODES_FIRST}:${{c}}${CODES_LAST}"
    return f'=SUMIFS({rng.format(c=val)},{rng.format(c=col)},"Y")'


def build_return(wb):
    ws = wb.create_sheet("VAT_Return")
    ws.sheet_view.showGridLines = False
    title(ws, "VAT 4 - Cyprus VAT return", "Transfer the amounts in column C to the VAT 4 return on the Tax For All (TFA) portal. Amounts in EUR.")
    info = [("Company", f"={set_ref(S_COMPANY)}", None), ("VAT number", f"={set_ref(S_VATNO)}", None),
            ("Period from", f"={set_ref(S_START)}", DATE), ("Period to", f"={set_ref(S_END)}", DATE),
            ("Return and payment due by", f"={set_ref(S_DUE)}", DATE)]
    for i, (lab, f, fmt) in enumerate(info):
        put(ws, f"A{3 + i}", lab, F_BOLD)
        c = put(ws, f"B{3 + i}", f, F_LINK, fmt=fmt)
        c.alignment = Alignment(horizontal="left")
    hdr(ws, 9, ["Box", "Description (VAT 4)", "Amount EUR", "What this box must include", "Fed by VAT codes"],
        [7, 48, 16, 90, 40], height=24)
    r = 10
    box_row = {}
    for box, desc, include, codes in BOXES:
        box_row[box] = r
        put(ws, f"A{r}", box, F_BOX, FILL_KEY).alignment = Alignment(horizontal="center", vertical="top")
        put(ws, f"B{r}", desc, F_BOLD, wrap=True)
        if box == "3":
            f = f"=C{box_row['1']}+C{box_row['2']}"
        elif box == "5":
            f = f"=C{box_row['3']}-C{box_row['4']}"
        else:
            f = box_sum(box)
        c = put(ws, f"C{r}", f, F_BOX, FILL_KEY, EUR)
        c.alignment = Alignment(vertical="top")
        put(ws, f"D{r}", include, F_BASE, wrap=True)
        put(ws, f"E{r}", codes, F_NOTE, wrap=True)
        ws.row_dimensions[r].height = max(30, 13 * (len(include) // 85 + 1))
        r += 1
    r += 1
    put(ws, f"B{r}", "Result", F_BOLD)
    put(ws, f"C{r}", f'=IF(C{box_row["5"]}>0,"VAT PAYABLE",IF(C{box_row["5"]}<0,"VAT CREDIT","NIL"))', F_BOLD, FILL_KEY)
    put(ws, f"D{r}", "Payable: pay through TFA / JCC by the due date. Credit: carried forward, or refund claim where conditions are met.", F_NOTE, wrap=True)
    r += 1
    put(ws, f"B{r}", "Exceptions outstanding (see Checks)", F_BOLD)
    put(ws, f"C{r}", "=Checks!C23", F_LINK, FILL_KEY)
    put(ws, f"D{r}", "Must be 0 before filing.", F_NOTE)
    r += 2
    put(ws, f"A{r}", "Sign-off", F_H2, border=False)
    for i, lab in enumerate(SIGN_OFF):
        put(ws, f"B{r + 1 + i}", lab, F_BOLD)
        put(ws, f"C{r + 1 + i}", None, F_INPUT, FILL_INPUT)
        ws.merge_cells(f"C{r + 1 + i}:D{r + 1 + i}")
    ws.freeze_panes = "A10"
    return box_row


def build_ic(wb):
    ws = wb.create_sheet("IC_Analysis")
    title(ws, "IC_Analysis - inter-company vs third-party, Cyprus / EU / third countries (return value EUR, excl. VAT)",
          "Use it to confirm every inter-company flow has the right treatment: EU group = EUS / AES19, non-EU group = OS3 / A3S19.")
    groups = [("Inter-company", "Y"), ("Third party", "N")]
    regions = ["CY", "EU", "3RD"]
    heads = ["Code", "Description", "Nature"] + [f"{g} - {rg}" for g, _ in groups for rg in regions] + \
            ["No country", "Total", "Agrees to VAT_Codes"]
    hdr(ws, 4, heads, [8, 44, 8] + [14] * 6 + [12, 15, 12], height=40)
    for i in range(CODES_LAST - CODES_FIRST + 1):
        r = 5 + i
        cr = CODES_FIRST + i
        put(ws, f"A{r}", f'=IF({CODES}!A{cr}="","",{CODES}!A{cr})', F_LINK)
        put(ws, f"B{r}", f'=IF({CODES}!B{cr}="","",{CODES}!B{cr})', F_LINK)
        put(ws, f"C{r}", f'=IF({CODES}!C{cr}="","",{CODES}!C{cr})', F_LINK)
        col = 4
        for _, flag in groups:
            for rg in regions:
                put(ws, f"{get_column_letter(col)}{r}",
                    f'=IF($A{r}="",0,SUMIFS({GL}!$U:$U,{GL}!$N:$N,$A{r},{GL}!$Q:$Q,"{flag}",{GL}!$P:$P,"{rg}"))',
                    F_BASE, fmt=EUR)
                col += 1
        put(ws, f"J{r}", f'=IF($A{r}="",0,SUMIFS({GL}!$U:$U,{GL}!$N:$N,$A{r},{GL}!$P:$P,""))', F_BASE, fmt=EUR)
        put(ws, f"K{r}", f"=SUM(D{r}:J{r})", F_BOLD, fmt=EUR)
        put(ws, f"L{r}", f'=IF(ROUND(K{r}-{CODES}!T{cr},2)=0,"OK","CHECK")', F_BASE)
    last = 5 + CODES_LAST - CODES_FIRST
    tr = last + 1
    put(ws, f"B{tr}", "TOTAL", F_BOLD)
    for col in range(4, 12):
        L = get_column_letter(col)
        put(ws, f"{L}{tr}", f"=SUM({L}5:{L}{last})", F_BOLD, FILL_KEY, EUR)
    tr += 2
    put(ws, f"B{tr}", "Summary by nature", F_H2, border=False)
    for j, nature in enumerate(["Output", "Input"]):
        rr = tr + 1 + j
        put(ws, f"B{rr}", f"Total {nature.lower()}s", F_BOLD)
        put(ws, f"C{rr}", nature)
        for col in range(4, 12):
            L = get_column_letter(col)
            put(ws, f"{L}{rr}", f'=SUMIFS({L}$5:{L}${last},$C$5:$C${last},$C{rr})', F_BOLD, fmt=EUR)
    ws.freeze_panes = "D5"


def build_vies(wb, n_rows=300):
    ws = wb.create_sheet("VIES")
    title(ws, "VIES - recapitulative statement amounts per EU customer (EUR)",
          "Filed monthly by the 15th of the following month. Rows follow the Counterparties list; filter Total <> 0.")
    heads = ["Counterparty", "Name", "Country", "VAT number", "IC",
             "Goods M1", "Goods M2", "Goods M3", "Triang. M1", "Triang. M2", "Triang. M3",
             "Services M1", "Services M2", "Services M3", "Total", "VAT no. check"]
    hdr(ws, 4, heads, [14, 30, 8, 18, 5] + [12] * 10 + [22], height=30)
    write_vies_rows(ws, n_rows)
    ws.freeze_panes = "F5"


def write_vies_rows(ws, n_rows):
    for i in range(n_rows):
        r = 5 + i
        cr = 5 + i
        for col, src in zip("ABCDE", "ABCDE"):
            put(ws, f"{col}{r}", f'=IF({CPS}!{src}{cr}="","",{CPS}!{src}{cr})', F_LINK)
        col = 6
        for vtype in ("G", "T", "S"):
            for m in (1, 2, 3):
                put(ws, f"{get_column_letter(col)}{r}",
                    f'=IF($A{r}="",0,SUMIFS({GL}!$U:$U,{GL}!$F:$F,$A{r},{GL}!$X:$X,"{vtype}",{GL}!$Y:$Y,{m}))',
                    F_BASE, fmt=EUR)
                col += 1
        put(ws, f"O{r}", f"=SUM(F{r}:N{r})", F_BOLD, fmt=EUR)
        put(ws, f"P{r}", f'=IF(O{r}=0,"",IF($D{r}="","MISSING VAT NO.",IF(OR($C{r}="",$C{r}="CY",{CPS}!G{cr}<>"EU"),"NOT AN EU CUSTOMER","OK")))', F_BASE)
    tr = 5 + n_rows
    put(ws, f"B{tr}", "TOTAL", F_BOLD)
    for col in range(6, 16):
        L = get_column_letter(col)
        put(ws, f"{L}{tr}", f"=SUM({L}5:{L}{tr - 1})", F_BOLD, FILL_KEY, EUR)
    return tr


def build_recon(wb, box_row):
    ws = wb.create_sheet("Reconciliation")
    ws.sheet_view.showGridLines = False
    title(ws, "Reconciliation - VAT return vs GL VAT control accounts (EUR)",
          "VAT control account codes are set in Settings. GL movements are for the lines in GL_Input (the quarter).")
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 64
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 70
    vr = "VAT_Return"
    rc_out = f'SUMIFS({CODES}!$U${CODES_FIRST}:$U${CODES_LAST},{CODES}!$I${CODES_FIRST}:$I${CODES_LAST},"Y",{CODES}!$C${CODES_FIRST}:$C${CODES_LAST},"Input")'
    rc_in = f'SUMIFS({CODES}!$V${CODES_FIRST}:$V${CODES_LAST},{CODES}!$K${CODES_FIRST}:$K${CODES_LAST},"Y",{CODES}!$I${CODES_FIRST}:$I${CODES_LAST},"Y")'
    rc_in2 = f'SUMIFS({CODES}!$V${CODES_FIRST}:$V${CODES_LAST},{CODES}!$K${CODES_FIRST}:$K${CODES_LAST},"Y",{CODES}!$J${CODES_FIRST}:$J${CODES_LAST},"Y")'
    ivat = f'SUMIFS({CODES}!$V${CODES_FIRST}:$V${CODES_LAST},{CODES}!$F${CODES_FIRST}:$F${CODES_LAST},"Y")'
    out_gl = "+".join(f'SUMIFS({GL}!$S:$S,{GL}!$C:$C,{SET}!$C${r}&"")*({SET}!$C${r}<>"")' for r in OUT_VAT_ROWS)
    in_gl = "+".join(f'SUMIFS({GL}!$S:$S,{GL}!$C:$C,{SET}!$C${r}&"")*({SET}!$C${r}<>"")' for r in IN_VAT_ROWS)
    rc_flag = set_ref(S_RC_IN_GL)
    tol = set_ref(S_TOL)
    rows = [
        ("h", "A. Output VAT", None, None),
        ("", "Box 1 per return", f"={vr}!C{box_row['1']}", ""),
        ("", "Less: reverse-charge VAT self-accounted in Box 1 (services from abroad, domestic reverse charge)", f"=-{rc_out}", ""),
        ("", "Add: Box 2 (acquisitions of goods) if booked in the output VAT account", f"=IF({rc_flag}=\"Y\",{vr}!C{box_row['2']},0)", "Only when reverse-charge VAT is posted in the GL (Settings C11 = Y)."),
        ("", "Add back reverse-charge VAT if booked in the GL VAT accounts", f"=IF({rc_flag}=\"Y\",{rc_out},0)", ""),
        ("b", "Output VAT expected in the GL", "=SUM(C5:C8)", ""),
        ("", "GL output VAT control accounts - net credit movement", f"=-({out_gl})", "Credit movements shown positive."),
        ("b", "Difference", "=ROUND(C9-C10,2)", ""),
        ("", "Status", f'=IF(ABS(C11)<={tol},"OK","INVESTIGATE")', "Typical causes: wrong account, VAT rate errors, invoices outside period, manual journals."),
        ("h", "B. Input VAT", None, None),
        ("", "Box 4 per return", f"={vr}!C{box_row['4']}", ""),
        ("", "Less: reclaimable reverse-charge VAT on services / domestic reverse charge", f"=-{rc_in}", ""),
        ("", "Less: reclaimable VAT on EU acquisitions of goods (Box 2)", f"=-{rc_in2}", ""),
        ("", "Less: import VAT (IVAT - separate account)", f"=-{ivat}", ""),
        ("", "Add back reverse-charge VAT if booked in the GL VAT accounts", f"=IF({rc_flag}=\"Y\",{rc_in}+{rc_in2},0)", ""),
        ("b", "Input VAT expected in the GL", "=SUM(C14:C18)", ""),
        ("", "GL input VAT control accounts - net debit movement", f"={in_gl}", "Blocked / irrecoverable VAT should be expensed, not left in the input VAT account."),
        ("b", "Difference", "=ROUND(C19-C20,2)", ""),
        ("", "Status", f'=IF(ABS(C21)<={tol},"OK","INVESTIGATE")', ""),
        ("h", "C. Values", None, None),
        ("", "Box 6 must be at least Boxes 8A + 8B + 9 + 10 + 11B", f'=IF({vr}!C{box_row["6"]}+0.005>={vr}!C{box_row["8A"]}+{vr}!C{box_row["8B"]}+{vr}!C{box_row["9"]}+{vr}!C{box_row["10"]}+{vr}!C{box_row["11B"]},"OK","CHECK")', ""),
        ("", "Box 7 must be at least Boxes 11A + 11B", f'=IF({vr}!C{box_row["7"]}+0.005>={vr}!C{box_row["11A"]}+{vr}!C{box_row["11B"]},"OK","CHECK")', ""),
        ("", "Total sales per GL (all output codes) vs Box 6", f'=ROUND(SUMIFS({CODES}!$T${CODES_FIRST}:$T${CODES_LAST},{CODES}!$C${CODES_FIRST}:$C${CODES_LAST},"Output")-({vr}!C{box_row["6"]}-SUMIFS({CODES}!$T${CODES_FIRST}:$T${CODES_LAST},{CODES}!$C${CODES_FIRST}:$C${CODES_LAST},"Input",{CODES}!$L${CODES_FIRST}:$L${CODES_LAST},"Y")),2)', "Should be 0 (Box 6 also contains the reverse-charge services values)."),
    ]
    r = 4
    for style, label, f, note in rows:
        if style == "h":
            put(ws, f"B{r}", label, F_H2, border=False)
        else:
            put(ws, f"B{r}", label, F_BOLD if style == "b" else F_BASE, wrap=True)
            fmt = None if label in ("Status",) or label.startswith("Box") and "must" in label else EUR
            put(ws, f"C{r}", f, F_BOLD if style == "b" else F_LINK, FILL_KEY if style == "b" else None, fmt)
            put(ws, f"D{r}", note, F_NOTE, wrap=True)
        r += 1


def build_checks(wb):
    ws = wb.create_sheet("Checks")
    ws.sheet_view.showGridLines = False
    title(ws, "Checks - exception report", "Every line must be 0 / OK before the return is filed.")
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 70
    G = GL
    checks = [
        ("GL lines with no VAT code (map the account or override in column K)", f'=COUNTIFS({G}!$Z:$Z,"NO VAT CODE")', "Filter GL_Input column Z."),
        ("GL lines with a code not in VAT_Codes", f'=COUNTIFS({G}!$Z:$Z,"UNKNOWN VAT CODE")', ""),
        ("GL lines with a missing FX rate", f'=COUNTIFS({G}!$Z:$Z,"MISSING FX RATE")', "Run the fx command or enter a manual override in FX_Rates."),
        ("GL lines dated outside the VAT period", f'=COUNTIFS({G}!$Z:$Z,"OUTSIDE PERIOD")', "Remove them or confirm the tax point is in this quarter."),
        ("GL lines needing a counterparty country (EU / 3rd-country codes)", f'=COUNTIFS({G}!$Z:$Z,"MISSING COUNTERPARTY COUNTRY")', "Add the counterparty code in GL column F and its country in Counterparties."),
        ("GL lines where the code does not fit the counterparty's region", f'=COUNTIFS({G}!$Z:$Z,"REGION MISMATCH")', "e.g. EUS on a US customer, A3S19 on a German supplier."),
        ("Accounts in Account_Mapping without a VAT code", f'=COUNTIFS({MAPS}!$A$5:$A$5000,"<>",{MAPS}!$C$5:$C$5000,"")', ""),
        ("GL accounts missing from Account_Mapping", f'=SUMPRODUCT(({G}!$C${GL_FIRST}:$C$60000<>"")*ISNA(MATCH({G}!$C${GL_FIRST}:$C$60000&"",{MAPS}!$A$5:$A$5000,0)))', "Run import-gl again or add the accounts."),
        ("Counterparties used in the GL without a country", f'=COUNTIFS({CPS}!$A$5:$A$5000,"<>",{CPS}!$C$5:$C$5000,"",{CPS}!$H$5:$H$5000,">0")', ""),
        ("EU customers in VIES without a VAT number / not in the EU", '=COUNTIFS(VIES!$P:$P,"MISSING VAT NO.")+COUNTIFS(VIES!$P:$P,"NOT AN EU CUSTOMER")', ""),
        ("VIES goods total less Box 8A", f'=ROUND(SUMIFS(VIES!F:F,VIES!A:A,"<>")+SUMIFS(VIES!G:G,VIES!A:A,"<>")+SUMIFS(VIES!H:H,VIES!A:A,"<>")-VAT_Return!C{{8A}},2)', "Should be 0."),
        ("VIES services total less Box 8B", f'=ROUND(SUMIFS(VIES!L:L,VIES!A:A,"<>")+SUMIFS(VIES!M:M,VIES!A:A,"<>")+SUMIFS(VIES!N:N,VIES!A:A,"<>")-VAT_Return!C{{8B}},2)', "Should be 0."),
        ("Output VAT reconciliation", '=Reconciliation!C12', ""),
        ("Input VAT reconciliation", '=Reconciliation!C22', ""),
        ("Box 6 / Box 7 completeness", '=IF(AND(Reconciliation!C24="OK",Reconciliation!C25="OK"),"OK","CHECK")', ""),
        ("Exchange rates still the EXAMPLE placeholders", f'=COUNTIF({FX}!$H:$H,"{EXAMPLE_TAG}*")', "Run the fx command (or enter manual overrides and change the note)."),
        ("Inter-company analysis agrees to VAT_Codes", f'=IF(COUNTIF(IC_Analysis!$L$5:$L${5 + CODES_LAST - CODES_FIRST},"CHECK")=0,"OK","CHECK")', ""),
    ]
    hdr_row = 4
    for i, t in enumerate(["Check", "Result", "What to do"]):
        c = ws.cell(row=hdr_row, column=2 + i, value=t)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
    return ws, checks


def finish_checks(ws, checks, box_row):
    r = 5
    for label, f, note in checks:
        f = f.replace("{8A}", str(box_row["8A"])).replace("{8B}", str(box_row["8B"]))
        put(ws, f"B{r}", label, wrap=True)
        put(ws, f"C{r}", f, F_LINK, fmt="#,##0.00;-#,##0.00;0")
        put(ws, f"D{r}", note, F_NOTE, wrap=True)
        r += 1
    # C23 = total number of open exceptions (referenced by VAT_Return)
    assert r <= 22, "too many checks for the fixed total row"
    put(ws, "B23", "TOTAL OPEN EXCEPTIONS", F_BOLD)
    put(ws, "C23", '=COUNTIF(C5:C22,">0.004")+COUNTIF(C5:C22,"<-0.004")+COUNTIF(C5:C22,"INVESTIGATE")+COUNTIF(C5:C22,"CHECK")',
        F_BOLD, FILL_KEY)
    ws.conditional_formatting.add("C5:C23", FormulaRule(
        formula=['OR(AND(ISNUMBER(C5),ABS(C5)>0.004),C5="CHECK",C5="INVESTIGATE")'],
        fill=PatternFill("solid", fgColor="F8CBAD")))


def build_lists(wb):
    ws = wb.create_sheet(LISTS)
    title(ws, "Lists - EU member states (other than Cyprus)", "Any country code not listed here (and not CY) is treated as a third country.")
    hdr(ws, 4, ["ISO code", "Member state"], [10, 34])
    for i, (c, n) in enumerate(EU_MEMBERS):
        put(ws, f"A{5 + i}", c, F_INPUT, FILL_INPUT)
        put(ws, f"B{5 + i}", n)
    put(ws, f"A{6 + len(EU_MEMBERS)}", "Note: the UK (GB) is a third country since 1.1.2021 (Northern Ireland - XI - stays in the EU VAT area for goods only).", F_NOTE, border=False)


def build_legal(wb):
    ws = wb.create_sheet("Legal_Notes")
    ws.sheet_view.showGridLines = False
    title(ws, "Legal_Notes - Cyprus VAT rules used in this working (as at October 2026)",
          "Summary for preparation purposes only. Check the current law and Tax Department announcements every quarter.")
    rows = [
        ("h", "Law"),
        ("", "Value Added Tax Law N.95(I)/2000 as amended, and the regulations / decisions issued under it; EU VAT Directive 2006/112/EC; Implementing Regulation 282/2011."),
        ("h", "VAT rates"),
        ("", "19% standard rate - all supplies not covered by a reduced, zero or exempt schedule."),
        ("", "9% reduced rate - e.g. hotel accommodation, restaurant and catering services (excluding alcoholic drinks, beer, wine, soft drinks), certain passenger transport."),
        ("", "5% reduced rate - e.g. foodstuffs, water, pharmaceuticals, vaccines, books (where not 3%), qualifying first residence (with conditions amended from 1 September 2026), renovation of private homes over 3 years old."),
        ("", "3% super-reduced rate - e.g. books, newspapers and periodicals, and some other items introduced in 2023."),
        ("", "0% - exports, intra-EU supplies of goods, qualifying ships / aircraft supplies, and items temporarily zero-rated by decree (several temporary zero rates on basic goods run through 2026-2027 - check the current decree for each item)."),
        ("h", "Filing and payment"),
        ("", "Quarterly VAT 4 return and payment due by the 10th day of the second month after the end of the quarter (e.g. Q3 Jul-Sep: 10 November). Filed electronically on the Tax For All (TFA) portal."),
        ("", "VIES recapitulative statement: monthly, by the 15th day after the end of each month in which intra-EU supplies of goods or general-rule B2B services were made. Corrections within one month."),
        ("", "Intrastat (Statistical Service) is separate and applies when intra-EU goods flows exceed the annual thresholds."),
        ("h", "Penalties (summary - verify the current amounts)"),
        ("", "Late return: EUR 100 per return. Late payment: additional tax of 10% of the VAT due plus interest at the public-interest rate set by the Minister of Finance. "
             "Late VIES statement: EUR 50; incorrect VIES not corrected: EUR 15. Failure to apply the reverse charge: reported as EUR 200 per period up to EUR 4,000 (secondary sources - verify)."),
        ("h", "Place of supply of services and reverse charge"),
        ("", "B2B general rule (s.14 / Art. 44): services are supplied where the customer is established. Cyprus company -> EU business customer: no Cyprus VAT, Box 6 + 8B + VIES, invoice marked 'Reverse charge'. "
             "Cyprus company -> non-EU business customer: outside scope with right to deduct, Box 6 + 10."),
        ("", "Services received by the Cyprus company from abroad (EU or non-EU) under the general rule: the Cyprus company self-accounts for VAT at the Cyprus rate (reverse charge, s.11): "
             "VAT in Box 1, deductible part in Box 4, value in Box 6 and Box 7; services from EU suppliers also in Box 11B."),
        ("", "Inter-company transactions are normal supplies for VAT (management fees, cost recharges, IT / licence fees, guarantees, secondments). Recharges at cost are still supplies. "
             "Inter-company loans: interest is an exempt financial service - to EU recipients: Box 6 only (no right to deduct); to non-EU recipients: Box 6 + 10 (with right to deduct). Dividends are outside the scope - not reported."),
        ("", "Exceptions to the general rule (check case by case): services connected with immovable property (where the property is), passenger transport, admission to events, restaurant/catering, short-term hire of means of transport, B2C supplies."),
        ("h", "Goods"),
        ("", "Intra-EU supply (0%): customer VAT number valid on VIES, goods leave Cyprus, evidence of transport (Art. 45a Reg. 282/2011), included in VIES. Otherwise Cyprus VAT applies."),
        ("", "Acquisition from an EU supplier: self-account VAT in Box 2, recover in Box 4, value in Boxes 7 and 11A."),
        ("", "Imports from third countries: import VAT is paid to / accounted with Customs (customs declaration is the evidence); recovered in Box 4, goods value in Box 7. Exports: 0%, Boxes 6 and 9, keep export evidence."),
        ("h", "Domestic reverse charge (s.11B / 11C and related decrees)"),
        ("", "The Cyprus customer accounts for the VAT on, among others: construction and related works to a taxable person, scrap and used metals, mobile phones, integrated-circuit devices, game consoles, tablets and laptops, "
             "raw and semi-finished precious metals, gas and electricity from non-established suppliers, and immovable property transferred in a loan restructuring. The supplier invoices without VAT."),
        ("h", "Input VAT - not deductible"),
        ("", "Business entertainment (except staff entertainment in specific cases), private / non-business use, saloon cars (purchase, hire, with limited exceptions) and their fuel where not wholly business, "
             "costs relating to exempt supplies without right to deduct, VAT without a valid tax invoice / import document, foreign VAT. Mixed-use costs: provisional pro-rata during the year, annual adjustment in the last return of the year."),
        ("h", "Exchange rates"),
        ("", "Art. 91(2) VAT Directive: foreign-currency amounts are converted at the latest selling rate of the most representative exchange market, or the latest ECB rate, at the time VAT becomes chargeable. "
             "The ECB euro reference rates are published daily (around 16:00 CET) and republished by the Central Bank of Cyprus. This working uses the quarterly average of those rates as a practical simplification - "
             "confirm acceptance with your adviser. Where an invoice already states the VAT in EUR, that amount should be used (enter it in GL_Input column J / L)."),
        ("h", "Record keeping"),
        ("", "Keep invoices, customs documents, VIES evidence and this working for at least 6 years."),
        ("h", "Sources"),
        ("", "Cyprus Tax Department - VAT guide 'Fill in your VAT Return': https://www.mof.gov.cy/mof/tax/taxdep.nsf (VAT > Guides)"),
        ("", "Cyprus Tax Department - VAT legislation and announcements: https://www.mof.gov.cy/mof/tax/taxdep.nsf"),
        ("", "Central Bank of Cyprus - euro reference rates: https://www.centralbank.cy/en/rates/euro-reference-rates"),
        ("", "European Central Bank - euro foreign exchange reference rates / data API: https://data.ecb.europa.eu/"),
        ("", "PwC Worldwide Tax Summaries - Cyprus, Other taxes: https://taxsummaries.pwc.com/cyprus/corporate/other-taxes"),
        ("", "KPMG - Cyprus reduced VAT rates for certain goods and services (2023): https://kpmg.com/us/en/home/insights/2023/07/tnf-cyprus-reduced-vat-rates-certain-goods-services.html"),
        ("", "EU VIES VAT number validation: https://ec.europa.eu/taxation_customs/vies/"),
    ]
    text_rows(ws, 4, rows)


# --------------------------------------------------------------------------
# Example data (shows the expected format; removed by import-gl)
# --------------------------------------------------------------------------
def example_data(start: dt.date):
    d = lambda m, day: dt.date(start.year, start.month + m, day)  # noqa: E731
    gl = [
        dict(date=d(0, 5), doc="SI-1001", account="4000", account_name="Sales - Cyprus customers", description="Consulting services July",
             cp_code="C-CY01", cp_name="Limassol Trading Ltd", currency="EUR", amount_fc=-10000),
        dict(date=d(0, 5), doc="SI-1001", account="2201", account_name="Output VAT", description="VAT on SI-1001",
             cp_code="C-CY01", cp_name="Limassol Trading Ltd", currency="EUR", amount_fc=-1900),
        dict(date=d(0, 28), doc="SI-1002", account="4100", account_name="Management fees - group", description="Q3 management fee",
             cp_code="IC-DE01", cp_name="Example GmbH (group)", currency="EUR", amount_fc=-25000),
        dict(date=d(1, 28), doc="SI-1003", account="4100", account_name="Management fees - group", description="Q3 management fee",
             cp_code="IC-AE01", cp_name="Example FZE Dubai (group)", currency="USD", amount_fc=-30000),
        dict(date=d(1, 12), doc="SI-1004", account="4300", account_name="Sales of goods - export / EU", description="Goods to Italy",
             cp_code="C-IT01", cp_name="Milano Srl", currency="EUR", amount_fc=-8000),
        dict(date=d(2, 3), doc="SI-1005", account="4300", account_name="Sales of goods - export / EU", description="Goods to USA",
             cp_code="C-US01", cp_name="Boston Inc", currency="USD", amount_fc=-12000),
        dict(date=d(0, 10), doc="PI-501", account="6000", account_name="Professional fees", description="Audit fee",
             cp_code="S-CY01", cp_name="Nicosia Audit Ltd", currency="EUR", amount_fc=2000),
        dict(date=d(0, 10), doc="PI-501", account="2101", account_name="Input VAT", description="VAT on PI-501",
             cp_code="S-CY01", cp_name="Nicosia Audit Ltd", currency="EUR", amount_fc=380),
        dict(date=d(1, 20), doc="PI-502", account="6100", account_name="Group recharges", description="IT recharge from parent",
             cp_code="IC-NL01", cp_name="Example BV (parent)", currency="EUR", amount_fc=15000),
        dict(date=d(1, 2), doc="PI-503", account="6000", account_name="Professional fees", description="Software subscription",
             cp_code="S-US01", cp_name="Cloudsoft Inc", currency="USD", amount_fc=1200, vat_code="A3S19"),
        dict(date=d(2, 8), doc="PI-504", account="5000", account_name="Purchases of goods", description="Goods from Germany",
             cp_code="S-DE01", cp_name="Hamburg Supplies GmbH", currency="EUR", amount_fc=5000),
        dict(date=d(2, 15), doc="PI-505", account="5010", account_name="Imports of goods", description="Goods from China",
             cp_code="S-CN01", cp_name="Shenzhen Mfg Co", currency="USD", amount_fc=4000),
        dict(date=d(2, 15), doc="C-77", account="2105", account_name="Import VAT paid to Customs", description="Import VAT SAD 26CY...",
             currency="EUR", amount_fc=700),
        dict(date=d(1, 25), doc="PI-506", account="6500", account_name="Entertainment", description="Client dinner",
             cp_code="S-CY02", cp_name="Paphos Restaurant Ltd", currency="EUR", amount_fc=357),
        dict(date=d(0, 28), doc="PAY-07", account="7000", account_name="Salaries", description="July payroll",
             currency="EUR", amount_fc=20000),
        dict(date=d(0, 28), doc="PAY-07", account="1200", account_name="Bank - EUR", description="July payroll",
             currency="EUR", amount_fc=-20000),
    ]
    accounts = [
        ("1200", "Bank - EUR", "NR", ""), ("2101", "Input VAT", "NR", "VAT control account"),
        ("2105", "Import VAT paid to Customs", "IVAT", "Line amount = import VAT"),
        ("2201", "Output VAT", "NR", "VAT control account"), ("4000", "Sales - Cyprus customers", "S19", ""),
        ("4100", "Management fees - group", "EUS", "EU group = EUS; override non-EU lines?"),
        ("4300", "Sales of goods - export / EU", "EUG", ""), ("5000", "Purchases of goods", "AEG19", ""),
        ("5010", "Imports of goods", "IMP", ""), ("6000", "Professional fees", "P19", ""),
        ("6100", "Group recharges", "AES19", ""), ("6500", "Entertainment", "PB19", "Blocked input VAT"),
        ("7000", "Salaries", "NR", ""),
    ]
    cps = [
        ("C-CY01", "Limassol Trading Ltd", "CY", "CY10111111A", "N", ""),
        ("IC-DE01", "Example GmbH (group)", "DE", "DE123456789", "Y", "Sister company"),
        ("IC-AE01", "Example FZE Dubai (group)", "AE", "", "Y", "Sister company"),
        ("C-IT01", "Milano Srl", "IT", "IT01234567890", "N", ""),
        ("C-US01", "Boston Inc", "US", "", "N", ""),
        ("S-CY01", "Nicosia Audit Ltd", "CY", "CY10222222B", "N", ""),
        ("IC-NL01", "Example BV (parent)", "NL", "NL123456789B01", "Y", "Parent company"),
        ("S-US01", "Cloudsoft Inc", "US", "", "N", ""),
        ("S-DE01", "Hamburg Supplies GmbH", "DE", "DE987654321", "N", ""),
        ("S-CN01", "Shenzhen Mfg Co", "CN", "", "N", ""),
        ("S-CY02", "Paphos Restaurant Ltd", "CY", "CY10333333C", "N", ""),
    ]
    # Line-level overrides that the mapping alone cannot express
    gl[3]["vat_code"] = "OS3"    # management fee to non-EU group company
    gl[5]["vat_code"] = "EXPG"   # goods to USA = export
    gl[13]["vat_fc"] = 57        # entertainment: actual VAT on invoice
    gl[13]["amount_fc"] = 300
    accounts = [(a, n, c, f"{EXAMPLE_TAG} {note}".strip()) for a, n, c, note in accounts]
    cps = [(*cp[:5], f"{EXAMPLE_TAG} {cp[5]}".strip()) for cp in cps]
    return gl, accounts, cps


# --------------------------------------------------------------------------
# Commands
# --------------------------------------------------------------------------
def cmd_new(args):
    start, end = quarter_dates(args.year, args.quarter)
    out = Path(args.workbook)
    if out.exists() and not args.force:
        raise SystemExit(f"{out} exists - use --force to overwrite")
    wb = Workbook()
    wb.remove(wb.active)
    gl, accounts, cps = example_data(start)
    build_instructions(wb)
    build_settings(wb, start, end, f"Q{args.quarter} {args.year}")
    box_row = build_return(wb)
    build_gl(wb, examples=gl)
    build_mapping(wb, [(*a, True) for a in accounts])
    build_counterparties(wb, [(*c, True) for c in cps])
    build_codes(wb)
    build_fx(wb)
    build_ic(wb)
    build_vies(wb)
    build_recon(wb, box_row)
    ws, checks = build_checks(wb)
    finish_checks(ws, checks, box_row)
    build_lists(wb)
    build_legal(wb)
    order = ["Instructions", SET, GL, MAPS, CPS, CODES, FX, FXD, "VAT_Return", "IC_Analysis", "VIES",
             "Reconciliation", "Checks", "Legal_Notes", LISTS]
    wb._sheets = [wb[n] for n in order]
    wb["VAT_Return"].sheet_properties.tabColor = "1F3864"
    wb["Instructions"].sheet_properties.tabColor = "00B050"
    wb["Checks"].sheet_properties.tabColor = "C00000"
    for n in (GL, MAPS, CPS, SET):
        wb[n].sheet_properties.tabColor = "FFC000"
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    print(f"Created {out}  (period {start:%d/%m/%Y} - {end:%d/%m/%Y})")


def read_settings_map(ws_set):
    m = {}
    r = MAP_FIRST
    while ws_set[f"B{r}"].value:
        key = str(ws_set[f"B{r}"].value).strip()
        val = ws_set[f"C{r}"].value
        m[key] = str(val).strip() if val not in (None, "") else None
        r += 1
    return m


def load_gl_file(path: Path, sheet: str | None):
    import pandas as pd
    if path.suffix.lower() in (".csv", ".txt"):
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    return pd.read_excel(path, sheet_name=sheet or 0, dtype=object)


def to_number(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return None if v != v else float(v)  # NaN check
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "nan", "NaT", "None"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        x = float(s)
    except ValueError:
        raise SystemExit(f"Cannot read amount '{v}' - check the GL column map in Settings")
    return -x if neg else x


def to_date(v):
    import pandas as pd
    if v is None or (isinstance(v, float) and v != v) or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date() if not isinstance(v, pd.Timestamp) else v.to_pydatetime().date()
    if isinstance(v, dt.date):
        return v
    ts = pd.to_datetime(str(v), dayfirst=True, errors="coerce")
    if pd.isna(ts):
        raise SystemExit(f"Cannot read date '{v}' - check the date column")
    return ts.date()


def clean_text(v):
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s or None


def cmd_import(args):
    wb = load_workbook(args.workbook)
    colmap = read_settings_map(wb[SET])
    df = load_gl_file(Path(args.gl), args.sheet)
    headers = {str(c).strip(): c for c in df.columns}

    def col(key):
        h = colmap.get(key)
        if not h:
            return None
        if h not in headers:
            raise SystemExit(f"Column '{h}' (Settings, field '{key}') not found in the GL file. Columns are: {list(headers)}")
        return headers[h]

    need = ["date", "account"]
    for k in need:
        if not col(k):
            raise SystemExit(f"Settings GL column map: '{k}' is required")
    if not col("amount_fc") and not (col("debit_fc") or col("credit_fc")) and not col("amount_eur"):
        raise SystemExit("Settings GL column map: give 'amount_fc', or 'debit_fc' and 'credit_fc', or 'amount_eur'")

    rows = []
    for rec in df.to_dict("records"):
        g = lambda k: rec.get(col(k)) if col(k) else None  # noqa: E731
        account = clean_text(g("account"))
        if not account:
            continue
        if col("amount_fc"):
            amt = to_number(g("amount_fc"))
        elif col("debit_fc") or col("credit_fc"):
            amt = (to_number(g("debit_fc")) or 0) - (to_number(g("credit_fc")) or 0)
        else:
            amt = None
        rows.append(dict(
            date=to_date(g("date")), doc=clean_text(g("doc")), account=account,
            account_name=clean_text(g("account_name")), description=clean_text(g("description")),
            cp_code=clean_text(g("cp_code")), cp_name=clean_text(g("cp_name")),
            currency=(clean_text(g("currency")) or "EUR").upper(), amount_fc=amt,
            amount_eur=to_number(g("amount_eur")), vat_code=clean_text(g("vat_code")),
            vat_fc=to_number(g("vat_fc")),
        ))
    if not rows:
        raise SystemExit("No GL lines found")

    # --- rewrite GL_Input
    ws = wb[GL]
    ws.auto_filter.ref = None
    if ws.max_row >= GL_FIRST:
        ws.delete_rows(GL_FIRST, ws.max_row - GL_FIRST + 1)
    for i, row in enumerate(rows):
        write_gl_row(ws, GL_FIRST + i, row)
    for i in range(len(rows), len(rows) + 200):          # spare formula rows for manual additions
        write_gl_row(ws, GL_FIRST + i, None)
    last = GL_FIRST + len(rows) + 199
    ws.auto_filter.ref = f"A{GL_HDR_ROW}:Z{last}"

    # --- account mapping: drop examples, keep user rows, add new accounts
    wm = wb[MAPS]
    existing = []
    for r in range(5, wm.max_row + 1):
        code = clean_text(wm[f"A{r}"].value)
        if not code:
            continue
        note = wm[f"D{r}"].value or ""
        if str(note).startswith(EXAMPLE_TAG):
            continue
        existing.append((code, wm[f"B{r}"].value, wm[f"C{r}"].value, note))
    known = {a[0] for a in existing}
    added_acc = []
    for row in rows:
        if row["account"] not in known:
            known.add(row["account"])
            existing.append((row["account"], row["account_name"], None, "NEW - assign VAT code"))
            added_acc.append(row["account"])
    if wm.max_row >= 5:
        wm.delete_rows(5, wm.max_row - 4)
    for i, a in enumerate(existing):
        write_mapping_row(wm, 5 + i, *a)

    # --- counterparties
    wc = wb[CPS]
    cps = []
    for r in range(5, wc.max_row + 1):
        code = clean_text(wc[f"A{r}"].value)
        if not code:
            continue
        note = wc[f"F{r}"].value or ""
        if str(note).startswith(EXAMPLE_TAG):
            continue
        cps.append((code, wc[f"B{r}"].value, wc[f"C{r}"].value, wc[f"D{r}"].value, wc[f"E{r}"].value or "N", note))
    knowncp = {c[0] for c in cps}
    added_cp = []
    for row in rows:
        c = row["cp_code"]
        if c and c not in knowncp:
            knowncp.add(c)
            cps.append((c, row["cp_name"], None, None, "N", "NEW - add country / VAT no. / IC flag"))
            added_cp.append(c)
    if wc.max_row >= 5:
        wc.delete_rows(5, wc.max_row - 4)
    for i, c in enumerate(cps):
        write_cp_row(wc, 5 + i, *c)

    # --- VIES rows must cover every counterparty
    wv = wb["VIES"]
    if wv.max_row >= 5:
        wv.delete_rows(5, wv.max_row - 4)
    write_vies_rows(wv, max(300, len(cps) + 100))

    # --- FX rows for currencies in the GL
    ensure_fx_currencies(wb, sorted({r["currency"] for r in rows if r["currency"] != "EUR"}))

    wb.calculation.fullCalcOnLoad = True
    wb.save(args.workbook)
    print(f"Imported {len(rows)} GL lines into {args.workbook}")
    if added_acc:
        print(f"  {len(added_acc)} new accounts added to Account_Mapping - assign VAT codes: {', '.join(added_acc[:15])}{' ...' if len(added_acc) > 15 else ''}")
    if added_cp:
        print(f"  {len(added_cp)} new counterparties added - add country / VAT no. / IC flag: {', '.join(added_cp[:15])}{' ...' if len(added_cp) > 15 else ''}")
    dates = [r["date"] for r in rows if r["date"]]
    if dates:
        print(f"  GL dates {min(dates):%d/%m/%Y} - {max(dates):%d/%m/%Y}")


def ensure_fx_currencies(wb, ccys, source_note=None):
    """Make FX_Rates / FX_Daily list exactly the currencies given (keeping manual overrides)."""
    wf, wd = wb[FX], wb[FXD]
    overrides = {}
    for r in range(5, wf.max_row + 1):
        c = wf[f"A{r}"].value
        if c:
            overrides[str(c)] = (wf[f"C{r}"].value, wf[f"H{r}"].value)
    old = [wf[f"A{r}"].value for r in range(5, wf.max_row + 1) if wf[f"A{r}"].value]
    ccys = list(dict.fromkeys(list(ccys) or old))
    if wf.max_row >= 5:
        wf.delete_rows(5, wf.max_row - 4)
    for i, c in enumerate(ccys):
        ov, note = overrides.get(c, (None, None))
        write_fx_row(wf, 5 + i, c, i, source_note or note or "Run the fx command to download the ECB rates")
        if ov not in (None, ""):
            wf[f"C{5 + i}"].value = ov
    # rebuild FX_Daily header for the new currency list (data rewritten by fx)
    existing = {}
    hdr_cols = {wd.cell(row=4, column=j).value: j for j in range(2, wd.max_column + 1)}
    for r in range(5, wd.max_row + 1):
        d = wd[f"A{r}"].value
        if d is None:
            continue
        existing[d] = {c: wd.cell(row=r, column=j).value for c, j in hdr_cols.items() if c}
    if wd.max_row >= 4:
        wd.delete_rows(4, wd.max_row - 3)
    hdr(wd, 4, ["Date"] + ccys, [12] + [10] * len(ccys))
    for i, (d, vals) in enumerate(sorted(existing.items())):
        put(wd, f"A{5 + i}", d, F_INPUT, fmt=DATE)
        for k, c in enumerate(ccys):
            put(wd, f"{get_column_letter(2 + k)}{5 + i}", vals.get(c), F_INPUT, fmt=RATE)
    return ccys


def fetch_ecb(ccys, start: dt.date, end: dt.date):
    """Daily ECB euro reference rates: {date: {ccy: rate}} (units of currency per 1 EUR)."""
    url = (f"https://data-api.ecb.europa.eu/service/data/EXR/D.{'+'.join(ccys)}.EUR.SP00.A"
           f"?startPeriod={start:%Y-%m-%d}&endPeriod={end:%Y-%m-%d}&format=csvdata")
    req = urllib.request.Request(url, headers={"Accept": "text/csv", "User-Agent": "cy-vat-working/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            text = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return {}, url
        raise SystemExit(f"ECB download failed ({e}). Use --csv with a file downloaded from the CBC / ECB site.")
    except urllib.error.URLError as e:
        raise SystemExit(f"ECB download failed ({e.reason}). Use --csv with a file downloaded from the CBC / ECB site.")
    return parse_ecb_csv(text), url


def parse_ecb_csv(text: str):
    out: dict[dt.date, dict[str, float]] = {}
    for row in csv.DictReader(io.StringIO(text)):
        v = row.get("OBS_VALUE")
        if not v:
            continue
        d = dt.date.fromisoformat(row["TIME_PERIOD"])
        out.setdefault(d, {})[row["CURRENCY"]] = float(v)
    return out


def parse_wide_csv(path: Path):
    """CSV with a date column first and one column per currency (as downloaded from CBC / ECB)."""
    import pandas as pd
    df = pd.read_csv(path)
    if {"TIME_PERIOD", "CURRENCY", "OBS_VALUE"} <= set(df.columns):
        return parse_ecb_csv(path.read_text(encoding="utf-8"))
    dcol = df.columns[0]
    out = {}
    for rec in df.to_dict("records"):
        d = to_date(rec[dcol])
        vals = {}
        for c in df.columns[1:]:
            code = str(c).strip().upper()[:3]
            x = to_number(rec[c])
            if x:
                vals[code] = x
        if d and vals:
            out[d] = vals
    return out


def cmd_fx(args):
    start, end = quarter_dates(args.year, args.quarter)
    wb = load_workbook(args.workbook)
    if args.currencies:
        ccys = [c.strip().upper() for c in args.currencies.split(",") if c.strip()]
    else:
        ws = wb[GL]
        ccys = sorted({str(ws[f"H{r}"].value).strip().upper() for r in range(GL_FIRST, ws.max_row + 1)
                       if ws[f"H{r}"].value and str(ws[f"H{r}"].value).strip().upper() != "EUR"})
    if not ccys:
        print("No foreign currencies in GL_Input - nothing to do.")
        return
    if args.csv:
        data = parse_wide_csv(Path(args.csv))
        source = f"File {Path(args.csv).name} loaded {dt.date.today():%d/%m/%Y}"
    else:
        data, url = fetch_ecb(ccys, start, end)
        source = f"ECB euro reference rates (data-api.ecb.europa.eu), retrieved {dt.date.today():%d/%m/%Y}"
    data = {d: v for d, v in data.items() if start <= d <= end}
    ccys = ensure_fx_currencies(wb, ccys, source)
    wd = wb[FXD]
    if wd.max_row >= 5:
        wd.delete_rows(5, wd.max_row - 4)
    for i, d in enumerate(sorted(data)):
        put(wd, f"A{5 + i}", d, F_INPUT, fmt=DATE)
        for k, c in enumerate(ccys):
            put(wd, f"{get_column_letter(2 + k)}{5 + i}", data[d].get(c), F_INPUT, fmt=RATE)
    wb.calculation.fullCalcOnLoad = True
    wb.save(args.workbook)
    print(f"{len(data)} business days of rates written for Q{args.quarter} {args.year}:")
    for c in ccys:
        vals = [v[c] for v in data.values() if c in v]
        if vals:
            print(f"  {c}: average {statistics.mean(vals):.4f} per EUR ({len(vals)} days)")
        else:
            print(f"  {c}: NOT PUBLISHED by the ECB for this period - enter a manual override in FX_Rates column C")


def cmd_roll(args):
    start, end = quarter_dates(args.year, args.quarter)
    wb = load_workbook(args.workbook)
    ws = wb[SET]
    ws[S_START].value, ws[S_END].value = start, end
    ws[S_QLABEL].value = f"Q{args.quarter} {args.year}"
    wr = wb["VAT_Return"]
    for r in range(1, wr.max_row + 1):          # clear last quarter's sign-off
        if wr[f"B{r}"].value in SIGN_OFF:
            wr[f"C{r}"].value = None
    wb.calculation.fullCalcOnLoad = True
    wb.save(args.workbook)
    print(f"{args.workbook} set to Q{args.quarter} {args.year} ({start:%d/%m/%Y} - {end:%d/%m/%Y}). "
          f"Next: import-gl, then fx.")


def main(argv=None):
    p = argparse.ArgumentParser(description="Cyprus VAT working (VAT 4) generator", formatter_class=argparse.RawDescriptionHelpFormatter,
                                epilog=__doc__.split("Examples")[1] if "Examples" in __doc__ else "")
    sub = p.add_subparsers(dest="cmd", required=True)
    today = dt.date.today()
    prev_q = (today.month - 1) // 3 or 4
    prev_y = today.year if (today.month - 1) // 3 else today.year - 1

    a = sub.add_parser("new", help="create a new working")
    a.add_argument("workbook")
    a.add_argument("--year", type=int, default=prev_y)
    a.add_argument("--quarter", type=int, default=prev_q)
    a.add_argument("--force", action="store_true")
    a.set_defaults(func=cmd_new)

    b = sub.add_parser("import-gl", help="import a GL export into the working")
    b.add_argument("workbook")
    b.add_argument("gl")
    b.add_argument("--sheet", help="sheet name in the GL workbook (default: first)")
    b.set_defaults(func=cmd_import)

    c = sub.add_parser("fx", help="download ECB rates and compute the quarterly average")
    c.add_argument("workbook")
    c.add_argument("--year", type=int, default=prev_y)
    c.add_argument("--quarter", type=int, default=prev_q)
    c.add_argument("--currencies", help="comma list, default = currencies found in GL_Input")
    c.add_argument("--csv", help="load rates from a CSV file instead of the ECB API")
    c.set_defaults(func=cmd_fx)

    d = sub.add_parser("roll", help="move the working to a new quarter")
    d.add_argument("workbook")
    d.add_argument("--year", type=int, required=True)
    d.add_argument("--quarter", type=int, required=True)
    d.set_defaults(func=cmd_roll)

    args = p.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    sys.exit(main())

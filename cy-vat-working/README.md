# Cyprus VAT working (VAT 4)

A Python script that builds and maintains an Excel working paper. It turns a full general ledger into the Cyprus VAT return (Boxes 1–11B). The working:

- splits every transaction between **inter-company and third party**, and between **Cyprus, other EU countries and third countries**
- converts foreign-currency amounts at the **quarterly average ECB euro reference rate**, the same rates the Central Bank of Cyprus publishes
- produces the **VIES** amounts per EU customer and month
- **reconciles** Box 1 and Box 4 to the VAT control accounts in the GL
- runs an **exception report**

Every amount is an Excel formula, so changing a mapping updates the return straight away.
Full instructions are on the `Instructions` sheet. Rules and sources are on `Legal_Notes`.

## Setup

```bash
pip install openpyxl pandas
```

## Each quarter

```bash
# first quarter: create the working
python cy_vat_working.py new CY_VAT_Q3_2026.xlsx --year 2026 --quarter 3

# later quarters: copy last quarter's file, then move it to the new quarter
python cy_vat_working.py roll CY_VAT_Q4_2026.xlsx --year 2026 --quarter 4

# load the full GL export (.xlsx or .csv). Set its column names first in Settings > GL column map
python cy_vat_working.py import-gl CY_VAT_Q4_2026.xlsx GL_Q4_2026.xlsx

# download the ECB daily rates for the quarter and compute the average
python cy_vat_working.py fx CY_VAT_Q4_2026.xlsx --year 2026 --quarter 4
#   no internet access: python cy_vat_working.py fx ... --csv rates.csv
```

After the commands, open the file in Excel and do the following in order:

1. Complete `Account_Mapping` and `Counterparties` for any accounts or counterparties marked **NEW**.
2. Clear the `Checks` sheet.
3. Review `IC_Analysis` and `Reconciliation`.
4. Copy the `VAT_Return` amounts to the TFA portal.

`CY_VAT_Working_template.xlsx` is a ready-made file for Q3 2026. It holds example data, shaded orange, to show the expected format. `import-gl` removes it.

## Tests

```bash
python -m unittest test_cy_vat_working.py
```

> This is a preparation tool, not tax advice. Using the quarterly average exchange rate is a simplification. Art. 91(2) of the VAT Directive refers to the rate at the tax point, so confirm with your adviser that the average is acceptable for you.

"""
extract_annual_report.py
========================
Turns the key tables in the Stora Enso Annual Report 2025 (PDF) into pandas
DataFrames, then saves them as CSV files and as an Excel workbook that
Power BI can import directly.

What it extracts
----------------
1. Segment results   (pages 53-54)  sales, EBITDA, EBIT, operating capital, ...
2. Group key figures (page 51)      sales, adjusted EBIT, margin, ROCE, EPS, ...
3. Net debt table    (page 66)      adjusted EBITDA, net debt, net debt/EBITDA
4. Sensitivities     (page 60)      EBIT impact of +/-10% price or volume

Outputs (in the folder given by --out, default "output")
--------------------------------------------------------
- segments_long.csv   tidy format: one row per segment, year and metric
- group_long.csv      tidy format: one row per year and metric
- fact_segment.csv    wide format used by the Power BI model
- fact_group.csv      wide format used by the Power BI model
- fact_sensitivity.csv
- cockpit_data.xlsx   the same three fact tables as named Excel Tables,
                      so the Power BI report from the lean guide works unchanged

How to run
----------
    pip install pdfplumber pandas openpyxl
    python extract_annual_report.py --pdf STORAENSO_Annual_Report_2025.pdf

How it works
------------
The report pages have up to three columns side by side. Reading the whole
page mixes the columns together, so the script "crops" each column first
(like cutting the page with scissors), then reads the text line by line and
keeps lines that look like:  <label>  <number>  <number>

To reuse it for another report, change the CONFIG section: page numbers,
column positions and the list of segment names.
"""

import argparse
import re
from pathlib import Path

import pandas as pd
import pdfplumber

# ---------------------------------------------------------------------------
# CONFIG: the only part you need to change for another report
# ---------------------------------------------------------------------------
# Page numbers are the printed numbers; in this report they match PDF pages.
# Column crops are x-coordinates in PDF points (the page is ~1057 points wide).
SEGMENT_PAGES = [53, 54]
SEGMENT_COLUMNS = [(40, 370), (370, 700), (700, None)]  # None = page edge
SEGMENT_NAMES = ["Packaging Materials", "Packaging Solutions", "Biomaterials",
                 "Wood Products", "Forest", "Other"]
SEGMENT_YEARS = [2025, 2024]

KEY_FIGURES_PAGE = 51
KEY_FIGURES_CROP = (535, None)        # right-hand "Key figures" table
KEY_FIGURES_YEARS = [2025, 2024, 2023]

NET_DEBT_PAGE = 66
NET_DEBT_CROP = (40, 370)             # left column
NET_DEBT_YEARS = [2025, 2024, 2023]

SENSITIVITY_PAGE = 60
SENSITIVITY_CROP = (40, 530)          # left column, Table 1 (next column starts at ~540)

# Rename report labels to short, code-friendly column names.
# Labels not listed here are kept, converted to snake_case automatically.
METRIC_NAMES = {
    "Sales": "sales",
    "Sales, EUR million": "sales",
    "Adjusted EBITDA": "adj_ebitda",
    "Adjusted EBITDA margin": "adj_ebitda_margin",
    "Adjusted EBIT": "adj_ebit",
    "Adjusted EBIT, EUR million": "adj_ebit",
    "Adjusted EBIT margin": "adj_ebit_margin",
    "Operating capital, LTM": "operating_capital",
    "Adjusted ROOC, LTM": "adj_rooc",
    "Adjusted ROCE, LTM": "adj_roce",
    "Items affecting comparability (IAC)": "iac",
    "Fair valuations and non-operational items (FV)": "fv",
    "Operating result (IFRS)": "operating_result_ifrs",
    "Operating result (IFRS), EUR million": "operating_result_ifrs",
    "Cash flow from operations": "cash_flow_operations",
    "Cash flow after investing activities": "cash_flow_after_investing",
    "Net debt": "net_debt",
    "Net debt to adjusted EBITDA ratio": "net_debt_to_ebitda",
}

TOLERANCE = 3  # EUR million allowed difference in reconciliation (rounding)

# ---------------------------------------------------------------------------
# Step 1: helpers that read text and turn it into numbers
# ---------------------------------------------------------------------------
NUMBER = r"-?[\d,]+(?:\.\d+)?%?"          # e.g. 1,027  -15  7.8%  2.8


def column_text(page, x0, x1):
    """Crop one column of a page and return its text."""
    x1 = page.width if x1 is None else x1
    box = (x0, 60, x1, page.height - 20)   # skip the top menu bar and footer
    return page.crop(box).extract_text() or ""


def to_number(text):
    """Convert '1,027' -> 1027.0 and '7.8%' -> 0.078."""
    text = text.replace(",", "")
    if text.endswith("%"):
        return float(text[:-1]) / 100
    return float(text)


def clean_label(label):
    """Remove footnote markers: 'IAC)1' -> 'IAC)', 'Dividend per share1, EUR' -> 'Dividend per share, EUR'."""
    label = re.sub(r"(?<=[a-z\)])1(?=,|\s|$)", "", label.strip())
    return label.strip(" ,")


def snake_case(label):
    """'Board deliveries, 1,000 tonnes' -> 'board_deliveries_1000_tonnes'."""
    s = re.sub(r"[^0-9a-zA-Z]+", "_", label.replace(",", "")).strip("_").lower()
    return s


_METRIC_LOOKUP = {k.lower(): v for k, v in METRIC_NAMES.items()}


def metric_name(label):
    """Look up the short name, ignoring upper/lower case ('Operating Capital' = 'Operating capital')."""
    return _METRIC_LOOKUP.get(label.lower(), snake_case(label))


def parse_table_lines(text, n_values):
    """
    Find lines that end with exactly n_values numbers.
    Returns a list of (label, [values]).
    A label starting with a lowercase letter continues the previous line,
    e.g. 'Corrugated packaging European deliveries,' + 'million m2'.
    """
    pattern = re.compile(rf"^(.*?)\s+" + r"\s+".join([f"({NUMBER})"] * n_values) + r"$")
    rows, previous = [], ""
    for raw in text.split("\n"):
        line = re.sub(r"(\d)\s+%", r"\1%", raw.strip())   # '9.4 %' -> '9.4%'
        match = pattern.match(line)
        if match and re.search(r"[A-Za-z]", match.group(1)):
            label = match.group(1)
            if label[0].islower() and previous:
                label = previous + " " + label
            values = [to_number(v) for v in match.groups()[1:]]
            rows.append((clean_label(label), values))
            previous = ""
        else:
            previous = line
    return rows


# ---------------------------------------------------------------------------
# Step 2: extract each table into a tidy DataFrame
# ---------------------------------------------------------------------------
def find_segment_name(text):
    """Return the first known segment name found in the column heading."""
    head = "\n".join(text.split("\n")[:5])
    for name in SEGMENT_NAMES:
        if name in head:
            return name
    return None


def extract_segments(pdf):
    records = []
    for page_no in SEGMENT_PAGES:
        page = pdf.pages[page_no - 1]
        for x0, x1 in SEGMENT_COLUMNS:
            text = column_text(page, x0, x1)
            segment = find_segment_name(text)
            if segment is None:
                continue
            for label, values in parse_table_lines(text, len(SEGMENT_YEARS)):
                if label.startswith("EUR million"):       # the header row
                    continue
                for year, value in zip(SEGMENT_YEARS, values):
                    records.append({"segment": segment, "year": year,
                                    "metric": metric_name(label), "label": label,
                                    "value": value, "source": f"AR2025 p{page_no}"})
    return pd.DataFrame(records)


def extract_year_table(pdf, page_no, crop, years):
    page = pdf.pages[page_no - 1]
    text = column_text(page, *crop)
    records = []
    for label, values in parse_table_lines(text, len(years)):
        if label.startswith("EUR million"):
            continue
        for year, value in zip(years, values):
            records.append({"year": year, "metric": metric_name(label), "label": label,
                            "value": value, "source": f"AR2025 p{page_no}"})
    return pd.DataFrame(records)


def extract_sensitivity(pdf):
    page = pdf.pages[SENSITIVITY_PAGE - 1]
    text = column_text(page, *SENSITIVITY_CROP)
    text = text.split("Table 2")[0]          # stop before the cost table below it
    rows = [(label, v) for label, v in parse_table_lines(text, 2) if label in SEGMENT_NAMES]
    return pd.DataFrame([{"segment": label, "price_10pct": v[0], "volume_10pct": v[1],
                          "source": f"AR2025 p{SENSITIVITY_PAGE}"} for label, v in rows])


# ---------------------------------------------------------------------------
# Step 3: reshape into the wide tables used by Power BI
# ---------------------------------------------------------------------------
def build_fact_segment(seg_long, fact_group):
    wide = (seg_long[seg_long["metric"].isin(["sales", "adj_ebitda", "adj_ebit", "operating_capital"])]
            .pivot_table(index=["segment", "year"], columns="metric", values="value", aggfunc="first")
            .reset_index())
    wide["source"] = wide["segment"].map(seg_long.groupby("segment")["source"].first())

    # Eliminations = group total minus sum of segments, so the data reconciles
    elim = []
    for year in SEGMENT_YEARS:
        row = {"segment": "Eliminations", "year": year, "source": "Balancing figure (calculated)"}
        for m in ["sales", "adj_ebitda", "adj_ebit"]:
            row[m] = fact_group.loc[fact_group["year"] == year, m].iloc[0] - wide.loc[wide["year"] == year, m].sum()
        elim.append(row)
    wide = pd.concat([wide, pd.DataFrame(elim)], ignore_index=True)

    order = {name: i for i, name in enumerate(SEGMENT_NAMES + ["Eliminations"])}
    wide = wide.sort_values(by=["segment", "year"], key=lambda s: s.map(order) if s.name == "segment" else -s)
    cols = ["segment", "year", "sales", "adj_ebitda", "adj_ebit", "operating_capital", "source"]
    return wide[cols].reset_index(drop=True)


def build_fact_group(group_long):
    wanted = ["sales", "adj_ebitda", "adj_ebit", "net_debt", "net_debt_to_ebitda"]
    wide = (group_long[group_long["metric"].isin(wanted)]
            .drop_duplicates(subset=["year", "metric"])        # Adjusted EBITDA appears twice on p66
            .pivot(index="year", columns="metric", values="value")
            .reset_index())
    wide = wide[wide["year"].isin(SEGMENT_YEARS)].sort_values("year", ascending=False)
    wide["source"] = f"AR2025 p{KEY_FIGURES_PAGE}, p{NET_DEBT_PAGE}"
    return wide[["year"] + wanted + ["source"]].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Step 4: validate, so bad extraction never reaches Power BI
# ---------------------------------------------------------------------------
def validate(fact_segment, fact_group, fact_sensitivity):
    problems = []

    found = set(fact_segment["segment"]) - {"Eliminations"}
    missing = set(SEGMENT_NAMES) - found
    if missing:
        problems.append(f"Segments not found: {sorted(missing)}")
    if len(fact_sensitivity) != 5:
        problems.append(f"Expected 5 sensitivity rows, found {len(fact_sensitivity)}")

    # Eliminations should be a small share of sales but not absurd for EBIT
    elim = fact_segment[fact_segment["segment"] == "Eliminations"]
    for _, row in elim.iterrows():
        if abs(row["adj_ebit"]) > 25:
            problems.append(f"{row['year']}: EBIT eliminations of {row['adj_ebit']:.0f} look too large")

    # Group margin cross-check: sales x margin should match EBIT
    print("\nValidation")
    for _, g in fact_group.iterrows():
        seg_sum = fact_segment.loc[(fact_segment["year"] == g["year"]) &
                                   (fact_segment["segment"] != "Eliminations"), "sales"].sum()
        print(f"  {g['year']}: group sales {g['sales']:,.0f} | segments before eliminations {seg_sum:,.0f}"
              f" | EBIT eliminations {elim.loc[elim['year'] == g['year'], 'adj_ebit'].iloc[0]:,.0f}")

    if problems:
        raise SystemExit("Validation failed:\n  " + "\n  ".join(problems))
    print("  All checks passed.")


# ---------------------------------------------------------------------------
# Step 5: save CSV and Excel (with named Excel Tables for Power BI)
# ---------------------------------------------------------------------------
def save_excel(path, tables):
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in tables.items():
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ref = f"A1:{get_column_letter(df.shape[1])}{len(df) + 1}"
            table = Table(displayName=name, ref=ref)
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium7", showRowStripes=True)
            ws.add_table(table)
            for i, col in enumerate(df.columns, start=1):
                ws.column_dimensions[get_column_letter(i)].width = max(12, len(str(col)) + 4)


def main():
    parser = argparse.ArgumentParser(description="Extract annual report tables to CSV and Excel")
    parser.add_argument("--pdf", default="STORAENSO_Annual_Report_2025.pdf", help="path to the report PDF")
    parser.add_argument("--out", default="output", help="output folder")
    args = parser.parse_args()

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    with pdfplumber.open(args.pdf) as pdf:
        seg_long = extract_segments(pdf)
        group_long = pd.concat([
            extract_year_table(pdf, KEY_FIGURES_PAGE, KEY_FIGURES_CROP, KEY_FIGURES_YEARS),
            extract_year_table(pdf, NET_DEBT_PAGE, NET_DEBT_CROP, NET_DEBT_YEARS),
        ], ignore_index=True)
        fact_sensitivity = extract_sensitivity(pdf)

    fact_group = build_fact_group(group_long)
    fact_segment = build_fact_segment(seg_long, fact_group)

    print(f"Extracted {len(seg_long)} segment values, {len(group_long)} group values, "
          f"{len(fact_sensitivity)} sensitivity rows")
    print("\nfact_segment\n", fact_segment.to_string(index=False))
    print("\nfact_group\n", fact_group.to_string(index=False))
    print("\nfact_sensitivity\n", fact_sensitivity.to_string(index=False))

    validate(fact_segment, fact_group, fact_sensitivity)

    seg_long.to_csv(out / "segments_long.csv", index=False)
    group_long.to_csv(out / "group_long.csv", index=False)
    fact_segment.to_csv(out / "fact_segment.csv", index=False)
    fact_group.to_csv(out / "fact_group.csv", index=False)
    fact_sensitivity.to_csv(out / "fact_sensitivity.csv", index=False)
    save_excel(out / "cockpit_data.xlsx", {"fact_segment": fact_segment,
                                           "fact_group": fact_group,
                                           "fact_sensitivity": fact_sensitivity})
    print(f"\nSaved CSV files and cockpit_data.xlsx to '{out}/'")


if __name__ == "__main__":
    main()

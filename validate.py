"""
validate.py - check that segment figures reconcile to group totals.

Why: controllers need one reliable source of truth. Before any number goes
into Power BI, we prove that segments + eliminations = reported group total.

Run from the folder that contains cockpit_data.xlsx:
    python validate.py
"""
import pandas as pd

FILE = "cockpit_data.xlsx"
TOLERANCE = 3  # EUR million; published figures are rounded

# 1. Load the two data sheets into DataFrames (tables of rows and columns)
seg = pd.read_excel(FILE, sheet_name="fact_segment", nrows=14)
grp = pd.read_excel(FILE, sheet_name="fact_group", nrows=2).set_index("year")

# 2. For each metric and year, compare group total with the sum of segments
errors = []
for metric in ["sales", "adj_ebitda", "adj_ebit"]:
    seg_total = seg.groupby("year")[metric].sum()      # add up all segments per year
    for year in grp.index:
        reported = grp.loc[year, metric]
        summed = seg_total[year]
        diff = reported - summed
        status = "OK" if abs(diff) <= TOLERANCE else "FAIL"
        print(f"{metric:<11} {year}: group {reported:>7,.0f} | "
              f"segments {summed:>7,.0f} | diff {diff:>3,.0f}  {status}")
        if status == "FAIL":
            errors.append((metric, year, diff))

# 3. Stop with a clear message if anything does not reconcile
if errors:
    raise SystemExit(f"\nValidation failed: {errors}")
print("\nAll checks passed. Data is ready for Power BI.")

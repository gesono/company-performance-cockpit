# Company Performance Cockpit

A Python pipeline that turns annual report PDFs into validated, analysis-ready
tables for pandas and Power BI.

Built and demonstrated on the [Stora Enso Annual Report 2025](https://www.storaenso.com/en/investors).
Independent portfolio project, not affiliated with Stora Enso.

## What it does

```
PDF ──► crop columns ──► parse lines ──► tidy tables ──► validate ──► CSV + Excel
```

From a 225-page PDF it extracts 156 segment values, 123 group values and 5
sensitivity rows, then checks that segment figures plus eliminations reconcile
to the published group totals before writing any output.

```
Validation
  2025: group sales 9,326 | segments before eliminations 12,186 | EBIT eliminations 2
  2024: group sales 9,049 | segments before eliminations 11,601 | EBIT eliminations -11
  All checks passed.
```

Nothing is saved unless the numbers tie.

## Run it

```bash
pip install -r requirements.txt
python extract_annual_report.py --pdf STORAENSO_Annual_Report_2025.pdf
```

The PDF is not included here (it is large and copyrighted). Download it from
the link above and place it in the project folder.

## Output

| File | Format | Use |
| --- | --- | --- |
| `segments_long.csv` | long | analysis in pandas |
| `group_long.csv` | long | analysis in pandas |
| `fact_segment.csv` | wide | BI model |
| `fact_group.csv` | wide | BI model |
| `fact_sensitivity.csv` | wide | scenario analysis |
| `cockpit_data.xlsx` | named Excel tables | direct Power BI import |

Every row carries a `source` column with the report page it came from.

## Reusing it for another report

All layout-specific settings live in one CONFIG block at the top of the script:
page numbers, column positions and segment names. Adapting it to a different
report means editing that block, not the parsing logic.

See [docs/how-it-works.md](docs/how-it-works.md) for the method, including how
to find column positions in any PDF.

## Learn from it

- [docs/how-it-works.md](docs/how-it-works.md) — how each part of the script works
- [explore_extraction.ipynb](explore_extraction.ipynb) — runnable walkthrough of each stage
- Practice exercises are at the end of both

## Limitations

- Works on text-based PDFs. Scanned documents need OCR first.
- Documents with ruled table borders may parse better with `page.extract_tables()`
  or `camelot`.
- Page numbers and column positions are layout-specific, which is why the
  validation gate exists.

## Built with

Python, pdfplumber, pandas, openpyxl

## Author

Gerald Esono 

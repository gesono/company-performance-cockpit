# How the data extraction works

This section explains `extract_annual_report.py`: the script that turns the
Stora Enso Annual Report 2025 PDF into tidy tables for Python and Power BI.

## The pipeline

```
PDF page ──► crop into columns ──► read text lines ──► keep "label + numbers" lines
        ──► long table (one row per value) ──► wide tables ──► validate ──► CSV + Excel
```

Each function in the script does one of these steps.

## Run it

```bash
pip install pdfplumber pandas openpyxl
python extract_annual_report.py --pdf STORAENSO_Annual_Report_2025.pdf
```

Expected output:

```
Extracted 156 segment values, 123 group values, 5 sensitivity rows

Validation
  2025: group sales 9,326 | segments before eliminations 12,186 | EBIT eliminations 2
  2024: group sales 9,049 | segments before eliminations 11,601 | EBIT eliminations -11
  All checks passed.

Saved CSV files and cockpit_data.xlsx to 'output/'
```

| File in `output/` | Format | Used for |
| --- | --- | --- |
| `segments_long.csv` | long | analysis in pandas |
| `group_long.csv` | long | analysis in pandas |
| `fact_segment.csv` | wide | Power BI model |
| `fact_group.csv` | wide | Power BI model |
| `fact_sensitivity.csv` | wide | Power BI what-if page |
| `cockpit_data.xlsx` | 3 named Excel Tables | direct import into Power BI |

## Section A — CONFIG

The only part to change when the report changes.

```python
SEGMENT_PAGES = [53, 54]
SEGMENT_COLUMNS = [(40, 370), (370, 700), (700, None)]
```

**Why crop columns?** Page 53 holds three segment tables side by side. Reading
the whole page reads straight across the page, mixing them:

```
Sales  4,478  4,502      Sales  1,027  987      Sales  1,458  1,587
```

Cropping each column first gives one clean table at a time.

The numbers are horizontal positions in PDF points (this page is ~1,057 points
wide). `None` means "to the right edge". I found them by listing where the word
"Sales" starts on the page (x = 45, 375 and 705) and cutting just to the left of
each. Notebook cell 3 shows how to do this for any report.

`METRIC_NAMES` maps report labels to short column names (`Adjusted EBITDA` →
`adj_ebitda`). Anything not listed is converted automatically, so
`Board deliveries, 1,000 tonnes` becomes `board_deliveries_1000_tonnes`.

## Section B — Helpers

**`column_text(page, x0, x1)`** crops a rectangle and returns its text. The box
starts 60 points down to skip the navigation bar and stops 20 points above the
bottom to skip the "Unaudited 53" footer.

**`to_number(text)`** converts report text into numbers:

| Input | Output |
| --- | --- |
| `"1,027"` | `1027.0` |
| `"-15"` | `-15.0` |
| `"7.8%"` | `0.078` |

Percentages become decimals, which is the format Power BI expects.

**`clean_label(label)`** removes footnote markers: `Items affecting
comparability (IAC)1` → `Items affecting comparability (IAC)`. It only strips a
`1` that follows a letter or bracket, so units such as `m2` survive.

**`parse_table_lines(text, n_values)`** is the core. It keeps lines made of a
label followed by exactly `n_values` numbers (2 for segment tables: 2025 and
2024; 3 for the key figures table).

```python
NUMBER = r"-?[\d,]+(?:\.\d+)?%?"
```

| Piece | Meaning |
| --- | --- |
| `-?` | optional minus sign |
| `[\d,]+` | digits and thousands commas, e.g. `1,027` |
| `(?:\.\d+)?` | optional decimals, e.g. `.8` |
| `%?` | optional percent sign |

Two cleaning tricks live here:

1. **Percent spacing.** Some lines extract as `9.4 %`.
   `re.sub(r"(\d)\s+%", r"\1%", line)` joins them back together.
2. **Labels split across two lines.**

   ```
   Corrugated packaging European deliveries,      <- no numbers, stored as "previous"
   million m2  1,228  1,217                       <- label starts lowercase
   ```

   A matched label starting with a lowercase letter is treated as a
   continuation and joined to the previous line.

## Section C — Extraction

`extract_segments()` loops over 2 pages x 3 columns = 6 tables.
`find_segment_name()` checks the first five lines of each column against the
known segment names, which is how "Segment Other" is recognised as "Other".
The header row (`EUR million 2025 2024`) matches the number pattern too, so it
is skipped explicitly.

Every value becomes one row — long format:

| segment | year | metric | value | source |
| --- | --- | --- | --- | --- |
| Packaging Solutions | 2025 | sales | 1027 | AR2025 p53 |
| Packaging Solutions | 2024 | sales | 987 | AR2025 p53 |

`extract_year_table()` does the same for the group tables (key figures p. 51,
net debt p. 66), with years in place of segments.

`extract_sensitivity()` contains one documented fix:

```python
text = text.split("Table 2")[0]   # stop before the cost table below it
```

The cost table sits directly under the sensitivity table, and its line
`Other 10% 10%` also matched the pattern. PDF extraction is usually 80% generic
logic and 20% small fixes like this — worth documenting rather than hiding.

## Section D — Reshape

`build_fact_segment()` pivots long to wide (one column per metric), then adds an
**Eliminations** row:

```
eliminations = group total - sum of segments
2025 sales: 9,326 - 12,186 = -2,860
```

Segment sales include internal sales (for example Forest selling wood to the
mills). The eliminations row removes that double counting so the data
reconciles to the published group figure.

`build_fact_group()` calls `.drop_duplicates()` because "Adjusted EBITDA"
appears twice on page 66; without it the pivot fails.

## Section E — Validation

`validate()` stops the run if:

- any of the six segments is missing,
- there are not exactly five sensitivity rows (this is what caught the "Other"
  bug above),
- EBIT eliminations exceed EUR 25 million, which would suggest a misread number.

A pipeline that fails loudly beats one that quietly delivers wrong numbers.

## Section F — Save

`save_excel()` writes each DataFrame to its own sheet and wraps it in a **named
Excel Table**. Power BI's Navigator lists those names (`fact_segment`,
`fact_group`, `fact_sensitivity`), so the report model picks them up unchanged.

## Section G — Command line

```bash
python extract_annual_report.py --pdf other_report.pdf --out output_upm
```

`argparse` exposes the PDF path and output folder, so no code edits are needed
to run the script on a different report.

## Explore it yourself

Open `explore_extraction.ipynb` for a four-step walkthrough: raw column text,
parsed lines, finding column positions in any PDF, and analysing the results in
pandas.

```bash
pip install pdfplumber pandas openpyxl matplotlib jupyter
jupyter notebook explore_extraction.ipynb
```

## Practice exercises

| # | Exercise | Skill |
| --- | --- | --- |
| 1 | Add `cash_flow_operations` to `fact_segment` (edit the metric list in `build_fact_segment`) | pivoting |
| 2 | Crop the wrong column on purpose, e.g. `(300, 700)`, and see which check catches it | why validation matters |
| 3 | Extract the cost table (Table 2, p. 60) into a `fact_cost_mix` DataFrame | writing your own extractor |
| 4 | Run the script on a quarterly interim report and adjust the CONFIG | reusability |

Hint for exercise 3: the values are percentages, and you only need the text
after "Table 2".

## Limitations

- Page numbers and column positions are specific to this report layout; they
  live in the CONFIG block so adapting is quick.
- Annual reports are the hard case. Quarterly interim reports have simpler
  layouts and usually parse with fewer adjustments.
- Extracted figures are always checked against the published totals before use.

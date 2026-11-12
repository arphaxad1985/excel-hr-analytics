# HR Workforce Analytics — Excel Portfolio

An Excel workbook demonstrating advanced formula-based HR analytics —
headcount, salary analysis, hiring trends, and interactive filtering —
built with dynamic array formulas and no pivot tables.

## Sheets

| Sheet | Purpose | Key formulas |
|-------|---------|--------------|
| `Data` | Employee records | — |
| `Test 1` | Headcount by department | `UNIQUE()`, `COUNTIF()` |
| `Test 2` | Job title breakdown + salary | `UNIQUE()`, `XLOOKUP()`, `COUNTIF()`, `AVERAGEIF()` |
| `Test 3` | Employee register (filtered) | `FILTER()`, `SORT()` |
| `Test 4` | Yearly starters | `COUNTIFS()` with date criteria |
| `Test 5` | Interactive drop-down filter | `FILTER()`, data validation |

## Skills demonstrated

- **Dynamic array formulas**: `UNIQUE`, `FILTER`, `SORT`, `SORTBY`
- **Lookups**: `XLOOKUP`, `INDEX/MATCH`
- **Conditional aggregation**: `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`
- **Date-based analysis**: year-on-year starter counts
- **Interactive reporting**: dependent drop-downs with `FILTER()`
- **Data cleaning**: deduplication, null handling, date standardisation

## Data

All data is **synthetic** — no real employee or company data is included.

## How to use

1. Download `hr_analyst_final.xlsx`
2. Open in Excel (dynamic arrays require Excel 365 or 2021+)
3. Start on the `Data` tab, then explore sheets `Test 1`–`Test 5`
4. Change the drop-downs on `Test 5` to see live filtering

## Tools

- Microsoft Excel (dynamic arrays, data validation, conditional formatting)

## Author

[Arphaxad Nguka Owange](https://github.com/arphaxad1985)
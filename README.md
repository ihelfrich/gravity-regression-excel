# Linear Regression in Excel — Gravity Trade Workbook

A self-contained, hands-on lesson in linear regression that you run yourself in Excel, built
on **200 real country-pair trade flows** from BACI 2022 and the CEPII Gravity database. No
simulated data, no black boxes. Every statistic in the workbook is a live Excel formula, so
the whole thing recomputes if you change the data.

## What you get

| File | What it is |
|---|---|
| **`Gravity_Regression_Workbook.xlsx`** | The course. Seven tabs, from data to a graded answer key. |
| **`LESSON.md`** | The written lesson: the theory and a walkthrough of every tab. |
| **`data/gravity_sample_200.csv`** | The raw 200-row dataset. |
| **`data/DATA_DICTIONARY.md`** | What every column means and where it came from. |
| **`build_workbook.py`** | The script that builds the workbook. Re-run it to rebuild from scratch. |

## How to use it

1. Download the repository. On GitHub, click the green **Code** button, then **Download ZIP**.
   (Or, from a terminal: `git clone https://github.com/ihelfrich/gravity-regression-excel.git`.)
2. Open `Gravity_Regression_Workbook.xlsx` in Excel. Google Sheets and LibreOffice Calc also
   work; both recalculate the formulas on open.
3. Start on the **Start Here** tab, then go through the lesson tabs in order:
   - **L1 Simple Regression** — one predictor, a chart with a trendline, a prediction calculator
   - **L2 Multiple Regression** — the full gravity equation with `LINEST`
   - **L3 Dummy Variables** — yes/no predictors read as percentage effects
4. Do the **Exercises** tab before you open the **Answer Key**.
5. Read `LESSON.md` alongside the workbook for the narrative and the interpretation.

## What it teaches

- Fitting and reading a simple regression with `SLOPE`, `INTERCEPT`, `RSQ`
- Why one strong predictor can still explain almost nothing (R-squared)
- Multiple regression with `LINEST`, and the reverse-order trap that catches everyone
- Reading coefficients as elasticities (logged variables) and as percentage effects (dummies)
- t-statistics, p-values, adjusted R-squared, and the F-test
- Three routes to the same answer: trendline, functions, and the Analysis ToolPak

The three models, estimated live in the workbook and cross-checked against Python
(`statsmodels`) on the Answer Key tab:

| Model | Predictors | R-squared |
|---|---|---:|
| 1 | distance | 0.06 |
| 2 | distance + both GDPs | 0.35 |
| 3 | + border, language, trade agreement | 0.41 |

## Data sources

Trade values from BACI 2022 (CEPII, built on UN Comtrade). Distance, contiguity, common
language, colonial history, and trade-agreement membership from the CEPII Gravity database.
GDP from the World Bank's World Development Indicators. See `data/DATA_DICTIONARY.md` for the
full provenance.

## Requirements

Any spreadsheet program that reads `.xlsx`: Microsoft Excel (2016 or later recommended for
`LINEST`), Google Sheets, or LibreOffice Calc. To rebuild the workbook yourself you need
Python with `pandas`, `openpyxl`, and `statsmodels`.

---

*Prepared as a teaching resource by Dr. Ian Helfrich. Built with AI assistance and verified
against an independent `statsmodels` estimation.*

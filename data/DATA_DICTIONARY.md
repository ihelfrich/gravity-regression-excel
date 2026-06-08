# Data dictionary — `gravity_sample_200.csv`

200 directed country-pair observations. Each row is one exporter (`i`) selling to one
importer (`j`). The sample is drawn from a full bilateral trade panel and trimmed to 200
clean pairs so it stays workable in a spreadsheet while remaining real data.

| Column | Type | Unit | Description |
|---|---|---|---|
| `iso3_i` | text | — | Exporter ISO-3 code |
| `country_i` | text | — | Exporter name |
| `iso3_j` | text | — | Importer ISO-3 code |
| `country_j` | text | — | Importer name |
| `trade_musd` | number | thousands of USD | Value of goods exported from `i` to `j` |
| `gdp_i_musd` | number | millions of USD | Exporter GDP |
| `gdp_j_musd` | number | millions of USD | Importer GDP |
| `distance` | number | kilometres | Population-weighted distance between `i` and `j` |
| `ln_trade` | number | log | Natural log of `trade_musd` — the outcome (Y) |
| `ln_dist` | number | log | Natural log of `distance` |
| `ln_gdp_i` | number | log | Natural log of exporter GDP |
| `ln_gdp_j` | number | log | Natural log of importer GDP |
| `contiguity` | 0/1 | dummy | 1 if the two countries share a land border |
| `common_language` | 0/1 | dummy | 1 if they share an official language |
| `agree_fta` | 0/1 | dummy | 1 if a trade agreement is in force between them |
| `colony_ever` | 0/1 | dummy | 1 if one ever colonised the other |

## Why the logs

Trade, GDP, and distance span several orders of magnitude (from a few thousand dollars to
hundreds of billions). Taking natural logs compresses that range and, more usefully, turns
the slope coefficients into **elasticities**: a coefficient of `b` on a logged predictor
means a 1 percent rise in that predictor goes with a `b` percent change in trade. This is
the standard "log-log" gravity specification.

## Sources

- **Trade values:** BACI 2022, CEPII. Harmonised bilateral trade reconciled from UN Comtrade.
- **Distance, contiguity, common language, colonial history, trade-agreement membership:**
  CEPII Gravity database.
- **GDP:** World Bank, World Development Indicators.

Assembled by Dr. Ian Helfrich from the project trade-data holdings. The 200-row teaching
sample and all derived log columns are reproducible from `build_workbook.py`.

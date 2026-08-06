# Tennis Homophily — Project Rules for Claude

## Mandatory sync rule
Any change to the regression models MUST update all four files atomically:

1. **`code/analysis/homophily.ipynb`** — re-execute the relevant cells after editing
2. **`homophily.do`** — keep in sync with the notebook (same spec, same controls, same FE)
3. **`report.html`** — update every affected table with new AME values, SEs, p-values, N
4. **`report.pdf`** — regenerate from report.html using:
   ```
   msedge.exe --headless --print-to-pdf="C:/Users/aldi/AppData/Local/Temp/report_out.pdf" \
     "file:///c:/Users/aldi/Documents/GitHub/tennis-homophily/report.html"
   cp report_out.pdf report.pdf
   ```

Never update one file without updating the others. After any model change, run a three-way comparison (Stata / notebook / report) before declaring done.

## Current specification (as of 2026-06)

**Sample:** Grand Slams only (Olympics excluded), N = 3,744 team-obs, 1,872 matches (2018–2025 incl. 2020).  
`exp_mean` imputed to 0 for 216 first-time GS participants (no prior appearances in dataset window),
then demeaned against each estimation sample's own mean to form `exp_mean_dm` (see Controls below).

**Outcome variables:**
- Table 3: `win` (match win, binary)
- Table 4: `won_tb` (tiebreak win, binary) — unit = one team per tiebreak, N = 2,296

**Culture measures (enter one at a time):**
- `same_country` — same nationality (binary)
- `same_language` — same official language (binary)
- `ling_prox` — ethnolinguistic proximity 0–1 (continuous)

**Controls:** `rank_mean`, `opp_rank_mean`, `single_top100`, `exp_mean_dm`, `exp_mean_dm_sq`
(`exp_mean_dm = exp_mean − mean`; `exp_mean_dm_sq = exp_mean_dm²`; demeaning is a pure
reparameterization of the quadratic and leaves the culture AMEs, other controls' AMEs, and
fitted model unchanged — it only shifts what the linear "GS appearances" AME represents, from
the effect at zero appearances to the effect at mean appearances)

**Fixed effects:** `C(tournament):C(year)` + `C(stage_code)` (tournament×year + round)

**Standard errors:** clustered by `match_id`

**Estimator:** logit; report Average Marginal Effects (AME, dP/dx) — NOT logit coefficients

**Heterogeneity tables:**
- Table 6a: culture × `exp_mean`
- Table 6b: culture × surface (clay/grass; hardcourt baseline)
- Table 6c: culture × `ic_team_dm` (Hofstede IDV, demeaned; mean = 62.99, SD = 20.87)

## Key file locations
- Data (GS panel): `data/atp/team_gs_panel.csv`
- Data (tiebreak panel): `data/atp/tiebreak_panel.csv`
- Stata results: `stata_homophily_results.txt`
- Report: `report.html` → `report.pdf`

## Notebook execution
```bash
cd code/analysis
python -m nbconvert --to notebook --execute --inplace \
  --ExecutePreprocessor.timeout=600 \
  --ExecutePreprocessor.kernel_name=python3 homophily.ipynb
```

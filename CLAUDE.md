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

## Current specification (as of 2026-08)

**Report scope (updated 2026-08-19):** `report.html`/`report.pdf` (the Grand-Slams
regression report) remain **Part 1, Grand Slams only** — zero Olympic data mixed into
any count, table, or cross-reference there (the Observation Breakdown funnel starts
from a GS-only raw count, not a GS+Olympics combined one).

The overleaf paper (`overleaf/main.tex`) is broader than `report.html` and now
**does** include Olympic-cycle descriptive team-formation evidence (Tables 2A/2B,
Tokyo and Paris cycles — nationality/language/ling_prox composition around each
Games, no performance outcomes) in its main-text §6.3, per Lingqing's 2026-08-19
rewrite. This supersedes the earlier "Olympics deferred entirely" rule for the paper
specifically; `report.html` itself is unaffected and still excludes Olympics
entirely, since it is Part 1 in the narrower sense (the regression tables only).

**Sample:** Grand Slams only, N = 3,752 team-obs, 1,876 matches (2018–2025 incl. 2020).
Of the 14 matches originally dropped for incomplete doubles ranking, 4 were retrieved and restored
(Guillermo García-López's ranking merge fixed for match_id 6/35/49; Alejandro Davidovich Fokina's
full profile reconstructed for match_id 913) — 10 remain dropped as unretrievable.
`exp_mean` imputed to 1 (not 0) for 213 rookies (turned pro in/after the tournament year, so raw
tenure ≤0) — a nominal first-year tenure rather than zero, per reviewer feedback; no observations
are dropped. Demeaned against each estimation sample's own mean to form `exp_mean_dm` (see
Controls below).

**Tiebreak classification:** A standard tiebreak is a 7-pt breaker at 6-6 (any set, 1–5). An
**advantage-set decider with no breaker played** (e.g. 8–6, natural 2-game margin) is **not** a
tiebreak — confirmed directly from the raw `winners_setN_tiebreak`/`losers_setN_tiebreak` score
columns, which are null for all 15 such cases in this dataset. It is excluded from every tiebreak
count and from Table 4 entirely. A **12-12 breaker** (deciding set reaches 12-12, then a real
breaker decides it, e.g. 13–12) genuinely IS a tiebreak — a real breaker score is recorded (e.g.
loser scores 4, 6, or 2 points) — but is not currently included in Table 4 either, since it isn't
a standard 7-pt format (3 such cases exist). No genuine 10-point super-tiebreak exists anywhere in
this GS dataset, so Table 4 currently has no valid robustness spec — main spec (7pt, sets 1–5) only.

**Outcome variables:**
- Table 3: `win` (match win, binary)
- Table 4: `won_tb` (tiebreak win, binary) — unit = one team per tiebreak, N = 2,364 (7pt tiebreaks, sets 1–5; no robustness spec — see above)

**Culture measures (enter one at a time):**
- `same_country` — same nationality (binary)
- `same_language` — same official language (binary)
- `ling_prox` — ethnolinguistic proximity 0–1 (continuous)

**Controls:** `rank_mean`, `opp_rank_mean`, `single_top100`, `exp_mean_dm`, `exp_mean_dm_sq`
(`exp_mean` = tournament year − year turned pro, i.e. years of professional tenure, averaged
across the two teammates — NOT a count of prior Grand Slam appearances, despite the variable
name; imputed to 1 for rookies whose raw tenure is ≤0. `exp_mean_dm = exp_mean − mean`;
`exp_mean_dm_sq = exp_mean_dm²`; demeaning is a pure reparameterization of the quadratic and
leaves the culture AMEs, other controls' AMEs, and fitted model unchanged — it only shifts
what the linear "years since turning pro" AME represents, from the effect at zero tenure to
the effect at mean tenure)

**Fixed effects:** `C(tournament):C(year)` + `C(stage_code)` (tournament×year + round)

**Standard errors:** clustered by `match_id`

**Estimator:** logit; report Average Marginal Effects (AME, dP/dx) — NOT logit coefficients

**Heterogeneity tables:**
- Table 5: culture × surface (two specs: grass-vs-rest, clay-vs-rest; no tournament×year FE),
  plus a robustness spec restoring tournament×year FE with only Culture×grass/Culture×clay
  (no separate surface main effect) — the two specs agree (interactions insignificant throughout)
- Table 6: culture × `ic_team_dm` (Hofstede IDV, demeaned; mean = 62.98, SD = 20.85), spec 2 only,
  plus a Tier-2-proxy-excluded robustness cut (N = 3,597; drops 155 team-obs)
- Table 6a: culture × `exp_mean` (years since turning pro, demeaned; mean = 11.90, SD = 5.17),
  plus a robustness spec adding the quadratic interaction Culture × `exp_mean_dm_sq`
  (insignificant throughout — linear-interaction finding holds)
- Table 6b (new, per reviewer request 2026-08-11): culture × `exp_gap_dm` (within-team
  experience gap, `|exp_i − exp_j|` between the two teammates' own tenure, demeaned; mean =
  5.48, SD = 4.92), plus a quadratic-interaction robustness spec. Interaction is positive
  throughout but insignificant (p 0.24–0.34) — no evidence culture compensates for an
  experience mismatch within a team.

**Section 6 (new, per reviewer request 2026-08-11): partner-selection sorting check.**
Descriptive comparison of actual vs. random-matching-benchmark same-nationality/language/
ling_prox rates among realized doubles partnerships (deduped to one row per tournament×team,
N = 2,101), split by All / both-top-100 / both-top-50 (ranking at time of tournament).
Random benchmark is closed-form (not simulated): Σ over C(n,2) pairs in that tournament's
actual field, using the same CEPII `comlang_off`/`comlang_ethno` country-pair lookup that
`same_language`/`ling_prox` are themselves built from. Finding: partner selection is far more
culturally assortative than chance at every skill level (6–8× benchmark for nationality,
3–4× for language), but the degree of excess is essentially flat across brackets — elite
players are not disproportionately more assortative. A follow-up continuous test (§6.1,
logit/OLS of same_country/same_language/ling_prox on the ego player's own ranking, 4,202
ego-rows, clustered by partnership) confirms this: the coefficient on own rank is positive
and significant for all three outcomes, meaning *worse*-ranked players sort into
same-culture partnerships slightly *more*, not less — the reverse of the "stronger players
have more choice and sort more" concern.

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

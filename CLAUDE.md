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

**Sample:** Grand Slams only, N = 3,728 team-obs, 1,864 matches (2018–2025 incl. 2020).
Of the 14 matches originally dropped for incomplete doubles ranking, 4 were retrieved and restored
(Guillermo García-López's ranking merge fixed for match_id 6/35/49; Alejandro Davidovich Fokina's
full profile reconstructed for match_id 913) — 10 remain dropped as unretrievable.
`exp_mean` imputed to 1 (not 0) for rookies (turned pro in/after the tournament year, so raw
tenure ≤0) — a nominal first-year tenure rather than zero, per reviewer feedback; no observations
are dropped. Demeaned against each estimation sample's own mean to form `exp_mean_dm` (see
Controls below).

**Sample correction (found 2026-09-03): Wimbledon 2018 qualifying-round contamination.**
The raw scrape for Wimbledon 2018 alone included 12 qualifying-round doubles matches (8 "1st
Round Qualifying" + 4 "2nd Round Qualifying") mixed in with the 63 main-draw matches — no other
tournament-year in the dataset has any qualifying-round rows. These are now excluded at the very
start of the pipeline (`homophily.ipynb`, cell ~1, `_qualifying_mask`), which is what moved the
sample from the previously-documented 1,876 matches / 3,752 team-obs down to the current 1,864 /
3,728 (Wimbledon: 437 → 425 matches). This is a genuine data-quality fix, not an error in either
the notebook or the paper — `overleaf/main.tex` (Lingqing's 2026-09 revision) already reflects the
corrected sample throughout and was independently verified against a fresh end-to-end re-execution
of `homophily.ipynb` (e.g. Table 3 same-nationality AME = 0.0361, matching exactly).
**Resync status (updated 2026-09-08): complete.** `report.html`/`report.pdf` were re-audited
table-by-table against a fresh notebook execution — every table (Section 2 breakdown, Tables
3–6b, Section 6/6.1/6.2 partner-selection, age/double-counting appendix tables) already carried
the corrected 1,864/3,728 sample with no stale numbers found; `report.pdf`'s text was extracted
and confirmed to match. `homophily.do` had one substantive gap — it never dropped the Wimbledon
2018 qualifying-round matches, so its regression sample would have computed to the stale
1,876/3,752 rather than 1,864/3,728 — fixed by adding the same `strpos(stage, "Qualifying")`
drop immediately after loading the raw Excel, mirroring the notebook's `_qualifying_mask`
(cell 1). A handful of cosmetic stale numbers in `homophily.do` comments/display strings
(1,876/1,886/3,752/4,202/4,198) were also corrected to 1,864/1,874/3,728/4,178/4,175.
`data/atp/partner_selection_ego.csv` was checked directly (header + row count) and is current
(4,175 data rows, matching the notebook's fresh export exactly) — the earlier claim above that
it was stale was itself outdated by the time it was checked.

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
- Table 4: `won_tb` (tiebreak win, binary) — unit = one team per tiebreak, N = 2,354 (7pt tiebreaks, sets 1–5; no robustness spec — see above; N post-2026-09-03 qualifying-round fix, was 2,364)

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
- Table 6 (report.html only — not in `main.tex`, which uses "Table 6" for the experience
  heterogeneity table below instead): culture × `ic_team_dm` (Hofstede IDV, demeaned; mean =
  62.98, SD = 20.85), spec 2 only, plus a Tier-2-proxy-excluded robustness cut (N = 3,577;
  drops 151 team-obs)
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
N = 2,089, post-2026-09-03 qualifying-round fix, was 2,101), split by All / both-top-100 /
both-top-50 (ranking at time of tournament).
Random benchmark is closed-form (not simulated): Σ over C(n,2) pairs in that tournament's
actual field, using the same CEPII `comlang_off`/`comlang_ethno` country-pair lookup that
`same_language`/`ling_prox` are themselves built from. Finding: partner selection is far more
culturally assortative than chance at every skill level (6–8× benchmark for nationality,
3–4× for language), but the degree of excess is essentially flat across brackets — elite
players are not disproportionately more assortative. A follow-up continuous test (§6.1,
logit/OLS of same_country/same_language/ling_prox on the ego player's own ranking, ~4,178
ego-rows, clustered by partnership) confirms this: the coefficient on own rank is positive
and significant for all three outcomes, meaning *worse*-ranked players sort into
same-culture partnerships slightly *more*, not less — the reverse of the "stronger players
have more choice and sort more" concern.

**Section 6.2 (new, per Lingqing's 2026-08 meeting notes; finalized 2026-09): tournament-field
composition.** Extends the §6.1 ego-row regression from `own_rank` alone to three specs: (1)
own_rank alone; (2) the outcome-matched field-composition variable alone; (3) own_rank + field
composition together. Field composition = share of every *other* player in that tournament-year's
full field (all entrants, not just rank-complete partnerships) sharing the focal player's
nationality (or the analogous mean pairwise same_language/ling_prox value). All three specs
exclude the 3 ego-rows whose own nationality is represented by exactly one player in the
ego-row sample (N=4,175, from 4,178). An own-nationality (country FE) specification — `own_rank`
replaced by `C(own_iso3)`, a 60-level categorical — was tried and dropped entirely, not just
partially reported: it produces no single reportable coefficient, and its binary-outcome logits
(same_country, same_language) fail to converge (quasi-complete separation) regardless. Only
Spec 1/2/3 above are estimated or reported anywhere. Reported in `main.tex` §6.2, `report.html`
§6.2, and `homophily.do` Section 8.

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

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
**⚠ This fix lives ONLY in the currently-committed `data/atp/men_matches_with_ranks_cleaned.xlsx`
— confirmed 2026-09-11 that it is not encoded anywhere in `code/` (checked `final_ds.ipynb`,
`code/merging/`, and `homophily.ipynb` — only a markdown cell in the latter describes it,
none of it is executable).** `men_matches_with_ranks.xlsx` (the raw input to `final_ds.ipynb`)
lacks it. **Never run `final_ds.ipynb` end-to-end and overwrite `men_matches_with_ranks_cleaned.xlsx`
with its output** — doing so silently regresses the sample to 1,860 matches / 3,720 team-obs
(discovered the hard way 2026-09-11, when adding the continuous-linguistic-proximity columns
via a full notebook re-run did exactly this; fixed by `git checkout` on the cleaned file plus a
standalone column-only patch script instead of a full re-run). If you need to add a column to
this file, read the already-cleaned file and add the column with a standalone script/cell,
the way the continuous linguistic-proximity columns were ultimately added — do not regenerate
it from `INPUT_FILE`.
`exp_mean` imputed to 1 (not 0) for rookies (turned pro in/after the tournament year, so raw
tenure ≤0) — a nominal first-year tenure rather than zero, per reviewer feedback; no observations
are dropped. Demeaned against each estimation sample's own mean to form `exp_mean_dm` (see
Controls below).

**Sample correction (found 2026-09-03): Wimbledon 2018 qualifying-round contamination.**
The raw scrape for Wimbledon 2018 alone included 12 qualifying-round doubles matches (8 "1st
Round Qualifying" + 4 "2nd Round Qualifying") mixed in with the 63 main-draw matches — no other
tournament-year in the dataset has any qualifying-round rows. These are now excluded, which is what moved the
sample from the previously-documented 1,876 matches / 3,752 team-obs down to the current 1,864 /
3,728 (Wimbledon: 437 → 425 matches). This is a genuine data-quality fix, not an error in either
the notebook or the paper — `overleaf/main.tex` (Lingqing's 2026-09 revision) already reflects the
corrected sample throughout and was independently verified against a fresh end-to-end re-execution
of `homophily.ipynb` (e.g. Table 3 same-nationality AME = 0.0361, matching exactly).
**(moved 2026-09-11)** This filter now lives in `code/cleaning/final_ds.ipynb` (per the project
principle that all cleaning/filtering/merging belongs in the cleaning pipeline, not the analysis
notebook) — `men_matches_with_ranks_cleaned.xlsx` is already qualifying-round-free (1,985 rows,
was 1,997). Applied via the same safe non-destructive patch pattern as the linguistic-proximity
columns (see the warning above) — NOT via a full `final_ds.ipynb` re-run — after confirming none
of the 4 manually-fixed matches (6, 35, 49, 913) were among the 12 dropped rows.
`homophily.ipynb` cell 1 now just asserts no qualifying rows remain, rather than filtering them.
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
- `ling_prox` — **changed 2026-09-10: now `prox1`, continuous on [0,1]**, not the binary
  CEPII `comlang_ethno` measure documented in earlier versions of this file. `prox1`
  \citep{melitztoubal2014} is Ethnologue-tree-based linguistic proximity — the same measure
  \citet{bekesottaviano2025} use for their own language-similarity variable, retrieved from
  Farid Toubal's site (`data/gravity/melitz_toubal_proxling.dta`; CEPII's own download link
  for this file is dead) since CEPII's public gravity table only ever exposed the
  pre-binarized `comlang_ethno`. Same-country pairs are forced to 1.0; Monaco (absent from
  the raw Melitz-Toubal table) uses France's value as a stand-in, or 1.0 against another
  French-official country; a handful of remaining pairs (mostly involving South Korea,
  entirely absent from this table) fall back to 0. 100% coverage on this sample.
  **(moved 2026-09-11)** All raw ingestion + Monaco/KOR fallback resolution now lives in
  `code/merging/merge_linguistic_proximity.ipynb`, which exports a complete, closed
  country-pair lookup to `data/gravity/ling_prox_pairs_final.csv` (65 countries × all pairs
  + self-pairs, no missing values). `homophily.ipynb` and `final_ds.ipynb` just read that
  file — no raw Melitz-Toubal data or fallback logic in either of them anymore. Similarly,
  `code/merging/merge_psw2024.ipynb` resolves the two PSW2024 robustness measures (below)
  into `data/gravity/psw2024_pairs_final.csv` (100% coverage directly from the raw file, no
  fallback needed). All four non-ATP reference datasets — `gravity_lang_lookup.xlsx`,
  `melitz_toubal_proxling.dta`, `linguistic_distance_PSW2024.csv`, and the two `_final.csv`
  outputs — live in `data/gravity/` alongside `hofstede.csv`, not `data/atp/` (which holds
  only ATP match/ranking data).
  **Estimator implication:** `ling_prox` is now continuous, so wherever `ling_prox` is the
  *outcome* variable (Section 6.1/6.2 partner-selection regressions, `homophily.ipynb`
  cells 59/61, `homophily.do` Section 8), it uses OLS/LPM, not logit — same_country/
  same_language there still use logit+AME as before. Wherever `ling_prox` is a *regressor*
  (Tables 3–6b), logit+AME is unchanged; only the AME's underlying regressor changed from
  binary to continuous, so the reported number is now "AME per full 0→1 change," directly
  comparable in magnitude to the old binary AME.
  **Appendix (Tables A3–A6b in `main.tex`, `\label{app:langprox}`):** the original binary
  `comlang_ethno` (now `ling_prox_binary`) plus two further continuous robustness measures
  from `data/gravity/linguistic_distance_PSW2024.csv` \citep{psw2024} — `ling_prox_psw_tree`
  (inverted tree distance) and `ling_prox_psw_cognet` (cognate/lexical proximity) — are
  relocated there. All tables agree qualitatively with the `prox1` main text with one
  exception: Table 6 (culture × `exp_mean`)'s interaction is significant (p=0.034) for the
  binary measure but *not* for `prox1` (p=0.211) or either PSW measure (p=0.305, p=0.265) —
  the experience-amplification finding holds for shared language but not for linguistic
  proximity once measured continuously. Every other table's conclusion is unchanged across
  all four measures. Distinct from `same_language` only in using linguistic closeness
  rather than official designation — the two can diverge, e.g. Italian/Spanish speakers
  score high `prox1` without sharing an official language.

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
logit AME of same_country/same_language/ling_prox on the ego player's own ranking, ~4,178
ego-rows, clustered by partnership) confirms this: the AME on own rank is positive
and significant for all three outcomes, meaning *worse*-ranked players sort into
same-culture partnerships slightly *more*, not less — the reverse of the "stronger players
have more choice and sort more" concern.
**Estimator fix (2026-09-10):** `ling_prox` was previously estimated with OLS/LPM in §6.1/6.2
while `same_country`/`same_language` used logit AME — an inconsistency, since `ling_prox` is
just as binary as the other two (see the "Culture measures" section above). All three outcomes
now use logit AME throughout §6.1/6.2, in `homophily.ipynb`, `homophily.do`, `report.html`, and
`main.tex`. The switch changes `ling_prox`'s coefficients modestly (e.g. §6.1 Spec 1 own_rank:
was coef=+0.000104 (OLS), now AME=+0.000146) but not the sign, significance, or qualitative
finding of any table.

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

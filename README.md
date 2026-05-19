# Tennis Doubles Homophily

Does cultural similarity between doubles partners affect match outcomes in Grand Slam tennis?

This project examines **homophily** — the tendency for teammates to share cultural traits — and its relationship with performance in ATP men's doubles at Grand Slam tournaments, including Olympic cycles.

---

## Research Question

We test whether teams whose players share **nationality**, **official language**, or **linguistic proximity** (ethnic language overlap) perform better than culturally mixed teams, and whether this effect is amplified or attenuated under pressure (tiebreaks, comebacks).

---

## Data

| Source | Description |
|---|---|
| Wikipedia bracket pages | Match scores and draw results scraped for each Grand Slam, 2018–2025 |
| ATP ranking archives | Doubles rankings merged at match date (within 1-year window) |
| Gravity linguistic dataset | Linguistic proximity index between nationalities |

- **Coverage:** Australian Open, Roland Garros, Wimbledon, US Open — 2018–2025 (excl. 2020); Olympics 2021 and 2024 included separately
- **Final sample:** 1,793 completed matches (1,840 raw; 47 dropped: 20 walkovers, 27 retirements)
- **Regression sample:** 1,599 GS matches, 3,198 team-obs (138 further dropped: 14 for ranking, 124 for nationality/language; Olympics excluded)
- **Unit of analysis:** team-match pair (two observations per match: winner team = 1, loser team = 0)

---

## Pipeline

```
code/scraping/matches_scraper.ipynb   ← Wikipedia wikitext bracket parser
code/merging/merge_matches.ipynb      ← consolidate scraped data
code/cleaning/matches_doubles.ipynb   ← clean scores, flag retirements/walkovers
code/merging/matches_ranks.ipynb      ← join ATP doubles rankings
code/cleaning/final_ds.ipynb          ← add homophily variables, final dataset
code/analysis/homophily_analysis.ipynb ← main analysis (descriptives + regressions)
```

`run_pipeline.py` executes the full chain in order via `nbclient`.

---

## Key Variables

| Variable | Description |
|---|---|
| `same_country` | Both players share the same nationality (0/1) |
| `same_language` | Both players share the same official language (0/1) |
| `ling_prox` | Continuous linguistic proximity index (ethnic language overlap, 0–1) |
| `rank_mean` | Team average ATP doubles ranking at match date |
| `rank_gap` | Absolute ranking gap between the two teammates |
| `single_top100` | At least one player was a top-100 singles player within 1 year (0/1) |

---

## Main Results

### Descriptive: Homophily rises around Olympic cycles

Team composition at Grand Slams shifts markedly in the lead-up to the Olympics, consistent with players forming nationally-homogeneous pairs to qualify for the Games.

**Paris 2024 cycle (Grand Slams only):**

| Period | Window | Same Nationality | Same Language | Ling. Proximity |
|---|---|---|---|---|
| Pre-Paris | AO22–USO22 | 39.3% | 54.2% | 55.1% |
| Paris Prep | AO23–Wim24 | 42.5% | 59.1% | 59.6% |
| Post-Paris | USO24–USO25 | 45.3% | 57.7% | 59.1% |

**Winner vs. loser comparison (Grand Slams):**

| Measure | Winners | Losers | Difference |
|---|---|---|---|
| Same nationality | 39.5% | 41.3% | −1.7 pp |
| Same language | 55.1% | 52.8% | +2.3 pp |
| Ling. proximity | 55.7% | 53.8% | +1.9 pp |

### Regressions

All models: logit, tournament×year fixed effects, round fixed effects, standard errors clustered by match. Culture measure: linguistic proximity (ethnic). Controls: team average ranking, opponent average ranking, teammate rank gap, top-100 singles indicator.

**Table 3 — Match Win** (N = 3,198 team-obs, 1,599 matches):

| Variable | Coeff | SE | p |
|---|---|---|---|
| Same nationality | +0.058 | 0.079 | 0.466 |
| Same language | +0.116 | 0.077 | 0.129 |
| Language proximity | +0.097 | 0.076 | 0.205 |
| Team avg. ranking | −0.004*** | 0.001 | <0.001 |
| Opponent avg. ranking | +0.003*** | 0.000 | <0.001 |

**Table 4 — Tiebreak Win** (N = 1,424 team-obs, 712 matches; 7-pt regular tiebreaks, sets 1–2):

| Variable | Coeff | SE | p |
|---|---|---|---|
| Same nationality | +0.104 | 0.120 | 0.384 |
| Same language | +0.059 | 0.116 | 0.610 |
| Language proximity | +0.041 | 0.116 | 0.723 |

**Table 5 — Comeback Win** (N = 1,599 obs, 1 per match; all teams that lost set 1):

| Variable | Coeff | SE | p |
|---|---|---|---|
| Same nationality | −0.237* | 0.137 | 0.083 |
| Same language | −0.326** | 0.132 | 0.014 |
| Language proximity | −0.332** | 0.133 | 0.012 |

### Interpretation

- Cultural similarity (nationality, language, linguistic proximity) **does not predict overall match wins** once ranking controls and fixed effects are included.
- No effect in tiebreaks — culturally similar teams show no advantage or disadvantage in 7-point sudden-death situations.
- Culturally similar teams are **less likely to mount a comeback** after losing the first set (same language p = 0.014; language proximity p = 0.012), suggesting they are more dominant-set-heavy and less resilient when behind.

---

## Observation Breakdown

| Step | Description | Matches | Team-obs |
|---|---|---|---|
| 1 | Raw dataset (2018–2025, excl. 2020) | 1,840 | 3,680 |
| 2 | Drop retirements / walkovers (20 WO · 9 ret. S1 · 18 ret. S2+) | 1,793 | 3,586 |
| 3 | Drop: ranking incomplete for ≥1 player | 1,779 | 3,558 |
| 4 | Drop: nationality/language missing for ≥1 player (birthplace fallback applied first) | 1,655 | 3,310 |
| 5 | Exclude Olympic matches (Tokyo 2021 · Paris 2024) | **1,599** | **3,198** |
| → | **Grand Slams regression sample (Tables 3–5)** | **1,599** | **3,198** |

Regression sub-samples: Table 4 — 712 matches, 1,424 team-obs (regular tiebreaks). Table 5 — 1,599 obs, 1 per match (all teams that lost set 1).

## Pressure Outcome Frequencies

| Outcome | N | Rate |
|---|---|---|
| Any tiebreak (all types) | 896 | 50.0% |
| Regular tiebreak (sets 1 or 2) | 804 | 44.8% |
| Match went to 3 sets | 787 | 43.9% |
| Comeback wins | 351 | 19.6% |

---

## Repository Structure

```
data/atp/                    ← raw and processed datasets (Excel/CSV)
code/scraping/               ← Wikipedia scraper
code/merging/                ← data merging notebooks
code/cleaning/               ← data cleaning notebooks
code/analysis/               ← main analysis notebook
ppt/                         ← figures and poster slides
run_pipeline.py              ← end-to-end pipeline runner
```

---

## Requirements

Python 3.10+. Key packages: `pandas`, `numpy`, `statsmodels`, `matplotlib`, `nbformat`, `nbclient`, `selenium` (for scraping).

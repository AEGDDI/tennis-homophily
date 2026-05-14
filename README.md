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
| Pre-Paris | AO22–USO22 | 37.4% | 51.2% | 52.3% |
| Paris Prep | AO23–Wim24 | 41.0% | 55.6% | 56.0% |
| Post-Paris | USO24–USO25 | 43.1% | 54.8% | 56.2% |

**Winner vs. loser comparison (Grand Slams):**

| Measure | Winners | Losers | Difference |
|---|---|---|---|
| Same nationality | 39.5% | 41.3% | −1.7 pp |
| Same language | 55.1% | 52.8% | +2.3 pp |
| Ling. proximity | 55.7% | 53.8% | +1.9 pp |

### Regressions

All models: logit, tournament×year fixed effects, round fixed effects, standard errors clustered by match. Culture measure: linguistic proximity (ethnic). Controls: team average ranking, opponent average ranking, teammate rank gap, top-100 singles indicator.

**Table 3 — Match Win** (N = 3,436 team-obs, 1,718 matches):

| Variable | Coeff | SE | p |
|---|---|---|---|
| Language proximity | +0.155** | 0.073 | 0.034 |
| Team avg. ranking | −0.004*** | 0.001 | <0.001 |
| Opponent avg. ranking | +0.003*** | 0.000 | <0.001 |

**Table 4 — Tiebreak Win** (N = 1,717, restricted to matches with any tiebreak):

| Variable | Coeff | SE | p |
|---|---|---|---|
| Language proximity | −0.026 | 0.108 | 0.808 |

**Table 5 — Comeback Win** (N = 762, restricted to 3-set matches where team lost set 1):

| Variable | Coeff | SE | p |
|---|---|---|---|
| Language proximity | −0.398** | 0.171 | 0.020 |

### Interpretation

- Linguistic proximity **positively predicts match wins** overall, consistent with a communication or coordination advantage.
- The effect **does not extend to tiebreaks** — in sudden-pressure situations (7-point tiebreak), cultural similarity offers no measurable edge.
- Counterintuitively, linguistically similar teams are **less likely to mount a comeback** after losing the first set, suggesting the homophily advantage may reflect a playing style that is more dominant-set-heavy and less resilient when behind.

---

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

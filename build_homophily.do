* build_homophily.do
* Tennis Doubles Homophily — Main Analysis
* Run this from the project root in STATA.
* Mirrors sections 1-5 from homophily_analysis.ipynb
*
* Data: Grand Slam doubles matches 2018–2025. N = 2,192 matches.
* Sections:
*   1. Variable construction
*   2. Pressure outcomes (tiebreaks, comebacks)
*   3. Olympic cycle descriptive evidence
*   4. Team panel construction (match-level → team-level)
*   5. Baseline regressions (match win / tiebreak win / comeback win)

capture log close _all
set more off
cd "C:\Users\aldi\Documents\GitHub\tennis-homophily"
clear

import excel using "data/atp/men_matches_with_ranks_cleaned.xlsx", sheet("players_list") firstrow clear

destring year winners_set1 winners_set2 winners_set3 losers_set1 losers_set2 losers_set3 ///
    winners_p1_top100_within_1y winners_p2_top100_within_1y losers_p1_top100_within_1y losers_p2_top100_within_1y ///
    rank_mean_winners rank_mean_losers rank_diff_winners rank_diff_losers, replace force

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 1: VARIABLE CONSTRUCTION
* ═══════════════════════════════════════════════════════════════════════════════

* ── Tiebreak flags ────────────────────────────────────────────────────────────
* Regular 7-point tiebreak: set reaches 7-6
generate byte tb_s1 = (winners_set1==7 & losers_set1==6) | (winners_set1==6 & losers_set1==7)
generate byte tb_s2 = (winners_set2==7 & losers_set2==6) | (winners_set2==6 & losers_set2==7)
generate byte regular_tb = tb_s1 | tb_s2

* Match tiebreak / super-tiebreak (10-point, set3 ≥ 8): Olympics and Wimbledon post-2018
generate byte tb_s3_regular = (winners_set3==7 & losers_set3==6) | (winners_set3==6 & losers_set3==7)
generate byte match_tb = (winners_set3 >= 8)
generate byte any_tb = regular_tb | match_tb | tb_s3_regular

* Who won each tiebreak (winner-team perspective)
generate byte w_won_tb_s1 = tb_s1 & (winners_set1==7)
generate byte w_won_tb_s2 = tb_s2 & (winners_set2==7)

* ── Match structure ───────────────────────────────────────────────────────────
generate byte three_sets = !missing(winners_set3)
generate byte winner_lost_s1 = (winners_set1 < losers_set1)
generate byte comeback = winner_lost_s1 & three_sets
* ── Olympic period flags ──────────────────────────────────────────────────────
* Tokyo 2021 cycle: Pre-Tokyo, Tokyo Prep, Post-Tokyo
* Paris 2024 cycle: Pre-Paris, Paris Prep, Post-Paris

generate str12 cycle_tokyo = ""
replace cycle_tokyo = "Pre-Tokyo" if inlist(year, 2018, 2019)
replace cycle_tokyo = "Tokyo Prep" if year == 2020 & tournament != "Wimbledon"
replace cycle_tokyo = "Tokyo Prep" if year == 2021 & inlist(tournament, "Australian Open", "Roland Garros", "Wimbledon")
replace cycle_tokyo = "Post-Tokyo" if (year == 2021 & tournament == "US Open") | year == 2022

generate str12 cycle_paris = ""
replace cycle_paris = "Pre-Paris" if year == 2022
replace cycle_paris = "Paris Prep" if year == 2023
replace cycle_paris = "Paris Prep" if year == 2024 & inlist(tournament, "Australian Open", "Roland Garros", "Wimbledon")
replace cycle_paris = "Post-Paris" if (year == 2024 & tournament == "US Open") | year == 2025

generate byte pre_olympic   = cycle_paris != ""
generate byte olympic_period = cycle_paris != "" | cycle_tokyo != ""
generate byte olympics_tourn = (tournament == "Olympics")

* ── Match-level homophily (averaged across both teams in each match) ──────────
generate double same_country_avg   = (same_country_winners + same_country_losers) / 2
generate double same_language_avg  = (winners_same_language + losers_same_language) / 2
generate double ling_prox_avg      = (winners_linguistic_proximity + losers_linguistic_proximity) / 2

display "Section 1 complete: Variable construction."

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 2: PRESSURE OUTCOMES — COUNTS AND INSPECTION
* ═══════════════════════════════════════════════════════════════════════════════

display ""
display "=== TABLE 1. Pressure Outcome Counts (all matches, N=2,192) ==="
display ""
count if tb_s1 == 1
local n_tb_s1 = r(N)
display "  Set-1 tiebreak (7-pt):                                 N=" %4.0f `n_tb_s1'

count if tb_s2 == 1
local n_tb_s2 = r(N)
display "  Set-2 tiebreak (7-pt):                                 N=" %4.0f `n_tb_s2'

count if tb_s3_regular == 1
local n_tb_s3 = r(N)
display "  Set-3 regular tiebreak (7-pt):                         N=" %4.0f `n_tb_s3'

count if match_tb == 1
local n_match_tb = r(N)
display "  Match tiebreak / super-tb (10-pt, set3 >= 8):          N=" %4.0f `n_match_tb'

count if regular_tb == 1
local n_regular_tb = r(N)
display "  Any regular tiebreak (sets 1 or 2):                    N=" %4.0f `n_regular_tb'

count if any_tb == 1
local n_any_tb = r(N)
display "  Any tiebreak (all types):                              N=" %4.0f `n_any_tb'

count if three_sets == 1
local n_three_sets = r(N)
display "  Match went to 3 sets:                                  N=" %4.0f `n_three_sets'

count if comeback == 1
local n_comeback = r(N)
display "  Comeback wins (winner lost set 1):                     N=" %4.0f `n_comeback'

display ""
display "Regular tiebreaks by tournament:"
tabstat regular_tb, by(tournament) stats(sum mean) format(%9.0f)

display ""
display "Comeback wins by tournament:"
tabstat comeback, by(tournament) stats(sum mean) format(%9.0f)

display "Section 2 complete: Pressure outcomes."
display "Section 2 complete: Pressure outcomes."

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 3: OLYMPIC CYCLE DESCRIPTIVE EVIDENCE
* ═══════════════════════════════════════════════════════════════════════════════

display ""
display "=== TABLE 2A. Tokyo 2021 cycle — Team composition across period ==="

preserve
keep if olympics_tourn==0
collapse (count) N_teams=match_id (mean) same_country_avg same_language_avg ling_prox_avg, by(cycle_tokyo)
destring N_teams, replace
replace same_country_avg = same_country_avg*100
replace same_language_avg = same_language_avg*100
replace ling_prox_avg = ling_prox_avg*100

replace N_teams = N_teams*2
order cycle_tokyo
list cycle_tokyo N_teams same_country_avg same_language_avg ling_prox_avg, abbreviate(14) separator(0)
restore

display ""
display "=== TABLE 2B. Paris 2024 cycle — Team composition across period ==="

preserve
keep if olympics_tourn==0
collapse (count) N_teams=match_id (mean) same_country_avg same_language_avg ling_prox_avg, by(cycle_paris)
destring N_teams, replace
replace same_country_avg = same_country_avg*100
replace same_language_avg = same_language_avg*100
replace ling_prox_avg = ling_prox_avg*100

replace N_teams = N_teams*2
order cycle_paris
list cycle_paris N_teams same_country_avg same_language_avg ling_prox_avg, abbreviate(14) separator(0)
restore

display ""
display "Winner vs. loser team homophily rates (Grand Slams):"
preserve
keep if olympics_tourn==0
display "Same nationality:"
tabstat same_country_winners same_country_losers, stats(mean) format(%9.3f)
display "Same language:"
tabstat winners_same_language losers_same_language, stats(mean) format(%9.3f)
display "Linguistic proximity:"
tabstat winners_linguistic_proximity losers_linguistic_proximity, stats(mean) format(%9.3f)
restore

display "Section 3 complete: Olympic cycle descriptive evidence."

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 4: TEAM PANEL CONSTRUCTION
* ═══════════════════════════════════════════════════════════════════════════════
* Load team-level panel from CSV (pre-constructed by homophily_analysis.ipynb)
* Each observation is a team in a match (2 obs per match: winner=1, loser=0)

clear
import delimited "data/atp/team_gs_panel.csv", clear

* Check structure
display "After import:"
describe
display ""
display "Check variable types and first row:"
list match_id tournament year win ling_prox rank_mean in 1/1

* Destring all variables
destring _all, replace force

display ""
display "After destring - Check variable types and contents:"
describe tournament year stage_code, short

display ""
display "Sample of tournament and year:"
list match_id tournament year win in 1/10

* If tournament is still string, encode it
capture confirm string variable tournament
if !_rc {
    display ""
    display "Tournament is STRING; encoding to numeric..."
    encode tournament, generate(tourn_code)
    drop tournament
    rename tourn_code tournament
    describe tournament, short
}

* If year is string, destring it again
capture confirm numeric variable year
if _rc {
    display "Year is STRING; destringifying..."
    destring year, replace force
}

display ""
display "Final check before creating ty:"
describe tournament year stage_code, short
list match_id tournament year win in 1/10

count
display "Team panel loaded: " r(N) " obs"
display ""
display "Section 4 complete: Team panel loaded."

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 5: BASELINE REGRESSIONS
* ═══════════════════════════════════════════════════════════════════════════════
* Each match contributes two observations (winner team = 1, loser team = 0)
* Culture measure: language proximity (ling_prox) entered one at a time
* Controls: team avg doubles rank, opponent rank, teammate rank gap, top-100 singles indicator
* Fixed effects: tournament×year interaction, round (stage_code)
* Standard errors: clustered by match

display ""
display "Diagnostic: Data check before regression"
count if !missing(win)
display "Non-missing win: " r(N)
display ""

* Replace missing stage_code with 0 for regression
replace stage_code = 0 if missing(stage_code)

display "Creating ty (tournament×year grouping)..."
describe tournament year, short
list tournament year in 1/5

capture drop ty
quietly egen ty = group(tournament year)

display "After egen:"
describe ty
tabstat ty, stats(count mean min max)
list match_id tournament year ty win in 1/10

display ""

log using "stata_homophily_results.txt", replace text

display ""
display "=== TABLE 3. Match Win — Logit ==="
display "FE: tournament×year (as ty) + stage_code | SE clustered by match"
logit win i.ty i.stage_code ling_prox rank_mean opp_rank_mean rank_gap single_top100, cluster(match_id)
estimates store win_ling_prox
estimates table win_ling_prox, b se star stats(N ll)

display ""
display "=== TABLE 4. Tiebreak Win — Logit ==="
display "Sample: matches with any tiebreak (7-pt or 10-pt)"
count if any_tb==1
display "  Obs with any_tb==1: " r(N)
logit won_any_tb i.ty i.stage_code ling_prox rank_mean opp_rank_mean rank_gap single_top100 if any_tb==1, cluster(match_id)
estimates store tb_ling_prox
estimates table tb_ling_prox, b se star stats(N ll)

display ""
display "=== TABLE 5. Comeback Win — Logit ==="
display "Outcome: team won after losing set 1 (P(win | lost set 1, 3 sets))"
display "Sample: 3-set Grand Slam matches, conditional on losing set 1"

preserve
keep if three_sets==1 & lost_set1==1
count
display "Raw sample: " r(N) " team-obs | " r(N)/2 " matches"

* Drop tournament×year cells with no outcome variation
quietly egen minwin = min(win), by(ty)
quietly egen maxwin = max(win), by(ty)
drop if minwin==maxwin
count
display "After dropping non-informative ty cells: " r(N) " team-obs | " r(N)/2 " matches"

logit win i.ty i.stage_code ling_prox rank_mean opp_rank_mean rank_gap single_top100, cluster(match_id)
estimates store cb_ling_prox
estimates table cb_ling_prox, b se star stats(N ll)
restore

log close

display ""
display "Section 5 complete: Baseline regressions."
display "Results saved to stata_homophily_results.txt"
estimates store cb_ling_prox
estimates table cb_ling_prox, b se star stats(N ll)
restore

log close

display ""
display "Section 5 complete: Baseline regressions."
display "Results saved to stata_homophily_results.txt"


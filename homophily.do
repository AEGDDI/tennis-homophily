

* homophily.do
* Tennis Doubles Homophily — Main Analysis
* Run this from the project root in STATA.
* Mirrors sections 1-5 from homophily_analysis.ipynb
*
* Data: Grand Slam doubles matches 2018–2025 (excl. 2020).
*       Raw: 1,840 matches. After dropping 47 retirements/walkovers: 1,793 matches.
*       GS panel (complete ranking + nationality/language): 1,716 matches, 3,432 team-obs.
*       Regression sample: 3,432 team-obs (exp_mean imputed to 0 for 210 first-time GS participants).
*       Nationality/language filled from pipeline (birthplace, surname lookup, Monaco fix,
*       manual overrides in manual_nationality.csv).
* Sections:
*   0. Observation breakdown (funnel from raw data to regression sample)
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

* Drop likely retirements/walkovers that are not explicitly flagged in the data.
* A valid completed set is 6-0 through 6-4, 7-5, 7-6, or an extended final set
* score with a two-game margin (e.g. 8-6, 9-7). If the first two sets are split,
* the third set must also be complete, either as a regular set or as a match tiebreak.
generate byte s1_complete = !missing(winners_set1, losers_set1) & ///
    ((max(winners_set1, losers_set1)==6 & min(winners_set1, losers_set1)<=4) | ///
     (max(winners_set1, losers_set1)==7 & inlist(min(winners_set1, losers_set1), 5, 6)) | ///
     (max(winners_set1, losers_set1)>=8 & max(winners_set1, losers_set1) - min(winners_set1, losers_set1)==2))

generate byte s2_complete = !missing(winners_set2, losers_set2) & ///
    ((max(winners_set2, losers_set2)==6 & min(winners_set2, losers_set2)<=4) | ///
     (max(winners_set2, losers_set2)==7 & inlist(min(winners_set2, losers_set2), 5, 6)) | ///
     (max(winners_set2, losers_set2)>=8 & max(winners_set2, losers_set2) - min(winners_set2, losers_set2)==2))

generate byte split_sets = !missing(winners_set1, losers_set1, winners_set2, losers_set2) & ///
    ((winners_set1 > losers_set1 & winners_set2 < losers_set2) | ///
     (winners_set1 < losers_set1 & winners_set2 > losers_set2))

generate byte s3_regular_complete = !missing(winners_set3, losers_set3) & ///
    ((max(winners_set3, losers_set3)==6 & min(winners_set3, losers_set3)<=4) | ///
     (max(winners_set3, losers_set3)==7 & inlist(min(winners_set3, losers_set3), 5, 6)) | ///
     (max(winners_set3, losers_set3)>=8 & max(winners_set3, losers_set3) - min(winners_set3, losers_set3)==2))

generate byte s3_match_tb_complete = !missing(winners_set3, losers_set3) & ///
    max(winners_set3, losers_set3) >= 10 & ///
    max(winners_set3, losers_set3) - min(winners_set3, losers_set3) >= 2

generate byte s3_complete = s3_regular_complete | s3_match_tb_complete
generate byte retired_or_incomplete = !s1_complete | !s2_complete | ///
    (split_sets & !s3_complete)

tempfile incomplete_matches
preserve
keep if retired_or_incomplete
keep match_id
duplicates drop
save `incomplete_matches'
restore

count if retired_or_incomplete
display "Dropping likely retirements/walkovers: " r(N) " matches"
drop if retired_or_incomplete

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 1: VARIABLE CONSTRUCTION
* ═══════════════════════════════════════════════════════════════════════════════

* ── Tiebreak flags ────────────────────────────────────────────────────────────
* Regular 7-point tiebreak: set reaches 7-6
generate byte tb_s1 = (winners_set1==7 & losers_set1==6) | (winners_set1==6 & losers_set1==7)
generate byte tb_s2 = (winners_set2==7 & losers_set2==6) | (winners_set2==6 & losers_set2==7)
generate byte regular_tb = tb_s1 | tb_s2

* Match tiebreak / super-tiebreak (10-point, set3 >= 10): Olympics and Wimbledon post-2018
generate byte tb_s3_regular = (winners_set3==7 & losers_set3==6) | (winners_set3==6 & losers_set3==7)
generate byte match_tb = s3_match_tb_complete
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

generate byte pre_olympic    = inlist(cycle_paris, "Pre-Paris", "Paris Prep")
generate byte olympic_period = inlist(cycle_paris, "Pre-Paris", "Paris Prep", "Post-Paris")
generate byte olympics_tourn = (tournament == "Olympics")

* ── Match-level homophily (averaged across both teams in each match) ──────────
generate double same_country_avg   = (same_country_winners + same_country_losers) / 2
generate double same_language_avg  = (winners_same_language + losers_same_language) / 2
generate double ling_prox_avg      = (winners_linguistic_proximity + losers_linguistic_proximity) / 2

display "Section 1 complete: Variable construction."

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 2: PRESSURE OUTCOMES — COUNTS AND INSPECTION
* ═══════════════════════════════════════════════════════════════════════════════

count
local N_all = r(N)
display ""
display "=== TABLE 1. Pressure Outcome Counts (all matches, N=" `N_all' ") ==="
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
display "  Match tiebreak / super-tb (10-pt, set3 >= 10):         N=" %4.0f `n_match_tb'

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
display "Match tiebreaks by tournament:"
tabstat match_tb, by(tournament) stats(sum mean) format(%9.0f)
egen tourn_year = group(tournament year), label
tabstat match_tb, by(tourn_year) stats(sum mean) format(%9.3f)

display ""
display "Comeback wins by tournament:"
tabstat comeback, by(tournament) stats(sum mean) format(%9.0f)

display "Section 2 complete: Pressure outcomes."

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 3: OLYMPIC CYCLE DESCRIPTIVE EVIDENCE
* ═══════════════════════════════════════════════════════════════════════════════

display ""
display "=== TABLE 2A. Tokyo 2021 cycle — Same nat/lang/langprox by period ==="

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
display "=== TABLE 2B. Paris 2024 cycle — Same nat/lang/langprox by period ==="

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
display "Overall same vs. diff homophily rates (Grand Slams, all teams):"
preserve
keep if olympics_tourn==0
display "Same nationality (same=avg, diff=1-avg):"
tabstat same_country_avg, stats(mean) format(%9.3f)
display "Same language (same=avg, diff=1-avg):"
tabstat same_language_avg, stats(mean) format(%9.3f)
display "Linguistic proximity (same=avg, diff=1-avg):"
tabstat ling_prox_avg, stats(mean) format(%9.3f)
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

* Preserve string categorical variables before destringing numeric fields.
capture confirm string variable tournament
if !_rc {
    display ""
    display "Tournament is STRING; encoding to numeric..."
    encode tournament, generate(tourn_code)
    drop tournament
    rename tourn_code tournament
    describe tournament, short
}

capture confirm string variable surface
if !_rc {
    encode surface, generate(surface_code)
    drop surface
    rename surface_code surface
}

* Destring numeric variables that may have been imported as text.
destring _all, replace force

* If Section 4 is run by itself, the tempfile from the raw-match section above
* does not exist. Rebuild it here so the team panel is filtered consistently.
if `"`incomplete_matches'"' == "" {
    tempfile incomplete_matches
    preserve
    import excel using "data/atp/men_matches_with_ranks_cleaned.xlsx", sheet("players_list") firstrow clear

    destring year winners_set1 winners_set2 winners_set3 losers_set1 losers_set2 losers_set3, replace force

    generate byte s1_complete = !missing(winners_set1, losers_set1) & ///
        ((max(winners_set1, losers_set1)==6 & min(winners_set1, losers_set1)<=4) | ///
         (max(winners_set1, losers_set1)==7 & inlist(min(winners_set1, losers_set1), 5, 6)))

    generate byte s2_complete = !missing(winners_set2, losers_set2) & ///
        ((max(winners_set2, losers_set2)==6 & min(winners_set2, losers_set2)<=4) | ///
         (max(winners_set2, losers_set2)==7 & inlist(min(winners_set2, losers_set2), 5, 6)))

    generate byte split_sets = !missing(winners_set1, losers_set1, winners_set2, losers_set2) & ///
        ((winners_set1 > losers_set1 & winners_set2 < losers_set2) | ///
         (winners_set1 < losers_set1 & winners_set2 > losers_set2))

    generate byte s3_regular_complete = !missing(winners_set3, losers_set3) & ///
        ((max(winners_set3, losers_set3)==6 & min(winners_set3, losers_set3)<=4) | ///
         (max(winners_set3, losers_set3)==7 & inlist(min(winners_set3, losers_set3), 5, 6)) | ///
         (max(winners_set3, losers_set3)>=8 & max(winners_set3, losers_set3) - min(winners_set3, losers_set3)==2))

    generate byte s3_match_tb_complete = !missing(winners_set3, losers_set3) & ///
        max(winners_set3, losers_set3) >= 10 & ///
        max(winners_set3, losers_set3) - min(winners_set3, losers_set3) >= 2

    generate byte s3_complete = s3_regular_complete | s3_match_tb_complete
    generate byte retired_or_incomplete = !s1_complete | !s2_complete | ///
        (split_sets & !s3_complete)

    keep if retired_or_incomplete
    keep match_id
    duplicates drop
    save `incomplete_matches'
    restore
}

merge m:1 match_id using `incomplete_matches', keep(master match)
drop if _merge == 3
drop _merge
count
display "Team panel after dropping likely retirements/walkovers: " r(N) " obs"

display ""
display "After destring - Check variable types and contents:"
describe tournament year stage_code, short

display ""
display "Sample of tournament and year:"
list match_id tournament year win in 1/10

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
* Controls: team avg doubles rank, opponent rank, top-100 singles indicator,
*           team avg GS appearances (exp_mean) and its square
* Fixed effects: tournament×year interaction, round (stage_code)
* Standard errors: clustered by match (Tables 3–4); HC3 robust (Table 5, one obs/match)
* Sensitivity tables (3b, 4b, 5b) add teammate rank gap to verify stability

display ""
display "Diagnostic: Data check before regression"
count if !missing(win)
display "Non-missing win: " r(N)
display ""

* Replace missing stage_code with 0 for regression
replace stage_code = 0 if missing(stage_code)

* Section 5 depends on numeric fixed-effect variables.
capture confirm string variable tournament
if !_rc {
    encode tournament, generate(tourn_code)
    drop tournament
    rename tourn_code tournament
}

capture confirm numeric variable year
if _rc {
    destring year, replace force
}

capture confirm numeric variable stage_code
if _rc {
    destring stage_code, replace force
    replace stage_code = 0 if missing(stage_code)
}

* Impute exp_mean = 0 for first-time GS participants (no prior appearances in dataset)
replace exp_mean = 0 if missing(exp_mean)

* Generate experience squared (may already exist in CSV; drop first to be safe)
capture drop exp_mean_sq
generate double exp_mean_sq = exp_mean ^ 2

display "Obs in Grand Slams panel: " _N " (exp_mean imputed to 0 for first-timers)"

display "Creating ty (tournament×year grouping)..."
capture drop ty
quietly egen ty = group(tournament year)

count if !missing(win, ty, stage_code, ling_prox, rank_mean, opp_rank_mean, single_top100, exp_mean)
display "Complete cases for Table 3: " r(N)
if r(N) == 0 {
    display as error "No complete observations for Table 3. Check missingness below:"
    misstable summarize win ty stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean
    error 2000
}

display ""

log using "stata_homophily_results.txt", replace text

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 0: OBSERVATION BREAKDOWN
* ═══════════════════════════════════════════════════════════════════════════════

display ""
display "=== SECTION 0. OBSERVATION BREAKDOWN ==="
display "(from raw scraped data to regression sample)"
display ""
display "  Step 1  Raw dataset (scraped + merged, 2018-2025 excl. 2020):  1,840 matches"
display "  Step 2  Drop retirements / walkovers:                         -47 matches"
display "          (20 walkovers | 9 retired in set 1 | 18 retired in set 2+)"
display "  ─────────────────────────────────────────────────────────────────────────────"
display "          Clean match dataset:                                  1,793 matches"
display ""
display "  Step 3  Expand to team-level (2 obs per match):               3,586 obs"
display "  Step 4  Drop: ranking incomplete for >=1 player:              -14 matches  (-28 obs)"
display "  Step 5  Drop: nationality/language missing for >=1 player:    0 matches"
display "          (pipeline fully resolves via birthplace, surname lookup, Monaco fix,"
display "           and manual_nationality.csv — 1 NaN-surname slot in a walkover, dropped in Step 2)"
display "  ─────────────────────────────────────────────────────────────────────────────"
display "          After completeness filter (incl. Olympics):           3,558 obs | 1,779 matches"
display ""
display "  Step 6  Exclude Olympic matches (2021 Tokyo + 2024 Paris):    -63 matches  (-126 obs)"
display "  ─────────────────────────────────────────────────────────────────────────────"
display "          Grand Slams only (pre-experience filter):             3,432 obs | 1,716 matches"
display ""
display "  Step 7  Drop: GS experience missing for >=1 player:           -210 obs"
display "  ─────────────────────────────────────────────────────────────────────────────"

count
local n_panel = r(N)
local n_matches = `n_panel' / 2
display "          Grand Slams regression sample:  `n_panel' obs (approx `n_matches' matches)"
display ""

count if missing(win, ty, stage_code, same_country, rank_mean, opp_rank_mean, single_top100, exp_mean)
display "  Records with any missing regression variable: " r(N) " (should be 0)"
display ""

display "  Regression sub-samples:"
count if regular_tb == 1
local n_tb   = r(N)
local n_tb_m = `n_tb' / 2
display "    Table 4  regular 7-pt tiebreaks (sets 1-2):     `n_tb' obs | `n_tb_m' matches"
count if any_tb == 1
local n_any_tb = r(N)
display "    Table 4  any tiebreak (all types):              `n_any_tb' obs | " `n_any_tb'/2 " matches"
count if lost_set1 == 1
local n_adv = r(N)
display "    Table 5  adversity sample (lost set 1):         `n_adv' obs = `n_adv' matches"
display ""

* ─────────────────────────────────────────────────────────────────────────────
* TABLE 3: MATCH WIN — LOGIT (Main specification)
* Controls: rank_mean, opp_rank_mean, single_top100, exp_mean, exp_mean_sq
* No rank_gap; see Table 3b for sensitivity with rank_gap.
* ─────────────────────────────────────────────────────────────────────────────
display "=== TABLE 3. Match Win — Logit ==="
display "FE: tournament×year (ty) + stage_code | SE clustered by match | Grand Slams only"
display "Controls: rank_mean, opp_rank_mean, single_top100, exp_mean, exp_mean_sq"
display ""

display "--- 3a. Same Nationality ---"
logit win i.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store win_same_country

display ""
display "--- 3b. Same Official Language ---"
logit win i.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store win_same_language

display ""
display "--- 3c. Language Proximity (ethnic) ---"
logit win i.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store win_ling_prox

estimates table win_same_country win_same_language win_ling_prox, b se stats(N ll)

* ─────────────────────────────────────────────────────────────────────────────
* SECTION 6: HETEROGENEITY ANALYSIS
* =============================================================================
* Three sets of interactions, each applied to the match-win logit (team panel).
*   6a. Culture x GS experience (exp_mean)
*   6b. Culture x surface (clay / grass; hardcourt = baseline)
*   6c. Culture x Hofstede individualism score, team-level, demeaned (ic_team_dm)
*       Spec 1: C + C*IC + controls + FE   (no main IC effect)
*       Spec 2: C + IC + C*IC + controls + FE
*
* Sample: same 3,432 obs as Tables 3-4 (exp_mean imputed to 0 for first-timers).
* Note: ic_team_dm is loaded from team_gs_panel.csv (computed in merge_hofstede.ipynb).
* =============================================================================

display "=== SECTION 6. HETEROGENEITY ANALYSIS ==="
display ""

* Surface dummies from the encoded surface variable (Clay=1, Grass=2, Hard=3)
* Use decode to safely generate string-based dummies.
decode surface, generate(surface_str)
generate byte clay  = (surface_str == "Clay")
generate byte grass = (surface_str == "Grass")
drop surface_str

display "Surface distribution:"
tabstat clay grass, stats(sum mean) format(%9.0f)

* ic_team_dm: demeaned Hofstede individualism score (mean = 0 by construction)
summarize ic_team_dm
display "ic_team_dm mean (should be ~0): " r(mean)

* ---------------------------------------------------------------------------
* TABLE 6a: Culture x GS Experience
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 6a. Heterogeneity: Culture x GS Experience (exp_mean) ==="
display "Controls: rank_mean, opp_rank_mean, single_top100, exp_mean, exp_mean_sq"
display "FE: ty + stage_code | SE: clustered by match"
display ""

display "--- 6a-i. Same Nationality x exp_mean ---"
generate double int_sc_exp  = same_country  * exp_mean
generate double int_sl_exp  = same_language * exp_mean
generate double int_lp_exp  = ling_prox     * exp_mean

logit win i.ty i.stage_code same_country  int_sc_exp rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country int_sc_exp rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_exp_sc

display ""
display "--- 6a-ii. Same Language x exp_mean ---"
logit win i.ty i.stage_code same_language  int_sl_exp rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language int_sl_exp rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_exp_sl

display ""
display "--- 6a-iii. Language Proximity x exp_mean ---"
logit win i.ty i.stage_code ling_prox  int_lp_exp rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox int_lp_exp rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_exp_lp

estimates table het_exp_sc het_exp_sl het_exp_lp, b se stats(N ll)

* ---------------------------------------------------------------------------
* TABLE 6b: Culture x Surface
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 6b. Heterogeneity: Culture x Surface ==="
display "Hardcourt (AO + USO) = baseline; clay = Roland Garros; grass = Wimbledon."
display "Note: surface is collinear with tournament but the interaction C x surface"
display "  is identified by cross-surface variation in cultural homophily."
display "FE: ty + stage_code | SE: clustered by match"
display ""

generate double int_sc_clay  = same_country  * clay
generate double int_sc_grass = same_country  * grass
generate double int_sl_clay  = same_language * clay
generate double int_sl_grass = same_language * grass
generate double int_lp_clay  = ling_prox     * clay
generate double int_lp_grass = ling_prox     * grass

display "--- 6b-i. Same Nationality x surface ---"
logit win i.ty i.stage_code same_country clay grass int_sc_clay int_sc_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country clay grass int_sc_clay int_sc_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_surf_sc

display ""
display "--- 6b-ii. Same Language x surface ---"
logit win i.ty i.stage_code same_language clay grass int_sl_clay int_sl_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language clay grass int_sl_clay int_sl_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_surf_sl

display ""
display "--- 6b-iii. Language Proximity x surface ---"
logit win i.ty i.stage_code ling_prox clay grass int_lp_clay int_lp_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox clay grass int_lp_clay int_lp_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_surf_lp

estimates table het_surf_sc het_surf_sl het_surf_lp, b se stats(N ll)

* ---------------------------------------------------------------------------
* TABLE 6c: Culture x Hofstede Individualism (team-level, demeaned)
* Hypothesis: high-individualism teams rely less on in-group bonding
*             => negative interaction theta.
* Spec 1: Y = beta*C + theta*(C*IC_dm) + controls + FE
* Spec 2: Y = beta*C + gamma*IC_dm + theta*(C*IC_dm) + controls + FE
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 6c. Heterogeneity: Culture x Hofstede Individualism (demeaned) ==="
display "ic_team_dm = team-avg IDV score - sample mean (demean so beta is at mean IC)"
display "Spec 1: no main IC effect  |  Spec 2: includes main IC effect"
display "FE: ty + stage_code | SE: clustered by match"
display ""

generate double int_sc_ic = same_country  * ic_team_dm
generate double int_sl_ic = same_language * ic_team_dm
generate double int_lp_ic = ling_prox     * ic_team_dm

* --- Spec 1: Same Nationality ---
display "--- 6c-i. Same Nationality, Spec 1 (no main IC) ---"
logit win i.ty i.stage_code same_country int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_sc_s1

display ""
display "--- 6c-ii. Same Nationality, Spec 2 (main IC included) ---"
logit win i.ty i.stage_code same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_sc_s2

display ""
display "--- 6c-iii. Same Language, Spec 1 (no main IC) ---"
logit win i.ty i.stage_code same_language int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_sl_s1

display ""
display "--- 6c-iv. Same Language, Spec 2 (main IC included) ---"
logit win i.ty i.stage_code same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_sl_s2

display ""
display "--- 6c-v. Language Proximity, Spec 1 (no main IC) ---"
logit win i.ty i.stage_code ling_prox int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_lp_s1

display ""
display "--- 6c-vi. Language Proximity, Spec 2 (main IC included) ---"
logit win i.ty i.stage_code ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_lp_s2

estimates table het_ic_sc_s1 het_ic_sc_s2, b se stats(N ll)
estimates table het_ic_sl_s1 het_ic_sl_s2, b se stats(N ll)
estimates table het_ic_lp_s1 het_ic_lp_s2, b se stats(N ll)

display ""
display "Section 6 complete: Heterogeneity analysis."

* =============================================================================
* SECTION 7: TABLE 4 — TIEBREAK WIN
* =============================================================================
* Unit of observation: one team in one specific tiebreak (2 obs per tiebreak).
* Outcome: won_tb — did this team win this tiebreak?
* Source: data/atp/tiebreak_panel.csv
*   940 tiebreaks | 1,880 team-obs (from 1,517 complete match pairs in GS sample)
* FE: tournament x year (ty) + stage_code | SE: clustered by match
* =============================================================================

display ""
display "=== SECTION 7. TABLE 4: TIEBREAK WIN ==="
display "Loading tiebreak panel..."

clear
import delimited "data/atp/tiebreak_panel.csv", clear

* Surface dummies before encoding
generate byte clay  = (surface == "Clay")
generate byte grass = (surface == "Grass")

* Encode string variables
encode tournament, generate(tourn_code)
drop tournament
rename tourn_code tournament

encode tb_type, generate(tb_type_code)
drop tb_type
rename tb_type_code tb_type

destring _all, replace force
replace stage_code = 0 if missing(stage_code)
replace exp_mean = 0 if missing(exp_mean)
capture drop exp_mean_sq
generate double exp_mean_sq = exp_mean ^ 2
quietly egen ty = group(tournament year)

count
display "Tiebreak team-obs loaded: " r(N)
display ""

display "Breakdown by tiebreak type:"
tabulate tb_type

display ""
display "=== TABLE 4. Tiebreak Win — Logit (all tiebreak types) ==="
display "Outcome: won_tb | Unit: team x tiebreak | Sample: GS only (no Olympics)"
display "Controls: rank_mean, opp_rank_mean, single_top100, exp_mean, exp_mean_sq"
display "FE: tournament x year (ty) + stage_code | SE: clustered by match"
display ""

display "--- 4a. Same Nationality ---"
logit won_tb i.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store tbn_same_country

display ""
display "--- 4b. Same Official Language ---"
logit won_tb i.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store tbn_same_language

display ""
display "--- 4c. Language Proximity (ethnic) ---"
logit won_tb i.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store tbn_ling_prox

estimates table tbn_same_country tbn_same_language tbn_ling_prox, b se stats(N ll)


display ""
display "Section 7 complete: Tiebreak win regressions (new Table 4)."

log close

display ""
display "All sections complete. Results saved to stata_homophily_results.txt"


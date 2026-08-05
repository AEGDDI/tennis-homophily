

* homophily.do
* Tennis Doubles Homophily — Main Analysis
* Run this from the project root in STATA.
* Mirrors sections 1-5 from homophily_analysis.ipynb
*
* Data: Grand Slam doubles matches 2018–2025.
*       Note: Wimbledon 2020 was cancelled; AO/RG/USO 2020 are included.
*       Raw matches, retirements/walkovers, GS panel size: updated after pipeline re-run.
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

* ─────────────────────────────────────────────────────────────────────────────
* 2.1 MATCH FORMAT: BEST-OF-3 VS BEST-OF-5
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== Match Format by Tournament-Year ==="
display "Wimbledon 2018/2019/2021/2022 played best-of-5; all other tournament-years"
display "  (AO/RG/USO always, Wimbledon 2023+, both Olympics) are best-of-3."
display ""

capture confirm numeric variable winners_set4
if _rc {
    destring winners_set4 losers_set4 winners_set5 losers_set5, replace force
}

capture drop reaches_4 reaches_5
generate byte reaches_4 = !missing(winners_set4)
generate byte reaches_5 = !missing(winners_set5)

tabstat reaches_4 reaches_5, by(tournament) stats(sum) format(%9.0f)

count if reaches_4 == 1
local n_bo5_extended = r(N)
display ""
display "Matches that actually required a 4th or 5th set: " %4.0f `n_bo5_extended' " (the rest of Wimbledon 2018/19/21/22 finished in straight sets)"

* ─────────────────────────────────────────────────────────────────────────────
* 2.2 TIEBREAKS AND DECIDING-SET OUTCOMES BY SET (1-5)
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== Tiebreak / Deciding-Set Outcomes by Set Number ==="
display "Standard tiebreak = set ends 7-6/6-7 (7-pt breaker at 6-6); the only outcome ever"
display "  observed at AO/RG/USO. 12-12 breaker = deciding set reaches 12-12, then a breaker"
display "  decides it (e.g. 13-12). Advantage set (no breaker) = deciding set won by a"
display "  natural 2-game margin at 8+ games (e.g. 8-6, 11-9, 22-20). The latter two are"
display "  observed only at Wimbledon (pre-2023) and Olympics."
display ""

forvalues s = 1/5 {
    capture confirm variable winners_set`s'
    if !_rc {
        capture drop hi_s`s' lo_s`s' std_tb_s`s' breaker1212_s`s' advset_s`s'
        generate double hi_s`s' = max(winners_set`s', losers_set`s')
        generate double lo_s`s' = min(winners_set`s', losers_set`s')
        generate byte std_tb_s`s'      = (winners_set`s'==7 & losers_set`s'==6) | (winners_set`s'==6 & losers_set`s'==7)
        generate byte breaker1212_s`s' = (lo_s`s'==12 & (hi_s`s'-lo_s`s')==1) if !missing(winners_set`s')
        generate byte advset_s`s'      = (hi_s`s'>=8 & (hi_s`s'-lo_s`s')==2) if !missing(winners_set`s')

        count if !missing(winners_set`s')
        local n_played = r(N)
        count if std_tb_s`s' == 1
        local n_std = r(N)
        count if breaker1212_s`s' == 1
        local n_brk = r(N)
        count if advset_s`s' == 1
        local n_adv = r(N)
        display "  Set `s':  N reaching=" %5.0f `n_played' "   Standard 7-6 TB=" %4.0f `n_std' "   12-12 breaker=" %3.0f `n_brk' "   Advantage(no brk)=" %3.0f `n_adv'
    }
}

* Cross-check against the fully raw (pre-retirement-filter) data: the working
* dataset here has already had the sets-1-3-only retirement filter applied
* (top of this do-file), which silently drops match_id 903 and 927 -- both
* genuine 12-12-breaker deciders finishing 13-12 in set 3. They are real,
* complete matches, not retirements; see the Section 4 reconciliation note.
preserve
import excel using "data/atp/men_matches_with_ranks_cleaned.xlsx", sheet("players_list") firstrow clear
destring winners_set3 losers_set3, replace force
generate double _hi3 = max(winners_set3, losers_set3)
generate double _lo3 = min(winners_set3, losers_set3)
count if _lo3==12 & (_hi3-_lo3)==1
display ""
display "Cross-check on fully raw data (before any retirement filtering): " r(N) " true 12-12-breaker"
display "  deciders in set 3. The 2 missing from the table above (match_id 903, 927) are"
display "  dropped upstream by the retirement-filter bug -- see Section 4 reconciliation note."
restore

display ""
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

* ─────────────────────────────────────────────────────────────────────────────
* 4.1 RECONCILIATION: PRESSURE-OUTCOME COUNTS VS. REGRESSION SAMPLE
* One known gap remains in the current pipeline, left as-is pending further review:
*   The retirement filter (top of this do-file) only inspects sets 1-3.
*   match_id 903, 927 (2021 Wimbledon, valid 13-12 set-3 finish) are wrongly
*   EXCLUDED from this panel; match_id 504, 1204 (genuine set-4 retirements,
*   2019/2022 Wimbledon) are wrongly INCLUDED.
* (The previous gap -- sets 4-5 tiebreaks missing from tiebreak_panel.csv -- is
*  now fixed: tiebreak_panel.ipynb covers sets 1-5, see Section 7.)
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== Reconciliation: Pressure-Outcome Counts vs. Regression Sample ==="
foreach mid in 903 927 504 1204 {
    count if match_id == `mid'
    local present = (r(N) > 0)
    display "  match_id `mid': present in team_gs_panel.csv = " `present'
}
display ""
display "This discrepancy is left as-is in this pass pending further guidance."

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
display "  Step 1  Raw dataset (scraped + merged, 2018-2025; Wimbledon 2020 cancelled, AO/RG/USO 2020 included)"
display "  Step 2  Drop retirements / walkovers (see count below)"
display "  ─────────────────────────────────────────────────────────────────────────────"
display "          Clean match dataset: see count below"
display ""
display "  Step 3  Expand to team-level (2 obs per match)"
display "  Step 4  Drop: ranking incomplete for >=1 player"
display "  Step 5  Drop: nationality/language missing for >=1 player:    0 matches"
display "          (pipeline fully resolves via birthplace, surname lookup, Monaco fix,"
display "           and manual_nationality.csv)"
display "  ─────────────────────────────────────────────────────────────────────────────"
display "  Step 6  Exclude Olympic matches (2021 Tokyo + 2024 Paris)"
display "  ─────────────────────────────────────────────────────────────────────────────"
display "  Step 7  Drop: GS experience missing for >=1 player (imputed to 0 for first-timers)"
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
*   Table 6a. Culture x GS experience (exp_mean, demeaned), FE: ty + stage_code
*   Table 5.  Culture x surface (clay / grass vs. rest, two specs), no ty FE
*   Table 6.  Culture x Hofstede individualism score, demeaned (ic_team_dm),
*             Spec 2 only (C + IC + C*IC + controls + FE)
*
* Sample: same 3,744 obs as Tables 3-4 (exp_mean imputed to 0 for first-timers).
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
display "=== TABLE 6a. Heterogeneity: Culture x GS Experience (exp_mean, demeaned) ==="
display "Controls: rank_mean, opp_rank_mean, single_top100, exp_mean_dm, exp_mean_sq"
display "  (exp_mean_sq is the same non-demeaned quadratic control as in Table 3, kept"
display "  alongside exp_mean_dm per the blanket 'all controls as in Table 3/4' rule.)"
display "exp_mean_dm = exp_mean - sample mean, so C is evaluated at mean GS experience"
display "  (consistent with how ic_team_dm is demeaned for Table 6)."
display "FE: ty + stage_code | SE: clustered by match"
display ""

quietly summarize exp_mean
generate double exp_mean_dm = exp_mean - r(mean)
display "exp_mean_dm: mean GS appearances = " r(mean) "  SD = " r(sd)

display "--- 6a-i. Same Nationality x exp_mean ---"
generate double int_sc_exp  = same_country  * exp_mean_dm
generate double int_sl_exp  = same_language * exp_mean_dm
generate double int_lp_exp  = ling_prox     * exp_mean_dm

logit win i.ty i.stage_code same_country  int_sc_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_sq, cluster(match_id)
margins, dydx(same_country int_sc_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_sq)
estimates store het_exp_sc

display ""
display "--- 6a-ii. Same Language x exp_mean ---"
logit win i.ty i.stage_code same_language  int_sl_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_sq, cluster(match_id)
margins, dydx(same_language int_sl_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_sq)
estimates store het_exp_sl

display ""
display "--- 6a-iii. Language Proximity x exp_mean ---"
logit win i.ty i.stage_code ling_prox  int_lp_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox int_lp_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_sq)
estimates store het_exp_lp

estimates table het_exp_sc het_exp_sl het_exp_lp, b se stats(N ll)

* ---------------------------------------------------------------------------
* TABLE 5: Culture x Surface (was Table 6b)
* Spec (a): grass + C*grass, baseline = hardcourt + clay pooled
* Spec (b): clay + C*clay, baseline = hardcourt + grass pooled
* No tournament x year FE (round FE only) -- surface is a deterministic function
* of tournament, so dropping the tournament x year FE lets the surface main
* effects be estimated directly instead of being absorbed/collinear.
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 5. Heterogeneity: Culture x Surface ==="
display "Hardcourt = AO + USO; clay = Roland Garros; grass = Wimbledon."
display "Spec (a): grass + C x grass, baseline = hardcourt+clay pooled"
display "Spec (b): clay + C x clay, baseline = hardcourt+grass pooled"
display "FE: stage_code only (no tournament x year FE) | SE: clustered by match"
display "Controls: same as Table 3 -- rank_mean, opp_rank_mean, single_top100, exp_mean, exp_mean_sq"
display ""

generate double int_sc_clay  = same_country  * clay
generate double int_sc_grass = same_country  * grass
generate double int_sl_clay  = same_language * clay
generate double int_sl_grass = same_language * grass
generate double int_lp_clay  = ling_prox     * clay
generate double int_lp_grass = ling_prox     * grass

display "--- 5a. Same Nationality: Spec (a) grass vs. rest ---"
logit win i.stage_code same_country grass int_sc_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country grass int_sc_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_surf_sc_grass

display ""
display "--- 5a. Same Nationality: Spec (b) clay vs. rest ---"
logit win i.stage_code same_country clay int_sc_clay rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country clay int_sc_clay rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_surf_sc_clay

display ""
display "--- 5b. Same Language: Spec (a) grass vs. rest ---"
logit win i.stage_code same_language grass int_sl_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language grass int_sl_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_surf_sl_grass

display ""
display "--- 5b. Same Language: Spec (b) clay vs. rest ---"
logit win i.stage_code same_language clay int_sl_clay rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language clay int_sl_clay rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_surf_sl_clay

display ""
display "--- 5c. Language Proximity: Spec (a) grass vs. rest ---"
logit win i.stage_code ling_prox grass int_lp_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox grass int_lp_grass rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_surf_lp_grass

display ""
display "--- 5c. Language Proximity: Spec (b) clay vs. rest ---"
logit win i.stage_code ling_prox clay int_lp_clay rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox clay int_lp_clay rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_surf_lp_clay

display ""
display "--- Comparison: Spec (a) grass vs. Spec (b) clay ---"
estimates table het_surf_sc_grass het_surf_sc_clay, b se stats(N ll)
estimates table het_surf_sl_grass het_surf_sl_clay, b se stats(N ll)
estimates table het_surf_lp_grass het_surf_lp_clay, b se stats(N ll)

* ---------------------------------------------------------------------------
* TABLE 6: Culture x Hofstede Individualism (team-level, demeaned) (was Table 6c)
* Hypothesis: high-individualism teams rely less on in-group bonding
*             => negative interaction theta.
* Spec 2 only (includes main IC effect): Y = beta*C + gamma*IC_dm + theta*(C*IC_dm) + controls + FE
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 6. Heterogeneity: Culture x Hofstede Individualism (demeaned) ==="
display "ic_team_dm = team-avg IDV score - sample mean (demean so beta is at mean IC)"
display "Spec 2 only (includes main IC effect) -- the reported specification."
display "FE: ty + stage_code | SE: clustered by match"
display ""

generate double int_sc_ic = same_country  * ic_team_dm
generate double int_sl_ic = same_language * ic_team_dm
generate double int_lp_ic = ling_prox     * ic_team_dm

display "--- 6a. Same Nationality ---"
logit win i.ty i.stage_code same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_sc_s2

display ""
display "--- 6b. Same Language ---"
logit win i.ty i.stage_code same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_sl_s2

display ""
display "--- 6c. Language Proximity ---"
logit win i.ty i.stage_code ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_lp_s2

estimates table het_ic_sc_s2 het_ic_sl_s2 het_ic_lp_s2, b se stats(N ll)

* ---------------------------------------------------------------------------
* TABLE 6 ROBUSTNESS: excluding Tier-2 Hofstede-proxy countries
* Tier-2 proxies (moderate/weaker confidence, per data/gravity/hofstede.csv):
*   BLR BOL DOM GEO KAZ LBN TUN UKR UZB
* Drops any team-obs where either player's nationality relies on a Tier-2 proxy.
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 6 ROBUSTNESS. Dropping Tier-2 Hofstede-proxy countries ==="
display "Tier-2 proxies: BLR BOL DOM GEO KAZ LBN TUN UKR UZB"
display "FE: ty + stage_code | SE: clustered by match"
display ""

preserve
import excel using "data/atp/men_matches_with_ranks_cleaned.xlsx", sheet("players_list") firstrow clear
keep match_id winners_p1_iso3 winners_p2_iso3 losers_p1_iso3 losers_p2_iso3
duplicates drop match_id, force
tempfile iso_lookup
save `iso_lookup'
restore

merge m:1 match_id using `iso_lookup', keep(master match) nogen

local tier2 "BLR BOL DOM GEO KAZ LBN TUN UKR UZB"
generate byte tier2_w = 0
generate byte tier2_l = 0
foreach c of local tier2 {
    replace tier2_w = 1 if winners_p1_iso3 == "`c'" | winners_p2_iso3 == "`c'"
    replace tier2_l = 1 if losers_p1_iso3  == "`c'" | losers_p2_iso3  == "`c'"
}
generate byte tier2_flag = tier2_w if win == 1
replace tier2_flag = tier2_l if win == 0

count if tier2_flag == 1
display "Team-obs dropped for Tier-2 proxy robustness: " r(N)

display ""
display "--- 6a, ex-Tier2. Same Nationality ---"
logit win i.ty i.stage_code same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq if tier2_flag == 0, cluster(match_id)
margins, dydx(same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_sc_s2_ex2

display ""
display "--- 6b, ex-Tier2. Same Language ---"
logit win i.ty i.stage_code same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq if tier2_flag == 0, cluster(match_id)
margins, dydx(same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_sl_s2_ex2

display ""
display "--- 6c, ex-Tier2. Language Proximity ---"
logit win i.ty i.stage_code ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq if tier2_flag == 0, cluster(match_id)
margins, dydx(ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store het_ic_lp_s2_ex2

display ""
display "--- Comparison: full sample vs. ex-Tier2 ---"
estimates table het_ic_sc_s2 het_ic_sc_s2_ex2, b se stats(N ll)
estimates table het_ic_sl_s2 het_ic_sl_s2_ex2, b se stats(N ll)
estimates table het_ic_lp_s2 het_ic_lp_s2_ex2, b se stats(N ll)

display ""
display "Section 6 complete: Heterogeneity analysis."

* =============================================================================
* SECTION 7: TABLE 4 — TIEBREAK WIN
* =============================================================================
* Unit of observation: one team in one specific tiebreak (2 obs per tiebreak).
* Outcome: won_tb — did this team win this tiebreak?
* Source: data/atp/tiebreak_panel.csv (now covers sets 1-5; rebuilt via
*   code/cleaning/tiebreak_panel.ipynb)
* Main spec:        7pt standard tiebreaks only, any set (1-5)
* Robustness spec:  adds the 10pt / advantage-set match-tiebreak decider category
* Controls: same as Table 3 | FE: tournament x year (ty) + stage_code | SE: clustered by match
* =============================================================================

display ""
display "=== SECTION 7. TABLE 4: TIEBREAK WIN ==="
display "Loading tiebreak panel..."

clear
import delimited "data/atp/tiebreak_panel.csv", clear

* Surface dummies before encoding
generate byte clay  = (surface == "Clay")
generate byte grass = (surface == "Grass")

* Encode tournament; keep tb_type as a STRING so it can be used to split the
* main (7pt-only) sample from the robustness (7pt+10pt) sample below.
encode tournament, generate(tourn_code)
drop tournament
rename tourn_code tournament

destring stage_code same_country same_language ling_prox rank_mean opp_rank_mean ///
    rank_gap single_top100 exp_mean win pre_olympic olympic_period tb_set won_tb year, replace force
replace stage_code = 0 if missing(stage_code)
replace exp_mean = 0 if missing(exp_mean)
capture drop exp_mean_sq
generate double exp_mean_sq = exp_mean ^ 2
quietly egen ty = group(tournament year)

count
display "Tiebreak team-obs loaded (7pt + 10pt, all sets 1-5): " r(N)
display ""

display "Breakdown by tiebreak type and set:"
tabulate tb_type
tabulate tb_set tb_type

display ""
display "=== TABLE 4. Tiebreak Win — Logit ==="
display "Outcome: won_tb | Unit: team x tiebreak | Sample: GS only (no Olympics)"
display "Main spec: 7pt standard tiebreaks only, any set. Robustness: adds 10pt/advantage-set deciders."
display "Controls: rank_mean, opp_rank_mean, single_top100, exp_mean, exp_mean_sq (same as Table 3)"
display "FE: tournament x year (ty) + stage_code | SE: clustered by match"
display ""

display "--- 4a. Same Nationality: Main (7pt only) ---"
logit won_tb i.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq if tb_type=="7pt", cluster(match_id)
margins, dydx(same_country rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store tbn_same_country_main

display ""
display "--- 4a. Same Nationality: Robustness (7pt + 10pt) ---"
logit won_tb i.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_country rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store tbn_same_country_rob

display ""
display "--- 4b. Same Official Language: Main (7pt only) ---"
logit won_tb i.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq if tb_type=="7pt", cluster(match_id)
margins, dydx(same_language rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store tbn_same_language_main

display ""
display "--- 4b. Same Official Language: Robustness (7pt + 10pt) ---"
logit won_tb i.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(same_language rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store tbn_same_language_rob

display ""
display "--- 4c. Language Proximity: Main (7pt only) ---"
logit won_tb i.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq if tb_type=="7pt", cluster(match_id)
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store tbn_ling_prox_main

display ""
display "--- 4c. Language Proximity: Robustness (7pt + 10pt) ---"
logit won_tb i.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq, cluster(match_id)
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 exp_mean exp_mean_sq)
estimates store tbn_ling_prox_rob

display ""
display "--- Comparison: Main (7pt) vs. Robustness (7pt+10pt) ---"
estimates table tbn_same_country_main tbn_same_country_rob, b se stats(N ll)
estimates table tbn_same_language_main tbn_same_language_rob, b se stats(N ll)
estimates table tbn_ling_prox_main tbn_ling_prox_rob, b se stats(N ll)


display ""
display "Section 7 complete: Tiebreak win regressions (Table 4)."

log close

display ""
display "All sections complete. Results saved to stata_homophily_results.txt"


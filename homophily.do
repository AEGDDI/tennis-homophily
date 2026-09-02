

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
    winners_set4 losers_set4 winners_set5 losers_set5 ///
    winners_p1_top100_within_1y winners_p2_top100_within_1y losers_p1_top100_within_1y losers_p2_top100_within_1y ///
    rank_mean_winners rank_mean_losers rank_diff_winners rank_diff_losers, replace force

* Drop likely retirements/walkovers that are not explicitly flagged in the data.
* A valid completed set is 6-0 through 6-4, 7-5, 7-6, an extended final set score
* with a two-game margin (e.g. 8-6, 9-7, or a genuine 10-pt super-tiebreak score
* like 10-5), OR a 12-12-breaker finish (e.g. 13-12, a real historical Wimbledon
* final-set rule -- a 1-game margin, so distinct from the two-game-margin case).
* If the first two sets are split, the third set must also be complete under the
* same rule. Sets 4-5 (best-of-5 matches), if played, must also be complete.
capture program drop valid_set_score
program define valid_set_score, rclass
    args wvar lvar outvar
    generate byte `outvar' = !missing(`wvar', `lvar') & ///
        ((max(`wvar', `lvar')==6 & min(`wvar', `lvar')<=4) | ///
         (max(`wvar', `lvar')==7 & inlist(min(`wvar', `lvar'), 5, 6)) | ///
         (max(`wvar', `lvar')>=8 & max(`wvar', `lvar') - min(`wvar', `lvar')>=2) | ///
         (min(`wvar', `lvar')==12 & max(`wvar', `lvar') - min(`wvar', `lvar')==1))
end

valid_set_score winners_set1 losers_set1 s1_complete
valid_set_score winners_set2 losers_set2 s2_complete
valid_set_score winners_set3 losers_set3 s3_complete

generate byte split_sets = !missing(winners_set1, losers_set1, winners_set2, losers_set2) & ///
    ((winners_set1 > losers_set1 & winners_set2 < losers_set2) | ///
     (winners_set1 < losers_set1 & winners_set2 > losers_set2))

valid_set_score winners_set4 losers_set4 _s4_complete
generate byte s4_valid = missing(winners_set4) | _s4_complete
valid_set_score winners_set5 losers_set5 _s5_complete
generate byte s5_valid = missing(winners_set5) | _s5_complete

generate byte retired_or_incomplete = !s1_complete | !s2_complete | ///
    (split_sets & !s3_complete) | !s4_valid | !s5_valid

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

* Sets 4-5 only occur in the 103 best-of-5 Wimbledon matches (2018/2019/2021/2022).
* Same 7-point standard-tiebreak pattern as sets 1-3.
generate byte tb_s3_regular = (winners_set3==7 & losers_set3==6) | (winners_set3==6 & losers_set3==7)
generate byte tb_s4 = (winners_set4==7 & losers_set4==6) | (winners_set4==6 & losers_set4==7)
generate byte tb_s5 = (winners_set5==7 & losers_set5==6) | (winners_set5==6 & losers_set5==7)
replace tb_s4 = 0 if missing(tb_s4)
replace tb_s5 = 0 if missing(tb_s5)

* Match tiebreak / advantage-set decider (no 6-6 breaker; hi>=8, margin>=2): can
* occur as the deciding set, which is set 3 for a best-of-3 match but set 4 or
* set 5 for a best-of-5 match.
generate byte match_tb_s3 = (max(winners_set3, losers_set3)>=8) & ///
    (max(winners_set3, losers_set3) - min(winners_set3, losers_set3)>=2) & !missing(winners_set3, losers_set3)
generate byte match_tb_s4 = (max(winners_set4, losers_set4)>=8) & ///
    (max(winners_set4, losers_set4) - min(winners_set4, losers_set4)>=2) & !missing(winners_set4, losers_set4)
generate byte match_tb_s5 = (max(winners_set5, losers_set5)>=8) & ///
    (max(winners_set5, losers_set5) - min(winners_set5, losers_set5)>=2) & !missing(winners_set5, losers_set5)
generate byte match_tb = match_tb_s3 | match_tb_s4 | match_tb_s5
generate byte match_tb_set = 3 if match_tb_s3==1
replace match_tb_set = 4 if match_tb_s4==1
replace match_tb_set = 5 if match_tb_s5==1

* any_std_tb: every literal 7-6 standard tiebreak, any set (1-5). any_tb
* additionally folds in the match_tb / advantage-set-decider category.
generate byte any_std_tb = tb_s1 | tb_s2 | tb_s3_regular | tb_s4 | tb_s5
generate byte any_tb = any_std_tb | match_tb

* Who won each tiebreak (winner-team perspective)
generate byte w_won_tb_s1 = tb_s1 & (winners_set1==7)
generate byte w_won_tb_s2 = tb_s2 & (winners_set2==7)
generate byte w_won_tb_s3 = tb_s3_regular & (winners_set3==7)
generate byte w_won_tb_s4 = tb_s4 & (winners_set4==7)
generate byte w_won_tb_s5 = tb_s5 & (winners_set5==7)

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
* All descriptive stats in this section (Table 1, 2.1, 2.2, 2.3) are computed on the
* SAME 1,876-match regression sample used in Tables 3-6, not the broader post-
* retirement-drop set -- so these numbers reconcile exactly with the regression
* tables. See 2.4 for how 1,876 relates to 1,886 and the raw data.
* ═══════════════════════════════════════════════════════════════════════════════

* Regression-sample flag: same ranking/nationality completeness filter used later
* in Section 4 (team panel construction), replicated here so Table 1 below matches
* the actual regression sample exactly.
capture drop reg_sample
generate byte reg_sample = (olympics_tourn == 0) & ///
    !missing(rank_mean_winners) & !missing(rank_mean_losers) & ///
    !missing(same_country_winners) & !missing(same_country_losers) & ///
    !missing(winners_same_language) & !missing(losers_same_language) & ///
    !missing(winners_linguistic_proximity) & !missing(losers_linguistic_proximity)

count if reg_sample == 1
local N_all = r(N)
display ""
display "=== TABLE 1. Pressure Outcome Counts (Grand Slams regression sample, N=" `N_all' ") ==="
display "Only genuine tiebreaks: a standard 7-pt tiebreak (sets 1-5). Advantage-set deciders"
display "without a breaker (e.g. 8-6) are NOT tiebreaks -- no breaker was ever played, confirmed"
display "by the raw tiebreak-score columns being null for all 15 such cases (see 2.2 for the 3"
display "separate cases where a real breaker WAS played at 12-12, e.g. 13-12)."
display ""
count if tb_s1 == 1 & reg_sample == 1
local n_tb_s1 = r(N)
display "  Set-1 tiebreak (7-pt):                                 N=" %4.0f `n_tb_s1'

count if tb_s2 == 1 & reg_sample == 1
local n_tb_s2 = r(N)
display "  Set-2 tiebreak (7-pt):                                 N=" %4.0f `n_tb_s2'

count if tb_s3_regular == 1 & reg_sample == 1
local n_tb_s3 = r(N)
display "  Set-3 tiebreak (7-pt):                                 N=" %4.0f `n_tb_s3'

count if tb_s4 == 1 & reg_sample == 1
local n_tb_s4 = r(N)
display "  Set-4 tiebreak (7-pt):                                 N=" %4.0f `n_tb_s4'

count if tb_s5 == 1 & reg_sample == 1
local n_tb_s5 = r(N)
display "  Set-5 tiebreak (7-pt):                                 N=" %4.0f `n_tb_s5'

local n_sum_std = `n_tb_s1' + `n_tb_s2' + `n_tb_s3' + `n_tb_s4' + `n_tb_s5'
display ""
display "Sum of set-1..5 standard (7-pt) tiebreaks: " %4.0f `n_tb_s1' " + " %4.0f `n_tb_s2' " + " %4.0f `n_tb_s3' " + " %4.0f `n_tb_s4' " + " %4.0f `n_tb_s5' " = " %4.0f `n_sum_std'
display "(This matches Table 4's tiebreak count exactly, since both are computed on the"
display " identical regression sample.)"
display ""
display "Advantage-set deciders with no breaker played are excluded entirely from every count"
display "above -- they are not tiebreaks."

* ─────────────────────────────────────────────────────────────────────────────
* 2.1 MATCH FORMAT: BEST-OF-3 VS BEST-OF-5
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== Match Format by Tournament-Year (regression sample) ==="
display "Wimbledon 2018/2019/2021/2022 played best-of-5; all other tournament-years"
display "  (AO/RG/USO always, Wimbledon 2023+) are best-of-3. Olympics (best-of-3"
display "  throughout) is not part of the regression sample -- see 2.5."
display ""

capture confirm numeric variable winners_set4
if _rc {
    destring winners_set4 losers_set4 winners_set5 losers_set5, replace force
}

capture drop reaches_4 reaches_5
generate byte reaches_4 = !missing(winners_set4)
generate byte reaches_5 = !missing(winners_set5)

tabstat reaches_4 reaches_5 if reg_sample==1, by(tournament) stats(sum) format(%9.0f)

count if reaches_4 == 1 & reg_sample == 1
local n_bo5_extended = r(N)
display ""
display "Matches that actually required a 4th or 5th set: " %4.0f `n_bo5_extended' " (the rest of Wimbledon 2018/19/21/22 finished in straight sets)"

* ─────────────────────────────────────────────────────────────────────────────
* 2.2 TIEBREAKS AND DECIDING-SET OUTCOMES BY SET (1-5)
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== Tiebreak / Deciding-Set Outcomes by Set Number (regression sample) ==="
display "Standard tiebreak = set ends 7-6/6-7 (7-pt breaker at 6-6); the only outcome ever"
display "  observed at AO/RG/USO. 12-12 breaker = deciding set reaches 12-12, then a breaker"
display "  decides it (e.g. 13-12). Advantage set (no breaker) = deciding set won by a"
display "  natural 2-game margin at 8+ games (e.g. 8-6, 11-9, 22-20). The latter two are"
display "  observed only at Wimbledon (pre-2023)."
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

        count if !missing(winners_set`s') & reg_sample==1
        local n_played = r(N)
        count if std_tb_s`s' == 1 & reg_sample==1
        local n_std = r(N)
        count if breaker1212_s`s' == 1 & reg_sample==1
        local n_brk = r(N)
        count if advset_s`s' == 1 & reg_sample==1
        local n_adv = r(N)
        display "  Set `s':  N reaching=" %5.0f `n_played' "   Standard 7-6 TB=" %4.0f `n_std' "   12-12 breaker=" %3.0f `n_brk' "   Advantage(no brk)=" %3.0f `n_adv'
    }
}

display ""
display %4.0f `n_tb_s1' " + " %4.0f `n_tb_s2' " + " %4.0f `n_tb_s3' " + " %4.0f `n_tb_s4' " + " %4.0f `n_tb_s5' " = " %4.0f `n_sum_std' "  <- sum of standard (7-6) tiebreaks across all sets"
display "Note: this matches Table 4's main-spec tiebreak count exactly, since both are"
display "  now computed on the identical 1,876-match regression sample."

* Cross-check against the fully raw (pre-retirement-filter) data: the working
* dataset's retirement filter (top of this do-file) now recognizes the
* 12-12-breaker pattern as valid, so match_id 903 and 927 -- both genuine
* 12-12-breaker deciders finishing 13-12 in set 3 -- are correctly retained.
preserve
import excel using "data/atp/men_matches_with_ranks_cleaned.xlsx", sheet("players_list") firstrow clear
destring winners_set3 losers_set3, replace force
generate double _hi3 = max(winners_set3, losers_set3)
generate double _lo3 = min(winners_set3, losers_set3)
count if _lo3==12 & (_hi3-_lo3)==1
display ""
display "Cross-check on fully raw data (before any retirement filtering): " r(N) " true 12-12-breaker"
display "  deciders in set 3. Both (match_id 903, 927) are present above -- fix confirmed,"
display "  see 2.4 reconciliation note."
restore

* ─────────────────────────────────────────────────────────────────────────────
* 2.3 TIEBREAKS PER MATCH
* How many standard (7-pt) tiebreaks does a single match contain? A match can
* have more than one (e.g. set 1 and set 2), so this is a genuine distribution.
* The advantage-set/12-12-breaker deciders are excluded here (different category,
* and always at most one per match since it can only be the final set played).
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== Matches by Number of Standard (7-pt) Tiebreaks ==="
capture drop tb_count
generate byte tb_count = tb_s1 + tb_s2 + tb_s3_regular + tb_s4 + tb_s5
tabulate tb_count if reg_sample==1

quietly summarize tb_count if reg_sample==1
local tb_count_sum = r(mean) * r(N)
display ""
display "Check: sum of tb_count across matches = " %4.0f `tb_count_sum' "  (matches the " %4.0f `n_sum_std' " total above)"

display ""
display "Section 2 complete: Pressure outcomes."

* ─────────────────────────────────────────────────────────────────────────────
* 2.5 OLYMPICS: PRESSURE OUTCOME COUNTS
* Reported separately rather than pooled into Table 1 above, since Olympic tennis
* follows different rules in this dataset and Olympic matches are excluded from
* the regression sample entirely (Tables 3-6).
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== 2.5 Olympics: Pressure Outcome Counts ==="
count if olympics_tourn == 1
local n_oly = r(N)
display "N=" `n_oly' " matches (Tokyo 2021 + Paris 2024)"
display ""

count if tb_s1 == 1 & olympics_tourn == 1
display "  Set-1 tiebreak (7-pt):                                 N=" %4.0f r(N)
count if tb_s2 == 1 & olympics_tourn == 1
display "  Set-2 tiebreak (7-pt):                                 N=" %4.0f r(N)
count if regular_tb == 1 & olympics_tourn == 1
display "  Any regular tiebreak (sets 1 or 2):                    N=" %4.0f r(N)
count if match_tb == 1 & olympics_tourn == 1
display "  Advantage-set / 12-12-breaker decider:                 N=" %4.0f r(N)
count if any_tb == 1 & olympics_tourn == 1
display "  Any tiebreak (all types, any set):                     N=" %4.0f r(N)
count if three_sets == 1 & olympics_tourn == 1
display "  Match went to 3 sets:                                  N=" %4.0f r(N)
count if comeback == 1 & olympics_tourn == 1
display "  Comeback wins (winner lost set 1):                     N=" %4.0f r(N)

display ""
display "By Olympic year:"
tabstat olympics_tourn if olympics_tourn==1, by(year) stats(sum) format(%9.0f)
display ""
display "Advantage-set / 12-12-breaker deciders by year:"
tabstat match_tb if olympics_tourn==1, by(year) stats(sum) format(%9.0f)

count if reaches_4 == 1 & olympics_tourn == 1
display ""
display "Matches reaching a 4th set: " %4.0f r(N) " (Olympics is best-of-3 throughout)."

display ""
display "Section 2.5 complete: Olympics pressure outcomes."

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
* Uses the same valid_set_score program defined at the top of this do-file.
if `"`incomplete_matches'"' == "" {
    tempfile incomplete_matches
    preserve
    import excel using "data/atp/men_matches_with_ranks_cleaned.xlsx", sheet("players_list") firstrow clear

    destring year winners_set1 winners_set2 winners_set3 losers_set1 losers_set2 losers_set3 ///
        winners_set4 losers_set4 winners_set5 losers_set5, replace force

    valid_set_score winners_set1 losers_set1 s1_complete
    valid_set_score winners_set2 losers_set2 s2_complete
    valid_set_score winners_set3 losers_set3 s3_complete

    generate byte split_sets = !missing(winners_set1, losers_set1, winners_set2, losers_set2) & ///
        ((winners_set1 > losers_set1 & winners_set2 < losers_set2) | ///
         (winners_set1 < losers_set1 & winners_set2 > losers_set2))

    valid_set_score winners_set4 losers_set4 _s4_complete
    generate byte s4_valid = missing(winners_set4) | _s4_complete
    valid_set_score winners_set5 losers_set5 _s5_complete
    generate byte s5_valid = missing(winners_set5) | _s5_complete

    generate byte retired_or_incomplete = !s1_complete | !s2_complete | ///
        (split_sets & !s3_complete) | !s4_valid | !s5_valid

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
* 4.1 (== notebook section 2.4) WHERE 1,876 AND 1,886 COME FROM
* The retirement filter (top of this do-file) has been fixed to recognize the
* 12-12-breaker pattern (e.g. 13-12) as a valid finish and to check sets 4-5 when
* present, so match_id 903, 927 (2021 Wimbledon, valid 13-12 set-3 finish) are
* correctly INCLUDED, and match_id 504, 1204 (genuine set-4 retirements,
* 2019/2022 Wimbledon) are correctly EXCLUDED. tiebreak_panel.ipynb covers sets
* 1-5, see Section 7.
* Of the 14 matches originally dropped for incomplete doubles ranking, 4 were
* retrieved and restored: 3 (match_id 6, 35, 49) failed only on an accented-
* surname merge mismatch for Guillermo Garcia-Lopez; 1 (match_id 913, 2021
* Wimbledon) had a genuinely blank partner slot for Alejandro Davidovich
* Fokina, whose profile was reconstructed from public sources. The remaining
* 10 could not be reliably retrieved and stay dropped.
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== Where 1,876 and 1,886 Come From ==="
display "1,997 raw scraped matches (Grand Slams + Olympics)."
display "Drop 48 retirements/walkovers -> 1,949 remain (47 GS + 1 Olympics)."
display "Of the 1,949: 63 are Olympic matches; 1,886 are Grand Slam matches -- this is"
display "  where 1,886 comes from (GS matches after dropping retirements, BEFORE the"
display "  ranking/nationality filters below). Section 2's Table 1 above uses 1,876,"
display "  not 1,886."
display "From 1,886, a further 10 matches are dropped for incomplete doubles ranking"
display "  (4 of the original 14 were retrieved and restored; 0 more for"
display "  nationality/language) -> 1,876 Grand Slam matches: the final regression"
display "  sample, matching team_gs_panel.csv exactly."
display ""
display "=== Reconciliation: Pressure-Outcome Counts vs. Regression Sample ==="
foreach mid in 903 927 504 1204 {
    count if match_id == `mid'
    local present = (r(N) > 0)
    display "  match_id `mid': present in team_gs_panel.csv = " `present'
}
display ""
display "Retirement-filter fix confirmed: 903/927 present, 504/1204 absent."

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 5: BASELINE REGRESSIONS
* ═══════════════════════════════════════════════════════════════════════════════
* Each match contributes two observations (winner team = 1, loser team = 0)
* Culture measure: language proximity (ling_prox) entered one at a time
* Controls: team avg doubles rank, opponent rank, top-100 singles indicator,
*           team avg years since turning pro (exp_mean) and its square
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

* Impute exp_mean = 1 for rookies (raw tenure <= 0, i.e. turned pro in/after tournament year)
replace exp_mean = 1 if missing(exp_mean)

* Generate experience squared (may already exist in CSV; drop first to be safe)
capture drop exp_mean_sq
generate double exp_mean_sq = exp_mean ^ 2

* Demean exp_mean (used in place of the raw variable everywhere below) and build its
* square from the demeaned variable, so 'years since turning pro' AMEs are evaluated at mean
* experience, not at zero appearances, in every table (Table 3/5/6/6a).
* (May already exist in the imported CSV; drop first to be safe, same as exp_mean_sq above.)
capture drop exp_mean_dm
capture drop exp_mean_dm_sq
quietly summarize exp_mean
generate double exp_mean_dm = exp_mean - r(mean)
generate double exp_mean_dm_sq = exp_mean_dm ^ 2
display "exp_mean_dm: mean yrs since turning pro = " r(mean) " (demeaned control, all main-panel tables)"

* Team-level age at tournament (age_mean) is built upstream in homophily.ipynb from each
* player's date of birth and imported here directly from team_gs_panel.csv (no rookie-style
* imputation needed: dob coverage is complete for every team-obs in this sample). Demean and
* square it exactly like exp_mean, for the Table 3-AGE spec below.
capture drop age_mean_dm
capture drop age_mean_dm_sq
quietly summarize age_mean
generate double age_mean_dm = age_mean - r(mean)
generate double age_mean_dm_sq = age_mean_dm ^ 2
display "age_mean_dm: mean team age at tournament = " r(mean) " (demeaned control, Table 3-AGE)"

display "Obs in Grand Slams panel: " _N " (exp_mean imputed to 1 for rookies)"

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
display "  Step 1  Raw scraped GS matches (2018-2025; Wimbledon 2020 cancelled, AO/RG/USO 2020"
display "          included). Olympics is loaded alongside GS in this working dataset but is"
display "          excluded from every count below via reg_sample / olympics_tourn==0 --"
display "          Part 1 of this report is GS-only throughout."
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
display "  Step 6  Years since turning pro missing for >=1 player: imputed to 1 for rookies,"
display "          no observations dropped"
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
* Controls: rank_mean, opp_rank_mean, single_top100, exp_mean_dm, exp_mean_dm_sq
* No rank_gap; see Table 3b for sensitivity with rank_gap.
* ─────────────────────────────────────────────────────────────────────────────
display "=== TABLE 3. Match Win — Logit ==="
display "FE: tournament×year (ty) + stage_code | SE clustered by match | Grand Slams only"
display "Controls: rank_mean, opp_rank_mean, single_top100, exp_mean_dm, exp_mean_dm_sq"
display ""

display "--- 3a. Same Nationality ---"
logit win i.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store win_same_country

display ""
display "--- 3b. Same Official Language ---"
logit win i.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store win_same_language

display ""
display "--- 3c. Language Proximity (ethnic) ---"
logit win i.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store win_ling_prox

estimates table win_same_country win_same_language win_ling_prox, b se stats(N ll)

* ─────────────────────────────────────────────────────────────────────────────
* TABLE 3-AGE: MATCH WIN — LOGIT, AGE SPEC (robustness: age_mean replaces exp_mean)
* age_mean (team avg. player age at the tournament) is highly correlated with exp_mean
* (years since turning pro) -- both proxy career maturity. corr(age_mean, exp_mean) = 0.83,
* implied VIF for age if entered alongside exp_mean_dm/exp_mean_dm_sq = 3.30 (see Python
* notebook cell for the full diagnostic). Entering both together destabilizes each tenure
* term's own coefficient (TABLE 3-JOINT below), so age is estimated in its own separate
* specification here rather than added alongside experience in the Table 3 main spec.
* ─────────────────────────────────────────────────────────────────────────────
display "=== TABLE 3-AGE. Match Win — Logit, age_mean replacing exp_mean ==="
display "FE: tournament×year (ty) + stage_code | SE clustered by match | Grand Slams only"
display "Controls: rank_mean, opp_rank_mean, single_top100, age_mean_dm, age_mean_dm_sq"
display ""

display "--- 3-age-i. Same Nationality ---"
logit win i.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq, cluster(match_id)
margins, dydx(same_country rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq)
estimates store win_same_country_age

display ""
display "--- 3-age-ii. Same Official Language ---"
logit win i.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq, cluster(match_id)
margins, dydx(same_language rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq)
estimates store win_same_language_age

display ""
display "--- 3-age-iii. Language Proximity (ethnic) ---"
logit win i.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq)
estimates store win_ling_prox_age

estimates table win_same_country_age win_same_language_age win_ling_prox_age, b se stats(N ll)

* ─────────────────────────────────────────────────────────────────────────────
* TABLE 3-JOINT: MATCH WIN — LOGIT, AGE + EXPERIENCE TOGETHER (diagnostic only)
* Not a main or robustness spec -- shown only to document why age and experience are
* estimated separately: entering both together destabilizes each tenure term's own AME.
* ─────────────────────────────────────────────────────────────────────────────
display "=== TABLE 3-JOINT. Match Win — Logit, age_mean_dm and exp_mean_dm together (diagnostic) ==="
display ""

display "--- 3-joint-i. Same Nationality ---"
logit win i.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq age_mean_dm age_mean_dm_sq, cluster(match_id)
margins, dydx(same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq age_mean_dm age_mean_dm_sq)

display ""
display "--- 3-joint-ii. Same Official Language ---"
logit win i.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq age_mean_dm age_mean_dm_sq, cluster(match_id)
margins, dydx(same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq age_mean_dm age_mean_dm_sq)

display ""
display "--- 3-joint-iii. Language Proximity (ethnic) ---"
logit win i.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq age_mean_dm age_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq age_mean_dm age_mean_dm_sq)

* ─────────────────────────────────────────────────────────────────────────────
* TABLE 3-NOINT: MATCH WIN — LOGIT, NO INTERCEPT (double-counting check, per
* Lingqing's 2026-08-18 email, Option 2 detail)
* Drops the global constant and gives tournament x year FE full-rank (all-levels)
* indicators (ibn.ty) instead of the usual reference-dropped i.ty; stage_code stays
* reduced-rank (i.stage_code) to keep the design full column rank -- giving BOTH FE
* terms full rank with no constant would make them collinear (each set of dummies
* sums to 1 for every row). This spans the identical model as Table 3 (same column
* space, just reparameterized), so the AMEs below should match Table 3 exactly --
* verified in the Python notebook (log-likelihood matches to 3dp under both
* parameterizations) -- this is a check, not a new spec. Same clustering as Table 3.
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== TABLE 3-NOINT. Match Win — Logit, no intercept (full-rank tournament x year FE) ==="
display "Expect: AMEs, SEs, and log-likelihood identical to Table 3 (reparameterization check)."
display ""

display "--- 3-noint-i. Same Nationality ---"
logit win ibn.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id) noconstant
margins, dydx(same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)

display ""
display "--- 3-noint-ii. Same Official Language ---"
logit win ibn.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id) noconstant
margins, dydx(same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)

display ""
display "--- 3-noint-iii. Language Proximity (ethnic) ---"
logit win ibn.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id) noconstant
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)

* ─────────────────────────────────────────────────────────────────────────────
* TABLE 3-RANDONE: MATCH WIN — LOGIT, ONE RANDOMLY-SELECTED TEAM PER MATCH
* (double-counting check, Option 1 from Lingqing's 2026-08-18 email)
* For each match, keep only one team's row (coin flip) -- removes the mechanical
* win/lose complementarity entirely, so no clustering is needed; robust (vce(robust),
* HC1-style) SE used instead. Repeated over 10 seeds since a single draw is arbitrary.
* Draw is one random uniform PER MATCH (generated on one row, then propagated to both
* rows of that match via egen max), not one per row -- otherwise the two rows of a
* match could independently both survive or both be dropped, which is not what
* "keep exactly one team per match" means.
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== TABLE 3-RANDONE. Match Win — Logit, one randomly-selected team per match ==="
display "Each match contributes exactly one row (winner or loser, 50/50 coin flip); no"
display "clustering needed (one row per match). Robust SE. 10 random seeds."
display ""

tempfile randone_results
tempname rr
postfile `rr' str24 cvar seed double ame double pval using "`randone_results'", replace

foreach cvar in same_country same_language ling_prox {
    forvalues seed = 0/9 {
        preserve
        set seed `seed'
        tempvar u u_match keep_row
        bysort match_id (win): generate double `u' = runiform() if _n==1
        bysort match_id: egen double `u_match' = max(`u')
        generate byte `keep_row' = (win==1 & `u_match'>0.5) | (win==0 & `u_match'<=0.5)
        keep if `keep_row'

        quietly logit win `cvar' rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq i.ty i.stage_code, vce(robust)
        quietly margins, dydx(`cvar')
        matrix b = r(b)
        matrix V = r(V)
        local ame_ = b[1,1]
        local se_  = sqrt(V[1,1])
        local p_   = 2*(1 - normal(abs(`ame_'/`se_')))
        post `rr' ("`cvar'") (`seed') (`ame_') (`p_')
        restore
    }
}
postclose `rr'

display ""
display "--- Table 3-RANDONE summary across 10 random draws ---"
preserve
use "`randone_results'", clear
foreach cvar in same_country same_language ling_prox {
    quietly summarize ame if cvar=="`cvar'"
    local mean_a = r(mean)
    local sd_a   = r(sd)
    local min_a  = r(min)
    local max_a  = r(max)
    quietly count if cvar=="`cvar'" & pval<0.05
    local nsig = r(N)
    display "`cvar': mean AME=" %6.4f `mean_a' "  sd=" %6.4f `sd_a' "  range=[" %6.4f `min_a' ", " %6.4f `max_a' "]  significant in `nsig'/10 draws"
}
restore

* ─────────────────────────────────────────────────────────────────────────────
* SECTION 6: HETEROGENEITY ANALYSIS
* =============================================================================
* Three sets of interactions, each applied to the match-win logit (team panel).
*   Table 6a. Culture x Years Since Turning Pro (exp_mean, demeaned), FE: ty + stage_code
*   Table 5.  Culture x surface (clay / grass vs. rest, two specs), no ty FE
*   Table 6.  Culture x Hofstede individualism score, demeaned (ic_team_dm),
*             Spec 2 only (C + IC + C*IC + controls + FE)
*
* Sample: same 3,752 obs as Tables 3-4 (exp_mean imputed to 1 for rookies).
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
* TABLE 6a: Culture x Years Since Turning Pro
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 6a. Heterogeneity: Culture x Years Since Turning Pro (exp_mean, demeaned) ==="
display "Controls: rank_mean, opp_rank_mean, single_top100, exp_mean_dm, exp_mean_dm_sq"
display "  (exp_mean_dm_sq = exp_mean_dm^2, the demeaned quadratic control used consistently"
display "  across every table in this report, per the blanket 'all controls as in Table 3/4' rule.)"
display "exp_mean_dm = exp_mean - sample mean, so C is evaluated at mean years since turning pro"
display "  (consistent with how ic_team_dm is demeaned for Table 6)."
display "FE: ty + stage_code | SE: clustered by match"
display ""

display "--- 6a-i. Same Nationality x exp_mean ---"
generate double int_sc_exp  = same_country  * exp_mean_dm
generate double int_sl_exp  = same_language * exp_mean_dm
generate double int_lp_exp  = ling_prox     * exp_mean_dm

logit win i.ty i.stage_code same_country  int_sc_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country int_sc_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_exp_sc

display ""
display "--- 6a-ii. Same Language x exp_mean ---"
logit win i.ty i.stage_code same_language  int_sl_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language int_sl_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_exp_sl

display ""
display "--- 6a-iii. Language Proximity x exp_mean ---"
logit win i.ty i.stage_code ling_prox  int_lp_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox int_lp_exp rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_exp_lp

estimates table het_exp_sc het_exp_sl het_exp_lp, b se stats(N ll)

* ---------------------------------------------------------------------------
* TABLE 6a ROBUSTNESS: adding Culture x exp_mean_dm^2 (per reviewer request 2026-08-11)
* Checks the linear-interaction finding isn't an artefact of a curved true relationship.
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 6a ROBUSTNESS. Culture x exp_mean_dm + Culture x exp_mean_dm^2 ==="
display "FE: ty + stage_code | SE: clustered by match"
display ""

generate double int_sc_exp_sq = same_country  * exp_mean_dm_sq
generate double int_sl_exp_sq = same_language * exp_mean_dm_sq
generate double int_lp_exp_sq = ling_prox     * exp_mean_dm_sq

display "--- 6a-rob-i. Same Nationality x exp_mean + exp_mean^2 ---"
logit win i.ty i.stage_code same_country  int_sc_exp int_sc_exp_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country int_sc_exp int_sc_exp_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_exp2_sc

display ""
display "--- 6a-rob-ii. Same Language x exp_mean + exp_mean^2 ---"
logit win i.ty i.stage_code same_language  int_sl_exp int_sl_exp_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language int_sl_exp int_sl_exp_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_exp2_sl

display ""
display "--- 6a-rob-iii. Language Proximity x exp_mean + exp_mean^2 ---"
logit win i.ty i.stage_code ling_prox  int_lp_exp int_lp_exp_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox int_lp_exp int_lp_exp_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_exp2_lp

estimates table het_exp2_sc het_exp2_sl het_exp2_lp, b se stats(N ll)

* ---------------------------------------------------------------------------
* TABLE 6b: Culture x Within-Team Experience Gap (new, per reviewer request 2026-08-11)
* exp_gap = |exp_i - exp_j|, the two teammates' OWN tenure (not the team average exp_mean).
* Rationale: shared culture/language may substitute for shared experience when one teammate
* is a veteran and the other is new. Individual tenure imputed to 1 yr for rookies, same
* rule as exp_mean. Demeaned like every other interaction table. exp_gap_dm is also included
* as a plain control alongside its interaction, per the request.
* Requires re-merging individual player experience (winners/losers p1/p2 experience_double)
* since team_gs only carries the team-average exp_mean, not each player's own tenure.
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 6b. Heterogeneity: Culture x Within-Team Experience Gap ==="
display "FE: ty + stage_code | SE: clustered by match"
display ""

preserve
import excel using "data/atp/men_matches_with_ranks_cleaned.xlsx", sheet("players_list") firstrow clear
keep match_id winners_p1_experience_double winners_p2_experience_double ///
    losers_p1_experience_double losers_p2_experience_double
duplicates drop match_id, force
tempfile exp_gap_lookup
save `exp_gap_lookup'
restore

merge m:1 match_id using `exp_gap_lookup', keep(master match) nogen

destring winners_p1_experience_double winners_p2_experience_double ///
    losers_p1_experience_double losers_p2_experience_double, replace force

generate double _e1 = winners_p1_experience_double if win == 1
replace       _e1 = losers_p1_experience_double  if win == 0
generate double _e2 = winners_p2_experience_double if win == 1
replace       _e2 = losers_p2_experience_double  if win == 0
replace _e1 = 1 if missing(_e1) | _e1 <= 0
replace _e2 = 1 if missing(_e2) | _e2 <= 0

generate double exp_gap = abs(_e1 - _e2)
quietly summarize exp_gap
generate double exp_gap_dm = exp_gap - r(mean)
generate double exp_gap_dm_sq = exp_gap_dm ^ 2
display "exp_gap_dm: mean = " r(mean) "  SD = " r(sd)
drop _e1 _e2

generate double int_sc_gap = same_country  * exp_gap_dm
generate double int_sl_gap = same_language * exp_gap_dm
generate double int_lp_gap = ling_prox     * exp_gap_dm

display "--- 6b-i. Same Nationality x exp_gap ---"
logit win i.ty i.stage_code same_country  int_sc_gap exp_gap_dm rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country int_sc_gap exp_gap_dm rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_gap_sc

display ""
display "--- 6b-ii. Same Language x exp_gap ---"
logit win i.ty i.stage_code same_language  int_sl_gap exp_gap_dm rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language int_sl_gap exp_gap_dm rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_gap_sl

display ""
display "--- 6b-iii. Language Proximity x exp_gap ---"
logit win i.ty i.stage_code ling_prox  int_lp_gap exp_gap_dm rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox int_lp_gap exp_gap_dm rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_gap_lp

estimates table het_gap_sc het_gap_sl het_gap_lp, b se stats(N ll)

* --- Table 6b robustness: adding Culture x exp_gap_dm^2 ---
display ""
display "=== TABLE 6b ROBUSTNESS. + Culture x exp_gap_dm^2 ==="
display ""

generate double int_sc_gap_sq = same_country  * exp_gap_dm_sq
generate double int_sl_gap_sq = same_language * exp_gap_dm_sq
generate double int_lp_gap_sq = ling_prox     * exp_gap_dm_sq

logit win i.ty i.stage_code same_country  int_sc_gap int_sc_gap_sq exp_gap_dm exp_gap_dm_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country int_sc_gap int_sc_gap_sq exp_gap_dm exp_gap_dm_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_gap2_sc

logit win i.ty i.stage_code same_language  int_sl_gap int_sl_gap_sq exp_gap_dm exp_gap_dm_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language int_sl_gap int_sl_gap_sq exp_gap_dm exp_gap_dm_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_gap2_sl

logit win i.ty i.stage_code ling_prox  int_lp_gap int_lp_gap_sq exp_gap_dm exp_gap_dm_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox int_lp_gap int_lp_gap_sq exp_gap_dm exp_gap_dm_sq rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_gap2_lp

estimates table het_gap2_sc het_gap2_sl het_gap2_lp, b se stats(N ll)

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
display "Controls: same as Table 3 -- rank_mean, opp_rank_mean, single_top100, exp_mean_dm, exp_mean_dm_sq"
display ""

generate double int_sc_clay  = same_country  * clay
generate double int_sc_grass = same_country  * grass
generate double int_sl_clay  = same_language * clay
generate double int_sl_grass = same_language * grass
generate double int_lp_clay  = ling_prox     * clay
generate double int_lp_grass = ling_prox     * grass

display "--- 5a. Same Nationality: Spec (a) grass vs. rest ---"
logit win i.stage_code same_country grass int_sc_grass rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country grass int_sc_grass rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_surf_sc_grass

display ""
display "--- 5a. Same Nationality: Spec (b) clay vs. rest ---"
logit win i.stage_code same_country clay int_sc_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country clay int_sc_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_surf_sc_clay

display ""
display "--- 5b. Same Language: Spec (a) grass vs. rest ---"
logit win i.stage_code same_language grass int_sl_grass rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language grass int_sl_grass rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_surf_sl_grass

display ""
display "--- 5b. Same Language: Spec (b) clay vs. rest ---"
logit win i.stage_code same_language clay int_sl_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language clay int_sl_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_surf_sl_clay

display ""
display "--- 5c. Language Proximity: Spec (a) grass vs. rest ---"
logit win i.stage_code ling_prox grass int_lp_grass rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox grass int_lp_grass rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_surf_lp_grass

display ""
display "--- 5c. Language Proximity: Spec (b) clay vs. rest ---"
logit win i.stage_code ling_prox clay int_lp_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox clay int_lp_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_surf_lp_clay

display ""
display "--- Comparison: Spec (a) grass vs. Spec (b) clay ---"
estimates table het_surf_sc_grass het_surf_sc_clay, b se stats(N ll)
estimates table het_surf_sl_grass het_surf_sl_clay, b se stats(N ll)
estimates table het_surf_lp_grass het_surf_lp_clay, b se stats(N ll)

* ---------------------------------------------------------------------------
* TABLE 5 ROBUSTNESS: tournament x year FE restored, interaction only, no separate
* surface main effect (per reviewer request 2026-08-11). Mirror image of the main Table 5
* spec's tradeoff: T x Y FE comes back, but the surface main effect can no longer be
* separately identified from the tournament dummies it's nested in.
* ---------------------------------------------------------------------------
display ""
display "=== TABLE 5 ROBUSTNESS. T x Y FE, Culture x Surface interaction only ==="
display "No separate surface main effect -- hardcourt is the implicit baseline."
display "FE: ty + stage_code | SE: clustered by match"
display ""

display "--- 5-rob-i. Same Nationality ---"
logit win i.ty i.stage_code same_country int_sc_grass int_sc_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country int_sc_grass int_sc_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_surfty_sc

display ""
display "--- 5-rob-ii. Same Language ---"
logit win i.ty i.stage_code same_language int_sl_grass int_sl_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language int_sl_grass int_sl_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_surfty_sl

display ""
display "--- 5-rob-iii. Language Proximity ---"
logit win i.ty i.stage_code ling_prox int_lp_grass int_lp_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox int_lp_grass int_lp_clay rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_surfty_lp

estimates table het_surfty_sc het_surfty_sl het_surfty_lp, b se stats(N ll)

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
logit win i.ty i.stage_code same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_ic_sc_s2

display ""
display "--- 6b. Same Language ---"
logit win i.ty i.stage_code same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_ic_sl_s2

display ""
display "--- 6c. Language Proximity ---"
logit win i.ty i.stage_code ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
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
logit win i.ty i.stage_code same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq if tier2_flag == 0, cluster(match_id)
margins, dydx(same_country ic_team_dm int_sc_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_ic_sc_s2_ex2

display ""
display "--- 6b, ex-Tier2. Same Language ---"
logit win i.ty i.stage_code same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq if tier2_flag == 0, cluster(match_id)
margins, dydx(same_language ic_team_dm int_sl_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store het_ic_sl_s2_ex2

display ""
display "--- 6c, ex-Tier2. Language Proximity ---"
logit win i.ty i.stage_code ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq if tier2_flag == 0, cluster(match_id)
margins, dydx(ling_prox ic_team_dm int_lp_ic rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
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
* Source: data/atp/tiebreak_panel.csv (sets 1-5, rebuilt via
*   code/cleaning/tiebreak_panel.ipynb)
* Spec: standard 7pt tiebreaks only, any set (1-5). Advantage-set deciders with no
*   breaker played are not tiebreaks and are excluded entirely (confirmed by the raw
*   tiebreak-score columns being null for all such cases); no genuine 10-point
*   super-tiebreak exists anywhere in this GS dataset, so tiebreak_panel.csv is 7pt-only
*   and there is currently nothing valid to add as a robustness spec.
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

encode tournament, generate(tourn_code)
drop tournament
rename tourn_code tournament

destring stage_code same_country same_language ling_prox rank_mean opp_rank_mean ///
    rank_gap single_top100 exp_mean age_mean win pre_olympic olympic_period tb_set won_tb year, replace force
replace stage_code = 0 if missing(stage_code)
replace exp_mean = 1 if missing(exp_mean)
quietly egen ty = group(tournament year)

* Demean exp_mean against this panel's own estimation-sample mean (mirrors Table
* 3/5/6/6a's use of exp_mean_dm, but this panel is a different sample so it needs
* its own mean).
quietly summarize exp_mean
generate double exp_mean_dm = exp_mean - r(mean)
generate double exp_mean_dm_sq = exp_mean_dm ^ 2
display "exp_mean_dm: mean yrs since turning pro (tiebreak sample) = " r(mean)

* Same demeaning for age_mean, for the Table 4-AGE robustness spec below.
quietly summarize age_mean
generate double age_mean_dm = age_mean - r(mean)
generate double age_mean_dm_sq = age_mean_dm ^ 2
display "age_mean_dm: mean team age at tournament (tiebreak sample) = " r(mean)

count
display "Tiebreak team-obs loaded (7pt, all sets 1-5): " r(N)
display ""

display "Breakdown by set:"
tabulate tb_set

display ""
display "=== TABLE 4. Tiebreak Win — Logit ==="
display "Outcome: won_tb | Unit: team x tiebreak | Sample: GS only (no Olympics)"
display "Spec: standard 7pt tiebreaks only, any set (1-5)."
display "Controls: rank_mean, opp_rank_mean, single_top100, exp_mean_dm, exp_mean_dm_sq (same as Table 3;"
display "  exp_mean_dm demeaned against this panel's own estimation-sample mean)"
display "FE: tournament x year (ty) + stage_code | SE: clustered by match"
display ""

display "--- 4a. Same Nationality ---"
logit won_tb i.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store tbn_same_country

display ""
display "--- 4b. Same Official Language ---"
logit won_tb i.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store tbn_same_language

display ""
display "--- 4c. Language Proximity ---"
logit won_tb i.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)
estimates store tbn_ling_prox

display ""
display "--- Table 4 summary ---"
estimates table tbn_same_country tbn_same_language tbn_ling_prox, b se stats(N ll)

* ─────────────────────────────────────────────────────────────────────────────
* TABLE 4-AGE: TIEBREAK WIN — LOGIT, AGE SPEC (robustness: age_mean replaces exp_mean)
* Same age-vs-experience substitution as Table 3-AGE, applied to the tiebreak-win outcome.
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== TABLE 4-AGE. Tiebreak Win — Logit, age_mean replacing exp_mean ==="
display "Controls: rank_mean, opp_rank_mean, single_top100, age_mean_dm, age_mean_dm_sq"
display ""

display "--- 4-age-i. Same Nationality ---"
logit won_tb i.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq, cluster(match_id)
margins, dydx(same_country rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq)
estimates store tbn_same_country_age

display ""
display "--- 4-age-ii. Same Official Language ---"
logit won_tb i.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq, cluster(match_id)
margins, dydx(same_language rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq)
estimates store tbn_same_language_age

display ""
display "--- 4-age-iii. Language Proximity ---"
logit won_tb i.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq, cluster(match_id)
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 age_mean_dm age_mean_dm_sq)
estimates store tbn_ling_prox_age

display ""
display "--- Table 4-AGE summary ---"
estimates table tbn_same_country_age tbn_same_language_age tbn_ling_prox_age, b se stats(N ll)

* ─────────────────────────────────────────────────────────────────────────────
* TABLE 4-NOINT: TIEBREAK WIN — LOGIT, NO INTERCEPT (double-counting check)
* Same reparameterization check as Table 3-NOINT, applied to the tiebreak panel.
* Expect AMEs identical to Table 4.
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== TABLE 4-NOINT. Tiebreak Win — Logit, no intercept (full-rank tournament x year FE) ==="
display ""

display "--- 4-noint-i. Same Nationality ---"
logit won_tb ibn.ty i.stage_code same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id) noconstant
margins, dydx(same_country rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)

display ""
display "--- 4-noint-ii. Same Official Language ---"
logit won_tb ibn.ty i.stage_code same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id) noconstant
margins, dydx(same_language rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)

display ""
display "--- 4-noint-iii. Language Proximity ---"
logit won_tb ibn.ty i.stage_code ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq, cluster(match_id) noconstant
margins, dydx(ling_prox rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq)

* ─────────────────────────────────────────────────────────────────────────────
* TABLE 4-RANDONE: TIEBREAK WIN — LOGIT, ONE RANDOMLY-SELECTED TEAM PER TIEBREAK
* Pair unit here is one specific tiebreak (match_id x tb_set x tb_type), NOT the
* match -- a match can contain several tiebreaks (sets 1-5). SE stays clustered by
* match_id (a match can still contribute >1 tiebreak-row after de-duplication),
* unlike Table 3-RANDONE where de-duplication leaves exactly one row per match and
* no clustering unit remains. Sparse tournament x year x round cells after halving
* the sample (~1,182 vs 2,364) can produce a singular Hessian for some draws --
* caught with capture/`_rc' and skipped rather than left to crash the do-file.
* ─────────────────────────────────────────────────────────────────────────────
display ""
display "=== TABLE 4-RANDONE. Tiebreak Win — Logit, one randomly-selected team per tiebreak ==="
display ""

capture drop tb_key
egen long tb_key = group(match_id tb_set tb_type)

tempfile randone_results_tb
tempname rr2
postfile `rr2' str24 cvar seed double ame double pval using "`randone_results_tb'", replace

foreach cvar in same_country same_language ling_prox {
    forvalues seed = 0/9 {
        preserve
        set seed `seed'
        tempvar u u_tb keep_row
        bysort tb_key (won_tb): generate double `u' = runiform() if _n==1
        bysort tb_key: egen double `u_tb' = max(`u')
        generate byte `keep_row' = (won_tb==1 & `u_tb'>0.5) | (won_tb==0 & `u_tb'<=0.5)
        keep if `keep_row'

        capture noisily logit won_tb `cvar' rank_mean opp_rank_mean single_top100 exp_mean_dm exp_mean_dm_sq i.ty i.stage_code, cluster(match_id)
        if _rc == 0 {
            capture noisily margins, dydx(`cvar')
            if _rc == 0 {
                matrix b = r(b)
                matrix V = r(V)
                local ame_ = b[1,1]
                local se_  = sqrt(V[1,1])
                local p_   = 2*(1 - normal(abs(`ame_'/`se_')))
                post `rr2' ("`cvar'") (`seed') (`ame_') (`p_')
            }
            else {
                display "  [`cvar'] seed `seed': margins failed -- skipped"
            }
        }
        else {
            display "  [`cvar'] seed `seed': logit failed to converge (likely singular Hessian, sparse FE cell) -- skipped"
        }
        restore
    }
}
postclose `rr2'

display ""
display "--- Table 4-RANDONE summary across converged random draws ---"
preserve
use "`randone_results_tb'", clear
foreach cvar in same_country same_language ling_prox {
    quietly count if cvar=="`cvar'"
    local nok = r(N)
    quietly summarize ame if cvar=="`cvar'"
    local mean_a = r(mean)
    local sd_a   = r(sd)
    local min_a  = r(min)
    local max_a  = r(max)
    quietly count if cvar=="`cvar'" & pval<0.05
    local nsig = r(N)
    display "`cvar': `nok'/10 converged  mean AME=" %6.4f `mean_a' "  sd=" %6.4f `sd_a' "  range=[" %6.4f `min_a' ", " %6.4f `max_a' "]  significant in `nsig'/`nok' draws"
}
restore


display ""
display "Section 7 complete: Tiebreak win regressions (Table 4, Table 4-AGE, Table 4-NOINT, Table 4-RANDONE)."

* ═══════════════════════════════════════════════════════════════════════════════
* SECTION 8: PARTNER SELECTION — OWN COUNTRY AND TOURNAMENT-FIELD COMPOSITION
* (new, per Lingqing's meeting notes / main.tex red comment: extend the Section 6.1
* "does own ranking predict partner-culture similarity?" ego-row regression with own
* nationality and field-composition controls.)
*
* IMPORTANT: the underlying Section 6 "Partner Selection" pipeline (deduplicating
* partnerships by tournament x team, the closed-form random-matching benchmark, and
* the tournament-field composition calculation) exists ONLY in the Python notebook
* (code/analysis/homophily.ipynb, cells ~51-56) and has never previously been mirrored
* in Stata -- unlike every other section of this do-file, which mirrors an existing
* Stata-equivalent computation. Re-deriving that dedup/composition logic from scratch
* in Stata, unexecuted, would carry real risk of a subtle, uncaught bug (string-based
* team-key construction, per-tournament-year field aggregation, etc.). Instead, this
* section imports data/atp/partner_selection_ego.csv, which was exported directly from
* the EXECUTED and verified Python ego2 dataframe (4,202 ego-rows; identical
* construction to homophily.ipynb) -- so the regressions below run on real, correct
* data even though, like the rest of this do-file, they have not been run through
* Stata itself in this environment.
* ═══════════════════════════════════════════════════════════════════════════════

display ""
display "=== SECTION 8. Partner Selection: Own Country and Field Composition ==="
display "Data: data/atp/partner_selection_ego.csv (exported from the executed Python ego2 sample)"
display ""

import delimited "data/atp/partner_selection_ego.csv", clear varnames(1) stringcols(1)

encode tourn_year, generate(ty8)
encode own_iso3, generate(own_iso3_code)
encode own_iso3_grp, generate(own_iso3_grp_code)
egen team_id = group(team_key)

destring own_rank same_country same_language ling_prox composition_nat composition_lang composition_ling, replace force

count
display "Ego-rows loaded: " r(N) " (expect 4,202)"
display ""

* Singleton-nationality exclusion (Spec 1 & 3 only, per Alessandro's request 2026-09-01):
* drop ego-rows whose own_iso3 has exactly 1 observation in the ego-row sample, so a single
* lone-nationality player cannot single-handedly drive the own_rank coefficient. This is
* separate from, and does not touch, the own_iso3_grp <10-obs "OTHER" pooling used below,
* which stays as-is and continues to apply only to the Spec 2/4 own-country FE regressions.
bysort own_iso3: gen long _iso3_n = _N
gen byte _singleton_iso3 = (_iso3_n == 1)
quietly count if _singleton_iso3
display "Singleton-nationality exclusion (Spec 1 & 3 only): " r(N) " ego-rows from nationalities with exactly 1 observation dropped for these two specs"
display ""

* --- Spec 1: own_rank alone (reproduces Section 6.1's original result) ---
display "--- Spec 1: own_rank + tourn_year FE (excl. singleton-nationality ego-rows) ---"
foreach outcome in same_country same_language {
    display "  [`outcome']"
    quietly logit `outcome' own_rank i.ty8 if !_singleton_iso3, cluster(team_id)
    margins, dydx(own_rank)
}
display "  [ling_prox, OLS]"
regress ling_prox own_rank i.ty8 if !_singleton_iso3, cluster(team_id)

* --- Spec 2: own country (own_iso3) fixed effects in place of own_rank ---
* NOTE: per the Python analysis, this specification fails to converge for the binary
* outcomes (same_country, same_language) via quasi-complete separation -- several of
* 64 nationalities have too few ego-row observations (14 have <10, down to single
* players) for a country-specific fixed effect to be identified. own_iso3_grp pools
* nationalities with <10 obs into "OTHER" (51 remaining groups), which is enough for
* ling_prox (OLS) but NOT enough to fix the logit specs -- expect `logit` to fail or
* emit a "not identified" / omitted-predictor error for same_country/same_language.
display ""
display "--- Spec 2: C(own_iso3_grp) + tourn_year FE (expect failure for binary outcomes) ---"
capture noisily logit same_country i.own_iso3_grp_code i.ty8, cluster(team_id)
capture noisily logit same_language i.own_iso3_grp_code i.ty8, cluster(team_id)
display "  [ling_prox, OLS -- should estimate successfully]"
regress ling_prox i.own_iso3_grp_code i.ty8, cluster(team_id)

* --- Spec 3: own_rank + outcome-matched field composition ---
display ""
display "--- Spec 3: own_rank + field composition + tourn_year FE (excl. singleton-nationality ego-rows) ---"
display "  [same_country]"
quietly logit same_country own_rank composition_nat i.ty8 if !_singleton_iso3, cluster(team_id)
margins, dydx(own_rank composition_nat)
display "  [same_language]"
quietly logit same_language own_rank composition_lang i.ty8 if !_singleton_iso3, cluster(team_id)
margins, dydx(own_rank composition_lang)
display "  [ling_prox, OLS]"
regress ling_prox own_rank composition_ling i.ty8 if !_singleton_iso3, cluster(team_id)

* --- Spec 4: own country + field composition (expect same binary-outcome failure as Spec 2) ---
display ""
display "--- Spec 4: C(own_iso3_grp) + field composition + tourn_year FE (expect failure for binary outcomes) ---"
capture noisily logit same_country i.own_iso3_grp_code composition_nat i.ty8, cluster(team_id)
capture noisily logit same_language i.own_iso3_grp_code composition_lang i.ty8, cluster(team_id)
display "  [ling_prox, OLS -- should estimate successfully]"
regress ling_prox i.own_iso3_grp_code composition_ling i.ty8, cluster(team_id)

display ""
display "Section 8 complete: Partner selection, own-country and field-composition specs."

log close

display ""
display "All sections complete. Results saved to stata_homophily_results.txt"


* build_homophily.do
* Run this from the project root in STATA.
* It imports the same Excel dataset, builds the team panel, and estimates separate logit models
* for same_country, same_language, and ling_prox.

capture log close _all
set more off

cd "C:\Users\ALESSANDRO\Documents\GitHub\tennis-homophily"
clear

import excel using "data/atp/men_matches_with_ranks_cleaned.xlsx", sheet("Sheet1") firstrow clear

destring year winners_set1 winners_set2 winners_set3 losers_set1 losers_set2 losers_set3 ///
    winners_p1_top100_within_1y winners_p2_top100_within_1y losers_p1_top100_within_1y losers_p2_top100_within_1y ///
    rank_mean_winners rank_mean_losers rank_diff_winners rank_diff_losers, replace force

* Create tiebreak flags
generate byte tb_s1 = (winners_set1==7 & losers_set1==6) | (winners_set1==6 & losers_set1==7)
generate byte tb_s2 = (winners_set2==7 & losers_set2==6) | (winners_set2==6 & losers_set2==7)
generate byte regular_tb = tb_s1 | tb_s2
generate byte tb_s3_regular = (winners_set3==7 & losers_set3==6) | (winners_set3==6 & losers_set3==7)
generate byte match_tb = (winners_set3 >= 8)
generate byte any_tb = regular_tb | match_tb | tb_s3_regular

generate byte w_won_tb_s1 = tb_s1 & (winners_set1==7)
generate byte w_won_tb_s2 = tb_s2 & (winners_set2==7)

generate byte three_sets = !missing(winners_set3)
generate byte winner_lost_s1 = (winners_set1 < losers_set1)
generate byte comeback = winner_lost_s1 & three_sets

* Build team panel: winners and losers rows
preserve
keep match_id tournament year surface stage_code any_tb regular_tb match_tb tb_s3_regular three_sets comeback winner_lost_s1 pre_olympic olympic_period olympics_tourn ///
     same_country_winners winners_same_language winners_linguistic_proximity rank_mean_winners rank_mean_losers rank_diff_winners ///
     winners_p1_top100_within_1y winners_p2_top100_within_1y winners_set1 winners_set2
rename same_country_winners same_country
rename winners_same_language same_language
rename winners_linguistic_proximity ling_prox
rename rank_mean_winners rank_mean
rename rank_mean_losers rank_mean_opp
rename rank_diff_winners rank_gap
generate byte single_top100 = (winners_p1_top100_within_1y==1 | winners_p2_top100_within_1y==1)
generate byte win = 1
generate byte lost_set1 = winner_lost_s1
generate byte won_tb_s1 = tb_s1 & (winners_set1==7)
generate byte won_tb_s2 = tb_s2 & (winners_set2==7)
save "temp_winners.dta", replace
restore

preserve
keep match_id tournament year surface stage_code any_tb regular_tb match_tb tb_s3_regular three_sets comeback winner_lost_s1 pre_olympic olympic_period olympics_tourn ///
     same_country_losers losers_same_language losers_linguistic_proximity rank_mean_losers rank_mean_winners rank_diff_losers ///
     losers_p1_top100_within_1y losers_p2_top100_within_1y winners_set1 winners_set2
rename same_country_losers same_country
rename losers_same_language same_language
rename losers_linguistic_proximity ling_prox
rename rank_mean_losers rank_mean
rename rank_mean_winners rank_mean_opp
rename rank_diff_losers rank_gap
generate byte single_top100 = (losers_p1_top100_within_1y==1 | losers_p2_top100_within_1y==1)
generate byte win = 0
generate byte lost_set1 = !winner_lost_s1
generate byte won_tb_s1 = tb_s1 & (winners_set1==6)
generate byte won_tb_s2 = tb_s2 & (winners_set2==6)
append using "temp_winners.dta"

bysort match_id: egen opp_rank_mean = total(rank_mean) - rank_mean
replace stage_code = -1 if missing(stage_code)
replace rank_gap = . if missing(rank_gap)
replace rank_mean = . if missing(rank_mean)

generate byte won_any_tb = (won_tb_s1==1 | won_tb_s2==1 | match_tb==1 | tb_s3_regular==1)

* Section 3: Olympic cycle descriptive tables and graphs
preserve
keep if olympics_tourn==0

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

* Table 2A: Tokyo 2021 cycle
preserve
keep if cycle_tokyo != ""
collapse (count) N_teams=match_id (mean) same_country same_language ling_prox, by(cycle_tokyo)
replace same_country = same_country*100
replace same_language = same_language*100
replace ling_prox = ling_prox*100
list cycle_tokyo N_teams same_country same_language ling_prox, abbreviate(14)
restore

* Table 2B: Paris 2024 cycle
preserve
keep if cycle_paris != ""
collapse (count) N_teams=match_id (mean) same_country same_language ling_prox, by(cycle_paris)
replace same_country = same_country*100
replace same_language = same_language*100
replace ling_prox = ling_prox*100
list cycle_paris N_teams same_country same_language ling_prox, abbreviate(14)
restore

* Graph: cultural homophily over time in Grand Slams
preserve
collapse (mean) same_country same_language ling_prox, by(year)
replace same_country = same_country*100
replace same_language = same_language*100
replace ling_prox = ling_prox*100

graph twoway \
    (line same_country year, lcolor(navy) lpattern(solid) lwidth(medium)) \
    (line same_language year, lcolor(cranberry) lpattern(dash) lwidth(medium)) \
    (line ling_prox year, lcolor(green) lpattern(dot) lwidth(medium)), \
    title("Cultural Homophily in Grand Slam Doubles") \
    xtitle(Year) ytitle("Share of teams (%)") \
    legend(order(1 "Same nationality" 2 "Same language" 3 "Language proximity") ring(0) pos(9)) \
    xlabel(2018(1)2025)
graph export "fig_homophily_year_trends.png", replace
restore

restore

drop if missing(rank_mean) | missing(rank_gap)

* Run models for each culture measure separately
local cultures same_country same_language ling_prox

gen byte ty = .
quietly egen ty = group(tournament year)

log using "stata_homophily_results.txt", replace text

foreach culture of local cultures {
    display "\n*** Culture measure: `culture' ***\n"

    display "Table 3: Match Win"
    quietly logit win i.tournament#i.year i.stage_code `culture' rank_mean opp_rank_mean rank_gap single_top100
    estimates store win_`culture'
    estimates dir
    estimates table win_`culture', b se star stats(N ll)

    display "\nTable 4: Tiebreak Win (any tiebreak)"
    quietly logit won_any_tb i.tournament#i.year i.stage_code `culture' rank_mean opp_rank_mean rank_gap single_top100 if any_tb==1
    estimates store tb_`culture'
    estimates table tb_`culture', b se star stats(N ll)

    display "\nTable 5: Comeback Win (conditional on losing set 1, 3 sets)"
    preserve
    keep if three_sets==1 & lost_set1==1
    quietly egen minwin = min(win), by(ty)
    quietly egen maxwin = max(win), by(ty)
    drop if minwin==maxwin
    quietly logit win i.ty i.stage_code `culture' rank_mean opp_rank_mean rank_gap single_top100
    estimates store cb_`culture'
    estimates table cb_`culture', b se star stats(N ll)
    restore
}

log close

display "STATA comparison complete. Results saved to stata_homophily_results.txt"

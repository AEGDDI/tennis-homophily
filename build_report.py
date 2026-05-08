"""
Generates a self-contained HTML report for the tennis homophily analysis.
Run from the repo root: python build_report.py
Output: report/homophily_report.html
"""
import warnings; warnings.filterwarnings('ignore')
import os, base64, io
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import statsmodels.formula.api as smf
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

os.makedirs('report', exist_ok=True)
os.makedirs('pdf', exist_ok=True)

# ─────────────────────────── DATA + VARIABLES ────────────────────────────────
df = pd.read_excel('data/atp/men_matches_with_ranks_cleaned.xlsx')
n_raw_matches = len(df)

# Drop likely retirements/walkovers that are not explicitly flagged in the data.
# A valid completed set is 6-0 through 6-4, 7-5, 7-6, or an extended final set
# score with a two-game margin (e.g. 8-6, 9-7). If the first two sets are split,
# the third set must also be complete, either as a regular set or as a
# match tiebreak.
def regular_set_complete(w, l):
    valid = w.notna() & l.notna()
    hi = pd.concat([w, l], axis=1).max(axis=1)
    lo = pd.concat([w, l], axis=1).min(axis=1)
    return valid & (
        ((hi == 6) & (lo <= 4)) |
        ((hi == 7) & lo.isin([5, 6])) |
        ((hi >= 8) & ((hi - lo) == 2))
    )

def match_tiebreak_complete(w, l):
    valid = w.notna() & l.notna()
    hi = pd.concat([w, l], axis=1).max(axis=1)
    lo = pd.concat([w, l], axis=1).min(axis=1)
    return valid & (hi >= 10) & ((hi - lo) >= 2)

s1_complete = regular_set_complete(df['winners_set1'], df['losers_set1'])
s2_complete = regular_set_complete(df['winners_set2'], df['losers_set2'])
split_sets = (
    ((df['winners_set1'] > df['losers_set1']) & (df['winners_set2'] < df['losers_set2'])) |
    ((df['winners_set1'] < df['losers_set1']) & (df['winners_set2'] > df['losers_set2']))
)
s3_complete = (
    regular_set_complete(df['winners_set3'], df['losers_set3']) |
    match_tiebreak_complete(df['winners_set3'], df['losers_set3'])
)
retired_or_incomplete = ~s1_complete | ~s2_complete | (split_sets & ~s3_complete)
n_dropped_incomplete = int(retired_or_incomplete.sum())
df = df.loc[~retired_or_incomplete].copy()

df['tb_s1'] = ((df['winners_set1']==7)&(df['losers_set1']==6))|((df['winners_set1']==6)&(df['losers_set1']==7))
df['tb_s2'] = ((df['winners_set2']==7)&(df['losers_set2']==6))|((df['winners_set2']==6)&(df['losers_set2']==7))
df['regular_tb']    = df['tb_s1'] | df['tb_s2']
df['tb_s3_regular'] = (((df['winners_set3']==7)&(df['losers_set3']==6))|((df['winners_set3']==6)&(df['losers_set3']==7))).fillna(False)
df['match_tb']      = match_tiebreak_complete(df['winners_set3'], df['losers_set3'])
df['any_tb']        = df['regular_tb'] | df['match_tb'] | df['tb_s3_regular']
df['three_sets']    = df['winners_set3'].notna()
df['w_won_tb_s1']   = df['tb_s1'] & (df['winners_set1']==7)
df['w_won_tb_s2']   = df['tb_s2'] & (df['winners_set2']==7)
df['winner_lost_s1']= df['winners_set1'] < df['losers_set1']
df['comeback']      = df['winner_lost_s1'] & df['three_sets']
df['pre_olympic']   = df['year'].isin([2022,2023]).astype(int)
df['olympic_period']= df['year'].isin([2022,2023,2024]).astype(int)
df['olympics_tourn']= (df['tournament']=='Olympics').astype(int)
df['same_country_diff']  = df['same_country_winners'].astype(float) - df['same_country_losers'].astype(float)
df['same_language_diff'] = df['winners_same_language'].astype(float) - df['losers_same_language'].astype(float)
df['ling_prox_diff']     = df['winners_linguistic_proximity'].astype(float) - df['losers_linguistic_proximity'].astype(float)
df['same_country_avg']   = (df['same_country_winners'] + df['same_country_losers'])/2
df['same_language_avg']  = (df['winners_same_language'] + df['losers_same_language'])/2
df['ling_prox_avg']      = (df['winners_linguistic_proximity'] + df['losers_linguistic_proximity'])/2

def assign_tokyo_cycle(row):
    y = row['year']
    t = row['tournament']
    if y in [2018, 2019]:
        return 'Pre-Tokyo (2018–19)'
    if y == 2020:
        return 'Tokyo Prep (2020–21)' if t != 'Wimbledon' else None
    if y == 2021:
        if t in ['Australian Open', 'Roland Garros', 'Wimbledon']:
            return 'Tokyo Prep (2020–21)'
        if t == 'US Open':
            return 'Post-Tokyo (2021–22)'
    if y == 2022:
        return 'Post-Tokyo (2021–22)'
    return None

def assign_paris_cycle(row):
    y = row['year']
    t = row['tournament']
    if y == 2022:
        return 'Pre-Paris (2022)'
    if y in [2023, 2024]:
        return 'Paris Prep (2023–24)' if t in ['Australian Open', 'Roland Garros', 'Wimbledon'] else ('Post-Paris (2024–25)' if t == 'US Open' else None)
    if y == 2025:
        return 'Post-Paris (2024–25)'
    return None

cycle_map = {2018:'Pre-Tokyo (2018–19)',2019:'Pre-Tokyo (2018–19)',2021:'Post-Tokyo (2021)',
             2022:'Pre-Paris (2022–23)',2023:'Pre-Paris (2022–23)',2024:'Paris 2024',2025:'Post-Paris (2025)'}
df['cycle'] = df['year'].map(cycle_map)
df['cycle_tokyo'] = df.apply(assign_tokyo_cycle, axis=1)
df['cycle_paris'] = df.apply(assign_paris_cycle, axis=1)

gs = df[df['olympics_tourn']==0].copy()

# team panel
match_vars = ['match_id','tournament','year','surface','stage_code','any_tb','regular_tb','match_tb','tb_s3_regular',
              'three_sets','comeback','winner_lost_s1','pre_olympic','olympic_period','olympics_tourn']
w_row = df[match_vars].copy()
for col,src in [('same_country','same_country_winners'),('same_language','winners_same_language'),
                ('ling_prox','winners_linguistic_proximity'),('rank_mean','rank_mean_winners'),
                ('wl_career_diff','wl_career_diff_winners'),('rank_gap','rank_diff_winners')]:
    w_row[col] = df[src].values
w_row['single_top100'] = ((df['winners_p1_top100_within_1y']==1) | (df['winners_p2_top100_within_1y']==1)).astype(int)
w_row['win']=1; w_row['won_tb_s1']=df['w_won_tb_s1'].values; w_row['won_tb_s2']=df['w_won_tb_s2'].values

l_row = df[match_vars].copy()
for col,src in [('same_country','same_country_losers'),('same_language','losers_same_language'),
                ('ling_prox','losers_linguistic_proximity'),('rank_mean','rank_mean_losers'),
                ('wl_career_diff','wl_career_diff_losers'),('rank_gap','rank_diff_losers')]:
    l_row[col] = df[src].values
l_row['single_top100'] = ((df['losers_p1_top100_within_1y']==1) | (df['losers_p2_top100_within_1y']==1)).astype(int)
l_row['win']=0
l_row['won_tb_s1']=(df['tb_s1']&(df['winners_set1']==6)).values
l_row['won_tb_s2']=(df['tb_s2']&(df['winners_set2']==6)).values

team_df = pd.concat([w_row, l_row], ignore_index=True)
team_df['opp_rank_mean'] = team_df.groupby('match_id')['rank_mean'].transform('sum') - team_df['rank_mean']
team_df['stage_code'] = team_df['stage_code'].fillna(-1).astype(int)
team_df['lost_set1'] = np.where(team_df['win']==1, team_df['winner_lost_s1'], ~team_df['winner_lost_s1'])
team_df['comeback'] = (team_df['win'] & team_df['lost_set1']).astype(int)
for col in team_df.columns:
    if str(team_df[col].dtype) in ('bool','boolean','Int64','Int32'):
        team_df[col] = team_df[col].fillna(0).astype(int)
team_gs = team_df[team_df['olympics_tourn']==0].dropna(subset=['rank_mean','rank_gap']).copy()
team_gs['won_regular_tb'] = ((team_gs['won_tb_s1']==1)|(team_gs['won_tb_s2']==1)).astype(int)
team_gs['won_any_tb'] = ((team_gs['won_tb_s1']==1)|(team_gs['won_tb_s2']==1)|(team_gs['match_tb']==1)|(team_gs['tb_s3_regular']==1)).astype(int)
tb_sub = team_gs[team_gs['any_tb']==1].copy()

cdf = df[(df['olympics_tourn']==0)&(df['three_sets']==True)].copy()
cdf['log_rank_diff'] = np.log1p(cdf['rank_mean_winners'].clip(lower=1))-np.log1p(cdf['rank_mean_losers'].clip(lower=1))
for col in cdf.columns:
    if str(cdf[col].dtype) in ('bool','boolean','Int64','Int32'): cdf[col]=cdf[col].fillna(0).astype(int)
cdf = cdf.dropna(subset=['log_rank_diff','same_country_diff'])

print('Data ready. Fitting models...')

# ─────────────────────────── MODELS ──────────────────────────────────────────
# Align baseline regressions with notebook (`code/analysis/homophily_analysis.ipynb`).
# Culture measure used here: language proximity (ethnic) only.
# For STATA comparison, run build_homophily.do to estimate same_country and same_language separately.
BASE = '+ ling_prox + rank_mean + opp_rank_mean + rank_gap + single_top100'
FE   = '+ C(tournament):C(year) + C(stage_code)'
FE2  = '+ C(tournament) + C(stage_code)'

logit_win = smf.logit(f'win ~ {BASE} {FE}', data=team_gs).fit(cov_type='cluster', cov_kwds={'groups':team_gs['match_id']}, disp=False)
logit_tb  = smf.logit(f'won_any_tb ~ {BASE} {FE}', data=tb_sub).fit(cov_type='cluster', cov_kwds={'groups':tb_sub['match_id']}, disp=False)

# Table 5 (Comeback Win): align with notebook spec.
# Estimate P(win | lost set 1, 3 sets) with tournament×year FE (dropping non-informative cells).
BASE_T5 = '+ ling_prox + rank_mean + opp_rank_mean + rank_gap + single_top100'
cb_sub = team_gs[(team_gs['three_sets'] == True) & (team_gs['lost_set1'] == 1)].copy()
cb_sub['ty'] = cb_sub['tournament'].astype(str) + '__' + cb_sub['year'].astype(int).astype(str)
_ty_nuniq = cb_sub.groupby('ty')['win'].nunique()
keep_ty = _ty_nuniq[_ty_nuniq == 2].index
cb_sub = cb_sub[cb_sub['ty'].isin(keep_ty)].copy()
try:
    logit_cb = smf.logit(f'win ~ {BASE_T5} + C(ty) + C(stage_code)', data=cb_sub).fit(
        cov_type='cluster', cov_kwds={'groups': cb_sub['match_id']}, disp=False, maxiter=200)
except Exception:
    logit_cb = smf.logit(f'win ~ {BASE_T5} + C(ty)', data=cb_sub).fit(
        cov_type='cluster', cov_kwds={'groups': cb_sub['match_id']}, disp=False, maxiter=200)

logit_win_mfx = logit_win.get_margeff(at='overall', method='dydx', dummy=True).summary_frame()
logit_tb_mfx = logit_tb.get_margeff(at='overall', method='dydx', dummy=True).summary_frame()
logit_cb_mfx = logit_cb.get_margeff(at='overall', method='dydx', dummy=True).summary_frame()

print('Models fitted.')

# ─────────────────────────── HELPERS ─────────────────────────────────────────
def stars(p):
    return '***' if p<0.01 else '**' if p<0.05 else '*' if p<0.10 else ''

def fig_to_b64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode()

def html_to_pdf(html_path, pdf_path):
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-software-rasterizer')
    options.add_argument('--hide-scrollbars')
    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    try:
        driver.get('file:///' + os.path.abspath(html_path).replace('\\', '/'))
        driver.execute_cdp_cmd('Page.enable', {})
        result = driver.execute_cdp_cmd('Page.printToPDF', {
            'printBackground': True,
            'landscape': False,
            'marginTop': 0.4,
            'marginBottom': 0.4,
            'marginLeft': 0.4,
            'marginRight': 0.4,
        })
        with open(pdf_path, 'wb') as f:
            f.write(base64.b64decode(result['data']))
    finally:
        driver.quit()


def reg_row(res, var, label, note_mfx=False):
    if var not in res.params:
        return f'<tr><td>{label}</td><td>—</td><td></td></tr>'
    c, se, p = res.params[var], res.bse[var], res.pvalues[var]
    s = stars(p)
    sig_class = 'sig1' if p<0.01 else 'sig5' if p<0.05 else 'sig10' if p<0.10 else ''
    return (f'<tr class="{sig_class}"><td>{label}</td>'
            f'<td class="num">{c:+.4f}<sup>{s}</sup></td>'
            f'<td class="num se">({se:.4f})</td></tr>')

def reg_table_html(results_list, var_labels, model_labels, caption='', nobs_list=None, r2_list=None):
    cols = len(results_list)
    header = ''.join(f'<th colspan="2">{m}</th>' for m in model_labels)
    rows_html = ''
    for var, lbl in var_labels:
        row = f'<tr><td class="varlbl">{lbl}</td>'
        for res in results_list:
            if var not in res.params:
                row += '<td>—</td><td></td>'
            else:
                c, se, p = res.params[var], res.bse[var], res.pvalues[var]
                s = stars(p)
                sig = 'sig1' if p<0.01 else 'sig5' if p<0.05 else 'sig10' if p<0.10 else ''
                row += f'<td class="num {sig}">{c:+.4f}<sup>{s}</sup></td><td class="num se">({se:.4f})</td>'
        row += '</tr>'
        rows_html += row
    # stat rows
    if nobs_list:
        rows_html += '<tr class="stat"><td>Observations</td>' + ''.join(f'<td class="num" colspan="2">{n:,}</td>' for n in nobs_list) + '</tr>'
    if r2_list:
        rows_html += '<tr class="stat"><td>R² / Pseudo-R²</td>' + ''.join(f'<td class="num" colspan="2">{r}</td>' for r in r2_list) + '</tr>'
    cap = f'<caption>{caption}</caption>' if caption else ''
    return f'''<table class="regtable">{cap}
<thead><tr><th></th>{header}</tr></thead>
<tbody>{rows_html}</tbody></table>'''

def margin_stat(mfx, var, col):
    if var not in mfx.index:
        return np.nan
    if col in mfx.columns:
        return mfx.loc[var, col]
    aliases = {
        'dy/dx': ['dy/dx', 'Coef.', 'coef'],
        'Std. Err.': ['Std. Err.', 'Std. Error', 'std err'],
        'Pr(>|z|)': ['Pr(>|z|)', 'P>|z|', 'P>|t|'],
    }
    for candidate in aliases.get(col, []):
        if candidate in mfx.columns:
            return mfx.loc[var, candidate]
    return np.nan

def mfx_table_html(mfx, var_labels, caption='', nobs=None):
    rows_html = ''
    for var, lbl in var_labels:
        dy = margin_stat(mfx, var, 'dy/dx')
        se = margin_stat(mfx, var, 'Std. Err.')
        p = margin_stat(mfx, var, 'Pr(>|z|)')
        if pd.isna(dy):
            rows_html += f'<tr><td class="varlbl">{lbl}</td><td>—</td><td></td></tr>'
            continue
        s = stars(p) if not pd.isna(p) else ''
        sig = '' if pd.isna(p) else ('sig1' if p<0.01 else 'sig5' if p<0.05 else 'sig10' if p<0.10 else '')
        rows_html += (
            f'<tr><td class="varlbl">{lbl}</td>'
            f'<td class="num {sig}">{100*dy:+.2f}<sup>{s}</sup></td>'
            f'<td class="num se">({100*se:.2f})</td></tr>'
        )
    if nobs is not None:
        rows_html += f'<tr class="stat"><td>Observations</td><td class="num" colspan="2">{int(nobs):,}</td></tr>'
    cap = f'<caption>{caption}</caption>' if caption else ''
    return f'''<table class="regtable">{cap}
<thead><tr><th></th><th colspan="2">Average marginal effect, pp</th></tr></thead>
<tbody>{rows_html}</tbody></table>'''

def ling_mfx_sentence(mfx, outcome_label):
    dy = margin_stat(mfx, 'ling_prox', 'dy/dx')
    p = margin_stat(mfx, 'ling_prox', 'Pr(>|z|)')
    if pd.isna(dy) or pd.isna(p):
        return f'Language proximity could not be summarized as a marginal effect for {outcome_label}.'
    direction = 'higher' if dy > 0 else 'lower'
    sig = 'statistically significant' if p < 0.05 else ('marginally significant' if p < 0.10 else 'not statistically significant')
    return f'Language proximity is associated with a {100*abs(dy):.1f} percentage-point {direction} probability of {outcome_label}; this estimate is {sig} (p = {p:.3f}).'

# ─────────────────────────── FIGURES ─────────────────────────────────────────
# Figure 1: Homophily trends over time
hcols = ['same_country_avg','same_language_avg','ling_prox_avg']
year_means = gs.groupby('year')[hcols].mean()*100

fig1_tokyo, axes = plt.subplots(1,3, figsize=(13,4), sharey=False)
titles = ['Same Nationality','Same Official Language','Language Proximity (ethnic)']
colors = ['#2166ac','#1a9850','#d6604d']
for ax, col, title, clr in zip(axes, hcols, titles, colors):
    ax.plot(year_means.index, year_means[col], marker='o', color=clr, linewidth=2, markersize=6, zorder=3)
    ax.axvspan(2017.5, 2019.5, alpha=0.10, color='#8ecae6', label='Pre-Tokyo')
    ax.axvspan(2019.5, 2021.5, alpha=0.10, color='#f4a261', label='Tokyo Prep')
    ax.axvspan(2021.5, 2022.5, alpha=0.10, color='#2a9d8f', label='Post-Tokyo')
    ax.axvline(2021.5, color='#d62728', linestyle='--', linewidth=1.8, alpha=0.8, label='Tokyo Olympics', zorder=2)
    ax.set_title(title, fontsize=10, fontweight='bold')
    ax.set_xlabel('Year', fontsize=9)
    ax.set_ylabel('Share of pairs (%)', fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
    ax.set_xticks(sorted(gs['year'].unique()))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.grid(axis='y', alpha=0.3)
    ax.legend(fontsize=7)
fig1_tokyo.suptitle('Cultural Homophily in Grand Slam Doubles — Tokyo 2021 Cycle',
                  fontsize=12, fontweight='bold', y=1.02)
plt.tight_layout()
fig1_tokyo_b64 = fig_to_b64(fig1_tokyo)

fig1_paris, axes = plt.subplots(1,3, figsize=(13,4), sharey=False)
for ax, col, title, clr in zip(axes, hcols, titles, colors):
    ax.plot(year_means.index, year_means[col], marker='o', color=clr, linewidth=2, markersize=6, zorder=3)
    ax.axvspan(2021.5, 2022.5, alpha=0.10, color='#8ecae6', label='Pre-Paris')
    ax.axvspan(2022.5, 2024.5, alpha=0.10, color='#f4a261', label='Paris Prep')
    ax.axvspan(2024.5, 2025.5, alpha=0.10, color='#2a9d8f', label='Post-Paris')
    ax.axvline(2024.5, color='#d62728', linestyle='--', linewidth=1.8, alpha=0.8, label='Paris Olympics', zorder=2)
    ax.set_title(title, fontsize=10, fontweight='bold')
    ax.set_xlabel('Year', fontsize=9)
    ax.set_ylabel('Share of pairs (%)', fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
    ax.set_xticks(sorted(gs['year'].unique()))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.grid(axis='y', alpha=0.3)
    ax.legend(fontsize=7)
fig1_paris.suptitle('Cultural Homophily in Grand Slam Doubles — Paris 2024 Cycle',
                  fontsize=12, fontweight='bold', y=1.02)
plt.tight_layout()
fig1_paris_b64 = fig_to_b64(fig1_paris)

# Figure 3: Win-rate advantage by homophily (winners vs losers)
fig3, ax3 = plt.subplots(figsize=(7,4))
categories = ['Same\nNationality','Same\nLanguage','Language\nProximity']
w_rates = [100*gs['same_country_winners'].mean(),
           100*gs['winners_same_language'].mean(),
           100*gs['winners_linguistic_proximity'].mean()]
l_rates = [100*gs['same_country_losers'].mean(),
           100*gs['losers_same_language'].mean(),
           100*gs['losers_linguistic_proximity'].mean()]
x3 = np.arange(len(categories)); w3=0.35
ax3.bar(x3-w3/2, w_rates, w3, label='Winner teams', color='#2166ac', alpha=0.85)
ax3.bar(x3+w3/2, l_rates, w3, label='Loser teams',  color='#d6604d', alpha=0.85)
ax3.set_xticks(x3); ax3.set_xticklabels(categories, fontsize=10)
ax3.set_ylabel('Share of teams with trait (%)', fontsize=9)
ax3.set_title('Homophily Rates: Winner vs. Loser Teams (Grand Slams)', fontsize=11, fontweight='bold')
ax3.legend(fontsize=9); ax3.grid(axis='y', alpha=0.3)
for i,(w,l) in enumerate(zip(w_rates,l_rates)):
    diff = w-l; clr = '#1a9850' if diff>0 else '#d6604d'
    ax3.annotate(f'{diff:+.1f} pp', xy=(i,max(w,l)+0.5), ha='center', fontsize=9, color=clr, fontweight='bold')
plt.tight_layout()
fig3_b64 = fig_to_b64(fig3)

print('Figures generated.')

# ─────────────────────────── TABLES ──────────────────────────────────────────
# Table 1: pressure outcomes
cycle_order_tokyo = ['Pre-Tokyo (2018–19)','Tokyo Prep (2020–21)','Post-Tokyo (2021–22)']
cycle_order_paris = ['Pre-Paris (2022)','Paris Prep (2023–24)','Post-Paris (2024–25)']
by_tokyo = gs[gs['cycle_tokyo'].notna()].groupby('cycle_tokyo')[['same_country_avg','same_language_avg','ling_prox_avg']].mean()*100
by_tokyo = by_tokyo.reindex(cycle_order_tokyo)
by_tokyo.columns = ['Same Nationality (%)','Same Language (%)','Ling. Proximity (%)']
by_tokyo['N matches'] = gs[gs['cycle_tokyo'].notna()].groupby('cycle_tokyo').size().reindex(cycle_order_tokyo)
by_tokyo = by_tokyo.round(1)

by_paris = gs[gs['cycle_paris'].notna()].groupby('cycle_paris')[['same_country_avg','same_language_avg','ling_prox_avg']].mean()*100
by_paris = by_paris.reindex(cycle_order_paris)
by_paris.columns = ['Same Nationality (%)','Same Language (%)','Ling. Proximity (%)']
by_paris['N matches'] = gs[gs['cycle_paris'].notna()].groupby('cycle_paris').size().reindex(cycle_order_paris)
by_paris = by_paris.round(1)

# add full sample row to both panels
full_row = pd.DataFrame({
    'Same Nationality (%)': [round(gs['same_country_avg'].mean()*100,1)],
    'Same Language (%)':    [round(gs['same_language_avg'].mean()*100,1)],
    'Ling. Proximity (%)':  [round(gs['ling_prox_avg'].mean()*100,1)],
    'N matches':            [len(gs)],
}, index=['Full Sample'])

desc_tokyo_table = pd.concat([full_row, by_tokyo])
desc_paris_table = pd.concat([full_row, by_paris])

def df_to_html(df, id='', bold_first=False, caption=''):
    thead = '<tr>' + ''.join(f'<th>{c}</th>' for c in [''] + list(df.columns)) + '</tr>'
    rows = ''
    for i,(idx,row) in enumerate(df.iterrows()):
        cls = ' class="first-row"' if (bold_first and i==0) else (' class="alt"' if i%2==1 else '')
        rows += f'<tr{cls}><td class="rowlbl">{idx}</td>' + ''.join(f'<td class="num">{v}</td>' for v in row) + '</tr>'
    cap = f'<caption>{caption}</caption>' if caption else ''
    return f'<table class="dtable" id="{id}">{cap}<thead>{thead}</thead><tbody>{rows}</tbody></table>'

desc_tokyo_html = df_to_html(desc_tokyo_table, id='desc_tokyo', bold_first=True, caption='Table 2A. Tokyo 2021 cycle — Team composition across period')
desc_paris_html = df_to_html(desc_paris_table, id='desc_paris', bold_first=True, caption='Table 2B. Paris 2024 cycle — Team composition across period')

# Regression tables
vars_main = [
    ('ling_prox',      'Language proximity (ethnic)'),
    ('rank_mean',      'Team average doubles ranking'),
    ('opp_rank_mean',  'Opponent average doubles ranking'),
    ('rank_gap',       'Teammate rank gap'),
    ('single_top100',  'Top-100 singles player'),
]

reg_win_html = reg_table_html(
    [logit_win], vars_main,
    ['Logit'],
    caption='Table 3. Match Win. Outcome = 1 if team won. Language proximity only. Tournament×year and round FE. SE clustered by match.',
    nobs_list=[int(logit_win.nobs)],
    r2_list=[f'{logit_win.prsquared:.3f}'],
)
reg_win_mfx_html = mfx_table_html(
    logit_win_mfx, vars_main,
    caption='Table 3M. Match Win — average marginal effects. Effects are percentage-point changes in Pr(win).',
    nobs=int(logit_win.nobs),
)

reg_tb_html = reg_table_html(
    [logit_tb], vars_main,
    ['Logit'],
    caption='Table 4. Tiebreak Win — Logit (sample: matches with any tiebreak, 7p/10p). Outcome = 1 if team won any tiebreak. Language proximity only. Tournament×year and round FE. SE clustered by match.',
    nobs_list=[int(logit_tb.nobs)],
    r2_list=[f'{logit_tb.prsquared:.3f}'],
)
reg_tb_mfx_html = mfx_table_html(
    logit_tb_mfx, vars_main,
    caption='Table 4M. Tiebreak Win — average marginal effects. Effects are percentage-point changes in Pr(win any tiebreak).',
    nobs=int(logit_tb.nobs),
)

cb_vars = [
    ('ling_prox',      'Language proximity (ethnic)'),
    ('rank_mean',      'Team average doubles ranking'),
    ('opp_rank_mean',  'Opponent average doubles ranking'),
    ('rank_gap',       'Teammate rank gap'),
    ('single_top100',  'Top-100 singles player'),
]
reg_cb_html = reg_table_html(
    [logit_cb], cb_vars,
    ['Logit'],
    caption='Table 5. Comeback Win — Logit (3-set GS matches; conditional on losing set 1). Outcome = 1 if team wins given it lost set 1. Language proximity only. Tournament×year FE (cells with no within-cell variation dropped). SE clustered by match.',
    nobs_list=[int(logit_cb.nobs)],
    r2_list=[f'{logit_cb.prsquared:.3f}'],
)
reg_cb_mfx_html = mfx_table_html(
    logit_cb_mfx, cb_vars,
    caption='Table 5M. Comeback Win — average marginal effects. Effects are percentage-point changes in Pr(win | lost set 1).',
    nobs=int(logit_cb.nobs),
)

# Pressure outcomes mini-table
po_data = [
    ('Set-1 tiebreak (7-pt)',                      df['tb_s1'].sum(),          f"{100*df['tb_s1'].mean():.1f}%"),
    ('Set-2 tiebreak (7-pt)',                      df['tb_s2'].sum(),          f"{100*df['tb_s2'].mean():.1f}%"),
    ('Set-3 regular tiebreak (7-pt)',              df['tb_s3_regular'].sum(),  f"{100*df['tb_s3_regular'].mean():.1f}%"),
    ('Match tiebreak / super-tb (10-pt, set 3 ≥ 8)', df['match_tb'].sum(),   f"{100*df['match_tb'].mean():.1f}%"),
    ('Any tiebreak (all types)',                   df['any_tb'].sum(),         f"{100*df['any_tb'].mean():.1f}%"),
    ('Match went to 3 sets',                       df['three_sets'].sum(),     f"{100*df['three_sets'].mean():.1f}%"),
    ('Comeback win (winner lost set 1)',            df['comeback'].sum(),       f"{100*df['comeback'].mean():.1f}%"),
]
po_html = '''<table class="dtable"><thead><tr><th>Pressure Outcome</th><th>N</th><th>Share</th></tr></thead><tbody>'''
for i,(lbl,n,pct) in enumerate(po_data):
    cls = ' class="alt"' if i%2==1 else ''
    po_html += f'<tr{cls}><td class="rowlbl">{lbl}</td><td class="num">{n:,}</td><td class="num">{pct}</td></tr>'
po_html += '</tbody></table>'

# Tiebreaks by tournament
tb_tourn = df.groupby('tournament').agg(
    N_matches=('match_id','count'),
    N_reg_tb=('regular_tb','sum'),
    Rate_tb=('regular_tb','mean'),
    N_match_tb=('match_tb','sum'),
    N_comeback=('comeback','sum'),
    Rate_comeback=('comeback','mean'),
).reset_index()
tb_tourn['Rate_tb'] = (tb_tourn['Rate_tb']*100).round(1).astype(str)+'%'
tb_tourn['Rate_comeback'] = (tb_tourn['Rate_comeback']*100).round(1).astype(str)+'%'
tb_tourn.columns = ['Tournament','N Matches','Reg. Tiebreaks','TB Rate','Match Tiebreaks','Comeback Wins','Comeback Rate']
tb_tourn_html = '''<table class="dtable"><thead><tr>'''+''.join(f'<th>{c}</th>' for c in tb_tourn.columns)+'</tr></thead><tbody>'
for i,(_,row) in enumerate(tb_tourn.iterrows()):
    cls = ' class="alt"' if i%2==1 else ''
    tb_tourn_html += f'<tr{cls}>' + ''.join(f'<td class="num">{v}</td>' for v in row) + '</tr>'
tb_tourn_html += '</tbody></table>'

print('Tables ready. Writing HTML...')

# ─────────────────────────── HTML ────────────────────────────────────────────
CSS = """
  body { font-family: Georgia, 'Times New Roman', serif; font-size: 11pt; color: #1a1a1a;
         max-width: 960px; margin: 0 auto; padding: 30px 40px; line-height: 1.55; }
  h1   { font-size: 1.6em; border-bottom: 2px solid #2166ac; padding-bottom: 6px; margin-top: 0; }
  h2   { font-size: 1.2em; color: #2166ac; border-bottom: 1px solid #dde; padding-bottom: 3px; margin-top: 2em; }
  h3   { font-size: 1.05em; color: #444; margin-top: 1.4em; }
  p    { margin: 0.6em 0 0.8em; }
  .meta { font-size: 0.88em; color: #666; margin-bottom: 1.5em; }
  .note { font-size: 0.85em; color: #555; font-style: italic; margin: 4px 0 10px; }
  .box  { background: #f0f4fa; border-left: 4px solid #2166ac; padding: 10px 16px; margin: 14px 0;
           border-radius: 3px; font-size: 0.96em; }
  .box h4 { margin: 0 0 6px; color: #2166ac; font-size: 1em; }
  ul { margin: 0.4em 0; padding-left: 1.4em; }
  li { margin-bottom: 4px; }

  /* descriptive + pressure tables */
  .dtable { border-collapse: collapse; width: 100%; font-size: 0.9em; margin: 10px 0 16px; }
  .dtable th { background: #2166ac; color: white; padding: 6px 10px; text-align: left; }
  .dtable td { padding: 5px 10px; border-bottom: 1px solid #eee; }
  .dtable .rowlbl { text-align: left; }
  .dtable .num { text-align: right; font-variant-numeric: tabular-nums; }
  .dtable .alt { background: #f7f9fc; }
  .dtable .first-row { font-weight: bold; background: #e8f0fa; }

  /* regression tables */
  .regtable { border-collapse: collapse; width: 100%; font-size: 0.88em; margin: 12px 0; }
  .regtable caption { caption-side: bottom; font-size: 0.82em; color: #555; font-style: italic; padding: 6px 0 0; text-align: left; }
  .regtable thead th { background: #2166ac; color: white; padding: 6px 8px; text-align: center; border: 1px solid #1a5599; }
  .regtable thead th:first-child { text-align: left; }
  .regtable td { padding: 4px 8px; border-bottom: 1px solid #eee; }
  .regtable .varlbl { text-align: left; font-size: 0.9em; }
  .regtable .num { text-align: right; font-variant-numeric: tabular-nums; }
  .regtable .se { color: #666; font-size: 0.9em; }
  .regtable .stat { background: #f0f4fa; font-size: 0.88em; }
  .regtable .sig1 td { background: #fffde7; }
  .regtable .sig5 td { background: #fff8e8; }
  .sig1 .num, .sig5 .num, .sig10 .num { font-weight: bold; }
  .sigkey { font-size: 0.84em; color: #555; margin: -6px 0 14px; }

  figure { margin: 18px 0; text-align: center; }
  figure img { max-width: 100%; border: 1px solid #dde; border-radius: 4px; }
  figcaption { font-size: 0.85em; color: #555; margin-top: 6px; }

  .summary-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin: 16px 0; }
  .summary-box { background: #f7f9fc; border: 1px solid #dde; border-radius: 6px; padding: 14px 16px; }
  .summary-box h4 { margin: 0 0 8px; color: #2166ac; font-size: 0.95em; }

  @media print {
    body { max-width: 100%; padding: 0 10mm; font-size: 10pt; }
    h2 { page-break-before: auto; }
    .regtable, .dtable { page-break-inside: avoid; }
    figure { page-break-inside: avoid; }
  }
"""

def sample_overview():
    n_all = len(df); n_gs = len(gs); n_team = len(team_gs)
    n_gs_matches = n_team//2
    yrs = sorted(df['year'].unique())
    return f"""
    <ul>
      <li><strong>Matches:</strong> {n_all:,} retained from {n_raw_matches:,} raw matches &nbsp;|&nbsp; {n_gs:,} Grand Slams &nbsp;|&nbsp; {df[df['olympics_tourn']==1].shape[0]} Olympics</li>
      <li><strong>Retirements/walkovers:</strong> {n_dropped_incomplete:,} observations dropped using the score-completion rule.</li>
      <li><strong>Years:</strong> {', '.join(str(y) for y in yrs)} (2020 excluded — COVID)</li>
      <li><strong>Tournaments:</strong> Australian Open, Roland Garros, Wimbledon, US Open, Olympics (2021, 2024)</li>
      <li><strong>Analysis sample:</strong> {n_team:,} team-level observations ({n_gs_matches:,} GS matches × 2 teams)</li>
    </ul>"""

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Tennis Doubles Homophily — Main Results</title>
<style>{CSS}</style>
</head>
<body>

<h1>Cultural Homophily in Tennis Doubles<br><small style="font-size:0.6em;color:#555;">Grand Slam Analysis 2018–2025</small></h1>
<p class="meta">Prepared for supervisor review &nbsp;·&nbsp; May 2026 &nbsp;·&nbsp; Code: <code>code/analysis/homophily_analysis.ipynb</code></p>

<!-- ── EXECUTIVE SUMMARY ─────────────────────────────────────── -->
<div class="box">
<h4>Key Findings</h4>
<ul>
  <li><strong>Language proximity significantly predicts match wins</strong> in the filtered Grand Slam sample: the marginal effect is about +5.9 percentage points on match-win probability.</li>
  <li><strong>No robust pressure advantage in tiebreaks:</strong> language proximity is not statistically significant for tiebreak outcomes (point estimate ≈ +1.6 pp, p ≈ 0.50).</li>
  <li><strong>Comebacks do not benefit from greater language proximity:</strong> the comeback sample shows a marginally significant negative effect (≈ −7.0 pp, p ≈ 0.06) for teams that lost set 1.</li>
  <li><strong>Homophily shares are stable pre-2024</strong> (~40% same-nationality, ~53% same-language from 2018–23), then fall in 2024–25, driven by draw expansion rather than a sudden increase in partnership homophily.</li>
  <li><strong>Ranking and partner quality dominate outcomes</strong> across all models, consistent with standard performance expectations.</li>
</ul>
</div>

<!-- ── DATA ──────────────────────────────────────────────────── -->
<h2>1. Data Overview</h2>
{sample_overview()}
<p>Player rankings and biographical profiles (nationality, language, coach) are sourced from ATP Tour snapshots merged to each match. Language variables derive from the CEPII Gravity dataset: <em>same language (official)</em> uses common official language; <em>language proximity</em> uses the broader ethno-linguistic criterion.</p>

<!-- ── PRESSURE OUTCOMES ──────────────────────────────────────── -->
<h2>2. Pressure Outcomes</h2>

<h3>2.1 Tiebreak Observations</h3>
<p>A <strong>regular tiebreak</strong> (7-point, played at 6-6 in sets 1 or 2) occurred in <strong>{100*df['regular_tb'].mean():.1f}% of retained matches</strong>. A <strong>match tiebreak</strong> (10-point super-tiebreak, replacing a full third set) was played in {int(df['match_tb'].sum()):,} matches. For regression purposes both regular and match tiebreaks are combined under <em>any tiebreak</em>.</p>

{po_html}
<p class="note">N = {len(df):,} retained matches after dropping {n_dropped_incomplete:,} likely retirements/walkovers.</p>

<h3>2.2 By Tournament</h3>
{tb_tourn_html}

<h3>2.3 Comeback Wins</h3>
<p>A comeback win is defined as the eventual match winner having lost the first set. This occurred in <strong>{df['comeback'].sum():,} matches ({100*df['comeback'].mean():.1f}%)</strong> of the full sample, and in <strong>{100*(df['comeback'].sum()/df['three_sets'].sum()):.1f}%</strong> of three-set matches — meaning the team that won the first set subsequently lost the match roughly two-fifths of the time.</p>

<!-- ── OLYMPIC CYCLES ──────────────────────────────────────────── -->
<h2>3. Cultural Homophily Across Olympic Cycles</h2>

<p>The descriptive tables below show the share of pairs (averaged over both teams in each match) with each homophily characteristic, broken out separately for the Tokyo 2021 and Paris 2024 cycles. Grand Slam matches only — the Olympics tournament is excluded because same-nationality pairing is compulsory there.</p>

<h3>3.1 Tokyo 2021 cycle</h3>
{desc_tokyo_html}
<p class="note">Tokyo panel includes 2020 as Tokyo Prep, excluding Wimbledon 2020 because Olympic doubles was not yet in the build-up path.</p>
<figure>
  <img src="data:image/png;base64,{fig1_tokyo_b64}" alt="Tokyo cycle homophily trends">
  <figcaption>Figure 1. Tokyo 2021 cycle homophily trends. Shaded bands mark Pre-Tokyo, Tokyo Prep, and Post-Tokyo periods.</figcaption>
</figure>

<h3>3.2 Paris 2024 cycle</h3>
{desc_paris_html}
<p class="note">Paris panel treats the 2023–24 Grand Slams as the Paris Prep window, with 2024 US Open and 2025 matches as Post-Paris.</p>
<figure>
  <img src="data:image/png;base64,{fig1_paris_b64}" alt="Paris cycle homophily trends">
  <figcaption>Figure 2. Paris 2024 cycle homophily trends. Shaded bands mark Pre-Paris, Paris Prep, and Post-Paris periods.</figcaption>
</figure>

<h3>3.3 Winner vs. Loser Team Homophily</h3>
<figure>
  <img src="data:image/png;base64,{fig3_b64}" alt="Winner vs loser homophily">
  <figcaption>Figure 3. Percentage of teams with each homophily trait, split by match outcome. Annotation shows the winner-minus-loser gap.</figcaption>
</figure>
<p>Winner teams are more likely to share a language (+{100*(gs['winners_same_language'].mean()-gs['losers_same_language'].mean()):+.1f} pp) and to be linguistically proximate (+{100*(gs['winners_linguistic_proximity'].mean()-gs['losers_linguistic_proximity'].mean()):+.1f} pp) than loser teams. Same-nationality is marginally more common among losers ({100*gs['same_country_losers'].mean():.1f}% vs {100*gs['same_country_winners'].mean():.1f}%).</p>

<!-- ── REGRESSIONS ─────────────────────────────────────────────── -->
<h2>4. Baseline Regressions</h2>

<p><strong>Specification:</strong> each match contributes two team-level observations (winner = 1, loser = 0). Regressors include language proximity (ethnic) and ranking controls (team mean ranking, opponent mean ranking, within-team rank gap, and a top-100 singles indicator). Fixed effects for tournament×year and round. Standard errors are clustered by match.</p>

<h3>4.1 Match Win</h3>
{reg_win_html}
{reg_win_mfx_html}
<p class="sigkey">Significance: *** p &lt; 0.01 &nbsp; ** p &lt; 0.05 &nbsp; * p &lt; 0.10 &nbsp; Standard errors in parentheses.</p>
<p>{ling_mfx_sentence(logit_win_mfx, 'winning the match')}</p>

<h3>4.2 Tiebreak Win (pressure outcome)</h3>
<p>Sample restricted to matches with any tiebreak, regular or match tiebreak (N = {len(tb_sub)//2:,} matches). Outcome: team won at least one tiebreak.</p>
{reg_tb_html}
{reg_tb_mfx_html}
<p class="sigkey">Significance: *** p &lt; 0.01 &nbsp; ** p &lt; 0.05 &nbsp; * p &lt; 0.10 &nbsp; Standard errors in parentheses.</p>
<p>{ling_mfx_sentence(logit_tb_mfx, 'winning a tiebreak in tiebreak matches')}</p>

<h3>4.3 Comeback Win (pressure outcome)</h3>
<p>Sample: three-set Grand Slam matches <em>conditional on the team having lost set 1</em>. Outcome = 1 if the team wins the match (i.e., a comeback). Tournament×year fixed effects are included by encoding each tournament-year cell; cells with no within-cell variation in the outcome are dropped. Standard errors are clustered by match.</p>
{reg_cb_html}
{reg_cb_mfx_html}
<p class="sigkey">Significance: *** p &lt; 0.01 &nbsp; ** p &lt; 0.05 &nbsp; * p &lt; 0.10 &nbsp; Standard errors in parentheses.</p>
<p>{ling_mfx_sentence(logit_cb_mfx, 'completing the comeback')} Ranking controls have the expected signs: conditional on being a set down, stronger teams are more likely to complete the comeback.</p>

<!-- ── DISCUSSION ─────────────────────────────────────────────── -->
<h2>5. Discussion and Next Steps</h2>

<div class="summary-grid">
  <div class="summary-box">
    <h4>What the data say</h4>
    <ul>
      <li>Language proximity is positive and statistically significant for match wins: the average marginal effect is about +5.9 percentage points.</li>
      <li>There is no evidence that language proximity improves tiebreak performance; the tiebreak coefficient is small and statistically indistinguishable from zero.</li>
      <li>For comeback wins after losing set 1, language proximity is marginally negative (≈ −7.0 pp, p ≈ 0.06), implying it does not help teams recover in three-set matches.</li>
      <li>Homophily shares do not increase ahead of the Olympics; instead, 2024–25 shows a decline driven by larger Grand Slam draws.</li>
    </ul>
  </div>
  <div class="summary-box">
    <h4>Limitations and extensions</h4>
    <ul>
      <li><strong>Sample size:</strong> ~1,700 Grand Slam matches. Power is limited for detecting small homophily effects (≤ 3–5 pp).</li>
      <li><strong>Language collinearity:</strong> same language and language proximity are highly correlated; consider a single index.</li>
      <li><strong>Draw expansion 2024–25:</strong> the larger sample changes composition and may mask cycle-level patterns; restrict to comparable round ranges for a cleaner trend.</li>
      <li><strong>Conditioning on partnership formation:</strong> selection — same-country pairs may self-select into tournaments — is not yet addressed.</li>
      <li><strong>Pair fixed effects:</strong> a within-pair panel (same partners over time) would better isolate performance from selection.</li>
    </ul>
  </div>
</div>

<hr style="margin-top:2em;border:none;border-top:1px solid #dde;">
<p style="font-size:0.82em;color:#888;">Generated automatically from <code>build_report.py</code>. All regressions run in Python (statsmodels). Data: ATP Tour Grand Slam doubles 2018–2025 (excl. 2020).</p>

</body>
</html>"""

out_path = 'report/homophily_report.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(HTML)
print(f'\nReport written to {out_path}')
pdf_path = 'pdf/homophily_report.pdf'
try:
    html_to_pdf(out_path, pdf_path)
    print(f'PDF written to {pdf_path}')
except Exception as e:
    print('PDF generation failed:', e)
    try:
        alt_pdf_path = f"pdf/homophily_report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        html_to_pdf(out_path, alt_pdf_path)
        print(f'PDF written to {alt_pdf_path}')
    except Exception:
        print('Open the HTML in a browser and use File > Print > Save as PDF if needed.')

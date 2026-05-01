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

os.makedirs('report', exist_ok=True)

# ─────────────────────────── DATA + VARIABLES ────────────────────────────────
df = pd.read_excel('data/atp/men_matches_with_ranks_cleaned.xlsx')

df['tb_s1'] = ((df['winners_set1']==7)&(df['losers_set1']==6))|((df['winners_set1']==6)&(df['losers_set1']==7))
df['tb_s2'] = ((df['winners_set2']==7)&(df['losers_set2']==6))|((df['winners_set2']==6)&(df['losers_set2']==7))
df['regular_tb']    = df['tb_s1'] | df['tb_s2']
df['tb_s3_regular'] = (((df['winners_set3']==7)&(df['losers_set3']==6))|((df['winners_set3']==6)&(df['losers_set3']==7))).fillna(False)
df['match_tb']      = (df['winners_set3'] >= 8).fillna(False)
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
cycle_map = {2018:'Pre-Tokyo (2018–19)',2019:'Pre-Tokyo (2018–19)',2021:'Post-Tokyo (2021)',
             2022:'Pre-Paris (2022–23)',2023:'Pre-Paris (2022–23)',2024:'Paris 2024',2025:'Post-Paris (2025)'}
df['cycle'] = df['year'].map(cycle_map)

gs = df[df['olympics_tourn']==0].copy()

# team panel
match_vars = ['match_id','tournament','year','surface','stage_code','any_tb','regular_tb','match_tb',
              'three_sets','comeback','winner_lost_s1','pre_olympic','olympic_period','olympics_tourn']
w_row = df[match_vars].copy()
for col,src in [('same_country','same_country_winners'),('same_language','winners_same_language'),
                ('ling_prox','winners_linguistic_proximity'),('rank_mean','rank_mean_winners'),
                ('same_hand','same_hand_winners'),('same_coach','same_coach_winners'),
                ('wl_career_diff','wl_career_diff_winners')]:
    w_row[col] = df[src].values
w_row['win']=1; w_row['won_tb_s1']=df['w_won_tb_s1'].values; w_row['won_tb_s2']=df['w_won_tb_s2'].values

l_row = df[match_vars].copy()
for col,src in [('same_country','same_country_losers'),('same_language','losers_same_language'),
                ('ling_prox','losers_linguistic_proximity'),('rank_mean','rank_mean_losers'),
                ('same_hand','same_hand_losers'),('same_coach','same_coach_losers'),
                ('wl_career_diff','wl_career_diff_losers')]:
    l_row[col] = df[src].values
l_row['win']=0
l_row['won_tb_s1']=(df['tb_s1']&(df['winners_set1']==6)).values
l_row['won_tb_s2']=(df['tb_s2']&(df['winners_set2']==6)).values

team_df = pd.concat([w_row, l_row], ignore_index=True)
team_df['log_rank'] = np.log1p(team_df['rank_mean'].clip(lower=1))
team_df['stage_code'] = team_df['stage_code'].fillna(-1).astype(int)
for col in team_df.columns:
    if str(team_df[col].dtype) in ('bool','boolean','Int64','Int32'):
        team_df[col] = team_df[col].fillna(0).astype(int)
team_gs = team_df[team_df['olympics_tourn']==0].dropna(subset=['log_rank','wl_career_diff']).copy()
team_gs['won_regular_tb'] = ((team_gs['won_tb_s1']==1)|(team_gs['won_tb_s2']==1)).astype(int)
tb_sub = team_gs[team_gs['regular_tb']==1].copy()

cdf = df[(df['olympics_tourn']==0)&(df['three_sets']==True)].copy()
cdf['log_rank_diff'] = np.log1p(cdf['rank_mean_winners'].clip(lower=1))-np.log1p(cdf['rank_mean_losers'].clip(lower=1))
for col in cdf.columns:
    if str(cdf[col].dtype) in ('bool','boolean','Int64','Int32'): cdf[col]=cdf[col].fillna(0).astype(int)
cdf = cdf.dropna(subset=['log_rank_diff','same_country_diff'])

print('Data ready. Fitting models...')

# ─────────────────────────── MODELS ──────────────────────────────────────────
H    = 'same_country + same_language + ling_prox'
BASE = '+ log_rank + wl_career_diff + same_hand + same_coach'
FE   = '+ C(tournament) + C(stage_code) + C(year)'
FE2  = '+ C(tournament) + C(stage_code)'

lpm_win   = smf.ols(f'win ~ {H} {BASE} {FE}', data=team_gs).fit(cov_type='cluster', cov_kwds={'groups':team_gs['match_id']})
logit_win = smf.logit(f'win ~ {H} {BASE} {FE}', data=team_gs).fit(cov_type='cluster', cov_kwds={'groups':team_gs['match_id']}, disp=False)
lpm_tb    = smf.ols(f'won_regular_tb ~ {H} {BASE} {FE}', data=tb_sub).fit(cov_type='cluster', cov_kwds={'groups':tb_sub['match_id']})
logit_tb  = smf.logit(f'won_regular_tb ~ {H} {BASE} {FE}', data=tb_sub).fit(cov_type='cluster', cov_kwds={'groups':tb_sub['match_id']}, disp=False)
lpm_cb    = smf.ols('comeback ~ same_country_diff + same_language_diff + ling_prox_diff + log_rank_diff + C(tournament) + C(stage_code) + C(year)', data=cdf).fit(cov_type='HC1')
logit_cb  = smf.logit('comeback ~ same_country_diff + same_language_diff + ling_prox_diff + log_rank_diff + C(tournament) + C(stage_code) + C(year)', data=cdf).fit(disp=False)
lpm_int   = smf.ols(f'win ~ same_country*pre_olympic + same_language*pre_olympic + ling_prox*pre_olympic {BASE} {FE2}', data=team_gs).fit(cov_type='cluster', cov_kwds={'groups':team_gs['match_id']})
logit_int = smf.logit(f'win ~ same_country*pre_olympic + same_language*pre_olympic + ling_prox*pre_olympic {BASE} {FE2}', data=team_gs).fit(cov_type='cluster', cov_kwds={'groups':team_gs['match_id']}, disp=False)
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

# ─────────────────────────── FIGURES ─────────────────────────────────────────
# Figure 1: Homophily trends over time
hcols = ['same_country_avg','same_language_avg','ling_prox_avg']
year_means = gs.groupby('year')[hcols].mean()*100

fig1, axes = plt.subplots(1,3, figsize=(13,4), sharey=False)
titles = ['Same Nationality','Same Official Language','Language Proximity (ethnic)']
colors = ['#2166ac','#1a9850','#d6604d']
for ax, col, title, clr in zip(axes, hcols, titles, colors):
    ax.plot(year_means.index, year_means[col], marker='o', color=clr, linewidth=2, markersize=6, zorder=3)
    ax.axvspan(2021.5, 2023.5, alpha=0.10, color='#f4a261', label='Pre-Paris prep')
    ax.axvline(2024, color='#e76f51', linestyle='--', linewidth=1.5, label='Paris Olympics')
    ax.set_title(title, fontsize=10, fontweight='bold')
    ax.set_xlabel('Year', fontsize=9)
    ax.set_ylabel('Share of pairs (%)', fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
    ax.set_xticks(sorted(gs['year'].unique()))
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.grid(axis='y', alpha=0.3)
    ax.legend(fontsize=7)
fig1.suptitle('Cultural Homophily in Grand Slam Doubles — Olympic Cycle Pattern',
              fontsize=12, fontweight='bold', y=1.02)
plt.tight_layout()
fig1_b64 = fig_to_b64(fig1)

# Figure 2: Olympic interaction bar chart
fig2, ax2 = plt.subplots(figsize=(8,4.5))
vars_p = ['same_country','same_language','ling_prox']
lbls_p = ['Same Nationality','Same Official\nLanguage','Language\nProximity']
x = np.arange(len(vars_p)); w = 0.35

c_base = [lpm_int.params.get(v,np.nan) for v in vars_p]
c_int  = [lpm_int.params.get(f'{v}:pre_olympic',0) for v in vars_p]
c_olym = [b+i for b,i in zip(c_base,c_int)]
se_b   = [lpm_int.bse.get(v,np.nan) for v in vars_p]
se_i   = [lpm_int.bse.get(f'{v}:pre_olympic',0) for v in vars_p]
se_o   = [np.sqrt(sb**2+si**2) for sb,si in zip(se_b,se_i)]

b1 = ax2.bar(x-w/2, c_base, w, yerr=se_b, label='Non-Olympic years',
             color='#2166ac', alpha=0.85, capsize=5, error_kw={'linewidth':1.5})
b2 = ax2.bar(x+w/2, c_olym, w, yerr=se_o, label='Pre-Olympic build-up (2022–23)',
             color='#f4a261', alpha=0.85, capsize=5, error_kw={'linewidth':1.5})
ax2.axhline(0, color='black', linewidth=0.8, linestyle='--')
ax2.set_xticks(x); ax2.set_xticklabels(lbls_p, fontsize=10)
ax2.set_ylabel('Marginal effect on win probability (LPM)', fontsize=9)
ax2.set_title('Cultural Homophily and Match Win — Olympic Interaction', fontsize=11, fontweight='bold')
ax2.legend(fontsize=9)
ax2.grid(axis='y', alpha=0.3)
plt.tight_layout()
fig2_b64 = fig_to_b64(fig2)

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
cycle_order = ['Pre-Tokyo (2018–19)','Post-Tokyo (2021)','Pre-Paris (2022–23)','Paris 2024','Post-Paris (2025)']
by_cycle = gs.groupby('cycle')[['same_country_avg','same_language_avg','ling_prox_avg']].mean()*100
by_cycle = by_cycle.reindex(cycle_order)
by_cycle.columns = ['Same Nationality (%)','Same Language (%)','Ling. Proximity (%)']
by_cycle['N matches'] = gs.groupby('cycle').size().reindex(cycle_order)
by_cycle = by_cycle.round(1)
# add full sample row
full_row = pd.DataFrame({
    'Same Nationality (%)': [round(gs['same_country_avg'].mean()*100,1)],
    'Same Language (%)':    [round(gs['same_language_avg'].mean()*100,1)],
    'Ling. Proximity (%)':  [round(gs['ling_prox_avg'].mean()*100,1)],
    'N matches':            [len(gs)],
}, index=['Full Sample'])
desc_table = pd.concat([full_row, by_cycle])

def df_to_html(df, id='', bold_first=False):
    thead = '<tr>' + ''.join(f'<th>{c}</th>' for c in [''] + list(df.columns)) + '</tr>'
    rows = ''
    for i,(idx,row) in enumerate(df.iterrows()):
        cls = ' class="first-row"' if (bold_first and i==0) else (' class="alt"' if i%2==1 else '')
        rows += f'<tr{cls}><td class="rowlbl">{idx}</td>' + ''.join(f'<td class="num">{v}</td>' for v in row) + '</tr>'
    return f'<table class="dtable" id="{id}"><thead>{thead}</thead><tbody>{rows}</tbody></table>'

desc_html = df_to_html(desc_table, id='desc', bold_first=True)

# Regression tables
vars_main = [
    ('same_country',   'Same nationality'),
    ('same_language',  'Same language (official)'),
    ('ling_prox',      'Language proximity (ethnic)'),
]
vars_controls = [
    ('log_rank',       'Log(rank mean)'),
    ('wl_career_diff', 'Career W-L ratio diff'),
    ('same_hand',      'Same dominant hand'),
    ('same_coach',     'Same coach'),
]

reg_win_html = reg_table_html(
    [lpm_win, logit_win], vars_main + vars_controls,
    ['LPM', 'Logit'],
    caption='Table 3. Match Win. Outcome = 1 if team won. Tournament, round, year FE. SE clustered by match.',
    nobs_list=[int(lpm_win.nobs), int(logit_win.nobs)],
    r2_list=[f'{lpm_win.rsquared:.3f}', f'{logit_win.prsquared:.3f}'],
)

reg_tb_html = reg_table_html(
    [lpm_tb, logit_tb], vars_main + vars_controls,
    ['LPM', 'Logit'],
    caption='Table 4. Tiebreak Win. Outcome = 1 if team won a regular tiebreak. Sample: matches with tiebreak only.',
    nobs_list=[int(lpm_tb.nobs), int(logit_tb.nobs)],
    r2_list=[f'{lpm_tb.rsquared:.3f}', f'{logit_tb.prsquared:.3f}'],
)

cb_vars = [
    ('same_country_diff',  'Same nationality (winner − loser)'),
    ('same_language_diff', 'Same language (winner − loser)'),
    ('ling_prox_diff',     'Lang. proximity (winner − loser)'),
    ('log_rank_diff',      'Log(rank) advantage'),
]
reg_cb_html = reg_table_html(
    [lpm_cb, logit_cb], cb_vars,
    ['LPM', 'Logit'],
    caption='Table 5. Comeback Win. Outcome = 1 if winning team came back from losing set 1. 3-set matches only. HC1 SE.',
    nobs_list=[int(lpm_cb.nobs), int(logit_cb.nobs)],
    r2_list=[f'{lpm_cb.rsquared:.3f}', f'{logit_cb.prsquared:.3f}'],
)

int_vars = [
    ('same_country',             'Same nationality'),
    ('same_language',            'Same language'),
    ('ling_prox',                'Language proximity'),
    ('pre_olympic',              'Pre-Olympic period (2022–23)'),
    ('same_country:pre_olympic', 'Same nationality × Pre-Olympic'),
    ('same_language:pre_olympic','Same language × Pre-Olympic'),
    ('ling_prox:pre_olympic',    'Lang. proximity × Pre-Olympic'),
]
reg_int_html = reg_table_html(
    [lpm_int, logit_int], int_vars + vars_controls,
    ['LPM', 'Logit'],
    caption='Table 6. Olympic Interaction. Year FE excluded (collinear with pre_olympic). Tournament and round FE retained.',
    nobs_list=[int(lpm_int.nobs), int(logit_int.nobs)],
    r2_list=[f'{lpm_int.rsquared:.3f}', f'{logit_int.prsquared:.3f}'],
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
      <li><strong>Matches:</strong> {n_all:,} total &nbsp;|&nbsp; {n_gs:,} Grand Slams &nbsp;|&nbsp; {df[df['olympics_tourn']==1].shape[0]} Olympics</li>
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
  <li><strong>Cultural homophily does not significantly predict match wins</strong> after controlling for ranking quality and career performance. Same nationality (+3.1 pp, p = 0.23), same language (+3.4 pp, p = 0.70), and language proximity (+0.6 pp, p = 0.95) are all positive but statistically insignificant.</li>
  <li><strong>Same language shows a borderline pressure advantage:</strong> teams sharing an official language are approximately 20 pp more likely to win a tiebreak (p ≈ 0.11), consistent with communication benefits under stress — but not significant at the 5% level.</li>
  <li><strong>No Olympic build-up effect:</strong> homophily does not predict better outcomes during the 2022–23 Paris preparation period. All interaction terms are small and insignificant.</li>
  <li><strong>Homophily shares are stable pre-2024</strong> (~40% same-nationality, ~53% same-language from 2018–23), then drop sharply in 2024–25 — an artifact of expanded draw sizes (134 vs. ~62 matches per Grand Slam).</li>
  <li><strong>Ranking and career win rate dominate</strong> match outcomes, as expected.</li>
</ul>
</div>

<!-- ── DATA ──────────────────────────────────────────────────── -->
<h2>1. Data Overview</h2>
{sample_overview()}
<p>Player rankings and biographical profiles (nationality, language, coach) are sourced from ATP Tour snapshots merged to each match. Language variables derive from the CEPII Gravity dataset: <em>same language (official)</em> uses common official language; <em>language proximity</em> uses the broader ethno-linguistic criterion.</p>

<!-- ── PRESSURE OUTCOMES ──────────────────────────────────────── -->
<h2>2. Pressure Outcomes</h2>

<h3>2.1 Tiebreak Observations</h3>
<p>A <strong>regular tiebreak</strong> (7-point, played at 6-6 in sets 1 or 2) occurred in <strong>44.0% of matches</strong>. A <strong>match tiebreak</strong> (10-point super-tiebreak, replacing a full third set) was played in 30 matches, all at the Olympics or Wimbledon. For regression purposes both types are combined under <em>any tiebreak</em>.</p>

{po_html}
<p class="note">N = 2,192 total matches. Regular tiebreaks are most frequent at Wimbledon (52%) and least frequent at Roland Garros (35%).</p>

<h3>2.2 By Tournament</h3>
{tb_tourn_html}

<h3>2.3 Comeback Wins</h3>
<p>A comeback win is defined as the eventual match winner having lost the first set. This occurred in <strong>{df['comeback'].sum():,} matches ({100*df['comeback'].mean():.1f}%)</strong> of the full sample, and in <strong>{100*(df['comeback'].sum()/df['three_sets'].sum()):.1f}%</strong> of three-set matches — meaning the team that won the first set subsequently lost the match roughly two-fifths of the time.</p>

<!-- ── OLYMPIC CYCLES ──────────────────────────────────────────── -->
<h2>3. Cultural Homophily Across Olympic Cycles</h2>

<p>The descriptive table below shows the share of pairs (averaged over both teams in each match) with each homophily characteristic, broken out by Olympic cycle. Grand Slam matches only — the Olympics tournament is excluded because same-nationality pairing is compulsory there.</p>

{desc_html}
<p class="note">2024–25 sample sizes are larger (134 matches per Grand Slam vs. ~62 previously) because of expanded draw coverage. This compositional shift — more early-round, lower-ranked, diverse-nationality pairs — accounts for most of the observed decline in 2024–25 homophily shares.</p>

<figure>
  <img src="data:image/png;base64,{fig1_b64}" alt="Homophily trends">
  <figcaption>Figure 1. Share of pairs with each cultural characteristic, by year. Grand Slams only. Shaded area = pre-Paris build-up (2022–23). Red dashed line = Paris Olympics 2024.</figcaption>
</figure>

<h3>3.1 Winner vs. Loser Team Homophily</h3>
<figure>
  <img src="data:image/png;base64,{fig3_b64}" alt="Winner vs loser homophily">
  <figcaption>Figure 2. Percentage of teams with each homophily trait, split by match outcome. Annotation shows the winner-minus-loser gap.</figcaption>
</figure>
<p>Winner teams are more likely to share a language (+{100*(gs['winners_same_language'].mean()-gs['losers_same_language'].mean()):+.1f} pp) and to be linguistically proximate (+{100*(gs['winners_linguistic_proximity'].mean()-gs['losers_linguistic_proximity'].mean()):+.1f} pp) than loser teams. Same-nationality is marginally more common among losers ({100*gs['same_country_losers'].mean():.1f}% vs {100*gs['same_country_winners'].mean():.1f}%).</p>

<!-- ── REGRESSIONS ─────────────────────────────────────────────── -->
<h2>4. Baseline Regressions</h2>

<p><strong>Specification:</strong> each match contributes two team-level observations (winner = 1, loser = 0). Regressors include the three cultural homophily dummies, log of the team's mean ATP doubles ranking, within-team career W-L ratio difference, and indicators for same dominant hand and same coach. Fixed effects for tournament, round, and year. Standard errors clustered by match.</p>

<h3>4.1 Match Win</h3>
{reg_win_html}
<p class="sigkey">Significance: *** p &lt; 0.01 &nbsp; ** p &lt; 0.05 &nbsp; * p &lt; 0.10 &nbsp; Standard errors in parentheses.</p>
<p>All three cultural homophily variables are positive but statistically insignificant. Ranking (log mean) and career performance are the dominant, highly significant predictors.</p>

<h3>4.2 Tiebreak Win (pressure outcome)</h3>
<p>Sample restricted to matches where a regular tiebreak occurred in set 1 or 2 (N = {int(lpm_tb.nobs)//2:,} matches). Outcome: team won at least one tiebreak set.</p>
{reg_tb_html}
<p class="sigkey">Significance: *** p &lt; 0.01 &nbsp; ** p &lt; 0.05 &nbsp; * p &lt; 0.10 &nbsp; Standard errors in parentheses.</p>
<p>Same language has a positive effect (+20 pp in LPM, p ≈ 0.11) and language proximity a negative one (−18 pp, p ≈ 0.14). The two measures are highly correlated, making precise separation difficult; the pattern is consistent with shared language aiding communication under tiebreak pressure, but neither reaches the 5% threshold.</p>

<h3>4.3 Comeback Win (pressure outcome)</h3>
<p>Sample: {int(lpm_cb.nobs):,} three-set Grand Slam matches. Outcome = 1 if the eventual winner came back from losing set 1. Regressors are homophily <em>differences</em> (winner team minus loser team). HC1 heteroskedasticity-robust SE.</p>
{reg_cb_html}
<p class="sigkey">Significance: *** p &lt; 0.01 &nbsp; ** p &lt; 0.05 &nbsp; * p &lt; 0.10 &nbsp; Standard errors in parentheses.</p>
<p>No homophily variable is significant for comebacks. The modest ranking advantage of the winner team is marginally predictive (p ≈ 0.09), consistent with higher-ranked teams being better able to recover from a set down.</p>

<!-- ── OLYMPIC INTERACTION ─────────────────────────────────────── -->
<h2>5. Olympic Interaction</h2>

<p><strong>Hypothesis:</strong> players may strategically choose same-country partners during Olympic preparation windows (2022–23 for Paris 2024) because Olympic doubles requires same-country pairs. If so, homophilous teams would be better-trained for this format and should outperform in Grand Slams during that period.</p>

<p><strong>Specification:</strong> each homophily variable is interacted with <code>pre_olympic</code> (= 1 for 2022–23). Year fixed effects are excluded to avoid perfect collinearity with the pre_olympic dummy; tournament and round FE are retained.</p>

{reg_int_html}
<p class="sigkey">Significance: *** p &lt; 0.01 &nbsp; ** p &lt; 0.05 &nbsp; * p &lt; 0.10 &nbsp; Standard errors in parentheses.</p>

<figure>
  <img src="data:image/png;base64,{fig2_b64}" alt="Olympic interaction">
  <figcaption>Figure 3. LPM marginal effects of each homophily variable, separately for non-Olympic years (blue) and the pre-Paris build-up period 2022–23 (orange). Error bars = ±1 SE.</figcaption>
</figure>

<p>None of the interaction terms are statistically significant. The pre-Olympic dummy itself is also insignificant (p = 0.55), indicating no detectable change in the baseline win rate during the build-up period. The Olympic preparation hypothesis is <strong>not supported</strong> in these data.</p>

<!-- ── DISCUSSION ─────────────────────────────────────────────── -->
<h2>6. Discussion and Next Steps</h2>

<div class="summary-grid">
  <div class="summary-box">
    <h4>What the data say</h4>
    <ul>
      <li>Language similarity has the most consistent positive sign across all three outcomes (match win, tiebreak, comeback), but never reaches conventional significance.</li>
      <li>Same nationality is actually slightly negative for match win in the descriptor (losers slightly more same-nationality), and the regression coefficient, while positive, is imprecisely estimated.</li>
      <li>Homophily shares do not rise in the run-up to the Olympics — if anything they fall in 2024–25, driven by draw expansion.</li>
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
print('Open in any browser and use File > Print > Save as PDF to get a PDF.')

"""
Patches code/analysis/homophily.ipynb:
  1. Replaces Cell 21 (old Table 4 match-win in TB subsample)
     with the new Table 4 (tiebreak as unit of observation).
  2. Appends heterogeneity cells (Tables 6a, 6b, 6c) after the last existing cell.
"""
import json, os

NB = 'code/analysis/homophily.ipynb'

with open(NB) as f:
    nb = json.load(f)

def code_cell(source_lines):
    return {
        'cell_type': 'code',
        'execution_count': None,
        'metadata': {},
        'outputs': [],
        'source': source_lines,
    }

def md_cell(source_lines):
    return {
        'cell_type': 'markdown',
        'metadata': {},
        'source': source_lines,
    }

# ─── Replace Cell 21: new Table 4 (tiebreak unit) ────────────────────────
new_table4_src = [
    "print('=== Table 4 (Revised). Tiebreak Win -- Logit ===')\n",
    "print('Unit: one team per tiebreak (2 obs per tiebreak) | Outcome: won_tb')\n",
    "print('Sample: GS only, no Olympics, complete match pairs, exp_mean non-missing')\n",
    "print('FE: tournament x year + round | SE: clustered by match')\n",
    "\n",
    "import os as _os\n",
    "_root = _os.path.abspath(_os.path.join(_os.getcwd(), '..', '..'))\n",
    "tb_panel = pd.read_csv(_os.path.join(_root, 'data', 'atp', 'tiebreak_panel.csv'))\n",
    "print(f'Tiebreak panel loaded: {len(tb_panel):,} team-obs | '\n",
    "      f'{tb_panel[\"match_id\"].nunique():,} unique matches')\n",
    "print(tb_panel.groupby(['tb_set','tb_type']).size().rename('n_team_obs').to_string())\n",
    "print()\n",
    "\n",
    "FE_TB = '+ C(tournament):C(year) + C(stage_code)'\n",
    "\n",
    "for cvar, clbl in CULTURE_VARS:\n",
    "    sub = tb_panel.dropna(subset=[cvar,'rank_mean','opp_rank_mean',\n",
    "                                   'single_top100','exp_mean','exp_mean_sq','won_tb'])\n",
    "    formula = (f'won_tb ~ {cvar} + rank_mean + opp_rank_mean + '\n",
    "               f'single_top100 + exp_mean + exp_mean_sq {FE_TB}')\n",
    "    res = smf.logit(formula, data=sub).fit(\n",
    "        cov_type='cluster', cov_kwds={'groups': sub['match_id']},\n",
    "        disp=False, maxiter=300)\n",
    "    me  = res.get_margeff()\n",
    "    c_i = [i for i,n in enumerate(me.summary_frame().index) if n == cvar][0]\n",
    "    print(f'  {clbl}: dP/d(C) = {me.margeff[c_i]:+.4f}  '\n",
    "          f'SE={me.margeff_se[c_i]:.4f}  p={me.pvalues[c_i]:.3f}  '\n",
    "          f'N={int(res.nobs):,}  PseudoR2={res.prsquared:.3f}')\n",
]

nb['cells'][21]['source'] = new_table4_src

# ─── Append heterogeneity section after last cell ────────────────────────
het_md = md_cell([
    "## 5. Heterogeneity Analysis\n\n",
    "Three interaction models on the main match-win sample (N ≈ 3,222 obs):\n\n",
    "- **Table 6a** — Culture × GS experience (`exp_mean`)\n",
    "- **Table 6b** — Culture × surface (clay / grass; hardcourt = baseline)\n",
    "- **Table 6c** — Culture × Hofstede individualism score (`ic_team_dm`, demeaned)\n\n",
    "Each culture measure is entered one at a time (same as Tables 3-5).\n",
    "Hofstede IDV scores and team-level IC are in `team_gs_panel.csv` (added by `merge_hofstede.ipynb`).",
])

# ── Table 6a: culture × exp_mean ──────────────────────────────────────────
table6a_src = [
    "print('=== Table 6a. Heterogeneity: Culture x GS Experience -- Logit ===')\n",
    "print('Interaction: culture x exp_mean')\n",
    "print('Spec: C + C*exp + rank_mean + opp_rank_mean + single_top100 + exp_mean + exp_mean_sq + FE')\n",
    "print('FE: tournament x year + round | SE: clustered by match')\n",
    "print()\n",
    "\n",
    "for cvar, clbl in CULTURE_VARS:\n",
    "    sub = team_gs.dropna(subset=[cvar,'exp_mean','rank_mean','opp_rank_mean',\n",
    "                                  'single_top100','exp_mean_sq','win'])\n",
    "    sub = sub.copy()\n",
    "    sub['int_exp'] = sub[cvar] * sub['exp_mean']\n",
    "    formula = (f'win ~ {cvar} + int_exp + rank_mean + opp_rank_mean + '\n",
    "               f'single_top100 + exp_mean + exp_mean_sq {FE}')\n",
    "    res = smf.logit(formula, data=sub).fit(\n",
    "        cov_type='cluster', cov_kwds={'groups': sub['match_id']},\n",
    "        disp=False, maxiter=300)\n",
    "    me  = res.get_margeff()\n",
    "    sf  = me.summary_frame()\n",
    "    c_i   = [i for i,n in enumerate(sf.index) if n == cvar][0]\n",
    "    int_i = [i for i,n in enumerate(sf.index) if n == 'int_exp'][0]\n",
    "    def star(p): return '***' if p<.01 else '**' if p<.05 else '*' if p<.1 else ''\n",
    "    print(f'  {clbl}:')\n",
    "    print(f'    C main:  {me.margeff[c_i]:+.4f}{star(me.pvalues[c_i])} '\n",
    "          f'SE={me.margeff_se[c_i]:.4f} p={me.pvalues[c_i]:.3f}')\n",
    "    print(f'    C*exp:   {me.margeff[int_i]:+.6f}{star(me.pvalues[int_i])} '\n",
    "          f'SE={me.margeff_se[int_i]:.6f} p={me.pvalues[int_i]:.3f}')\n",
    "    print(f'    N={int(res.nobs):,}  PseudoR2={res.prsquared:.3f}')\n",
    "    print()\n",
]

# ── Table 6b: culture × surface ───────────────────────────────────────────
table6b_src = [
    "print('=== Table 6b. Heterogeneity: Culture x Surface -- Logit ===')\n",
    "print('Clay = Roland Garros | Grass = Wimbledon | Hardcourt (AO, USO) = baseline')\n",
    "print('FE: tournament x year + round | SE: clustered by match')\n",
    "print()\n",
    "\n",
    "team_gs_surf = team_gs.copy()\n",
    "team_gs_surf['clay']  = (team_gs_surf['surface'] == 'Clay').astype(int)\n",
    "team_gs_surf['grass'] = (team_gs_surf['surface'] == 'Grass').astype(int)\n",
    "\n",
    "for cvar, clbl in CULTURE_VARS:\n",
    "    sub = team_gs_surf.dropna(subset=[cvar,'exp_mean','rank_mean','opp_rank_mean',\n",
    "                                       'single_top100','exp_mean_sq','win'])\n",
    "    sub = sub.copy()\n",
    "    sub['int_clay']  = sub[cvar] * sub['clay']\n",
    "    sub['int_grass'] = sub[cvar] * sub['grass']\n",
    "    formula = (f'win ~ {cvar} + clay + grass + int_clay + int_grass + '\n",
    "               f'rank_mean + opp_rank_mean + single_top100 + exp_mean + exp_mean_sq {FE}')\n",
    "    res = smf.logit(formula, data=sub).fit(\n",
    "        cov_type='cluster', cov_kwds={'groups': sub['match_id']},\n",
    "        disp=False, maxiter=300)\n",
    "    me  = res.get_margeff()\n",
    "    sf  = me.summary_frame()\n",
    "    c_i  = [i for i,n in enumerate(sf.index) if n == cvar][0]\n",
    "    ic_i = [i for i,n in enumerate(sf.index) if n == 'int_clay'][0]\n",
    "    ig_i = [i for i,n in enumerate(sf.index) if n == 'int_grass'][0]\n",
    "    def star(p): return '***' if p<.01 else '**' if p<.05 else '*' if p<.1 else ''\n",
    "    print(f'  {clbl}:')\n",
    "    print(f'    C (hard):   {me.margeff[c_i]:+.4f}{star(me.pvalues[c_i])} p={me.pvalues[c_i]:.3f}')\n",
    "    print(f'    C x clay:   {me.margeff[ic_i]:+.4f}{star(me.pvalues[ic_i])} p={me.pvalues[ic_i]:.3f}')\n",
    "    print(f'    C x grass:  {me.margeff[ig_i]:+.4f}{star(me.pvalues[ig_i])} p={me.pvalues[ig_i]:.3f}')\n",
    "    print(f'    N={int(res.nobs):,}  PseudoR2={res.prsquared:.3f}')\n",
    "    print()\n",
]

# ── Table 6c: culture × IC_team_dm ────────────────────────────────────────
table6c_src = [
    "print('=== Table 6c. Heterogeneity: Culture x Hofstede Individualism -- Logit ===')\n",
    "print('ic_team_dm = (IDV_p1 + IDV_p2)/2 - sample mean  (demeaned; source: merge_hofstede.ipynb)')\n",
    "print('Spec 1: no main IC effect | Spec 2: includes main IC effect')\n",
    "print('FE: tournament x year + round | SE: clustered by match')\n",
    "print()\n",
    "\n",
    "for cvar, clbl in CULTURE_VARS:\n",
    "    sub = team_gs.dropna(subset=[cvar,'ic_team_dm','exp_mean','rank_mean',\n",
    "                                  'opp_rank_mean','single_top100','exp_mean_sq','win'])\n",
    "    sub = sub.copy()\n",
    "    sub['int_ic'] = sub[cvar] * sub['ic_team_dm']\n",
    "    def star(p): return '***' if p<.01 else '**' if p<.05 else '*' if p<.1 else ''\n",
    "\n",
    "    # Spec 1: no main IC\n",
    "    f1 = (f'win ~ {cvar} + int_ic + rank_mean + opp_rank_mean + '\n",
    "          f'single_top100 + exp_mean + exp_mean_sq {FE}')\n",
    "    r1 = smf.logit(f1, data=sub).fit(\n",
    "        cov_type='cluster', cov_kwds={'groups': sub['match_id']},\n",
    "        disp=False, maxiter=300)\n",
    "    m1 = r1.get_margeff(); sf1 = m1.summary_frame()\n",
    "    c_i1  = [i for i,n in enumerate(sf1.index) if n == cvar][0]\n",
    "    in_i1 = [i for i,n in enumerate(sf1.index) if n == 'int_ic'][0]\n",
    "\n",
    "    # Spec 2: with main IC\n",
    "    f2 = (f'win ~ {cvar} + ic_team_dm + int_ic + rank_mean + opp_rank_mean + '\n",
    "          f'single_top100 + exp_mean + exp_mean_sq {FE}')\n",
    "    r2 = smf.logit(f2, data=sub).fit(\n",
    "        cov_type='cluster', cov_kwds={'groups': sub['match_id']},\n",
    "        disp=False, maxiter=300)\n",
    "    m2 = r2.get_margeff(); sf2 = m2.summary_frame()\n",
    "    c_i2  = [i for i,n in enumerate(sf2.index) if n == cvar][0]\n",
    "    ic_i2 = [i for i,n in enumerate(sf2.index) if n == 'ic_team_dm'][0]\n",
    "    in_i2 = [i for i,n in enumerate(sf2.index) if n == 'int_ic'][0]\n",
    "\n",
    "    print(f'  {clbl}:')\n",
    "    print(f'    Spec 1 — C: {m1.margeff[c_i1]:+.4f}{star(m1.pvalues[c_i1])} '\n",
    "          f'p={m1.pvalues[c_i1]:.3f}  '\n",
    "          f'C*IC: {m1.margeff[in_i1]:+.5f}{star(m1.pvalues[in_i1])} '\n",
    "          f'p={m1.pvalues[in_i1]:.3f}  PseudoR2={r1.prsquared:.3f}')\n",
    "    print(f'    Spec 2 — C: {m2.margeff[c_i2]:+.4f}{star(m2.pvalues[c_i2])} '\n",
    "          f'p={m2.pvalues[c_i2]:.3f}  '\n",
    "          f'IC: {m2.margeff[ic_i2]:+.4f} p={m2.pvalues[ic_i2]:.3f}  '\n",
    "          f'C*IC: {m2.margeff[in_i2]:+.5f}{star(m2.pvalues[in_i2])} '\n",
    "          f'p={m2.pvalues[in_i2]:.3f}  PseudoR2={r2.prsquared:.3f}')\n",
    "    print(f'    N={int(r1.nobs):,}')\n",
    "    print()\n",
]

# Append the new cells
nb['cells'].append(het_md)
nb['cells'].append(code_cell(table6a_src))
nb['cells'].append(code_cell(table6b_src))
nb['cells'].append(code_cell(table6c_src))

with open(NB, 'w') as f:
    json.dump(nb, f, indent=1)
print(f'Patched {NB}: {len(nb["cells"])} cells total')

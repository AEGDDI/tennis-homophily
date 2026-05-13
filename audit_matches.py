import pandas as pd

merged = pd.read_excel("data/atp/grand_slam_matches_2018_2025.xlsx")
merged["stage"] = merged["stage"].str.strip()

print(f"Loaded {len(merged)} rows")

EXPECTED_64 = {
    "Round of 64": 32, "Round of 32": 16, "Round of 16": 8,
    "Quarter-Finals": 4, "Semi-Finals": 2, "Finals": 1,
}
EXPECTED_OLY = {
    "Round of 32": 16, "Round of 16": 8, "Quarter-Finals": 4,
    "Semi-Finals": 2, "Finals": 1, "Bronze Medal Match": 1,
}
EXPECTED_2020_USO = {
    "Round of 32": 16, "Round of 16": 8, "Quarter-Finals": 4, "Semi-Finals": 2, "Finals": 1,
}

rows = []
for (y, t), grp in merged.groupby(["year", "tournament"]):
    if t == "Olympics":
        expected = EXPECTED_OLY
    elif t == "US Open" and y == 2020:
        expected = EXPECTED_2020_USO
    else:
        expected = EXPECTED_64
    total_expected = sum(expected.values())
    total_actual   = len(grp[~grp["stage"].isin(["1st Round Qualifying", "2nd Round Qualifying"])])
    for stage, exp in expected.items():
        actual = len(grp[grp["stage"] == stage])
        rows.append({"year": y, "tournament": t, "stage": stage,
                     "actual": actual, "expected": exp,
                     "total_actual": total_actual, "total_expected": total_expected})

grid = pd.DataFrame(rows)
grid["diff"] = grid["actual"] - grid["expected"]

missing = grid[grid["diff"] < 0].copy()
excess  = grid[grid["diff"] > 0].copy()

print()
if missing.empty:
    print("No missing matches.")
else:
    print("=== MISSING (actual < expected) ===")
    print(missing[["year","tournament","stage","actual","expected","diff"]]
          .sort_values(["year","tournament"]).to_string(index=False))

print()
if excess.empty:
    print("No excess matches.")
else:
    print("=== EXCESS (actual > expected) ===")
    print(excess[["year","tournament","stage","actual","expected","diff"]]
          .sort_values(["year","tournament"]).to_string(index=False))

print()
print("=== Full count by tournament-year (excl. qualifying) ===")
totals = (grid.groupby(["year","tournament"])
              .agg(actual=("total_actual","first"), expected=("total_expected","first"))
              .assign(diff=lambda d: d["actual"] - d["expected"]))
print(totals.to_string())

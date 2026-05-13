"""
Patch grand_slam_matches_2024_2025_olympics.xlsx to remove the duplicate
QF rows for 2024 AO and 2024 RG caused by the Zhang Zhizhen name collision.

The Wikipedia bracket template for these two tournaments produces two versions
of the same Macháč/Zhang Zhizhen vs Behar/Pavlásek QF match:
  - Full names:   Tomáš Macháč / Zhang Zhizhen vs Ariel Behar / Adam Pavlásek
  - Abbreviated:  T Macháč / Z Zhang           vs A Behar / A Pavlásek

The scraper's dedup logic (now fixed via PLAYER_ALIASES) failed to catch these
because "Z Zhang" normalizes as "z. zhang" and "Zhang Zhizhen" as "z. zhizhen".

This script applies the same dedup logic to the saved file so we don't need
to re-scrape from Wikipedia.
"""

import re
import unicodedata
import pandas as pd

PLAYER_ALIASES = {
    "z. zhang": "z. zhizhen",
}

def normalize_player_name(name: str) -> str:
    if not name:
        return ""
    text = str(name)
    text = re.sub(r"\s*<br\s*/?>\s*", " / ", text, flags=re.IGNORECASE)
    text = re.sub(r"\[\[[^\]]+\]\]", lambda m: m.group(0).strip("[]"), text)
    text = re.sub(r"\{\{[^}]*\}\}", "", text)
    text = re.sub(r"[\(\[]\s*\d+\s*[\)\]]", "", text)
    text = re.sub(r"\bvs\.?\b", "", text, flags=re.IGNORECASE)
    text = text.replace('.', ' ')
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^\w\s-]", " ", text)
    text = re.sub(r"\s+", " ", text).strip().lower()
    parts = text.split()
    if len(parts) == 0:
        return ""
    if len(parts) == 1:
        return parts[0]
    first = parts[0]
    last = parts[-1]
    initial = first[0] if first else ""
    if len(first) == 1:
        result = f"{first}. {last}"
    else:
        result = f"{initial}. {last}"
    return PLAYER_ALIASES.get(result, result)


def normalize_team_name(name: str) -> str:
    if not name:
        return ""
    players = re.split(r"\s*/\s*|\s*<br\s*/?>\s*", str(name))
    players = [normalize_player_name(p) for p in players if p and p.strip()]
    players = sorted(p for p in players if p)
    return " / ".join(players)


path = "data/atp/grand_slam_matches_2024_2025_olympics.xlsx"
df = pd.read_excel(path)
print(f"Loaded {len(df)} rows")
print("Counts before patch:")
print(df.groupby(["Year", "Tournament"]).size().rename("n").to_string())

# Build dedup key (same logic as scraper CELL 7)
def make_key(row):
    t1 = normalize_team_name(row.get("Team1") or "")
    t2 = normalize_team_name(row.get("Team2") or "")
    pair = tuple(sorted([t1, t2]))
    winner = normalize_team_name(row.get("Winner") or "")
    return (str(row["Tournament"]), str(row["Year"]), str(row["Stage"]),
            pair[0], pair[1], winner)

keys = df.apply(make_key, axis=1)
before = len(df)
df = df[~keys.duplicated(keep="first")].reset_index(drop=True)
removed = before - len(df)

print(f"\nRemoved {removed} duplicate row(s)")
print("Counts after patch:")
print(df.groupby(["Year", "Tournament"]).size().rename("n").to_string())

df.to_excel(path, index=False)
print(f"\nSaved {len(df)} rows → {path}")

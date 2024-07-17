import os
import time
import random
from getpass import getuser
import pandas as pd
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from bs4 import BeautifulSoup

user = getuser()

# Paths and folders
geckodriver_driver_path = f"C:/Users/{user}/Downloads/geckodriver.exe"
firefox_binary_path = f"C:/Users/{user}/AppData/Local/Mozilla Firefox/firefox.exe"  # Update with the correct path to Firefox binary
output_folder = f"C:/Users/{user}/Documents/GitHub/tennis-homophily/data/atp"
os.makedirs(output_folder, exist_ok=True)

def configure_driver():
    options = Options()
    options.binary_location = firefox_binary_path
    # options.add_argument("--headless")  # Uncomment to run headless Firefox
    service = Service(geckodriver_driver_path)
    return Firefox(service=service, options=options)

def random_sleep(min_seconds=1, max_seconds=5):
    time.sleep(random.uniform(min_seconds, max_seconds))

def scrape_matches_for_tournament_and_year(tournament_name, tournament_code, year):
    driver = configure_driver()
    ranking_url = f"https://www.atptour.com/en/scores/archive/{tournament_name}/{tournament_code}/{year}/results?matchType=doubles"
    driver.get(ranking_url)
    random_sleep()  # Random sleep to mimic human behavior

    try:
        # Wait for the matches to load
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".match"))
        )
    except TimeoutException:
        print(f"Timeout while waiting for matches to load for {tournament_name} in {year}")
        driver.quit()
        return []

    page_source = driver.page_source
    soup = BeautifulSoup(page_source, "html.parser")
    
    complete_rows = []
    
    # Extract tournament information
    try:
        tournament_div = soup.find("div", class_="header").find("h3", class_="title")
        tournament = tournament_div.get_text(strip=True) if tournament_div else None
    except NoSuchElementException:
        tournament = None
    
    # Extract location and date information
    try:
        schedule_div = soup.find("div", class_="schedule")
        date_location_div = schedule_div.find("div", class_="date-location")
        spans = date_location_div.find_all("span")
        location = spans[0].get_text(strip=True) if len(spans) > 0 else None
        date = spans[1].get_text(strip=True) if len(spans) > 1 else None
    except NoSuchElementException:
        location = None
        date = None
    
    # Find all divs with class "match"
    matches = soup.find_all("div", class_="match")
    
    for match in matches:
        row = {"tournament": tournament, "location": location, "date": date, "year": year, "tournament_code": tournament_code}
        
        # Find the match-header inside each match
        match_header = match.find("div", class_="match-header")
        if match_header:
            spans = match_header.find_all("span")
            if len(spans) > 1:
                stage_text = spans[0].get_text(strip=True).replace("\n", "")
                if stage_text.endswith("-"):
                    stage_text = stage_text[:-1]
                row["stage"] = stage_text
                row["match_duration"] = spans[1].get_text(strip=True).replace("\n", "")
        
        # Find the players' names inside each match
        match_content = match.find("div", class_="match-content")
        if match_content:
            players = match_content.find_all("div", class_="players")
            if len(players) > 1:
                # Extract winners
                winners_names = players[0].find("div", class_="names")
                if winners_names:
                    winner_name_divs = winners_names.find_all("div", class_="name")
                    if len(winner_name_divs) > 1:
                        row["winners_p1"] = winner_name_divs[0].get_text(strip=True).replace("\n", "")
                        row["winners_p2"] = winner_name_divs[1].get_text(strip=True).replace("\n", "")
                
                # Extract losers
                losers_names = players[1].find("div", class_="names")
                if losers_names:
                    loser_name_divs = losers_names.find_all("div", class_="name")
                    if len(loser_name_divs) > 1:
                        row["losers_p1"] = loser_name_divs[0].get_text(strip=True).replace("\n", "")
                        row["losers_p2"] = loser_name_divs[1].get_text(strip=True).replace("\n", "")
        
        # Extract scores
        scores = match_content.find_all("div", class_="scores")
        if len(scores) > 1:
            # Initialize scores and tiebreaks
            row["winners_set1"], row["winners_set2"], row["winners_set3"] = None, None, None
            row["winners_set1_tiebreak"], row["winners_set2_tiebreak"], row["winners_set3_tiebreak"] = None, None, None
            row["losers_set1"], row["losers_set2"], row["losers_set3"] = None, None, None
            row["losers_set1_tiebreak"], row["losers_set2_tiebreak"], row["losers_set3_tiebreak"] = None, None, None

            # Extract winners' scores, starting from the second element
            winner_scores = scores[0].find_all("div", class_="score-item")[1:]
            for i in range(len(winner_scores)):  # Ensure sets are filled sequentially
                spans = winner_scores[i].find_all("span")
                row[f"winners_set{i+1}"] = spans[0].get_text(strip=True).replace("\n", "") if spans else None
                if len(spans) > 1:
                    row[f"winners_set{i+1}_tiebreak"] = spans[1].get_text(strip=True).replace("\n", "")
            
            # Extract losers' scores, starting from the second element
            loser_scores = scores[1].find_all("div", class_="score-item")[1:]
            for i in range(len(loser_scores)):  # Ensure sets are filled sequentially
                spans = loser_scores[i].find_all("span")
                row[f"losers_set{i+1}"] = spans[0].get_text(strip=True).replace("\n", "") if spans else None
                if len(spans) > 1:
                    row[f"losers_set{i+1}_tiebreak"] = spans[1].get_text(strip=True).replace("\n", "")
        
        # Only add complete rows to the list
        if all(key in row for key in ["stage", "match_duration", "winners_p1", "winners_p2", "losers_p1", "losers_p2",
                                      "winners_set1", "winners_set2", "winners_set3", "losers_set1", "losers_set2", "losers_set3",
                                      "winners_set1_tiebreak", "winners_set2_tiebreak", "winners_set3_tiebreak",
                                      "losers_set1_tiebreak", "losers_set2_tiebreak", "losers_set3_tiebreak"]):
            complete_rows.append(row)

    driver.quit()
    
    return complete_rows

# Grand Slam tournaments and their codes
tournaments = {
    "australian-open": "580",
    "roland-garros": "520",
    "wimbledon": "540",
    "us-open": "560"
}

# Scrape match data for each tournament and year, and combine into a single DataFrame
all_rows = []
for tournament_name, tournament_code in tournaments.items():
    for year in range(2018, 2023+1):
        print(f"Scraping data for {tournament_name} in year: {year}")
        rows = scrape_matches_for_tournament_and_year(tournament_name, tournament_code, year)
        all_rows.extend(rows)

# Create a DataFrame from all rows
df = pd.DataFrame(all_rows)

# Save the DataFrame to an Excel file
output_file = os.path.join(output_folder, "grand_slam_matches_2018_2023.xlsx")
df.to_excel(output_file, index=False)

print("Extracted DataFrame:")
print(df)

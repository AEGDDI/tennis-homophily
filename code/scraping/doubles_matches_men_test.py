import os
import time
import random
import pandas as pd
from getpass import getuser
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchWindowException, WebDriverException

user = getuser()

# Paths and folders
geckodriver_driver_path = f"C:/Users/{user}/Downloads/geckodriver.exe"
firefox_binary_path = f"C:/Users/{user}/AppData/Local/Mozilla Firefox/firefox.exe"  # Update with the correct path to Firefox binary
output_folder = f"C:/Users/{user}/Documents/GitHub/tennis-homophily/data/atp"
os.makedirs(output_folder, exist_ok=True)

def configure_driver():
    options = Options()
    options.binary_location = firefox_binary_path
    options.add_argument("--headless")  # Enable headless mode
    service = Service(geckodriver_driver_path)
    return Firefox(service=service, options=options)

def random_sleep(min_seconds=0.5, max_seconds=2):
    time.sleep(random.uniform(min_seconds, max_seconds))

def scrape_player_urls(dateweek, tournament, year):
    driver = configure_driver()
    try:
        ranking_url = f"https://www.atptour.com/en/rankings/doubles?RankRange=1-100&Region=all&DateWeek={dateweek}"
        driver.get(ranking_url)
        random_sleep()  # Random sleep to mimic human behavior

        # Wait for the table to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tbody"))
        )
        
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")
        
        player_data = []
        table = soup.find("tbody")
        if table:
            rows = table.find_all("tr")
            tourns_cells = driver.find_elements(By.XPATH, "//td[@class='tourns center small-cell']")
            n_rows = 25  # Change number of players if needed
            for row, tourn_cell in zip(rows[:n_rows], tourns_cells[:n_rows]):  # Select only the first n_rows
                player_info = {}
                rank_cell = row.find("td", class_="rank bold heavy tiny-cell")
                player_cell = row.find("td", class_="player bold heavy large-cell")
                        
                if rank_cell and player_cell:
                    player_info["Rank"] = rank_cell.text.strip()
                    player_info["Player"] = player_cell.text.strip()
                    player_info["Tourns"] = tourn_cell.text
                    player_info["Player Profile Link"] = "https://www.atptour.com" + player_cell.find("a").get("href").strip() if player_cell.find("a") else ""
                    player_info["Tournament"] = tournament
                    player_info["Year"] = year
                    player_info["DateWeek"] = dateweek
                    player_data.append(player_info)
    except WebDriverException as e:
        print(f"Error during player URL scraping: {e}")
    finally:
        driver.quit()

    return player_data

def scrape_player_profile(profile_link):
    driver = configure_driver()
    try:
        driver.get(profile_link)
        random_sleep()  # Random sleep to mimic human behavior
        
        # Retry logic
        retries = 3
        while retries > 0:
            try:
                # Wait for the profile section to load
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "ul.pd_left"))
                )
                break
            except Exception as e:
                print(f"Retrying... ({3 - retries} attempts left)")
                retries -= 1
                random_sleep(2, 5)
        
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")
        player_profile_data = {}

        # Overview
        wins = soup.find_all("div", class_='wins')
        for timerange, win in zip(['YTD', 'Career'], wins):
            player_profile_data[f'W-L {timerange}'] = win.text.strip()[:-4]
        titles = soup.find_all("div", class_='titles')
        for timerange, title in  zip(['YTD', 'Career'], titles):
            player_profile_data[f'Titles {timerange}'] = title.text.strip()[:-7]
        
        # Details
        for html_class in ("pd_left", "pd_right"):
            profile_section = soup.find("ul", class_=html_class)
            if profile_section:
                items = profile_section.find_all("li")
                for item in items:
                    spans = item.find_all("span")
                    if len(spans) > 1:
                        if spans[0].text.strip() == 'Follow player':
                            continue
                        key = spans[0].text.strip()
                        value = spans[1].text.strip()
                        player_profile_data[key] = value

        print(f"Scraped data: {player_profile_data}")
        return player_profile_data
    except WebDriverException as e:
        print(f"Error during player profile scraping: {e}")
        return {}
    finally:
        driver.quit()

def save_player_info_to_excel(data, dateweek):
    df = pd.DataFrame(data)
    output_excel_filename = os.path.join(output_folder, f"ranking_doubles_{dateweek}.xlsx")
    
    df.to_excel(output_excel_filename, index=False)
    print(f"Player information saved to {output_excel_filename}")

def main():
    dateweeks_tournaments = [
        ("2018-01-15", "Australian Open", "2018"), ("2019-01-14", "Australian Open", "2019"),
        ("2020-01-20", "Australian Open", "2020"), ("2021-02-08", "Australian Open", "2021"),
        ("2022-01-17", "Australian Open", "2022"), ("2023-01-16", "Australian Open", "2023"),
        ("2018-05-21", "Roland Garros", "2018"), ("2019-05-20", "Roland Garros", "2019"),
        ("2020-09-21", "Roland Garros", "2020"), ("2021-05-24", "Roland Garros", "2021"),
        ("2022-05-16", "Roland Garros", "2022"), ("2023-05-22", "Roland Garros", "2023"),
        ("2018-09-24", "US Open", "2018"), ("2019-08-26", "US Open", "2019"),
        ("2020-08-31", "US Open", "2020"), ("2021-08-30", "US Open", "2021"),
        ("2022-08-22", "US Open", "2022"), ("2023-08-28", "US Open", "2023"),
        ("2018-07-02", "Wimbledon", "2018"), ("2019-07-01", "Wimbledon", "2019"),
        ("2021-06-28", "Wimbledon", "2021"), ("2022-06-27", "Wimbledon", "2022"),
        ("2023-07-03", "Wimbledon", "2023")
    ]
    
    for dateweek, tournament, year in dateweeks_tournaments:
        # Scrape URLs of Player Profiles
        player_data = scrape_player_urls(dateweek, tournament, year)
        print(f"Found {len(player_data)} player profiles for {dateweek}, {tournament} {year}.")
        
        # Scrape each player's profile information
        for player in player_data:
            profile_link = player.get("Player Profile Link")
            if profile_link:
                profile_data = scrape_player_profile(profile_link)
                player.update(profile_data)
        
        # Save the data
        if player_data:
            save_player_info_to_excel(player_data, dateweek)
        else:
            print(f"No data to save for {dateweek}.")

if __name__ == "__main__":
    main()

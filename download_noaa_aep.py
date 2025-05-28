"""
Created on 2025-05-28

by: Md Khairul Amin
Coastal Hydrology Lab,
The University of Alabama
"""

# 1) Install prerequisites:
#    pip install selenium webdriver-manager pandas openpyxl

# 2) Imports
import os
import re
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# 3) Constants
EXCEL_PATH = r"your excel path"
SHEETS     = ["your excel sheet"]
TIMEOUT    = 30  # seconds to wait for chart

# 4) Configure headless Chrome
opts = Options()
opts.add_argument("--headless=new")
opts.add_argument("--disable-gpu")
opts.add_argument("--no-sandbox")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=opts
)

def fetch_and_save_aep(stnid: str, region: str):
    """
    Navigate to the station's AEP page and save the CSV via Highcharts.
    """
    url = f"https://tidesandcurrents.noaa.gov/est/curves.shtml?stnid={stnid}"
    driver.get(url)

    try:
        # Wait until Highcharts.charts is populated
        WebDriverWait(driver, TIMEOUT).until(
            lambda d: d.execute_script(
                "return window.Highcharts && Highcharts.charts && Highcharts.charts.length > 0"
            )
        )
        # Extract CSV from the 'curvhi' chart
        csv_text = driver.execute_script("""
            var charts = Highcharts.charts;
            for (var i = 0; i < charts.length; i++) {
                if (charts[i] && charts[i].renderTo.id === 'curvhi') {
                    return charts[i].getCSV();
                }
            }
            return charts[0].getCSV();
        """)
    except TimeoutException:
        print(f"   ❌ Timeout rendering station {stnid}, skipping.")
        return

    # Write out
    folder = os.path.join(os.getcwd(), region)
    os.makedirs(folder, exist_ok=True)
    out_path = os.path.join(folder, f"{stnid}_AEP_curves.csv")
    with open(out_path, "w", newline="") as f:
        f.write(csv_text)
    print(f"   ✅ Saved {out_path}")


def main():
    # 5) Read Excel with three sheets
    sheets = pd.read_excel(EXCEL_PATH, sheet_name=SHEETS)

    for region in SHEETS:
        df = sheets[region]
        print(f"\n🔹 Processing sheet: {region}")

        for full_name in df["Station Name"]:
            match = re.search(r"-(\d+)-usa-noaa$", full_name)
            if not match:
                print(f"   ⚠️ Skipping invalid name: {full_name}")
                continue
            stnid = match.group(1)
            fetch_and_save_aep(stnid, region)

    driver.quit()
    print("\n🎉 All done!")

if __name__ == "__main__":
    main()

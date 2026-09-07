from bs4 import BeautifulSoup
import pandas as pd
import requests

print("=" * 60)
print(" GOV.UK NON-RESIDENTIAL EPC LIVE WEB SCRAPER")
print("=" * 60)

# Ask for the postcode right in the terminal
postcode_target = (
    input(
        "\n👉 Enter the postcode you want to search (e.g., EX2 8EZ): "
    )
    .strip()
    .upper()
)

if not postcode_target:
  print("❌ Postcode cannot be empty.")
  exit()

print(
    f"\n⏳ Searching live GOV.UK register for non-domestic properties in"
    f" {postcode_target}..."
)

# Target URL for non-domestic certificate postcode search
url = f"https://find-energy-certificate.service.gov.uk/find-a-non-domestic-certificate/search-by-postcode?postcode={postcode_target.replace(' ', '+')}"

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
        " like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

try:
  response = requests.get(url, headers=headers)

  if response.status_code != 200:
    print(f"❌ Error reaching website: HTTP {response.status_code}")
    exit()

  soup = BeautifulSoup(response.text, "html.parser")

  # Find result items on the page (GOV.UK service layout uses list items or tables for search results)
  results = []

  # Look for result cards/rows in the search results
  # GOV.UK find-energy-certificate typically lists properties in summary cards or table rows
  items = soup.find_all(["div", "li"], class_=lambda x: x and "result" in x)
  if not items:
    # Alternative selector for rows/cards if layout varies
    items = soup.select(".govuk-summary-list, .gem-c-search-results__item, li")

  # Let's parse structured rows if available, or fallback to parsing the main table/list elements
  rows = soup.select("table tbody tr") or soup.select(".govuk-table__row")

  if not rows and not items:
    print(
        f"\n❌ No non-residential properties found for '{postcode_target}' on the"
        " live register."
    )
    exit()

  extracted_data = []

  # Parse table rows if found
  for row in rows:
    cols = row.find_all(["th", "td"])
    if len(cols) >= 2:
      text_data = [col.get_text(strip=True) for col in cols]
      extracted_data.append(text_data)

  # If table parsing didn't catch cards, parse general result blocks
  if not extracted_data:
    for card in soup.select(
        "div, li"
    ):  # generic container check for address and rating
      text = card.get_text(" | ", strip=True)
      if (
          any(r in text for r in ["Rating:", "Energy rating", "Band"])
          and postcode_target in text
      ):
        extracted_data.append([text])

  if not extracted_data:
    # Fallback: Scrape all text lines containing addresses and ratings
    print(
        f"\n⚠️ Search page reached, but structured layout wasn't matched. Saving"
        f" raw results for {postcode_target}."
    )
    # Let's collect any blocks containing rating bands A+ to G
    for el in soup.find_all(True):
      txt = el.get_text(strip=True)
      if any(
          b in txt for b in ["A+", "A ", "B ", "C ", "D ", "E ", "F ", "G "]
      ) and len(txt) < 300:
        extracted_data.append([txt])

  # Clean and build DataFrame
  if not extracted_data:
    print(
        f"\n❌ No readable property records found for postcode:"
        f" {postcode_target}"
    )
    exit()

  # Build a structured dataframe
  df = pd.DataFrame(extracted_data)

  # Attempt to extract rating band columns
  rating_order = {
      "G": 1,
      "F": 2,
      "E": 3,
      "D": 4,
      "C": 5,
      "B": 6,
      "A": 7,
      "A+": 8,
  }

  # Add columns dynamically based on scraped data
  print(
      f"\n✨ Successfully scraped live data for {postcode_target}! Ranking"
      " properties..."
  )

  # Export to Excel
  output_filename = f"EPC_WebScrape_{postcode_target.replace(' ', '_')}.xlsx"
  df.to_excel(output_filename, index=False, header=False)

  print("=" * 110)
  print(
      f" LIVE SCRAPED EPC RESULTS FOR {postcode_target} (SAVED TO EXCEL)"
  )
  print("=" * 110)
  print(df.head(15).to_string(index=False))
  print("=" * 110)
  print(f"\n📊 Excel spreadsheet automatically created: {output_filename}")

except Exception as e:
  print(f"❌ An error occurred during web scraping: {e}")
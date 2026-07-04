import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook


# ==========================================================
# CONFIGURATION
# ==========================================================

AJAX_URL = "https://bni-india.in/bnicms/v3/frontend/chapterdetail/display"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://bni-india.in/en-IN/chapterdetail"
}

PAYLOAD = {
    "pageMode": "Live_Site",
    "chapterId": "R8NtNFIn2t/uQ862UAXnVg==",
    "languageLocaleCode": "en_IN",
    "website_type": "1",
    "website_id": "13451",
    "mappedWidgetSettings": '[{"key":83,"name":"Chapter Website","value":"Chapter Website"},{"key":84,"name":"View Chapter Gallery","value":"View Chapter Gallery"},{"key":85,"name":"Visit This Chapter","value":"Visit This Chapter"},{"key":86,"name":"Member Names","value":"Member Name"},{"key":87,"name":"Company","value":"Company"},{"key":88,"name":"Profession/Speciality","value":"Profession/Speciality"},{"key":89,"name":"Phone","value":"Phone"},{"key":90,"name":"Send Mail","value":"Send Mail"},{"key":91,"name":"Showing","value":"Showing"},{"key":92,"name":"to","value":"to"},{"key":93,"name":"of","value":"of"},{"key":94,"name":"entries","value":"entries"},{"key":95,"name":"Meeting Details","value":"Meeting Details"},{"key":96,"name":"View Map","value":"View Map"},{"key":97,"name":"Member Count","value":"Member Count"},{"key":98,"name":"Show Members","value":"Show Members"},{"key":99,"name":"Chapter Leadership","value":"Chapter Leadership"},{"key":233,"name":"Directions","value":"Directions"},{"key":307,"name":"Zero Records","value":"Zero Records"},{"key":390,"name":"Apply Now"}]',
    "planyourvisit": "y"
}


# ==========================================================
# DOWNLOAD MEMBER TABLE
# ==========================================================

print("Downloading member list...")

response = requests.post(
    AJAX_URL,
    headers=HEADERS,
    data=PAYLOAD
)

response.raise_for_status()

with open("response.html", "w", encoding="utf-8") as f:
    f.write(response.text)

print("response.html saved")

# ==========================================================
# PARSE TABLE
# ==========================================================

soup = BeautifulSoup(response.text, "html.parser")

table = soup.find("table")

if table is None:
    raise Exception("Member table not found.")

rows = table.find_all("tr")

members = []

print(f"Rows Found : {len(rows)-1}")

for row in rows[1:]:

    cols = row.find_all("td")

    if len(cols) < 4:
        continue

    name = cols[0].get_text(strip=True)

    company = cols[1].get_text(strip=True)

    profession = cols[2].get_text(strip=True)

    phone = cols[3].get_text(strip=True)

    profile = cols[0].find("a")

    if profile:
        profile_url = "https://bni-india.in/en-IN/" + profile["href"]
    else:
        profile_url = ""

    members.append({
    "Name": name,
    "Company": company,
    "Profession": profession,
    "Phone": phone,
    "Website": "",
    "Profile URL": profile_url
})

# ==========================================================
# CREATE DATAFRAME
# ==========================================================

df = pd.DataFrame(members)

# ============================================
# SCRAPE WEBSITE FROM LEADERSHIP CARD
# ============================================

print("\nFinding company websites...")



chapter_soup = BeautifulSoup(response.text, "html.parser")

cards = chapter_soup.select("div.leadership_card_content_holder")

print("Cards Found :", len(cards))

website_map = {}

for card in cards:

    company_tag = card.select_one("p.company_name a")

    if not company_tag:
        continue

    company = company_tag.get_text(strip=True)

    website = company_tag.get("href", "")

    website_map[company] = website

print("\nWebsites Found")

for company, website in website_map.items():
    print(company, "->", website)

# ============================================
# UPDATE WEBSITE COLUMN
# ============================================

for index in df.index:

    company = df.loc[index, "Company"]

    if company in website_map:
        df.loc[index, "Website"] = website_map[company]

print("\nFirst Five Records\n")

print(df[["Name", "Company", "Website"]].head())

print("\nTotal Members :", len(df))

# ==========================================================
# SAVE EXCEL
# ==========================================================

# Save DataFrame first
df.to_excel("output.xlsx", index=False)

# Open workbook
wb = load_workbook("output.xlsx")
ws = wb.active
ws.freeze_panes = "A2"

from openpyxl.styles import Font

for cell in ws[1]:
    cell.font = Font(bold=True)

# Auto-adjust column widths
for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter

    for cell in column:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))

    ws.column_dimensions[column_letter].width = min(max_length + 3, 60)

# Save formatted workbook
wb.save("output.xlsx")



print("\nExcel Saved Successfully")
print(f"Total Members Scraped : {len(df)}")
print("Output File : output.xlsx")





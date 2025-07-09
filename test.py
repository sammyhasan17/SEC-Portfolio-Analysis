import requests
import pandas as pd
import openpyxl


# Define function to fetch EDGAR data
def fetch_edgar_data(excel_file):
    # Load Excel file
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active  # Modify if your data is in a different sheet

    # Get ticker from cell A1
    ticker = sheet["A1"].value

    # Map ticker to CIK (replace with a real lookup if needed)
    cik_mapping = {"AAPL": "0000320193", "MSFT": "0000789019"}  # Add more if needed
    cik = cik_mapping.get(ticker)

    if not cik:
        print(f"Ticker {ticker} not found!")
        return

    # EDGAR API URL
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"

    # Headers to avoid rate-limiting
    headers = {
        "User-Agent": "your-email@example.com"  # Replace with your email
    }

    # API Request
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print("Error fetching data")
        return

    data = response.json()

    # Extract 10-K financials (example: Assets)
    if "us-gaap" in data["facts"] and "Assets" in data["facts"]["us-gaap"]:
        assets_data = data["facts"]["us-gaap"]["Assets"]["units"]["USD"]

        # Extract most recent values
        records = [{"date": item["end"], "assets": item["val"]} for item in assets_data]
        df = pd.DataFrame(records)

        # Write data back to Excel
        with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name="EDGAR Data", index=False)

        print("Data updated successfully!")


# Run function
fetch_edgar_data("ticker_file.xlsx")

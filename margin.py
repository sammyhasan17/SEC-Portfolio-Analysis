import requests

# Company name and CIK mapping
company_map = {
    "0000077360": "PENTAIR plc (PNR)",
    "0000945841": "POOL CORP (POOL)",
    "0000091142": "SMITH A O CORP (AOS)",
    "0000795403": "WATTS WATER TECHNOLOGIES INC (WTS)",
    "0001834622": "Hayward Holdings, Inc. (HAYW)",
    "0001833197": "Latham Group, Inc. (SWIM)",
    "0001821806": "Leslie's, Inc. (LESL)"
}

# US-GAAP keys to search for each metric
metrics = {
    "Net Sales": ["Revenues", "SalesRevenueNet", "RevenueFromContractWithCustomerExcludingAssessedTax"],
    "Gross Profit": ["GrossProfit"],  # Used for calculating Gross Margin
    "EBITDA": ["EarningsBeforeInterestTaxesDepreciationAndAmortization"],  # Only use if directly reported
    "COGS": ["CostOfGoodsSold"],  # Cost of Goods Sold for EBITDA calculation
    "SG&A": ["SellingGeneralAndAdministrativeExpense"],
    "Net Cashflow from Operations": ["NetCashProvidedByUsedInOperatingActivities"]
}

headers = {"User-Agent": "Sam Hasan sam@example.com"}

# Loop through companies
for cik in company_map:
    cik_padded = cik.zfill(10)
    name = company_map[cik]
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik_padded}.json"
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print(f"\n\n==============================")
        print(f"{name} | CIK: {cik_padded} | Failed to fetch data.")
        print(f"==============================")
        continue

    data = response.json()

    # To hold the most recent shared metadata and values
    header_details = None
    results = {}

    # Fetch and process the required financial metrics
    for label, possible_keys in metrics.items():
        most_recent = None

        for key in possible_keys:
            try:
                records = data["facts"]["us-gaap"][key]["units"]["USD"]
                valid_records = [r for r in records if r.get("form") in ["10-K", "10-Q"] and "end" in r]
                if not valid_records:
                    continue
                valid_records.sort(key=lambda x: x["end"], reverse=True)
                most_recent = valid_records[0]
                break
            except KeyError:
                continue

        if most_recent:
            if not header_details:
                header_details = {
                    "form": most_recent.get("form", "N/A"),
                    "fy": most_recent.get("fy", "N/A"),
                    "fp": most_recent.get("fp", "N/A"),
                    "end": most_recent.get("end", "N/A")
                }
            results[label] = most_recent.get("val")
        else:
            results[label] = "N/A"

    # Calculate Gross Margin as a percentage if not directly available
    gross_profit = results.get("Gross Profit", None)
    net_sales = results.get("Net Sales", None)

    if gross_profit and net_sales:
        try:
            gross_margin = (gross_profit / net_sales) * 100
            results["Gross Margin"] = f"{gross_margin:.2f}%"
        except ZeroDivisionError:
            results["Gross Margin"] = "N/A"

    # Calculate EBITDA as Net Sales - COGS - SG&A if not directly available
    net_sales = results.get("Net Sales", None)
    cogs = results.get("COGS", None)
    sg_a = results.get("SG&A", None)

    if net_sales and cogs and sg_a:
        try:
            ebitda = net_sales - cogs - sg_a
            results["EBITDA"] = f"${ebitda:,.0f}"
        except TypeError:
            results["EBITDA"] = "N/A"

    # Display the header line with most recent data context
    if header_details:
        print(f"\n\n==============================")
        print(
            f"{name} | CIK: {cik_padded} | {header_details['form']} | FY: {header_details['fy']} | Period: {header_details['fp']} | End: {header_details['end']}")
        print(f"==============================")

        # Display only the selected metrics
        for label in ["Net Sales", "Gross Margin", "EBITDA", "SG&A", "Net Cashflow from Operations"]:
            val = results.get(label, "N/A")
            val_str = f"${val:,.0f}" if isinstance(val, (int, float)) else val
            print(f"{label}: {val_str}")

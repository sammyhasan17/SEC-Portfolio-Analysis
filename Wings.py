import requests
import xlwings as xw
import msal


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
    "Gross Profit": ["GrossProfit"],
    "EBITDA": ["EarningsBeforeInterestTaxesDepreciationAndAmortization"],  # Only use if directly reported
    "SG&A": ["SellingGeneralAndAdministrativeExpense"],
    "Net Cashflow from Operations": ["NetCashProvidedByUsedInOperatingActivities"] # Net Cash used in operating activities
}

# Keys used for calculating EBITDA if not directly available
ebitda_components = {
    "Net Income": ["NetIncomeLoss", "ProfitLoss"],
    "Interest": ["InterestExpense", "InterestAndDebtExpense"],
    "Taxes": ["IncomeTaxExpenseBenefit"],
    "Depreciation": ["Depreciation", "DepreciationAndAmortization", "DepreciationDepletionAndAmortization"],
    "Amortization": ["AmortizationOfIntangibleAssets", "Amortization"]
}

headers = {"User-Agent": "Sam Hasan sam@example.com"}

all_company_data = {}
# main loop
for cik in company_map:
    cik_padded = cik.zfill(10)
    name = company_map[cik]
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik_padded}.json"
    response = requests.get(url, headers=headers)

# error response
    if response.status_code != 200:
        print(f"\n\n==============================")
        print(f"{name} | CIK: {cik_padded} | Failed to fetch data.")
        print(f"==============================")
        continue

    data = response.json()
    header_details = None
    results = []

    # Try to get reported EBITDA first
    reported_ebitda = None
    for key in metrics["EBITDA"]:
        try:
            records = data["facts"]["us-gaap"][key]["units"]["USD"]
            valid_records = [r for r in records if r.get("form") in ["10-K", "10-Q"] and "end" in r]
            if valid_records:
                valid_records.sort(key=lambda x: x["end"], reverse=True)
                reported_ebitda = valid_records[0]
                if not header_details:
                    header_details = {
                        "form": reported_ebitda.get("form", "N/A"),
                        "fy": reported_ebitda.get("fy", "N/A"),
                        "fp": reported_ebitda.get("fp", "N/A"),
                        "end": reported_ebitda.get("end", "N/A")
                    }
                break
        except KeyError:
            continue

    # Fetch the main metrics (but not EBITDA)
    for label, possible_keys in metrics.items():
        if label == "EBITDA":
            continue  # skip, already handled above or below
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
            results.append((label, most_recent.get("val")))
        else:
            results.append((label, "N/A"))


    # Handle EBITDA: use reported or estimate it
    if reported_ebitda:
        results.append(("Reported EBITDA", reported_ebitda.get("val")))
    else:
        # Try estimating EBITDA
        ebitda_vals = {}
        for label, possible_keys in ebitda_components.items():
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
                ebitda_vals[label] = most_recent.get("val")

        # Calculate estimated EBITDA (Note: see readME for how EBITDA has been handled and why)
        required_parts = ["Net Income", "Interest", "Taxes", "Depreciation", "Amortization"]
        if all(part in ebitda_vals for part in required_parts):
            estimated_ebitda = sum(ebitda_vals[part] for part in required_parts)
            results.append(("Estimated EBITDA", estimated_ebitda))
        else:
            results.append(("Estimated EBITDA", "N/A"))

    # Calculate Gross Margin if Gross Profit and Net Sales are available
    gross_profit = next((val for label, val in results if label.startswith("Gross Profit") and isinstance(val, (int, float))), None)
    net_sales = next((val for label, val in results if label == "Net Sales" and isinstance(val, (int, float))), None)

    if gross_profit is not None and net_sales:
        gross_margin = (gross_profit / net_sales) * 100
        results.append(("Gross Margin (%)", round(gross_margin, 2)))
    else:
        results.append(("Gross Margin (%)", "N/A"))

    if header_details:
    # Print a header for this company's data in the console
        print(f"\n\n==============================")
        print(f"{name} | CIK: {cik_padded} | {header_details['form']} | FY: {header_details['fy']} | Period: {header_details['fp']} | End: {header_details['end']}")
        print(f"==============================")
    
    # Print each metric and its value in the console
    for label, val in results:
        if isinstance(val, (int, float)):
            val_str = f"{val:.0f}%" if "Margin" in label else f"${val:,.0f}"
        else:
            val_str = val
        print(f"{label}: {val_str}")

    # Open the Excel workbook
    wk = xw.books.open(r'C:\Users\Sammy\OneDrive\Documents\GitHub\SEC-to-EXCEL\ticker_file.xlsm')

    # Select the 'Data' sheet
    sheet = wk.sheets('Data')
    start_row = 1
    start_col = 1  # Column A

    # Find the last used row in the first column to append new data after a blank row
    last_row = sheet.range((sheet.cells.last_cell.row, start_col)).end('up').row
    write_row = last_row + 2 if last_row >= start_row else start_row

    # Write the company header info to the sheet
    sheet.range((write_row, start_col)).value = [
        f"{name} | CIK: {cik_padded} | {header_details['form']} | FY: {header_details['fy']} | Period: {header_details['fp']} | End: {header_details['end']}"
    ]
    write_row += 1

    # Write each metric and its value to the sheet, one per row
    for label, val in results:
        if isinstance(val, (int, float)):
            val_str = f"{val:.0f}%" if "Margin" in label else f"${val:,.0f}"
        else:
            val_str = val
        sheet.range((write_row, start_col)).value = [label, val_str]
        write_row += 1

# success output


print('####################')
print('Program completed')
print('####################')



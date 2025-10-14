
import requests
import xlwings as xw
from datetime import datetime

# Company name and CIK mapping
company_map = {
    "0000732717": "AT&T (T)",                     # Telecommunications & media
    "0000354950": "HOME DEPOT INC. (HD)",         # Retail - home improvement & building materials
    "0001318605": "Tesla Inc. (TSLA)",            # Automobiles & clean energy
    "0000936468": "Lockheed (LKH)",               # Aerospace & defense
    "0001045810": "NVIDIA (NVDA)",                # Semiconductors - GPUs & AI hardware
    "0000320187": "NIKE (NKE)",                   # Consumer goods - apparel & footwear
    "0000320193": "Apple (AAPL)",                 # Consumer electronics & software
    "0000789019": "Intel Corp. (INTC)",           # Semiconductors
    "0000104169": "Caterpillar Inc. (CAT)",       # Heavy machinery
    "0000815097": "3M Company (MMM)",             # Industrial & materials
    "0000021344": "General Electric (GE)",        # Diversified manufacturing
    "0000018230": "Ford Motor Co. (F)",           # Automobiles
    "0000051143": "Boeing (BA)",                  # Aerospace
    "0000066740": "General Motors (GM)",          # Automobiles
    "0001467858": "Honeywell (HON)",              # Industrial tech & systems
    "0000002488": "Advanced Micro Devices (AMD)"  # Semiconductors - CPUs & GPUs
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

all_company_data = []
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

        # FORMATTING LOGIC
        if isinstance(val, (int, float)):
            val_str = f"{val:.0f}%" if "Margin" in label else f"${val:,.0f}"
        else:
            val_str = val
            # Print to console
        print(f"{label}: {val_str}")


        ###############################################################
        #  POPULATE DATA STRCUTURE FOR OUR DATA TO TRANSPOSE LATER   #
        ###############################################################
        
        # testing global variable
        # Build a row dict for this company
    row_dict = {
        "Company": name,
        "CIK": cik_padded,
        "Form": header_details.get("form", "N/A"),
        "FY": header_details.get("fy", "N/A"),
        "Period": header_details.get("fp", "N/A"),
        "End Date": header_details.get("end", "N/A"),
    }

    # add metrics from results
    for label, val in results:
        row_dict[label] = val


    # append to global list
    all_company_data.append(row_dict)
    # could we optimze this so we could re-use results list?


    # Open the Excel workbook
    wk = xw.books.open(r'C:\Users\dele2\OneDrive\Documents\GitHub\SEC-Portfolio-Analysis\ticker_file.xlsm')


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
        # add it to a global list

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
print('\n')

# ---------- Write transposed row to 'PowerBI_Data' (headers across row 1) ----------

# 0) Get/create the destination sheet
try:
    sheet2 = wk.sheets['PowerBI_Data']
except Exception:
    sheet2 = wk.sheets.add('PowerBI_Data', after=wk.sheets[-1])

# 1) Define the header labels (left to right on row 1)
# headers = [
#     "Company","Form","FY","Period","End Date",
#     "Net Sales","Gross Profit","SG&A","Net Cashflow from Operations",
#     "Reported EBITDA","Estimated EBITDA","Gross Margin (%)"
# ]

# sheet2.range("A1").value = [headers]   # <- writes horizontally across row 1

# Write each metric and its value to the sheet, one per row
print(all_company_data) # JSON Formatted list



import csv
import sys
import requests

headers = list(all_company_data[0].keys())
writer = csv.DictWriter(sys.stdout, fieldnames=headers)
writer.writeheader()
writer.writerows(all_company_data) # CSV Output

# Select the 'Data' sheet
sheet = wk.sheets('PowerBI_Data')
start_row = 1
start_col = 1  # Column A

# sheet2['A1'].value = "hello world"

# Define column order explicitly so Excel columns are consistent
headers = [
    "Company","CIK","Form","FY","Period","End Date",
    "Net Sales","Gross Profit","SG&A","Net Cashflow from Operations",
    "Estimated EBITDA","Gross Margin (%)"
]
# Write headers into the first row of the sheet (A1 → across to the right).
# Wrapping headers in [ ... ] makes xlwings treat it as a single row instead of a column.
sheet2['A1'].value = [headers]   # row of headers

# Create an empty list that will hold each company's data row
rows = []

# Loop through every company dictionary in all_company_data
for company in all_company_data:
    # Build a row list following the exact header order
    # row_dict.get(col, "N/A") → get the value for header col,
    # if missing, put "N/A" instead so all rows have same length
    row = [company.get(col, "N/A") for col in headers]
    
    # Add this row to the rows list
    rows.append(row)

# Write all rows to Excel starting at A2 (below headers).
# xlwings will expand this 2D list into multiple rows/columns automatically.
sheet2['A2'].value = rows


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



# ========================================================================
# 🔗 PUSH DATA TO POWER BI (REST API / Azure AD (Azure Active Directory) & Authentication with MSAL)
# ========================================================================

# Python script → Azure AD (Azure Active Directory) → Access Token → Power BI REST API (datasets/reports/refresh) 
import msal # Microsfot Authentication Library

# set up our .ENV file & loading env variables
from dotenv import load_dotenv
import os
# env variables
load_dotenv()

# defining our credentials
TENANT_ID = "aba21ed8-6044-4926-826d-5c9eb6d37ead" # Tenant ID → tells Azure which organization directory the app belongs to.
CLIENT_ID = "c8e887ce-d043-4928-aad1-0e0baf3d4c6d" # Client ID → the username of your app.
CLIENT_SECRET = os.getenv("CLIENT_SECRET")  # Client Secret VALUE → the password of your app.

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]

# creating an app variable
app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)


token = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" in token:
    access_token = token["access_token"]
    # Print only first/last 30 chars to confirm
    print("✅ Got token:")
    print(access_token[:30] + "..." + access_token[-30:])
else:
    print("❌ Failed to get token:", token.get("error_description", token))

# =============================================================================
# 🚀 CREATE NEW DATASET IN FABRIC (via REST API + Access Token)
# =============================================================================
# This script:
# 1. Uses your Azure AD access token (from MSAL) to authenticate.
# 2. Creates a new push dataset inside the specified Power BI workspace.
# 3. Defines a table schema 
# 
# After running:
# - You’ll see some "DataSetName" appear in your Power BI Service (Fabric) workspace.
# - Reports in Power BI Service can now connect to this dataset.
# - Rows can be pushed into this dataset using the Push Rows API.
# =============================================================================

import requests


workspace_id = "4be12a96-fc3d-4ec3-b048-6feefac3f861"   # Fabric Workspace
dataset_name = "PythonDataset" # TODO: change name to something more useful 
table_name = 'CompanyData'

# url for our structure
url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets"


for row in all_company_data:
    row.pop("CIK", None)



# ✅ Headers must be a dict
headers = {
    "Authorization": f"Bearer {access_token}", 
    "Content-Type": "application/json"
}

# Example headers list from your schema
headers_list = ["Company","CIK","Form","FY","Period","End Date",
    "Net Sales","Gross Profit","SG&A","Net Cashflow from Operations",
    "Estimated EBITDA","Gross Margin (%)"]

# =====================================================================
# 🧹 Clean Data Before Pushing Into Payload JSON
#
# This step ensures our rows_payload is clean and consistent with the
# dataset schema before sending it to the REST API.
# =====================================================================

for row in all_company_data:
    for col in ["Gross Profit", "SG&A", "Net Cashflow from Operations",
                "Estimated EBITDA", "Gross Margin (%)"]:
        if row[col] in ["N/A", "", None]:
            row[col] = None


        

# ✅ Build schema - 
fields = []
for col in headers_list:
    if col in ["FY", "Net Sales", "Gross Profit", "SG&A", 
               "Net Cashflow from Operations", "Estimated EBITDA"]:
        dtype = "Int64" # whole numbers
    elif col == "Gross Margin (%)":
        dtype = "Double" # decimals
    else:
        dtype = "String"
    fields.append({"name": col, "dataType": dtype})

# Auto-name dataset with today's month-day (avoids duplicates)
#  Add month-day to dataset name

today = datetime.today().strftime("%m-%d")
dataset_name = f"SECDataset {today}"

# ============================================================================
#  Build JSON payload for dataset creation
#
# "defaultMode": "Push" means rows are inserted via API (not from files)
# ============================================================================


payload = {
    "name": dataset_name,
    "defaultMode": "Push",
    "tables": [
        {
            "name": "CompanyData",
            "columns": fields,

        }
    ]
}

# wrap our data (list of dictionaries) into rows
rows_payload = {"rows": all_company_data} # { "rows": [ {...}, {...}, {...} ] }



# ============================================================================
#  Create dataset in Fabric workspace (via Power BI REST API)
# ============================================================================
# - POST request goes to the Power BI REST API endpoint
# - Authorization header carries your Azure AD access token
# - Payload is the dataset definition (name + tables + schema)
# - If successful → API returns status_code 201 and JSON with dataset ID
# - If failed → status_code + error details help debug (401 = bad token, 
#   403 = missing permission, 400 = payload problem, etc.)
# ============================================================================

# ✅ Make requests

# fields
response = requests.post(url, headers=headers, json=payload)
print("Create Fields:", response.status_code, response.text)

# Handle different responses
if response.status_code == 201:
    dataset_id = response.json()["id"]
    print(f"✅ Dataset '{dataset_name}' created successfully!")
    print("   Dataset ID:", dataset_id)
elif response.status_code == 401:
    print("❌ Unauthorized (check your access token).")
elif response.status_code == 403:
    print("❌ Forbidden (you may not have rights in this workspace).")
elif response.status_code == 400:
    print("❌ Bad Request (check payload format / schema).")
else:
    print("⚠️ Unexpected error:", response.status_code, response.text)


# get dataset id from response
dataset_id = response.json()["id"]
print(dataset_id, "DATASET ID ! ")
# url for our records
rows_url  = f'https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}/tables/{table_name}/rows'


# records
row_response = requests.post(rows_url,headers=headers, json=rows_payload)
print("Insert rows:", row_response.status_code, row_response.text)


# Handle row responses
if row_response.status_code == 200:
    print("✅ Rows inserted successfully!")
elif row_response.status_code == 400:
    print("❌ Bad Request (check your row payload, column names, datatypes).")
elif row_response.status_code == 401:
    print("❌ Unauthorized (access token invalid/expired).")
elif row_response.status_code == 403:
    print("❌ Forbidden (no permission to insert rows in this dataset).")
elif row_response.status_code == 404:
    print("❌ Not Found (dataset_id or table_name is wrong).")
elif row_response.status_code >= 500:
    print("❌ Server error at Power BI side, try again later.")
else:
    print("⚠️ Unexpected rows response:", row_response.status_code, row_response.text)



# We got records to popoulate in Fabric! now use fabric to make dashboards in the cloud!



dashboard_name = "Market Overview"

dash_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/dashboards"
dash_payload = {"name": dashboard_name}

dash_resp = requests.post(dash_url, headers=headers, json=dash_payload)
print("Create dashboard:", dash_resp.status_code, dash_resp.text)

if dash_resp.status_code in (200, 201):
    dashboard_id = dash_resp.json()["id"]
    print("✅ Dashboard created:", dashboard_name, dashboard_id)
else:
    raise RuntimeError(f"Dashboard creation failed: {dash_resp.status_code} {dash_resp.text}")



# ---------------------------------------------------------------------
# • Q&A tiles   → created by Power BI’s natural-language engine.
#                 Example: ask “total [Net Sales] by [Company]”
#                 No report required, but can fail if Q&A is disabled
#                 or column names aren’t recognized.
#
# • Report tiles → pinned visuals that already exist in a Power BI
#                  report.  100 % reliable and works in every tenant.
#                  Requires a reportId and the visual/page names.
#
# This script uses the REPORT-BASED approach.
# ---------------------------------------------------------------------


# CREATES TILES THAT BUILD UP OUR DASHBOARD  # from your push-dataset creation step
# tiles_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/dashboards/{dashboard_id}/tiles"
# do we need this also ?


report_id = 'e780a33a-edc4-49c2-b5a1-102bda83408f' # this will change ?


# we want to automate the report creation (last 10%) 

# 1. bring in dataset to fabric -> 2. create a report in fabric with that data 
# (optinally) update that same report with new data


# ============================================================================
# 🆕 NEW SECTION: AUTOMATED REPORT CREATION
# ============================================================================

def create_report(workspace_id, dataset_id, report_name, access_token):
    """Creates a new blank report connected to the dataset."""
    report_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "name": report_name,
        "datasetId": dataset_id
    }
    
    response = requests.post(report_url, headers=headers, json=payload)
    
    if response.status_code in (200, 201):
        report_id = response.json()["id"]
        print(f"\n✅ Report '{report_name}' created successfully!")
        print(f"   Report ID: {report_id}")
        return report_id
    else:
        print(f"\n❌ Failed to create report: {response.status_code}")
        print(response.text)
        return None


def clone_report_from_template(workspace_id, template_report_id, dataset_id, report_name, access_token):
    """
    Clones an existing template report and rebinds it to the new dataset.
    This is the RECOMMENDED approach for production automation.
    """
    clone_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports/{template_report_id}/Clone"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "name": report_name,
        "targetModelId": dataset_id,
        "targetWorkspaceId": workspace_id
    }
    
    response = requests.post(clone_url, headers=headers, json=payload)
    
    if response.status_code in (200, 201):
        report_id = response.json()["id"]
        print(f"\n✅ Report cloned from template!")
        print(f"   New Report ID: {report_id}")
        return report_id
    else:
        print(f"\n❌ Clone failed: {response.status_code}")
        print(response.text)
        return None


def pin_report_to_dashboard(workspace_id, report_id, dashboard_id, access_token):
    """
    Pins the entire report to the dashboard.
    Note: For individual visuals, you need page/visual names.
    """
    # Get report pages first
    pages_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports/{report_id}/pages"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    response = requests.get(pages_url, headers=headers)
    
    if response.status_code == 200:
        pages = response.json().get("value", [])
        if pages:
            first_page = pages[0]
            page_name = first_page.get("name", "ReportSection")
            print(f"   Found report page: {page_name}")
            
            # Note: To pin specific visuals, you'd need visual IDs
            # For now, we'll just confirm the report is ready
            print(f"   Report is ready to be manually pinned to dashboard")
            print(f"   Or use Power BI Service UI to pin visuals")
            return True
    else:
        print(f"   Could not retrieve report pages: {response.status_code}")
        return False


# ============================================================================
# EXECUTE REPORT AUTOMATION
# ============================================================================

print("\n" + "="*60)
print("🎨 STARTING REPORT AUTOMATION")
print("="*60)

report_name = f"SEC Analysis Report {today}"

# # Option 1: Create blank report (uncomment to use)
# report_id = create_report(workspace_id, dataset_id, report_name, access_token)

# Option 2: Clone from template (RECOMMENDED - uncomment and add template ID)
template_report_id = "e780a33a-edc4-49c2-b5a1-102bda83408f" # long string after /reports/

clone_report_from_template(
    workspace_id, 
    template_report_id,      # Your original report's structure
    dataset_id,              # YOUR NEW DATASET (created today with fresh SEC data)
    report_name,             # New name like "SEC Analysis Report 10-07"
    access_token
)

if report_id:
    # Try to get report structure
    pin_report_to_dashboard(workspace_id, report_id, dashboard_id, access_token)
    
    print("\n" + "="*60)
    print("✅ AUTOMATION COMPLETE!")
    print("="*60)
    print(f"📊 Dataset: {dataset_name} (ID: {dataset_id})")
    print(f"📈 Dashboard: {dashboard_name} (ID: {dashboard_id})")
    print(f"📋 Report: {report_name} (ID: {report_id})")
    print("\n🔗 Next Steps:")
    print("   1. Open Power BI Service and navigate to your workspace")
    print("   2. Open the report and add visuals manually")
    print("   3. OR create a template report and use clone_report_from_template()")
    print("   4. Pin visuals from the report to your dashboard")
    print("="*60)
else:
    print("\n❌ Report creation failed. Check errors above.")

print("\n✅ Script execution finished!")
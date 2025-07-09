import pandas as pd
import requests
import xlwings as xw
#
wk = xw.books.open(r'C:\Users\sam.hasan\PyCharmMiscProject\ticker_file.xlsm')
sheet = wk.sheets('Sheet3')
wk.sheets["Sheet3"].range("N8").value = 'hi2'

#
# df = sheet.range("A1:A7").options(pd.DataFrame).value
#
# # xw.view(df) # opens a new excel workbook with our df
# # wb = xw.Book(df) # put df in out workbook called sheet
# wk.sheets["Sheet3"].range("C1").value = "Loaded from xlwings!"
#
# wk.save
#
# # create request header
# headers = {'User-Agent': "sammyhasan17@gmail.com"}
#
# # get all companies data
# companyTickers = requests.get(
#     "https://www.sec.gov/files/company_tickers.json",
#     headers=headers
#     # todo update tickers to competitors only (list in peers.txt)
#     )
#
# # review response / keys
# print(companyTickers.json().keys())
#
# # format response to dictionary and get first key/value
# firstEntry = companyTickers.json()['0']
#
# # parse CIK // without leading zeros
# directCik = companyTickers.json()['0']['cik_str']
#
# # dictionary to dataframe
# companyData = pd.DataFrame.from_dict(companyTickers.json(),
#                                      orient='index')
#
# # add leading zeros to CIK
# companyData['cik_str'] = companyData['cik_str'].astype(
#                            str).str.zfill(10)
#
# # review data
# print(companyData[:20])
#
# wk.sheets["Sheet3"].range("N8").value = companyData[:11]
#
# # Apple CIK assigned
# cik = companyData[0:1].cik_str[0]
#
# # get a company's specific filing metadata (cik)
# filingMetadata = requests.get(
#     f'https://data.sec.gov/submissions/CIK{cik}.json',
#     headers=headers
#     )
#
# # review json
# print(filingMetadata.json().keys())
# filingMetadata.json()['filings']
# filingMetadata.json()['filings'].keys()
# filingMetadata.json()['filings']['recent']
# filingMetadata.json()['filings']['recent'].keys()
#
# # dictionary to dataframe
# allForms = pd.DataFrame.from_dict(
#              filingMetadata.json()['filings']['recent']
#              )
#
# # # review columns
# allForms.columns
#
# # Other SEC API calls ...
# #
# # 10-Q metadata
# allForms.iloc[11]
#
#
#
# # get company facts data
# companyFacts = requests.get(
#     f'https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json',
#     headers=headers
#     )
#
# #review data
# companyFacts.json().keys()
# companyFacts.json()['facts']
# companyFacts.json()['facts'].keys()
#
# # filing metadata
# companyFacts.json()['facts']['dei'][
#     'EntityCommonStockSharesOutstanding']
# companyFacts.json()['facts']['dei'][
#     'EntityCommonStockSharesOutstanding'].keys()
# companyFacts.json()['facts']['dei'][
#     'EntityCommonStockSharesOutstanding']['units']
# companyFacts.json()['facts']['dei'][
#     'EntityCommonStockSharesOutstanding']['units']['shares']
# companyFacts.json()['facts']['dei'][
#     'EntityCommonStockSharesOutstanding']['units']['shares'][0]
#
# # concept data // financial statement line items
# companyFacts.json()['facts']['us-gaap']
# companyFacts.json()['facts']['us-gaap'].keys()
#
# # different amounts of data available per concept
# companyFacts.json()['facts']['us-gaap']['AccountsPayable']
# companyFacts.json()['facts']['us-gaap']['Revenues']
# companyFacts.json()['facts']['us-gaap']['Assets']
# companyFacts.json()['facts']['us-gaap']['CostOfGoodsAndServicesSold']
#
# # get company concept data
# companyConcept = requests.get(
#     (
#     f'https://data.sec.gov/api/xbrl/companyconcept/CIK{cik}'
#      f'/us-gaap/Assets.json'
#     ),
#     headers=headers
#     )
#
# # review data
# companyConcept.json().keys()
# companyConcept.json()['units']
# companyConcept.json()['units'].keys()
# companyConcept.json()['units']['USD']
# companyConcept.json()['units']['USD'][0]
#
# # parse assets from single filing
# companyConcept.json()['units']['USD'][0]['val']
#
# # get all filings data
# assetsData = pd.DataFrame.from_dict((
#                companyConcept.json()['units']['USD']))
#
# # review data
# assetsData.columns
# assetsData.form
#
# # get assets from 10Q forms and reset index
# assets10Q = assetsData[assetsData.form == '10-Q']
# assets10Q = assets10Q.reset_index(drop=True)
#
# # plot
# assets10Q.plot(x='end', y='val')
#
#
#
# print(companyData)
#
# # Define the list of CIKs to keep (We have everything except Fluidra)
# target_ciks = ['0000945841','0000091142','0001852345', '0000795403', '0001821806','0000077360', '0001833197','0001834622']
#
# # Format and filter CIKs
# filteredData = companyData[companyData['cik_str'].isin(target_ciks)]
#
# print(filteredData)
#
# # Loop through the filtered CIKs and get filing metadata
# #
#
# wk.sheets["Sheet3"].range("N22").value = filteredData
#
#
# filteredData.pop('cik_str')  # removes 'cik_str' if it exists, avoids KeyError
# filteredData.pop('ticker')
# # print(type(filteredData))
#
#
#
# print(filteredData)
#
# filteredData = filteredData.values.tolist()
#
# print(filteredData)
#
# wk.sheets["Sheet3"].range("O32").value = filteredData
#
# # wk.sheets["Sheet3"].range("N22").value = filteredData
#
# # wk.sheets["Sheet3"].range("N22").value = df
# # wk.sheets["Sheet3"].range("N22").value
#
# # assetsData = pd.DataFrame.from_dict((
# #                companyConcept.json()['units']['USD']))
#
# # testing code (POOL CORP)
#
# import requests
#
# CIK = "0000945841"
# url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json"
#
# headers = {
#     "User-Agent": "Your Name (your_email@example.com)"
# }
#
# response = requests.get(url, headers=headers)
# data = response.json()
#
# # List of potential revenue keys to check, we only care about net sales = net revenue for now
# revenue_keys = [
#     # "Revenues",
#     "SalesRevenueNet",
#     # "RevenueFromContractWithCustomerExcludingAssessedTax"
# ]
#
# for key in revenue_keys:
#     try:
#         items = data["facts"]["us-gaap"][key]["units"]["USD"]
#         print(f"\n--- {key} ---")
#         for item in items:
#             if item["form"] == "10-K" and item.get("fy") == 2018:
#                 print(f"Year: {item['fy']} - Revenue: ${item['val']:,}")
#     except KeyError:
#         continue
#
# # testing edgar tools
# # 1. Import the necessary functions from edgartools
# import edgar
# from edgar import Company, set_identity
# # 2. Tell the SEC who you are
# set_identity("sammyhasan17@gmail.com")
#
# # 3. Start using the library
#
#
# import requests
#
# # List of CIKs to target is above ' tagrget ciks'
#
# # Financial metrics and corresponding US-GAAP tags
# metrics = {
#     "Net Sales": ["Revenues", "SalesRevenueNet", "RevenueFromContractWithCustomerExcludingAssessedTax"],
#     "Gross Margin": ["GrossProfit"],
#     "EBITDA": ["EarningsBeforeInterestTaxesDepreciationAndAmortization"],
#     "SG&A": ["SellingGeneralAndAdministrativeExpense"],
#     "Net Cashflow from Operations": ["NetCashProvidedByUsedInOperatingActivities"]
# }
#
# # Loop through each company
# for cik in target_ciks:
#     print(f"\n\n====== CIK: {cik} ======\n")
#     url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik.zfill(10)}.json"
#     headers = {"User-Agent": "YourName Contact@domain.com"}
#     response = requests.get(url, headers=headers)
#
#     if response.status_code != 200:
#         print(f"Failed to fetch data for CIK {cik}")
#         continue
#
#     data = response.json()
#
#     for label, keys in metrics.items():
#         found = False
#         print(f"\n--- {label} ---")
#
#         for key in keys:
#             try:
#                 entries = data["facts"]["us-gaap"][key]["units"]["USD"]
#                 found = True
#
#                 for item in entries:
#                     form = item.get("form")
#                     fy = item.get("fy")
#                     fp = item.get("fp")  # fiscal period: Q1, Q2, FY, etc.
#                     val = item.get("val")
#                     if form in ["10-K", "10-Q"]:  # limit to annual and quarterly
#                         print(f"{form} | FY: {fy} | Period: {fp} | Value: ${val:,}")
#             except KeyError:
#                 continue
#
#         if not found:
#             print("No data found.")
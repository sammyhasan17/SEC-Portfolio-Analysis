Here is the final **README in full Markdown format**, with the **code snippets removed** from the end, while keeping the plain-English explanations of each financial metric:

````markdown
# 📈 SEC Filing EBITDA Extractor & Excel Automation

Automates extraction or estimation of EBITDA from SEC filings and writes results to Excel using Python and VBA.

---

## 📦 Dependencies

Install the required Python libraries:

```bash
pip install requests xlwings msal
````

---

## 📊 EBITDA Handling Logic

### ✅ If Directly Reported

Use the XBRL tag:

```
EarningsBeforeInterestTaxesDepreciationAndAmortization
```

### ⚙️ If Not Reported

Estimate EBITDA using this standard formula:

```
EBITDA = Net Income
       + Interest
       + Taxes
       + Depreciation
       + Amortization
```

---

### 🔁 Alternative Formula (for Verification)

```
EBITDA = Net Sales
       – Operating Expenses (excluding Depreciation & Amortization)
```

---

## ❌ Why We Do NOT Use Adjusted EBITDA

* Adjusted EBITDA includes custom, non-standard company-specific adjustments (e.g., stock-based compensation, restructuring)
* These adjustments are inconsistent across companies
* Not reported in machine-readable XBRL data — typically only available in press releases or presentations

---

## 🧮 Calculation Notes

* **Q4 Calculation**

  ```
  Q4 = Annual - Q1 - Q2 - Q3
  ```

* **Gross Margin**
  Rounded to the nearest integer

* **Net Cashflow from Operations**
  Uses the tag:

  ```
  NetCashProvidedByUsedInOperatingActivities
  ```

---

## ⚙️ Configuration Notes

If the program stops working:

* ✅ Re-save the Excel file to the same location (overwrite it)
* ✅ Ensure no other programs are using the file
* ✅ Close all Excel windows (multiple instances can break VBA)

### 🛠 Root Cause + Fix

Python couldn’t locate the file due to a relative path issue. Re-saving (Save As) refreshed the file reference.

> ✅ **Best Practice:** Use absolute file paths for reliability.

---

## 📌 Assumptions

We compare companies using their **most recent SEC filings**, regardless of fiscal calendar.

### Why?

* Keeps comparisons fresh and aligned to the same time periods
* Ensures data consistency for dashboards updated quarterly

---

## 🧠 What Do These Metrics Mean?

Net Sales → topline growth

Gross Profit → production efficiency

EBITDA → core operating health

SG&A → overhead efficiency

Net Cash from Ops → true financial strength

# Usefuleness
1. Acquisition Targeting
Insight: Find companies with declining revenue but strong EBITDA or cash flow
Action: Flag as potential acquisition targets — cost-cutting and growth opportunities

2. Overhead Efficiency
Insight: Benchmark SG&A % of revenue vs. competitors
Action: If ours is higher, reduce overhead through org streamlining or vendor renegotiation

3. Margin & Pricing Strategy
Insight: Compare gross profit and EBITDA margins across competitors
Action: If others have better margins, explore price increases or cost savings; if we lead, scale high-margin segments

4. Cash Flow Risk Monitoring
Insight: Spot competitors with strong EBITDA but weak or negative operating cash flow
Action: These firms may be unstable — capture their customers, recruit their laid-off staff, or prepare to acquire assets if they go under
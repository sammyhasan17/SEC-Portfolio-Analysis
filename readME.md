

# 📈 SEC Filing EBITDA Extractor & Excel Automation

Automates extraction or estimation of EBITDA from SEC filings and writes results to Excel using Python and VBA.

---
## 📊 Flowchart 

![Flowchart Automated SEC Dashboard](flowchart.png)


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
* These adjustments are not standardized in the SEC’s US-GAAP taxonomy.
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

* Ensures comparisons are across similar calendar windows
* Keeps analysis consistent and fresh
* Designed for Quarterly & Annual Reports


---

## 🧠 What Do These Metrics Mean?

Net Sales → how much we are earning

Gross Profit → production efficiency

EBITDA → core operating health

SG&A → overhead efficiency

Net Cash from Ops → true financial strength

# Why is this useful?

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


Net Cash Flow from Operations corresponds to "Net cash used in operating activities" on SEC Filings


# Assumptions
We assume it's best to compare companies using their most recent SEC filings, rather than aligning by fiscal quarter. Our dashboard updates quarterly, so comparing performance across the most recent calendar windows ensures a consistent and timely view—despite different fiscal calendars across companies.
We assume comparing the best way is to use the most recent SEC filings from each company rather than aligning by fiscal quarter,
because our dashboard is updated quarterly with new data.
This approach ensures we're comparing performance across similar calendar windows,
even though companies may operate on different fiscal calendars.

# Demo
https://github.com/user-attachments/assets/7a2ac980-1004-40d8-bc28-75a8ee6a81bc


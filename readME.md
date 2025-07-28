Absolutely — here’s the **complete `README.md` content**, perfectly formatted to paste directly into a Markdown file or a terminal that supports Markdown rendering:

---

````markdown
# 📈 SEC Filing EBITDA Extractor & Excel Automation

Automates the extraction or estimation of EBITDA from SEC filings and writes results to Excel using Python and VBA.

---

## 📦 Dependencies

Install the required Python libraries:

```bash
pip install requests xlwings msal
````

---

## 📊 EBITDA Handling Logic

### ✅ If Directly Reported

Use the `EarningsBeforeInterestTaxesDepreciationAndAmortization` tag from SEC XBRL filings.

### ⚙️ If Not Reported

Estimate EBITDA using the standard formula:

```
EBITDA = Net Income
       + Interest
       + Taxes
       + Depreciation
       + Amortization
```

---

### ❌ Why We Do NOT Use Adjusted EBITDA

* Includes non-standard custom adjustments (e.g., stock comp, restructuring).
* Not consistent across companies or available in XBRL.
* Often found only in **press releases** or **investor presentations**.

---

## ⚙️ Configuration Notes

If the program stops working:

* ✅ Re-save the Excel file to the same location (overwrite it).
* ✅ Make sure **no other programs** are using the file.
* ✅ **Close all Excel windows** – multiple instances may cause issues.

### 🛠 Root Cause + Fix

Python couldn't find the file because of a relative path issue. Re-saving (Save As) made the file readable.
**✅ Recommended: Use absolute paths for reliability.**

---

## 🧮 Calculation Notes

* **Q4 Calculation**:

  ```
  Q4 = Annual - Q3 - Q2 - Q1
  ```

* **Gross Margin**: Rounded to the nearest integer

* **Net Cashflow from Operations**: Uses `Net cash used in operating activities`

### 🔁 Alternative EBITDA Formula (for verification)

```
EBITDA = Net Sales
       – Operating Expenses (excluding Depreciation & Amortization)
```

---

## 📌 Assumptions

We compare companies using their **most recent SEC filings**, regardless of fiscal calendars.

### Why?

* Ensures comparisons are across **similar calendar windows**
* Keeps analysis consistent and fresh
* Our **dashboard is updated quarterly** with new filings

---

```

Let me know if you’d like this saved into a file or adapted for Jupyter, Notion, or a portfolio site.
```



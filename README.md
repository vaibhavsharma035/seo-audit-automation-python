# SEO Audit Automation using Python

## 📌 Overview

This project automates SEO validation for web pages by comparing **expected vs actual** values of:

* H1 Tags
* Title Tags
* Meta Descriptions

It reads input URLs from an Excel sheet and generates a detailed SEO audit report with pass/fail results.

---

## 🚀 Features

* Reads URLs from Excel (Ops Center sheet)
* Extracts:

  * H1 tag (and count)
  * Page title
  * Meta description
* Compares with expected values
* Generates Excel report with:

  * PASS / FAIL / MISSING status
* Highlights failed results in red
* Handles errors & broken URLs
* Adds delay to prevent blocking

---

## 🛠️ Tech Stack

* Python
* Pandas
* Requests
* BeautifulSoup
* OpenPyXL

---

## 📂 Input File Format

Excel sheet must contain the following columns:

* `42Works - Staging Site Link`
* `New H1`
* `Live Title Tag`
* `Live Meta Description`

Sheet Name:

```
Ops Center
```

---

## ▶️ How to Run

### 1. Clone the repository

```
git clone https://github.com/your-username/seo-audit-automation-python.git
cd seo-audit-automation-python
```

### 2. Install dependencies

```
pip install -r requirements.txt
```

### 3. Add your Excel file

Update file name inside script:

```python
input_file = "your_file.xlsx"
```

### 4. Run the script

```
python seo_audit.py
```

---

## 📊 Output

* Generates file:

```
seo_results_YYYYMMDD_HHMMSS.xlsx
```

### Output includes:

* Status Code
* Actual H1 / Title / Meta
* H1 Count
* PASS / FAIL Results
* ❌ Failed cells highlighted in red

---

## ⚠️ Notes

* SSL verification is disabled for flexibility
* Some websites may block scraping
* Delay added to avoid rate limiting

---

## 💡 Future Improvements

* Add parallel processing for speed
* Add logging system
* Export results to dashboard (Power BI / Streamlit)
* Add keyword density checks
* Handle JavaScript-rendered pages (Selenium)

---

##  Author

Vaibhav Sharma


## ⭐ If you found this useful, give it a star!


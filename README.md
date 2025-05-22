# Tender-Flow: Automated Tender Monitoring & Reporting System with Excel Report Generator

TenderFlow is a daily Python automation technique that scrapes railway tender data from the [IREPS.gov.in](https://www.ireps.gov.in) website, formats it into a clean Excel report, and emails it to clients or product managers. It's designed for companies or individuals looking to automate tender monitoring and reporting.

---

## 🚀 Features

* Scrapes railway tender data daily using Selenium
* Extracts key info like PO value, Supplier, Description, Qty, etc.
* Saves well-formatted Excel reports (one file per day)
* Emails the report automatically using yagmail
* Scheduled via Windows Task Scheduler to run daily
* No need to open the website manually every day!

---

## 📂 Folder Structure

```
TenderFlow/
├── python_code.ipynb      # Python script
├── requirements.txt       # Python dependencies
├── README.md              # This file
├── .env.example           # Sample for email credentials (optional)
├── Terms and conditions   # MIT License
├── ppt/presentation       # to explain about this project
```


## 🛠️ Technologies Used

* **Python 3**
* **Selenium** for browser automation
* **BeautifulSoup** for parsing HTML
* **Pandas** for data handling
* **OpenPyXL** for Excel formatting
* **Yagmail** for email automation

---

## 📦 Installation


## 🧪 How It Works (Step-by-Step)

1. Open [IREPS tender page](https://www.ireps.gov.in/epsn/anonymSearchPO.do)
2. Enter PO Value range (1 to 99999)
3. Select date = Yesterday (automatically picked by script)
4. Handle Captcha (based on first 6 characters of a URL param)
5. Submit form and scrape the result table
6. Format and save to Excel with:

   * Column widths
   * Filters
   * Wrapped text
   * Date formatting
7. Add a new column: `Direction/Location`
8. Email report to receiver using yagmail
9. Run daily using **Task Scheduler**

---

## 📧 Email Configuration

Update these lines in your code with your info:

```python
sender_email = "your_email@gmail.com"
app_password = "your_app_password"
receiver_emails = ["client@example.com"]
```

Or store these values in a `.env` file (optional).

---

## ⏰ Automate Daily Run (Task Scheduler Setup)

### 1. Open Task Scheduler

Search "Task Scheduler" in Windows Start and open it.

### 2. Click on "Create Task..."

Use the right-side panel to select `Create Task...`

### 3. General Tab:

* Name your task (e.g., `TenderFlowAutomation`)
* Check: `Run only when user is logged on`

### 4. Triggers Tab:

* Click **New\...**
* Set **Begin task:** `On a schedule`
* Choose Daily, Repeat every 1 day, Time: `08:00 AM`

### 5. Actions Tab:

* Click **New\...**
* Action: `Start a program`
* Program/script:

  ```
  C:\Path\To\Python\python.exe
  ```
* Add arguments:

  ```
  "C:\Path\To\Your\Script\python_code.py"
  ```

### 6. Save and you're done!

📌 *Make sure paths are correct and Python is added to your system path.*

📸 Screenshot example provided in `assets/task_scheduler.png` or below:

❗[Task Scheduler Setup](https://youtu.be/4n2fC97MNac?si=Ff31ofXqwum3CKce)

---

## ✅ Output Example (Excel file)

File is saved as:

```text
20_05_2025_results_696.xlsx
```

Columns:

* Sr
* Supplier Name
* PO Value
* Item Description
* PO Sl
* Consignee
* PO Qty/Unit
* T.U.R.
* Dely.Dt
* Direction/Location

---
**Libraries and other requirements**

gmail app password how to generate/create [click-here](https://youtu.be/MkLX85XU5rU?si=dfXn9D9QaoupY4kV)
selenium
beautifulsoup4
pandas
openpyxl
yagmail
python-dotenv
and some basic libraries like pandas, request, time etc...
**we also need chrome browser**

## 🙋‍♂️ About

This general-purpose tool is a real-world project to showcase:

* Data extraction
* Automation using Python
* Excel file formatting
* Scheduled task execution

Feel free to fork and customize it for your own use cases!

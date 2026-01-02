# CTCAC Tax Credit Application Parser

## 1. The Assignment

**Goal:** Automate the extraction of financing cost data from 165+ Excel applications for the **California Tax Credit Allocation Committee (CTCAC)**.

**Specific Requirements:**

1. **Download:** Programmatically fetch all Excel files from the CTCAC 2025 Third Round website.
2. **Extraction:** For each file, extract:
* **Construction Financing Costs:** (Interest, Fees, Bond Premiums, Title/Recording, etc.).
* **Permanent Financing Costs:** (Origination Fees, Credit Enhancement, etc.).
* **"Other" Costs:** Capture dynamic descriptions (e.g., "Deputy Inspection").
* **Project Metrics:** Total Units, Square Footage (SF), and Hard Costs.


3. **Benchmarking:** Calculate specific metrics:
* Financing Cost per Unit.
* Financing Cost per Square Foot.
* Financing Costs as a % of Hard Costs.


4. **Validation:** * Verify that the line items sum up to the Total listed in the sheet.
* Flag any files where data is missing or math does not reconcile.
* Handle files where data is located in different tabs or formats.



---

## 2. Methodology (The Solution)

This script uses a robust, multi-stage approach to ensure 100% data coverage.

### A. Intelligent Scanning

* **Smart Unit Search (The "Anti-9" Filter):** The script scans the application for "Total Units." Crucially, it mathematically ignores the question label (e.g., *"9. Total Units"*) and years (e.g., *"2025"*) to correctly identify the actual project size (e.g., *75*).
* **Global Search (for Square Footage):** Since SF data appears in inconsistent tabs (sometimes "Application," sometimes "Points System"), the script scans the *entire workbook* to find the maximum valid square footage value.
* **Section Parsing (for Costs):** The script locates the "Sources & Uses" budget, identifies the "Construction" and "Permanent" sections, and strictly pulls line items until it hits the "Total" line.

### B. Smart Fallbacks

* **Hard Costs (Bottom-Up Calculation):** The script first looks for the "Total New Construction Costs" line. If that line is missing or empty (0), the script **automatically recalculates the total** by summing the individual components (Site Work + Structures + Contractor Fees). This ensures no data is lost even if the summary line is blank.
* **"Other" Costs:** Instead of creating 100+ messy columns for every unique description, the script aggregates them into two clean columns:
* `Other Costs`: The total dollar amount.
* `Other Details`: A text string listing the descriptions (e.g., *"Legal Fees; Deputy Inspection"*).



### C. Analyst-Grade Validation

The script does not just return "Error." It distinguishes between:

* **Missing Data:** The applicant left the "Total" cell blank (Flag: `Sheet Total Missing`).
* **Math Errors:** The line items do not sum to the total provided (Flag: `Variance ($+5,000)`).
* **Empty Files:** The section is completely blank (Flag: `Costs Empty`).

---

## 3. Project Structure

```text
/project-folder
│
├── ctcac_parser.py        # The main Python script (Downloader + Processor)
├── summary_output.csv     # The final clean dataset (Generated after running)
├── README.md              # This documentation file
│
└── Downloaded files/      # Folder created automatically to store the .xlsx files
    ├── 25-825.xlsx
    ├── 25-826.xlsx
    └── ... (165+ files)

```

---

## 4. How to Run

### Prerequisites

You need Python installed. Install the required libraries using pip:

```bash
pip install requests beautifulsoup4 pandas openpyxl

```

### Execution

Run the script from your terminal:

```bash
python ctcac_parser.py

```

### Output

1. The script will check for the `Downloaded files` directory.
2. It will download any missing files from the CA Treasurer's website.
3. It will process every file, printing a progress bar (e.g., `[5/165] Processing...`).
4. It saves the final data to **`summary_output_final.csv`**.

---

## 5. Explanation in Basic Terms (ELI5)

Here is exactly what the robot (the code) is doing, step-by-step:

1. **The Shopping Trip (Download):** The code goes to the government website, finds every link that looks like an Excel file, and saves a copy to your computer.
2. **The Accountant (Opening Files):** It opens the Excel files one by one. It specifically looks at the "Values" (the result of the math), not the "Formulas."
3. **The Detective (Finding Data):** * It looks for "Total Units." If it sees "9. Total Units," it knows "9" is just the question number and keeps looking until it finds the real answer (like "45").
* It looks for "Square Footage." Since people hide this in different places, it looks through every page of the Excel file until it finds the biggest number that looks like a building size.
* It looks for the "Budget" page. It reads down the list of costs (Interest, Fees, Taxes) just like a human would, adding them up as it goes.


4. **The Auditor (Checking Math):** It adds up the costs it found and compares them to the "Total" number written at the bottom of the page.
* If they match? Great.
* If they don't match? It writes a "Flag" note in the final report saying, *"Hey, the math in this file is off by $50."*


5. **The Report (Output):** It saves all this information into a neat CSV table that is ready for analysis.

---

## 6. Output Columns Dictionary

The final CSV contains **30 columns** optimized for benchmarking:

* **Identifiers:** `File Name`, `Flag`
* **Benchmarks:** `Combined Financing Costs`, `Cost per Unit`, `Cost per SF`, `% of Hard Costs`
* **Project Stats:** `Total Units`, `Total SF`, `Hard Costs`
* **Construction Details:** `Const Loan Interest`, `Origination Fee`, `Taxes`, `Insurance`, `Other Costs` (Aggregated), `Other Details` (Text), `Const Total`.
* **Permanent Details:** `Perm Loan Origination`, `Taxes`, `Insurance`, `Other Costs`, `Other Details`, `Perm Total`.

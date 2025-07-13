# 🧪 Load Test to Form Submission & Reporting with Playwright

This Python script automates form submission across multiple URLs using Playwright, logs results and errors, and saves both detailed results and summary statistics into an Excel file with multiple sheets.

---

## 🛠️ Prerequisites

Ensure the following before running the script:

- ✅ Python **3.8 or newer**
- ✅ Internet connection (to install dependencies and load remote URLs)
- ✅ A folder named `URL` located in the same directory as the script
- ✅ At least one CSV file inside the `URL` folder containing a column named **`URL`** with the list of links to process

---

## 📦 Auto-Installed Dependencies

The script automatically installs these Python packages if missing:

| Library       | Purpose                              |
|---------------|------------------------------------|
| `pandas`      | Read/write CSV and Excel files     |
| `playwright`  | Headless browser automation        |
| `openpyxl`    | Writing Excel files with multiple sheets |

---

## 🚀 What the Script Does

1. Automatically finds and uses the **first CSV file (alphabetically)** from the `URL` folder as input.
2. Prompts the user to input the number of concurrent browser tabs (workers).
3. Loads the URLs from the selected CSV file.
4. For each URL:
   - Opens it in **Chromium** using **Playwright**.
   - Clicks form elements (radio buttons, text fields).
   - Submits the form.
   - Waits for a specific success message in Japanese.
   - Logs the result (`Success` or `Failed`), time taken, and any browser console errors.
5. After all tasks complete:
   - Saves detailed test results and a summary into a multi-sheet Excel `.xlsx` file.
   - Displays a summary of the test in the terminal.

---

## 📤 Output

The output Excel file is saved in the `Results/` directory with this naming pattern:
Results__tabs-<TAB_COUNT>__urls-<URL_COUNT>__YYYY-MM-DD__HH-MM-SS.xlsx


### 🔹 Sheet 1: `Test Results`

Contains one row per tested URL:

| url              | duration_sec | result  | message                  | console_errors         |
|------------------|--------------|---------|--------------------------|------------------------|
| https://example1  | 3.21         | Success | 安否確認へのご回答...     |                        |
| https://example2  | 0.00         | Failed  | TimeoutError: ...        | TypeError: Something   |

---

### 🔹 Sheet 2: `Summary`

Contains high-level statistics:

| Total URLs | Number of Tabs Used | Success Count | Failure Count | Max Duration (sec) | Min Duration (sec) | Average Duration (sec) | Total Execution Time (sec) |
|------------|---------------------|---------------|---------------|--------------------|--------------------|-----------------------|----------------------------|
| 100        | 50                  | 97            | 3             | 9.50               | 0.45               | 3.21                  | 185.34                     |

---

## ▶️ How to Run

1. Place your input CSV files inside a folder named `URL` located next to this script.
2. Run the script:

```bash
python your_script_name.py
```

3. Enter the number of concurrent tabs when prompted (e.g., 50).

4. The script will automatically pick the first CSV file alphabetically from the URL folder and begin processing.

5. Results will be saved inside the Results/ folder.


## 📁 Project Folder Structure

```text
project/
│
├── URL/
│   ├── input_file_1.csv
│   ├── input_file_2.csv
│   └── ...                   # Place one or more CSV files here (with 'URL' column)
│
├── your_script_name.py       # Main Python script to run the automation
│
└── Results/
    └── Results__tabs-50__urls-100__2025-07-10__14-31-22.xlsx
                              # Output Excel file containing both results and summary





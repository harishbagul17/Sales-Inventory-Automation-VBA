# 📊 Sales Inventory Automation in Excel (VBA)

This project is a fully automated Excel-based solution for generating, cleaning, analyzing, and reporting sales inventory data using **VBA macros**.

## ✅ Features

- Generates a random dataset of 100–200 sales entries
- Cleans and validates data (e.g., removes blanks, trims spaces, checks for valid values)
- Calculates profit for each sale
- Produces summary reports (total profit, profit by product)
- All steps are automated with reusable and modular VBA macros

---

## 📂 Project Structure

| File/Folder                | Description                                      |
|----------------------------|--------------------------------------------------|
| `SalesInventoryAutomation.xlsm` | Main Excel workbook with all macros included |
| `VBA_Code/`                | Exported `.bas` files for each module (see below) |

---

## 📄 VBA Modules

The code is organized into 4 modules:

1. `modGenerateData` – Creates random sales inventory data  
2. `modCleanData` – Cleans and validates data  
3. `modCalculateProfit` – Calculates profit for each sale  
4. `modSummaryReport` – Generates summary report by product and total

> ✅ Optional: A `RunAll` macro calls all steps in order for full automation.

---

## 🚀 How to Use

1. Download or clone this repo.
2. Open the `.xlsm` file in Excel.
3. Enable macros.
4. Press `Alt + F8`, run:
   - `GenerateSalesData`
   - `CleanSalesData`
   - `CalculateProfit`
   - `GenerateSummaryReport`
   - Or run `RunAll` to do all at once!

---

## 📌 Skills Demonstrated

- Excel VBA macro development
- Data cleaning and automation
- Summary reporting
- Project modularization
- GitHub version control

---

## 📷 Screenshots *(optional)*

> You can upload screenshots or GIFs here to show your workbook in action.

---

## 📜 License

This project is licensed under the [MIT License](LICENSE).

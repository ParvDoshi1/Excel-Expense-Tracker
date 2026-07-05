# 💰 Expense Tracker (Excel Mini Project)

A simple, single-sheet expense tracker built in Microsoft Excel to log daily transactions and automatically calculate key spending metrics using built-in formulas.

## 📋 Overview

This project logs day-to-day expenses under categories like **Food, Office, Fun, and Investment**, and uses Excel formulas to break down spending patterns — no macros, no VBA, just clean formula-driven logic.

## ✨ Features

- **Transaction Log** — Serial number, description, category (type), and amount for each entry
- **Category-wise Food Tracking** — Uses `COUNTIF` to isolate and calculate food-related expenses
- **Automated Summary Panel**
  - **Total for the Day** — Sum of all transaction amounts
  - **Average Transaction Value** — Mean spend per transaction
  - **Food Expense** — Total amount spent specifically on food

## 🧮 Formulas Used

| Metric | Formula |
|---|---|
| Food amount per row | `=COUNTIF(D3,"Food")*E3` |
| Total for the Day | `=SUM(E3:E12)` |
| Average Transaction Value | `=AVERAGE(E3:E12)` |
| Food Expense | `=SUM(F3:F12)` |

## 🗂️ Sheet Structure

| Column | Field |
|---|---|
| B | S.No |
| C | Description |
| D | Type (Food / Office / Fun / Investment) |
| E | Amount |
| F | Food amount (helper column) |
| I–J | Summary metrics |

## 🚀 How to Use

1. Download `Expense_Tracker.xlsx` from this repository
2. Open it in Excel (or Google Sheets)
3. Add new transactions in the next empty row, filling in the **Description**, **Type**, and **Amount**
4. Drag the formula in column F down to match the new row
5. Update the summary formula ranges (`E3:E12`, `F3:F12`) if you add rows beyond row 12

## 🔧 Tech Used

- Microsoft Excel
- Formulas: `SUM`, `AVERAGE`, `COUNTIF`

## 📈 Possible Future Improvements

- Extend category tracking beyond "Food" to all categories (Office, Fun, Investment) using the same `COUNTIF` logic
- Add a pivot table or chart to visualize spending by category
- Add monthly/weekly grouping with date column
- Automate the row-range in summary formulas using dynamic ranges (e.g., `SUM(E3:E1000)` or Excel Tables)

## 👤 Author

Parv


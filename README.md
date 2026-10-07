<img width="1536" height="1024" alt="Excel Analyzer Dashboard Infographic" src="https://github.com/user-attachments/assets/967b05fa-9afd-44db-b0c8-669af7561264" />


# 📊 Excel Analyzer Project

## 📌 Project Overview

**Excel Analyzer** is a practical Excel-based data analysis project designed to analyze, organize, and visualize business data using Microsoft Excel.

This project demonstrates important Excel skills such as **data cleaning, formulas, conditional formatting, data analysis, KPI calculation, sorting/filtering, charts, and dashboard creation**.

The main goal of this project is to transform raw data into meaningful information that can help users understand business performance and make better decisions.

---

## 🎯 Project Objectives

* Analyze raw Excel datasets
* Clean and organize data
* Apply Excel formulas and functions
* Use Conditional Formatting
* Calculate important KPIs
* Perform data analysis
* Create charts and visualizations
* Identify trends and patterns
* Build an interactive Excel analysis report
* Improve practical Excel skills

---

## 🛠️ Tools & Technologies

| Tool                   | Purpose                        |
| ---------------------- | ------------------------------ |
| Microsoft Excel        | Data analysis and calculations |
| Excel Formulas         | Automated calculations         |
| Conditional Formatting | Highlight important data       |
| Charts                 | Data visualization             |
| Pivot Tables           | Data summarization             |
| KPI Analysis           | Performance measurement        |
| Filters & Sorting      | Data exploration               |

---

## 📂 Project Structure

```text
Excel-Analyzer/
│
├── README.md
├── Excel_Analyzer.xlsx
├── Dataset/
│   └── dataset.xlsx
│
├── Screenshots/
│   ├── dataset.png
│   ├── analysis.png
│   └── dashboard.png
│
└── Documentation/
    └── Project_Explanation.pdf
```

---

## 📊 Project Workflow

```text
Raw Dataset
     ↓
Data Cleaning
     ↓
Data Formatting
     ↓
Excel Formulas
     ↓
Conditional Formatting
     ↓
Data Analysis
     ↓
KPI Calculation
     ↓
Charts & Visualization
     ↓
Final Dashboard
```

---

# 📝 Tasks Performed

## 🔹 Task 1 — Conditional Formatting

Conditional Formatting is used to automatically highlight cells based on specific conditions.

### Examples

* Highlight high sales
* Highlight low sales
* Identify duplicate values
* Highlight specific categories
* Use color scales for better visualization

### Example

```text
Sales > 50000 → Highlight
Sales < 10000 → Highlight
```

This makes important information easier to identify.

---

## 🔹 Task 2 — Text/Data Analysis

Excel text functions can be used to clean and analyze text-based information.

Common functions include:

```excel
=LEFT()
=RIGHT()
=MID()
=LEN()
=TRIM()
=UPPER()
=LOWER()
=PROPER()
```

### Example

```excel
=PROPER(A2)
```

This converts text into proper capitalization.

---

## 🔹 Task 3 — Data Analysis Using Formulas

Excel formulas are used to calculate useful information from the dataset.

Common functions used:

```excel
=SUM()
=AVERAGE()
=COUNT()
=COUNTA()
=MAX()
=MIN()
```

### Example

```excel
=SUM(B2:B20)
```

This calculates the total of the selected values.

---

## 🔹 Task 4 — Sorting

Sorting is used to arrange data in a specific order.

Examples:

* A → Z
* Z → A
* Smallest → Largest
* Largest → Smallest

Sorting helps analyze the highest and lowest values quickly.

---

## 🔹 Task 5 — Filtering

Filters are used to display only the data required for analysis.

For example:

```text
Region = West
Category = Technology
Sales > 10000
```

Filtering makes large datasets easier to analyze.

---

## 🔹 Task 6 — SUMIFS Analysis

`SUMIFS` is used to calculate a total based on multiple conditions.

### Syntax

```excel
=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2)
```

### Example

```excel
=SUMIFS(Sales, Region, "West", Product, "Laptop")
```

This calculates total sales for **Laptop products in the West region**.

---

## 🔹 Task 7 — Data Validation

Data Validation controls what users can enter into a cell.

Examples:

* Drop-down lists
* Number restrictions
* Date restrictions
* Text length restrictions

### Example

A Region cell can contain:

```text
East
West
North
South
```

This reduces incorrect data entry.

---

## 🔹 Task 8 — Data Analysis

The dataset is analyzed to identify:

* Total Sales
* Average Sales
* Maximum Sales
* Minimum Sales
* Number of Orders
* Product performance
* Regional performance

Example:

```excel
=AVERAGE(B2:B100)
```

---

## 🔹 Task 9 — Charts & Visualization

Charts convert numerical data into visual information.

Charts that can be used:

* 📊 Column Chart
* 📈 Line Chart
* 🥧 Pie Chart
* 📊 Bar Chart
* 🔵 Scatter Chart

Charts make trends and comparisons easier to understand.

---

# 🔹 Task 10 — KPI Analysis

## What is KPI?

**KPI = Key Performance Indicator**

A KPI is a measurable value used to understand how well a business or process is performing.

### Example KPIs

| KPI           | Meaning                 |
| ------------- | ----------------------- |
| Total Sales   | Overall sales amount    |
| Total Orders  | Number of orders        |
| Average Sales | Average sales per order |
| Maximum Sales | Highest sales           |
| Minimum Sales | Lowest sales            |
| Profit        | Overall profit          |

### Example formulas

**Total Sales**

```excel
=SUM(Sales_Range)
```

**Average Sales**

```excel
=AVERAGE(Sales_Range)
```

**Total Orders**

```excel
=COUNTA(Order_ID_Range)
```

**Maximum Sales**

```excel
=MAX(Sales_Range)
```

**Minimum Sales**

```excel
=MIN(Sales_Range)
```

---

# 📈 Dashboard

The final Excel Analyzer can contain a dashboard showing important information in one place.

### Dashboard Components

```text
┌─────────────────────────────────────────┐
│             EXCEL ANALYZER              │
├─────────────┬─────────────┬─────────────┤
│ Total Sales │Total Orders │ Avg. Sales  │
├─────────────┴─────────────┴─────────────┤
│                                         │
│          Sales Performance Chart        │
│                                         │
├──────────────────────┬──────────────────┤
│ Regional Analysis    │ Product Analysis │
│                      │                  │
└──────────────────────┴──────────────────┘
```

The dashboard provides a quick overview of the dataset.

---

# 📊 Key Excel Functions Used

| Function     | Purpose                         |
| ------------ | ------------------------------- |
| `SUM()`      | Calculate total                 |
| `AVERAGE()`  | Calculate average               |
| `COUNT()`    | Count numbers                   |
| `COUNTA()`   | Count non-empty cells           |
| `MAX()`      | Find maximum                    |
| `MIN()`      | Find minimum                    |
| `IF()`       | Apply conditions                |
| `SUMIF()`    | Sum using one condition         |
| `SUMIFS()`   | Sum using multiple conditions   |
| `COUNTIF()`  | Count using one condition       |
| `COUNTIFS()` | Count using multiple conditions |
| `LEFT()`     | Extract characters from left    |
| `RIGHT()`    | Extract characters from right   |
| `MID()`      | Extract characters from middle  |
| `LEN()`      | Count characters                |
| `TRIM()`     | Remove extra spaces             |
| `PROPER()`   | Format text                     |
| `UPPER()`    | Convert text to uppercase       |
| `LOWER()`    | Convert text to lowercase       |

---

# 📚 Skills Learned

Through this project, I practiced:

* ✅ Excel Fundamentals
* ✅ Data Cleaning
* ✅ Data Formatting
* ✅ Conditional Formatting
* ✅ Excel Formulas
* ✅ Text Functions
* ✅ Logical Functions
* ✅ SUMIFS
* ✅ COUNTIFS
* ✅ Sorting
* ✅ Filtering
* ✅ Data Validation
* ✅ KPI Analysis
* ✅ Charts
* ✅ Data Visualization
* ✅ Dashboard Creation
* ✅ Business Data Analysis

---

# 🚀 How to Use

1. Download or clone this repository.
2. Open `Excel_Analyzer.xlsx`.
3. Open the **Dataset** sheet.
4. Review the raw data.
5. Follow the analysis sheets.
6. Check the formulas and results.
7. Open the Dashboard sheet to view the final analysis.

---

# 💡 Project Outcome

The **Excel Analyzer Project** demonstrates how raw data can be converted into useful business insights using Microsoft Excel.

The project helps understand:

```text
Raw Data
   ↓
Clean Data
   ↓
Analyze Data
   ↓
Calculate KPIs
   ↓
Visualize Data
   ↓
Make Decisions
```

---

# 👩‍💻 Author

**Excel Analyzer Project**

Created as a practical Excel data analysis project for learning and demonstrating Excel skills.

---

## ⭐ Future Improvements

Future versions of this project may include:

* Interactive dashboards
* Pivot Tables
* Pivot Charts
* Slicers
* Advanced Excel formulas
* Power Query
* Power Pivot
* Automated reports
* More advanced data analysis

---

## 📌 Conclusion

**Excel Analyzer** is a complete practical project for learning how to use Excel for data analysis and visualization.

It combines Excel formulas, formatting, KPIs, charts, and dashboards to turn raw data into meaningful insights.

⭐ **If you find this project useful, consider giving the repository a star!**

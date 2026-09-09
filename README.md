# 📊 Sales Executive Performance Dashboard – Excel

## 📌 Project Overview

The **Sales Executive Performance Dashboard** is an interactive Excel-based reporting solution developed to analyze and monitor the performance of sales executives against their assigned targets.

The project uses **Microsoft Excel, formulas, charts, and VBA macros** to transform raw sales data into a visual performance report. The dashboard helps identify high-performing executives, executives who are below target, and the gap between actual sales and expected performance.

A key feature of this project is the **Macro Report Connection**, which connects the underlying raw data with the dashboard reporting process and helps refresh/update the performance report efficiently.

## 🎯 Project Objective

The main objective of this dashboard is to provide a simple and effective way to:

* Monitor sales executive performance
* Compare actual sales with assigned targets
* Identify executives who achieve or fall short of their targets
* Analyze daily sales performance
* Measure Target Hit %
* Measure the percentage away from the target
* Provide a visual summary for faster management decision-making
* Reduce repetitive manual reporting through Excel automation and VBA macros

## 🗂️ Workbook Structure

The Excel workbook contains two primary worksheets:

### 1. Raw Data

The **Raw Data** sheet contains the underlying sales information for each sales executive.

Key fields include:

* Employee Code
* Sales Executive
* Region
* Day 1 Sales
* Day 2 Sales
* Day 3 Sales
* Day 4 Sales
* Day 5 Sales
* Total Sales
* Target
* Target Hit %
* Away From Target %

The raw dataset contains sales information for multiple executives across different regions.

### 2. Dashboard

The **DASHBOARD** sheet converts the raw sales information into a visual performance report.

The dashboard focuses on metrics such as:

* Total Sales
* Target Hit %
* Away From Target %
* Sales Executive-wise performance
* Comparison of executives based on sales performance

Charts are used to make the comparison easier and allow management to quickly identify performance trends and gaps.

## ⚙️ Key Calculations

### Total Sales

Total sales are calculated by aggregating the sales generated across the five reporting days.

**Total Sales = Day 1 + Day 2 + Day 3 + Day 4 + Day 5**

### Target Hit %

Target achievement is measured by comparing actual sales against the assigned target.

**Target Hit % = Total Sales / Target**

### Away From Target %

The dashboard also calculates how far an executive is from achieving the assigned target.

**Away From Target % = 100% − Target Hit %**

These metrics allow performance to be evaluated from both achievement and gap perspectives.

## 🔄 Macro Report Connection

The project includes a **VBA macro-based report connection** to support the dashboard reporting workflow.

The macro functionality helps connect the underlying reporting data with the performance dashboard and reduces the need for repetitive manual updates.

The `.xlsm` format is used because the workbook contains embedded VBA functionality.

> **Note:** Macros must be enabled in Microsoft Excel for the automation functionality to operate correctly.


## 📊 Dashboard Insights

The dashboard is designed to answer important business questions such as:

* Which sales executives generated the highest sales?
* Which executives achieved their assigned targets?
* Which executives are performing below expectations?
* What percentage of the target has been achieved?
* How much does each executive remain away from their target?
* How does performance vary across sales executives and regions?
* Which executives may require additional attention or performance improvement?

## 🛠️ Tools & Technologies

* **Microsoft Excel**
* **Excel Formulas**
* **VBA / Macros**
* **Excel Charts**
* **Data Analysis**
* **Dashboard Design**

## 🔄 Project Workflow

The overall workflow of the project is:

**Raw Sales Data**
↓
**Data Preparation & Calculations**
↓
**Total Sales & Target Metrics**
↓
**Macro Report Connection**
↓
**Performance Dashboard**
↓
**Sales Performance Analysis & Insights**

## 💡 Business Value

This dashboard provides a centralized view of sales executive performance and makes it easier for managers to identify performance gaps.

Instead of manually reviewing individual sales records, users can use the dashboard to quickly compare executives based on actual sales, target achievement, and target gaps.

The macro-enabled reporting workflow also helps make recurring performance reporting more efficient and less dependent on repetitive manual work.

## 📁 Project File

**Sales Executive Analysis.xlsm**

The `.xlsm` file contains:

* Raw sales dataset
* Calculated performance metrics
* Dashboard
* Charts
* VBA macro functionality


## 🚀 How to Use

1. Download the `Sales Executive Analysis.xlsm` file.
2. Open the workbook using **Microsoft Excel**.
3. Enable macros if prompted.
4. Review the **Raw Data** sheet to understand the underlying dataset.
5. Navigate to the **DASHBOARD** sheet.
6. Use the dashboard to analyze sales executive performance.
7. Update the source data when new performance records are available and run the macro/report workflow as required.


## 📌 Skills Demonstrated

This project demonstrates practical experience in:

* Excel Dashboard Development
* Data Cleaning & Preparation
* Excel Data Analysis
* Formula-Based Calculations
* KPI Development
* Sales Performance Analysis
* Target vs Actual Analysis
* Data Visualization
* Excel Chart Development
* VBA Macro Automation
* Business Reporting
* Management Dashboard Design

## 📈 Conclusion

The **Sales Executive Performance Dashboard** demonstrates how Excel can be used not only for basic data analysis but also as a complete reporting and visualization solution.

By combining raw sales data, calculated KPIs, visual dashboards, and VBA-based report automation, the project provides a structured approach to monitoring sales performance and identifying areas that require management attention.

## 👤 Author

**Soumyaranjan**

Data Analyst | Excel | SQL | Power BI | Python

---

⭐ If you find this project useful, feel free to explore the repository and connect with me for feedback or collaboration.

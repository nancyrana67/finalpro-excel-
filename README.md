# 📊 Customer Transactions Analysis & Executive Dashboard

An end-to-end Excel data analysis solution analyzing 250 customer transactions to uncover purchasing trends, regional revenue drivers, customer segmentation insights, predictive models, and dynamic executive dashboards.

---

## 📌 Project Objective

The goal of this project is to create a complete analysis and reporting solution using Excel to evaluate customer transactions, generate actionable business insights, perform predictive regression modeling, and build interactive dashboards using Pivot Tables, Charts, and Slicers.

---

## 📁 Workbook Architecture

The Excel workbook (`finalpro-excel.xlsx`) is structured into five core worksheets:

| Sheet Name | Description & Purpose |
| :--- | :--- |
| **`Sheet1` (Raw Data)** | Contains 250 transaction logs with details on Customer ID, Products, Categories, Regional data, Order Quantities, Pricing, Customer Segments, and calculated Helper Columns (Tenure, EOM Due Dates, Product Codes, VIP Match). |
| **`Analysis`** | Contains summary metrics, formula-driven KPI calculations (`INDEX-MATCH`, `MODE`, `SUMIF`, `COUNTIF`), and high-value customer segmentation criteria. |
| **`Visualizations`** | Houses advanced analytical models including **Linear Regression** (Data Analysis Toolpak) and **What-If Analysis** (Scenario Manager & Goal Seek). |
| **`dashbord`** | Clean presentation sheet featuring Pivot Tables summarizing regional revenue and category product quantities. |
| **`Scenario Summary`** | Auto-generated What-If scenario summary report comparing Base, Best Case, and Worst Case business outcomes. |

---

## 🛠️ Key Topics & Functions Implemented

### 1. Date & Time Functions
* **`TODAY()` & `NOW()`**: Calculated real-time report update timestamps.
* **`DATEDIF()`**: Computed precise customer tenure in years based on `Customer_Since` and `Date`.
* **`EOMONTH()`**: Derived payment due dates based on net-30 business billing terms.

### 2. Lookup & Advanced Dynamic Formulas
* **`FILTER()`**: Extracted multi-value returns for regional transactions (e.g., filtering East region transactions).
* **Mode & Matching (`INDEX` + `MODE` + `MATCH`)**: Derived the most frequently purchased product name programmatically.
* **List Comparison (`IF` + `ISNUMBER` + `MATCH`)**: Identified high-value matching customers against target VIP lists.
* **Text Manipulation (`UPPER` + `LEFT`)**: Created standardized 3-letter uppercase product abbreviations (`Product_Code`).
* **Conditional Logic (`SUMIF` + `IF`)**: Summarized spent amounts per customer to tag **High-Value Customers** (spending > $5,000).

### 3. Business Analysis & Predictive Modeling
* **Linear Regression (Data Analysis Toolpak)**: Analyzed the impact of purchase quantity on total transaction revenue ($R^2 = 0.21$).
* **What-If Analysis (Scenario Manager & Goal Seek)**: Modeled best-case (+15% growth) and worst-case (-10% contraction) sales projections.

### 4. Interactive Dashboard & Visualization
* **Pivot Tables & Charts**: Column charts for Regional Revenue and Pie Charts for Category Product Volumes.
* **KPI Indicators**: Styled metrics highlighting key performance drivers.
* **Slicers & Timelines**: Synchronized interactive filters allowing dynamic data filtering across payment methods, regions, and dates.

---

## 📈 Key Summary Metrics & Findings

* **Total Business Revenue:** **$229,192.47**
* **Total Order Volume:** **250 Transactions** (753 total units sold)
* **Average Order Value (AOV):** **$916.77**
* **Top Category by Revenue:** **Electronics** ($171,756.05 | ~74.9% of total revenue)
* **Top Product by Unit Sales:** **Bookshelf** (102 units sold)
* **Top Performing Region:** **East** ($59,288.39 in total sales)
* **Top Individual Spender:** **Mark Carter** ($15,659.65 total spent)

---

## 💡 Strategic Business Recommendations

1. **VIP Loyalty Program:** Launch tailored rewards and account management for top spenders like Mark Carter and Edward Mitchell to boost customer retention.
2. **Regional Expansion:** Allocate higher marketing budgets toward the **East** and **North** regions, which consistently deliver the highest transaction revenue.
3. **Cross-Selling Opportunities:** Bundle top-selling furniture items (Bookshelves) with high-ticket electronics (Monitors, Keyboards, Laptops) to maximize basket size.

---

## ⚙️ How to Run / Inspect the Workbook

1. Open `finalpro-excel.xlsx` in Microsoft Excel (2016 or newer recommended).
2. Ensure macros/iterative calculations are enabled if viewing dynamic timestamp calculations (**File > Options > Formulas > Enable iterative calculation**).
3. Ensure the **Data Analysis Toolpak** add-in is enabled (**File > Options > Add-ins**) to re-run or edit regression models.
4. Interact with the Dashboard using the slicers on the `dashbord` tab.

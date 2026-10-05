# 🛒 E-Commerce Sales Analysis Using Microsoft Excel

## 📌 Project Overview

This project is an **E-Commerce Sales Data Analysis project developed using Microsoft Excel** as part of my learning journey as a **Data Analyst**.

The main objective of this project is to transform raw e-commerce data into a clean, structured, and analysis-ready dataset and then use Excel's data analysis and visualization features to generate meaningful business insights.

The project covers the complete data analysis workflow, including:

* Data import and organization
* Data cleaning and preprocessing
* Missing-value handling
* Duplicate checking
* Data standardization
* Excel formulas
* XLOOKUP
* Pivot Tables
* Descriptive statistics
* Data visualization
* PivotCharts
* Interactive dashboard
* Business insights and recommendations

The project documentation states that the objective was to improve the accuracy, consistency, and usability of the raw data through cleaning, transformation, formatting, and validation.

---

## 🎯 Project Objectives

The key objectives of this project are:

1. Import and organize the e-commerce dataset in Excel.
2. Convert raw data into structured Excel tables.
3. Identify and handle missing values.
4. Find and remove duplicate records.
5. Correct inconsistent text formatting and data-entry errors.
6. Remove irrelevant columns where applicable.
7. Apply Excel formulas for data transformation and calculations.
8. Use **XLOOKUP/VLOOKUP** to combine related tables.
9. Create Pivot Tables for business analysis.
10. Calculate descriptive statistics such as:

* Mean
* Median
* Mode
* Standard deviation

11. Create different types of charts.
12. Build an interactive Excel dashboard.
13. Generate business insights from the cleaned data.
14. Document the complete data-cleaning and analysis process.

These activities are aligned with the assignment requirements and the project documentation.

---

## 📂 Dataset Description

The project uses an e-commerce sales dataset containing information about customers, products, stores, and sales transactions.

### Customer Information

| Column        | Description                    |
| ------------- | ------------------------------ |
| Customer ID   | Unique customer identification |
| Name          | Customer name                  |
| Age           | Customer age                   |
| Gender        | Customer gender                |
| City          | Customer city                  |
| State         | Customer state                 |
| Country       | Customer country               |
| Loyalty Level | Customer loyalty level         |

### Product Information

| Column       | Description            |
| ------------ | ---------------------- |
| Product ID   | Product identification |
| Product Name | Product name           |
| Category     | Product category       |
| Brand        | Product brand          |
| Cost         | Product cost           |
| Stock        | Available stock        |

### Store Information

| Column      | Description          |
| ----------- | -------------------- |
| Store ID    | Store identification |
| Store Name  | Store name           |
| Region Name | Store region         |
| Store Type  | Type of store        |

### Sales Information

| Column       | Description                      |
| ------------ | -------------------------------- |
| Sales ID     | Sales transaction identification |
| Order Date   | Date of sales order              |
| Quantity     | Quantity purchased               |
| Unit Price   | Price per unit                   |
| Discount     | Discount applied                 |
| Payment Type | Payment method                   |
| Total Amount | Total transaction amount         |

The project documentation provides these field descriptions across the customer, product, store, and sales sections.

---

# 🧹 Data Cleaning & Preprocessing

Data cleaning was an important part of this project because the raw dataset contained missing values and inconsistent formatting.

The following cleaning activities were performed.

### 1. Text Cleaning

Excel functions such as:

```excel
CLEAN()
TRIM()
PROPER()
```

were used to standardize text fields.

For example:

```excel
PROPER(TRIM(CLEAN(C2)))
```

was used to improve the formatting of store names.

### 2. Customer Name Cleaning

The `Name` column was cleaned using:

```excel
CLEAN()
TRIM()
PROPER()
```

to remove unwanted spaces and standardize capitalization.

### 3. Country Standardization

Inconsistent country values were standardized.

For example:

```text
United States of America → USA
```

was handled using Excel's **Find and Replace** functionality.

### 4. Missing Loyalty Level

Missing values in the Loyalty Level column were replaced with:

```excel
=IF(ISBLANK([@[Loyalty_Level]]),"Unknown",[@[Loyalty_Level]])
```

This ensured that blank loyalty values were represented consistently.

### 5. Product ID Standardization

The Product ID values were checked and modified using Excel's Find and Replace functionality where required.

### 6. Duplicate Check

The dataset was checked for duplicate records.

**Result: No duplicate values were found.**

### 7. Brand Standardization

The Brand column contained inconsistent formatting.

The following functions were used:

```excel
CLEAN()
TRIM()
PROPER()
```

to standardize the values.

### 8. Missing Stock Values

Missing stock values were handled using an average-based replacement approach.

Example:

```excel
=IF(ISBLANK(G6),AVERAGE(G6:G105),G6)
```

### 9. Missing Cost Values

Missing cost values were replaced using the average cost:

```excel
=IF(ISBLANK(F6),AVERAGE(F6:F105),F6)
```

### 10. Store Name Cleaning

Store names were standardized using:

```excel
=PROPER(TRIM(CLEAN(C2)))
```

### 11. Order Date Formatting

The Order Date column was converted into the appropriate date format.

### 12. Missing Quantity

Missing quantity values were handled using an average-based approach:

```excel
=IF(ISBLANK(F2),AVERAGE(F2:F2001),F2)
```

### 13. Missing Unit Price

Missing Unit Price values were handled using:

```excel
=IF(ISBLANK(G2),AVERAGE(G2:G2001),G2)
```

### 14. Missing Discount

Blank discount values were replaced with zero:

```excel
=IF(ISBLANK(H2),"0",H2)
```

### 15. Total Amount Calculation

The Total Amount column was recalculated using:

```excel
=(G2-H2)*F2
```

### 16. Irrelevant Columns

Columns that were considered irrelevant to the analysis were removed.

These cleaning and transformation steps are documented in the project report.

---

# 🔎 Excel Formulas Used

Some of the important Excel functions used in this project include:

| Function       | Purpose                         |
| -------------- | ------------------------------- |
| `IF()`         | Handle missing values           |
| `ISBLANK()`    | Identify blank cells            |
| `AVERAGE()`    | Calculate average values        |
| `CLEAN()`      | Remove unwanted characters      |
| `TRIM()`       | Remove unnecessary spaces       |
| `PROPER()`     | Standardize capitalization      |
| `SUM()`        | Calculate totals                |
| `XLOOKUP()`    | Combine related tables          |
| Find & Replace | Standardize inconsistent values |

---

# 🔗 XLOOKUP

XLOOKUP was used to combine information from related tables.

Example used in the project:

```excel
=XLOOKUP(
    Table4[@[Customer_ID]],
    Table1[Customer_ID],
    Table1[Name],
    "Not Found",
    0
)
```

This formula searches for the Customer ID in another table and returns the corresponding customer name.

This helped demonstrate how different tables can be connected using a common key.

---

# 📊 Pivot Table Analysis

Pivot Tables were used to summarize the cleaned dataset and identify important business patterns.

The analysis included:

* Total sales
* Average price
* Total cost
* Category performance
* Gender-wise sales
* Age-wise sales
* Product performance

The project documentation specifically notes the use of Pivot Tables for analysing total sales, average price, total cost, and demographic sales patterns.

---

# 📈 Descriptive Statistics

Descriptive statistical analysis was performed using Excel formulas and the **Analysis ToolPak**.

The analysis included:

* Mean
* Median
* Mode
* Standard deviation
* Other descriptive measures

The Analysis ToolPak was used for Unit Price in the Sales Fact data and Cost in the Product Dimension data.

---

# 📊 Data Visualization

The project includes multiple visualization techniques to communicate the results clearly.

### Charts

At least three different chart types were created as part of the assignment requirements.

Examples include:

* Column/Bar Chart
* Line Chart
* Pie/Doughnut Chart

### PivotCharts

PivotCharts were used to visually represent summarized Pivot Table information.

### Interactive Dashboard

An Excel dashboard was created to provide a summarized view of the key business metrics and findings.

---

# 💡 Key Business Insights

The analysis generated several important findings.

### 🥇 Category Performance

**Electronics and Sports** were identified as the highest revenue-generating/contributing categories.

This indicates that these categories are strong contributors to overall sales performance.

### 👥 Customer Demographics

The analysis showed that customers between **25 and 35 years** were significant purchasers of electronic products across both male and female customer groups.

### 💰 Price Analysis

The predictive analysis indicated a price range of approximately **₹200–₹300** for the upcoming period based on the analysis performed in the project.

### 🛍️ Marketing Opportunity

Beauty products were identified as an area where additional promotional activities could help improve performance.

The project recommends using additional marketing and discount strategies for lower-performing products.

---

# 🎯 Business Recommendations

Based on the analysis, the following recommendations were identified:

1. Continue focusing on strong-performing categories such as **Electronics and Sports**.
2. Develop targeted marketing campaigns for customers aged **25–35**.
3. Consider promotional strategies for lower-performing products.
4. Provide additional discounts where appropriate to increase product demand.
5. Increase promotional activities for **Beauty products**.
6. Consider additional business-related promotional activities in **Tier-2 cities**.
7. Monitor pricing trends and customer purchasing behaviour regularly.

The final project conclusion similarly recommends additional discounts for non-performing products and promotional activities in Tier-2 cities.

# 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Tables
* Excel Formulas
* XLOOKUP
* Pivot Tables
* PivotCharts
* Analysis ToolPak
* Conditional Formatting
* Data Cleaning & Transformation
* Data Visualization
* Dashboard Design

---

# 📚 Skills Demonstrated

Through this project, I practiced the following Data Analyst skills:

### Data Preparation

* Data cleaning
* Data validation
* Missing-value handling
* Duplicate checking
* Data standardization
* Data transformation

### Excel

* IF
* ISBLANK
* AVERAGE
* SUM
* CLEAN
* TRIM
* PROPER
* XLOOKUP
* Excel Tables
* Pivot Tables
* PivotCharts
* Analysis ToolPak

### Data Analysis

* Descriptive statistics
* Customer analysis
* Product analysis
* Category analysis
* Sales analysis
* Demographic analysis

### Data Visualization

* Charts
* PivotCharts
* Interactive dashboard
* Business KPI presentation

### Business Analytics

* Descriptive analysis
* Diagnostic analysis
* Predictive analysis
* Prescriptive recommendations

--


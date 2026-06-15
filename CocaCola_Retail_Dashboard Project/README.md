# 🥤 Coca-Cola Retail Sales Dashboard



> ![Dashboard Overview](screenshots/dashboard_overview.png)




## 📌 Project Overview

The Coca-Cola Retail Sales Dashboard is an interactive Business Intelligence solution developed in Microsoft Excel using VBA, Pivot Tables, Pivot Charts, and Slicers. The dashboard enables users to analyze retail sales performance, monitor key business metrics, and gain actionable insights through dynamic visualizations and automated reporting.

The project demonstrates expertise in Excel-based analytics, dashboard design, automation, and business reporting.

---

## 🎯 Business Objective

Retail businesses generate large volumes of transactional data across multiple products, retailers, and regions. Analyzing this data manually can be time-consuming and prone to errors.

This dashboard was developed to:

* Automate sales reporting processes
* Provide real-time business insights
* Improve decision-making through data visualization
* Enable interactive analysis using filters and slicers
* Reduce manual effort in report generation

---

## 📊 Dashboard Features

### Executive KPI Dashboard

* Interactive KPI Cards
* Sales Performance Monitoring
* Trend Analysis
* Comparative Analysis

### Interactive Filtering

Users can filter reports dynamically using:

* Beverage Brand
* Retailer
* Invoice Month

### Automated Data Refresh

One-click refresh updates:

* Pivot Tables
* Charts
* KPI Metrics
* Dashboard Reports

### Enhanced User Experience

VBA automation is used to:

* Hide Excel Ribbon
* Hide Gridlines
* Hide Worksheet Tabs
* Open directly on the Homepage
* Provide a dashboard-like interface

---

# 🏗️ Solution Architecture

```text
                    +-------------------+
                    |   Source Data     |
                    | Retail Sales Data |
                    +---------+---------+
                              |
                              v
                    +-------------------+
                    |   Pivot Tables    |
                    | Data Aggregation  |
                    +---------+---------+
                              |
                              v
                    +-------------------+
                    | Pivot Charts      |
                    | KPI Calculations  |
                    +---------+---------+
                              |
                              v
                    +-------------------+
                    | Dashboard Layer   |
                    | Visual Analytics  |
                    +---------+---------+
                              |
                              v
                    +-------------------+
                    | VBA Automation    |
                    | Refresh & Control |
                    +-------------------+
```

---

## 📂 Workbook Structure

| Sheet Name  | Description                             |
| ----------- | --------------------------------------- |
| Homepage    | Landing page and dashboard navigation   |
| Dashboard   | Interactive KPI and visualization layer |
| Analysis    | Detailed sales analysis and trends      |
| Pivot Table | Backend aggregation engine              |
| Source Data | Raw retail sales dataset                |

---

## ⚙️ VBA & Macros Used

### Workbook_Open()

Purpose:

* Opens workbook on Homepage
* Displays welcome message
* Configures dashboard mode
* Customizes Excel interface

Example:

```vb
Private Sub Workbook_Open()

Worksheets("Homepage").Activate

MsgBox "Welcome to the Coca Cola Retail Dashboard"

End Sub
```

---

### Workbook_Deactivate()

Purpose:

* Restores Excel settings
* Re-enables Ribbon and standard interface options

---

### Refresh_button_click()

Purpose:

* Refresh all Pivot Tables
* Update Dashboard Visuals
* Clear active slicer filters

Example:

```vb
Sub Refresh_button_click()

ActiveWorkbook.RefreshAll

ActiveWorkbook.SlicerCaches("Slicer_Beverage_Brand").ClearManualFilter
ActiveWorkbook.SlicerCaches("Slicer_Retailer").ClearManualFilter
ActiveWorkbook.SlicerCaches("Slicer_Months__Invoice_Date").ClearManualFilter

End Sub
```

---

## 📈 Key Insights Generated

The dashboard helps stakeholders:

* Monitor total sales performance
* Analyze beverage brand trends
* Compare retailer performance
* Track monthly sales patterns
* Identify top-performing segments
* Support business decision-making

---

## 🛠️ Technology Stack

| Technology      | Purpose                             |
| --------------- | ----------------------------------- |
| Microsoft Excel | Dashboard Development               |
| VBA             | Automation & User Interface Control |
| Pivot Tables    | Data Aggregation                    |
| Pivot Charts    | Visualization                       |
| Slicers         | Interactive Filtering               |
| Excel Macros    | Report Automation                   |

---

# 📷 Dashboard Screenshots

## Homepage

> **Insert Screenshot Here**
>
> File Path Example:
>
> ```md
> ![Homepage](screenshots/homepage.png)
> ```

---

## Analysis Page

> **Insert Analysis Screenshot Here**
>
> File Path Example:
>
> ```md
> ![Analysis](screenshots/analysis.png)
> ```

---

## Pivot Tables & Backend Data Model

> **Insert Pivot Table Screenshot Here**
>
> File Path Example:
>
> ```md
> ![Pivot Tables](screenshots/pivot_tables.png)
> ```

---

## Refresh Macro Demonstration

> **Insert Screenshot Showing Refresh Button or VBA Automation**
>
> File Path Example:
>
> ```md
> ![Refresh Macro](screenshots/refresh_macro.png)
> ```

---

## 💼 Resume Project Highlights

* Developed an interactive Coca-Cola Retail Sales Dashboard using Microsoft Excel, VBA, Pivot Tables, and Pivot Charts.
* Automated reporting workflows through VBA macros, reducing manual effort and improving report efficiency.
* Designed dynamic KPI dashboards for monitoring retailer and beverage brand performance.
* Implemented slicer-based filtering for interactive business analysis and decision support.
* Built a scalable reporting solution capable of handling large retail sales datasets.

---

## 🧠 Skills Demonstrated

* Data Analysis
* Business Intelligence
* Dashboard Development
* Excel VBA
* Data Visualization
* KPI Reporting
* Report Automation
* Sales Analytics
* Problem Solving
* Stakeholder Reporting

---

## 🚀 Future Enhancements

* SQL Database Integration
* Power BI Migration
* Automated Email Reporting
* Forecasting & Predictive Analytics
* Machine Learning-Based Demand Forecasting

---

## 👩‍💻 Author

**Reshma**

Data Analytics & Business Intelligence Portfolio Project

This project showcases practical skills in Excel-based analytics, dashboard development, automation, and business reporting commonly used in Data Analyst and Business Analyst roles.

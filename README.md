# Customer Service Performance Analysis Dashboard in Excel - An Interactive Analysis of Operations, Representative Performance & Customer Insights

## **Introduction**

This project is an interactive Excel dashboard built to analyze customer service performance across calls, purchase amounts, call duration, satisfaction ratings, and representative-level performance. By combining KPI tracking, time-series trends, representative comparisons, and satisfaction analysis, the dashboard provides a comprehensive view of call centre operations and customer service outcomes. Interactive **Slicers**, **PivotTables**, **PivotCharts**, and **Conditional Formatting** enable dynamic filtering and visual exploration of representative performance, while supporting elements such as **XLOOKUP-based representative profiles**, ranking metrics, and custom chart data enhance the dashboard’s usability and interactivity. Designed with a clean, minimalistic, and whitespace-oriented layout, the dashboard transforms operational data into clear and accessible performance insights.

![Customer Service Performance Dashboard](Image%20-%20Customer%20Service%20Performance%20Excel%20Dashboard/Customer%20Service%20Performance%20Dashboard%20-%20Slicer%20State%201.png)

### Access the Dashboards

If you’d like to directly explore the interactive dashboards and project files, you can access them here:

[Google Drive link](https://drive.google.com/drive/u/0/folders/1NDn0yIfdXEhtsOwgGHCB7uynSZsxg6xv) : https://drive.google.com/drive/u/0/folders/1NDn0yIfdXEhtsOwgGHCB7uynSZsxg6xv

[Interactive dashboard video walkthrough](https://drive.google.com/file/d/1g-aL7QsLqk4vXGJRoSG-7CQD3_qItyGX/view): 
https://drive.google.com/file/d/1g-aL7QsLqk4vXGJRoSG-7CQD3_qItyGX/view

For the complete project details, including dataset context, analysis workflow, and documented insights, continue with this repository.

## **Project Overview**

The **Customer Service Performance Analysis Dashboard** project aims to:

* **Monitor Customer Service Operations**: Track key performance indicators such as total calls handled, purchase amount, call duration, and average satisfaction rating to provide a comprehensive view of call centre performance
* **Evaluate Representative Performance**: Compare representatives based on call volume, purchase amount, call duration, satisfaction ratings, and relative performance rankings
* **Analyze Customer Behavior & Experience**: Explore call patterns across months, days of the week, satisfaction ratings, cities, and customer demographics to identify meaningful operational trends
* **Enable Interactive Representative Analysis**: Use slicers and dynamically linked report elements to highlight selected representatives, compare their performance, and examine representative-level customer purchase activity
* **Support Data-Driven Decisions**: Transform raw customer service data into an interactive, visually structured dashboard that enables stakeholders to identify performance patterns, compare representatives, and make informed operational decisions

## **Dataset Structure**

The project is built on two interconnected datasets: a **Customer Dataset** containing customer demographics and location information, and a **Customer Service Dataset** capturing call-level interactions, representative assignments, purchase amounts, and satisfaction ratings. Together, they provide the foundation for analyzing customer service operations, representative performance, and customer experience across the 2023–2024 period.

**Customer Dataset:**  
Provides demographic and geographical information for each customer. Key features include:

* `Customer ID`: Unique identifier used to link customers with their service interactions
* `Gender`: Customer gender
* `Age`: Customer age
* `City`: Customer's city, covering Cincinnati, Cleveland, and Columbus

**Customer Service Dataset:**  
Captures individual customer service calls and their associated operational and satisfaction metrics. Key features include:

* `Call number`: Unique identifier for each customer service call
* `Customer ID`: Identifier linking each call to the corresponding customer
* `Duration`: Duration of the call in minutes
* `Representative`: Representative assigned to handle the call
* `Date of Call`: Date on which the customer interaction occurred
* `Purchase Amount`: Purchase amount associated with the call
* `Satisfaction Rating`: Customer satisfaction score for the interaction
* `FY`: Fiscal year of the call, covering 2023 and 2024
* `Day of week`: Day on which the call took place
* `Duration Bucket`: Categorized call-duration range used for distribution analysis
* `Rounded Rating`: Rounded satisfaction rating used for rating-distribution analysis and PivotCharts

The **Customer ID** serves as the key relationship between the two datasets, allowing customer attributes such as gender, age, and city to be analyzed alongside call activity, representative performance, purchase amounts, and satisfaction ratings. Additional derived fields such as `Duration Bucket`, `Day of week`, `FY`, and `Rounded Rating` support time-based, distributional, and operational analysis within the Excel dashboard.

## **Dashboard Development & Analytical Workflow**

- **Defined Analytical Objectives**  
  Established the key performance dimensions for the call-centre analysis, focusing on call volume, purchase amount, call duration, customer satisfaction, representative performance, and customer demographics across cities and gender.

- **Data Preparation & Feature Engineering**  
  - Structured the underlying call-centre and customer data into dedicated Excel tables.
  - Created analytical fields such as `FY`, `Day of week`, `Duration Bucket`, and `Rounded Rating` to support time-based, duration, and satisfaction analysis.
  - Built a Power Pivot data model connecting the `Call_Centre`, `Customers`, and `Dim_Ratings` tables through defined relationships.

- **Pivot-Based Analytical Modeling**  
  - Developed PivotTables to calculate overall and representative-level KPIs, including call count, purchase amount, call duration, average satisfaction rating, and 5★ calls.
  - Created monthly and weekly call trends, satisfaction-rating distributions, gender breakdowns by city, and representative performance summaries.
  - Built a customer-by-representative purchase matrix to examine purchase amounts across customers, cities, and representatives.
  - Developed supporting chart-data and calculation tables to control the dashboard's dynamic visual behavior.

- **Interactive Dashboard Construction**  
  - Designed a single-page **Customer Centre Report** combining KPI cards, time-series analysis, categorical comparisons, representative performance charts, and a detailed customer/city matrix.
  - Added a representative slicer for interactive filtering and performance exploration.
  - Implemented dynamic highlighting so that selected representatives are visually emphasized across representative charts and the customer/city matrix.
  - Extended the slicer logic to support **multiple representative selections**, including combined KPI values, collective representative imagery, multiple-selection labels, and simultaneous highlighting of selected representatives.
  - Used `XLOOKUP` to dynamically retrieve representative images and conditional formatting to create data-bar visualizations within the detailed matrix.
  - Applied a consistent visual design with KPI cards, rounded containers, whitespace, icons, and a cohesive blue-gray color palette.

- **Dynamic Selection & Reporting Logic**  
  - Developed helper formulas using functions such as `COUNTIF`, `IF`, `XLOOKUP`, and `NA()` to control which representatives appear as selected in charts and tables.
  - Implemented selection-aware labels and ranking logic, displaying individual call and amount ranks for single selections while switching to a **Multiple Selected** state when several representatives are chosen.
  - Created a collective representative image state for multi-selection scenarios rather than displaying a single representative image.

- **Dashboard Validation & Presentation**  
  - Tested the dashboard across individual representative selections and collective selections to ensure that KPIs, charts, conditional formatting, images, rankings, and table highlighting responded correctly.
  - Organized the workbook into four functional worksheets: **Data**, **Pivot Tables**, **Customer Centre Report**, and **Assets**, separating the data, analytical, presentation, and visual-resource layers.
 
## **Dashboard Previews**  

### Customer Service Performance Dashboard - Slicer State 1

![Customer Service Performance Dashboard - Slicer State 1](Image%20-%20Customer%20Service%20Performance%20Excel%20Dashboard/Customer%20Service%20Performance%20Dashboard%20-%20Slicer%20State%201.png)

### Customer Service Performance Dashboard - Slicer State 2

![Customer Service Performance Dashboard - Slicer State 2](Image%20-%20Customer%20Service%20Performance%20Excel%20Dashboard/Customer%20Service%20Performance%20Dashboard%20-%20Slicer%20State%202.png)

### Customer Service Performance Dashboard - Slicer State 3

![Customer Service Performance Dashboard - Slicer State 3](Image%20-%20Customer%20Service%20Performance%20Excel%20Dashboard/Customer%20Service%20Performance%20Dashboard%20-%20Slicer%20State%203.png)

## **Key Insights**

### **Overall Customer Service Performance**
- **Call Volume & Value**: The call centre handled **1,000 customer calls**, generating **₹96,623** in purchase value and **89,850 minutes** of total call duration.
- **Customer Satisfaction**: The overall average satisfaction rating was **3.9/5**, with **307 calls receiving a 5-star rating**.
- **Positive Rating Concentration**: Ratings of **4 and 5 accounted for 73.5% of all calls** (735 out of 1,000), indicating that the majority of customer interactions resulted in positive satisfaction outcomes.
- **Low-Rating Incidence**: Only **68 calls (6.8%)** received ratings of 0–2, suggesting that highly dissatisfied interactions represented a relatively small share of the overall call volume.

### **Representative Performance**
- **Call Volume Leader**: **R02** handled the highest number of calls with **218 calls (21.8%)**, making it the busiest representative by interaction volume.
- **Purchase Value Leader**: **R03** generated the highest purchase amount at **₹20,872**, despite handling fewer calls than R02, indicating stronger purchase value per interaction.
- **Performance Distribution**: Call volumes were relatively balanced across representatives, ranging from **186 calls for R04** to **218 calls for R02**, with no single representative handling a disproportionate share of the workload.
- **Representative Comparison**: The dashboard's interactive representative selector enables individual performance to be compared against overall call volume, purchase amount, call duration, and satisfaction distribution.

### **Customer & Geographic Patterns**
- **Gender Distribution**: Female customers accounted for **599 calls (59.9%)**, compared with **401 calls (40.1%)** from male customers.
- **City-Level Call Volume**: **Cleveland** recorded the highest call volume with **389 calls**, followed by **Columbus (335)** and **Cincinnati (276)**.
- **Gender Variation by City**: Cleveland showed a strong female customer concentration (**326 female vs. 63 male calls**), while Columbus displayed the opposite pattern (**129 female vs. 206 male calls**). Cincinnati was comparatively balanced at **144 female vs. 132 male calls**.
- **Customer-Level Purchase Analysis**: The customer-by-representative matrix provides a granular view of purchase amounts, allowing individual customer spending patterns to be compared across representatives and cities.

### **Temporal & Operational Trends**
- **Monthly Call Peaks**: Call volume peaked in **March with 155 calls**, followed by **April with 136 calls**. The lowest monthly volume occurred in **August with 50 calls**, indicating substantial variation throughout the year.
- **Weekly Workload Pattern**: **Saturday** recorded the highest call volume at **161 calls**, followed by **Wednesday (153)** and **Sunday (146)**. **Monday** had the lowest weekday volume at **133 calls**.
- **Seasonal Workload Variation**: The monthly trend shows notable peaks in **March–April** and again in **October**, while call volumes declined considerably during **August**, highlighting periods of increased and reduced service demand.
- **Representative Workload Planning**: The concentration of calls on Saturdays and during peak months can help inform staffing allocation and workload planning across customer-service representatives.

## **Project Highlights**

* Developed a **fully interactive Excel customer service performance dashboard** to provide a consolidated view of call volume, purchase amount, call duration, satisfaction, representative performance, and customer demographics.
* Adopted a **question-driven analytical approach**, leveraging PivotTables, PivotCharts, Slicers, XLOOKUPs, Conditional Formatting, and supporting calculations to transform raw call-centre data into an interactive reporting solution.
* Engineered **custom KPIs and representative-level metrics**, including call share, purchase amount, call duration, call and amount rankings, and 5-star call counts to evaluate individual representative performance.
* Implemented **dynamic representative selection**, allowing users to select individual representatives or multiple representatives and instantly update the dashboard's KPIs, trends, charts, representative highlights, and profile imagery.
* Conducted **temporal, satisfaction, demographic, and geographic analysis**, revealing monthly and weekly call patterns, satisfaction-rating distributions, gender differences across cities, and customer purchase amounts by representative.
* Built a **Power Pivot data model** connecting the primary call-centre dataset with customer and rating dimension tables, enabling structured relationship-based analysis and PivotTable reporting.
* Created a **customer-level purchase matrix** with conditional formatting to compare representative-specific purchase amounts across customers and cities while retaining overall totals.
* Implemented a **clean, user-centric dashboard design** with KPI cards, interactive charts, representative selectors, custom icons, representative imagery, and a consistent visual theme for intuitive stakeholder exploration.
* Delivered a **scalable Excel-based analytical solution** that connects data preparation, relational modeling, PivotTable analysis, interactive visualization, and reporting within a single workbook.

This project demonstrates how **Microsoft Excel can be used to build an interactive, model-driven customer service performance dashboard**, combining advanced spreadsheet functionality with structured business analysis to transform raw call-centre data into an accessible and decision-oriented reporting solution.

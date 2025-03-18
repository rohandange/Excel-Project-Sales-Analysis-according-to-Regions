# Excel Project for Sales Analysis by Region

  This Excel project performs a comprehensive **Sales Analysis** for different sales executives based on multiple key performance indicators, including **Total Sales**, **Target Hit Percentage**, and **Away from Target Hit Percentage**. The project incorporates data cleaning, pivot table creation, data visualization, and automation using macros to analyze sales data across different regions. The objective of this project is to provide a user-friendly dashboard for efficiently tracking and analyzing the performance of sales executives.

# Project Overview

The project follows a 4-step process to clean, analyze, and report sales data:

1. **Understanding and Cleaning the Data**
    - The dataset consists of 142 rows and 12 columns, containing information on sales executives, total sales, target hit percentages, and regional sales data.
    - The data was cleaned by changing data types, merging columns, and ensuring the proper format for all entries.
    - The dataset was converted from CSV to XLSX format to preserve any formatting and formula-based changes.

2. **Building the Analysis Logic**
    - The project focuses on analyzing sales performance at both the top and bottom ends. Sales executives are ranked based on total sales and target hit       
      percentage, with the goal of identifying key performers and areas for improvement.

3. **Applying Formulas and Pivot Tables**
    - Several formulas were applied for data transformation, including merging columns and formatting names to a proper "Full Name" format.
    - Pivot tables were created for the following analyses:
        - **Top 5 Sales Executives**: A pivot table was used to display the top 5 sales executives based on total sales.
        - **Bottom 5 Sales Executives**: A pivot table was created to show the bottom 5 sales executives based on total sales.
        - **Top 5 Target Hit Percentage**: Pivot tables display the top 5 sales executives based on their target hit percentage.
        - **Top 5 Away from Target Percentage**: This table highlights sales executives who are furthest from hitting their targets.

4. **Creating the Report and Dashboard**
    - **Pivot Charts**: Each pivot table was visualized using pivot charts to provide better insights into the data.
    - **Slicers**: Slicers were inserted to allow users to filter data by regions for targeted analysis.
    - **Buttons and Checkboxes**: Buttons were created for each pivot table, with form controls used to toggle and filter the data dynamically.
    - **Macro Automation**: A macro was recorded to automate the process of connecting slicers to pivot tables and toggling filter options. The macro was then 
    enhanced with VBA (Visual Basic for Applications) code to automate the report generation process for each sales region.

# Key Features
- **Data Cleaning & Transformation**: Improved data structure and integrity through proper formatting and cleaning.
- **Dynamic Pivot Tables**: Created pivot tables for analyzing top and bottom performers based on sales and target achievements.
- **Interactive Dashboard**: Implemented slicers, buttons, and pivot charts to make the report interactive and user-friendly.
- **Macro Automation**: Recorded and customized macros using VBA to automate common tasks and enhance report generation efficiency.

# How to Use
1. Open the Excel file in Microsoft Excel.
2. Use the **slicer** to filter sales data by region.
3. Click on the **buttons** to view the corresponding pivot tables for sales performance, top sales executives, and target achievements.
4. Use the **checkbox** controls to toggle between different views and filter settings.
5. The macros will automatically update the report based on your selection.

# Technologies Used
- Microsoft Excel (Pivot Tables, Pivot Charts, Slicers, Macros)
- VBA (Visual Basic for Applications)

# Conclusion
This project enables efficient analysis of sales performance across regions and provides insights into top-performing sales executives and areas needing attention. The combination of pivot tables, charts, and automated macros makes this solution powerful for any sales team looking to monitor and enhance their performance metrics.

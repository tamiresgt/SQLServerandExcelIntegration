# SQLServer and Excel Integration Project

This project aims to integrate data from a SQL Server database into Excel, creating readable tables and interactive dashboards for analyzing the company's online sales **Adventure Works** during 2013. The focus is on using **SQL** to extract and manipulate data and then using **Excel** to present visual and analytical insights.

![dashboard](https://github.com/user-attachments/assets/5d574f6b-88e4-4f51-be60-3bea0eb325a3)


**- Tools and Software:** SQLServer, Excel and PowerQuery  
**- Database:** AdventureWorks2014 - Microsoft example database

## Project Steps

- Step 1: Download the Microsoft database: AdventureWorksDW2014.bak using the link:
https://docs.microsoft.com/pt-br/sql/samples/adventureworks-install-configure?view=sql-server-ver16&tabs=ssms

- Step 2: Restore database in SQL Server

- Step 3 - Project definition: Creation of a data panel in Excel, which presents a detailed report on the **online sales** of AdventureWorks in 2013. The following indicators will be analyzed:

  - Total revenue versus total cost by country,
  - Sales evolution by month,
  - Sales by category,
  - Sales by Gender.
 
  **Note:** The impact of currency will be disregarded for simplification in this exercise.

- Step 4: Database Exploration and Defining the tables to be analyzed

  - DimCustomer
  - FactInternetSales
  - DimSalesTerritory
  - DimProduct
  - DimProductCategory
  - DimProductSubcategory

- Step 5: Define the necessary columns of the fact table to create the relationships

- Setp 6: Perform JOINS with the other tables to assemble the final table and create a view

- Step 7: Integration with Excel

- Step 8: Perform simple transformations using PowerQuery

- Step 9: Create PivotTables and the dashboard

## Analysis
- **Monthly Online Sales Growth:**  
Adventure Works' online sales increased month on month throughout 2013, highlighting a consistent growth trend. This may indicate an effective sales strategy or greater adoption of online shopping by consumers.

- **Product Categories:**  
The Accessories category was the best seller, while Bikes generated the most revenue. The analysis suggests that although accessories are popular in terms of sales volume, bikes have a greater impact on the company's profit. So, increasing sales of Bikes could be an effective strategy for maximizing profit.

- **Clothing Category Performance:**  
The clothing category is underperforming, both in terms of sales and profit margin. This could be an opportunity for improvement, whether in therms of price adjustments, promotions or marketing strategies. A market and consumer study could help to understand better the consumers needs in this category.

- **Performance by Geographical Region**  
Australia and United States are the countries with the best revenue versus product cost, indicating greater operational efficiency in these markets. On the other hand, other countriespresent opportunities for improvement both in terms of cost reduction and revenue increase.

- **Sales by Gender:**  
Both men and women buy from Adventure Works, demonstrating that marketing campaings should be inclusive reaching both audiences effectively.

## Conclusion
This project demonstrates how to integrate and analyze large volumes of data using tools such as SQL Server and Excel. By creating an interactive visualization, it is possible to extract meaningful insights that can inform strategic decisions in the company, such as product category optimization, geographic sales analysis and audience segmentation.

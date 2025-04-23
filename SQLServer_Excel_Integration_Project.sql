-- SQL Server and Excel Integration Project

-- Step 1: Download the Microsoft database: AdventureWorksDW2014.bak using the link:
--https://docs.microsoft.com/pt-br/sql/samples/adventureworks-install-configure?view=sql-server-ver16&tabs=ssms

-- Step 2: Restore database in SQL Server

-- Step 3 - Project definition: 
--Build a data panel in Excel that presents a report on the results of AdventureWorks' internet sales, showing the following indicators:

-- Total revenue versus total cost by country,
-- Sales evolution by month,
-- Sales by category,
-- Sales by Gender

-- The year analyzed will be 2013
-- The currency issue should be ignored for this exercise

-- Step 4: Getting to know the database and Defining the tables to be analyzed

SELECT * FROM DimCustomer
SELECT * FROM FactInternetSales
SELECT * FROM DimSalesTerritory
SELECT * FROM DimProduct
SELECT * FROM DimProductCategory
SELECT * FROM DimProductSubcategory

-- Step 5: Define the necessary columns of the fact table to create the relationships

SELECT 
	SalesOrderNumber,
	ProductKey,
	CAST(OrderDate AS DATE) as 'OrderDate',
	DATENAME(MONTH, OrderDate)  as 'MonthOrderDate',
	CustomerKey,
	SalesTerritoryKey,
	OrderQuantity
	TotalProductCost, 
	SalesAmount
FROM FactInternetSales

-- Setp 6: Perform JOINS with the other tables to assemble the final table and create a view

CREATE OR ALTER VIEW vwInternetSales2013 AS
SELECT
	FactInternetSales.SalesOrderNumber,
	FactInternetSales.OrderDate,
	DimProduct.EnglishProductName AS 'ProductName',
	DimProductCategory.EnglishProductCategoryName AS 'ProductCategoryName',
	CASE
		WHEN DimCustomer.MiddleName IS NOT NULL THEN CONCAT(DimCustomer.FirstName,' ',DimCustomer.MiddleName,' ',DimCustomer.LastName)
		ELSE CONCAT(DimCustomer.FirstName,' ',DimCustomer.LastName)
	END AS 'CustomerName',
	CASE
		WHEN DimCustomer.Gender = 'M' THEN 'Male'
		ELSE 'Female'
	END AS 'Gender',
	DimSalesTerritory.SalesTerritoryCountry AS 'Country',
	FactInternetSales.OrderQuantity,
	FactInternetSales.TotalProductCost,
	FactInternetSales.SalesAmount

FROM FactInternetSales
INNER JOIN DimProduct ON FactInternetSales.ProductKey = DimProduct.ProductKey
	INNER JOIN DimProductSubcategory ON DimProduct.ProductSubcategoryKey = DimProductSubcategory.ProductSubcategoryKey
		INNER JOIN DimProductCategory ON DimProductSubcategory.ProductCategoryKey = DimProductCategory.ProductCategoryKey
INNER JOIN DimCustomer ON FactInternetSales.CustomerKey = DimCustomer.CustomerKey
INNER JOIN DimSalesTerritory ON FactInternetSales.SalesTerritoryKey = DimSalesTerritory.SalesTerritoryKey
WHERE YEAR(OrderDate) = 2013

SELECT * FROM vwInternetSales2013

-- Step 7: Integration with Excel
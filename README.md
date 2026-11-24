# Sales Management Dashboard | SQL + Power BI
 
An end-to-end data analytics project that turns raw AdventureWorks sales data into an interactive executive sales dashboard for sales managers. The data is cleaned and modelled with **SQL Server**, combined with an **Excel** budget file, and visualised in **Power BI**.
 
## Dashboard Preview
 
### Sales Overview
 
![Sales Overview](Sales%20Overview.png)
 
### Star Schema (Power BI Model)
 
![Star Schema](Star%20Schema.png)
 
---
 
## Table of Contents
 
- [Business Request & User Stories](#business-request--user-stories)
- [Tools Used](#tools-used)
- [Data Cleansing & Transformation (SQL)](#data-cleansing--transformation-sql)
- [Data Model](#data-model)
- [Sales Management Dashboard](#sales-management-dashboard)
- [More Screenshots](#more-screenshots)
- [Credits](#credits)
---
 
## Business Request & User Stories
 
The business asked for an executive sales report for sales managers. To deliver it and keep the work on track, the request was broken down into the user stories below, each with its own acceptance criteria.
 
| # | As a (role) | I want (request / demand) | So that I (user value) | Acceptance Criteria |
|---|-------------|---------------------------|------------------------|---------------------|
| 1 | Sales Manager | To get a dashboard overview of internet sales | Can follow better which customers and products sell the best | A Power BI dashboard which updates data once a day |
| 2 | Sales Representative | A detailed overview of Internet Sales per Customers | Can follow up my customers that buy the most and who we can sell more to | A Power BI dashboard which allows me to filter data for each customer |
| 3 | Sales Representative | A detailed overview of Internet Sales per Products | Can follow up my products that sell the most | A Power BI dashboard which allows me to filter data for each product |
| 4 | Sales Manager | A dashboard overview of internet sales | Follow sales over time against budget | A Power BI dashboard with graphs and KPIs comparing against budget |
 
---
 
## Tools Used
 
- **SQL Server** (AdventureWorksDW2019) for data extraction, cleansing and transformation
- **Excel** for the sales budget data source
- **Power BI** for data modelling and dashboard design
---
 
## Data Cleansing & Transformation (SQL)
 
To build the data model needed to meet the user stories, the following tables were extracted and cleaned using SQL.
 
The sales budget was provided separately as an Excel file and was connected to the data model in a later step.
 
### DIM_Calendar
 
```sql
-- Cleansed DIM_Date Table --
SELECT 
  [DateKey], 
  [FullDateAlternateKey] AS Date, 
  --[DayNumberOfWeek], 
  [EnglishDayNameOfWeek] AS Day, 
  --[SpanishDayNameOfWeek], 
  --[FrenchDayNameOfWeek], 
  --[DayNumberOfMonth], 
  --[DayNumberOfYear], 
  --[WeekNumberOfYear],
  [EnglishMonthName] AS Month, 
  Left([EnglishMonthName], 3) AS MonthShort,   -- Useful for front end date navigation and front end graphs.
  --[SpanishMonthName], 
  --[FrenchMonthName], 
  [MonthNumberOfYear] AS MonthNo, 
  [CalendarQuarter] AS Quarter, 
  [CalendarYear] AS Year --[CalendarSemester], 
  --[FiscalQuarter], 
  --[FiscalYear], 
  --[FiscalSemester] 
FROM 
 [AdventureWorksDW2019].[dbo].[DimDate]
WHERE 
  CalendarYear >= 2019
```
 
### DIM_Customers
 
```sql
-- Cleansed DIM_Customers Table --
SELECT 
  c.customerkey AS CustomerKey, 
  --      ,[GeographyKey]
  --      ,[CustomerAlternateKey]
  --      ,[Title]
  c.firstname AS [First Name], 
  --      ,[MiddleName]
  c.lastname AS [Last Name], 
  c.firstname + ' ' + lastname AS [Full Name], 
  -- Combined First and Last Name
  --      ,[NameStyle]
  --      ,[BirthDate]
  --      ,[MaritalStatus]
  --      ,[Suffix]
  CASE c.gender WHEN 'M' THEN 'Male' WHEN 'F' THEN 'Female' END AS Gender,
  --      ,[EmailAddress]
  --      ,[YearlyIncome]
  --      ,[TotalChildren]
  --      ,[NumberChildrenAtHome]
  --      ,[EnglishEducation]
  --      ,[SpanishEducation]
  --      ,[FrenchEducation]
  --      ,[EnglishOccupation]
  --      ,[SpanishOccupation]
  --      ,[FrenchOccupation]
  --      ,[HouseOwnerFlag]
  --      ,[NumberCarsOwned]
  --      ,[AddressLine1]
  --      ,[AddressLine2]
  --      ,[Phone]
  c.datefirstpurchase AS DateFirstPurchase, 
  --      ,[CommuteDistance]
  g.city AS [Customer City] -- Joined in Customer City from Geography Table
FROM 
  [AdventureWorksDW2019].[dbo].[DimCustomer] as c
  LEFT JOIN dbo.dimgeography AS g ON g.geographykey = c.geographykey 
ORDER BY 
  CustomerKey ASC -- Ordered List by CustomerKey
```
 
### DIM_Products
 
```sql
-- Cleansed DIM_Products Table --
SELECT 
  p.[ProductKey], 
  p.[ProductAlternateKey] AS ProductItemCode, 
  --      ,[ProductSubcategoryKey], 
  --      ,[WeightUnitMeasureCode]
  --      ,[SizeUnitMeasureCode] 
  p.[EnglishProductName] AS [Product Name], 
  ps.EnglishProductSubcategoryName AS [Sub Category], -- Joined in from Sub Category Table
  pc.EnglishProductCategoryName AS [Product Category], -- Joined in from Category Table
  --      ,[SpanishProductName]
  --      ,[FrenchProductName]
  --      ,[StandardCost]
  --      ,[FinishedGoodsFlag] 
  p.[Color] AS [Product Color], 
  --      ,[SafetyStockLevel]
  --      ,[ReorderPoint]
  --      ,[ListPrice] 
  p.[Size] AS [Product Size], 
  --      ,[SizeRange]
  --      ,[Weight]
  --      ,[DaysToManufacture]
  p.[ProductLine] AS [Product Line], 
  --     ,[DealerPrice]
  --      ,[Class]
  --      ,[Style] 
  p.[ModelName] AS [Product Model Name], 
  --      ,[LargePhoto]
  p.[EnglishDescription] AS [Product Description], 
  --      ,[FrenchDescription]
  --      ,[ChineseDescription]
  --      ,[ArabicDescription]
  --      ,[HebrewDescription]
  --      ,[ThaiDescription]
  --      ,[GermanDescription]
  --      ,[JapaneseDescription]
  --      ,[TurkishDescription]
  --      ,[StartDate], 
  --      ,[EndDate], 
  ISNULL (p.Status, 'Outdated') AS [Product Status] 
FROM 
  [AdventureWorksDW2019].[dbo].[DimProduct] as p
  LEFT JOIN dbo.DimProductSubcategory AS ps ON ps.ProductSubcategoryKey = p.ProductSubcategoryKey 
  LEFT JOIN dbo.DimProductCategory AS pc ON ps.ProductCategoryKey = pc.ProductCategoryKey 
order by 
  p.ProductKey asc
```
 
### FACT_InternetSales
 
```sql
-- Cleansed FACT_InternetSales Table --
SELECT 
  [ProductKey], 
  [OrderDateKey], 
  [DueDateKey], 
  [ShipDateKey], 
  [CustomerKey], 
  --  ,[PromotionKey]
  --  ,[CurrencyKey]
  --  ,[SalesTerritoryKey]
  [SalesOrderNumber], 
  --  [SalesOrderLineNumber], 
  --  ,[RevisionNumber]
  --  ,[OrderQuantity], 
  --  ,[UnitPrice], 
  --  ,[ExtendedAmount]
  --  ,[UnitPriceDiscountPct]
  --  ,[DiscountAmount] 
  --  ,[ProductStandardCost]
  --  ,[TotalProductCost] 
  [SalesAmount] --  ,[TaxAmt]
  --  ,[Freight]
  --  ,[CarrierTrackingNumber] 
  --  ,[CustomerPONumber] 
  --  ,[OrderDate] 
  --  ,[DueDate] 
  --  ,[ShipDate] 
FROM 
  [AdventureWorksDW2019].[dbo].[FactInternetSales]
WHERE 
  LEFT (OrderDateKey, 4) >= YEAR(GETDATE()) -2 -- Ensures we always only bring two years of date from extraction.
ORDER BY
  OrderDateKey ASC
```
 
---
 
## Data Model
 
After the cleansed tables were loaded into Power BI, they were joined into a star-schema data model. The budget table (`FACT_Budget`) was connected to `FACT_InternetSales` and the other dimension tables so that actual sales could be compared against budget.
 
The full star schema diagram is shown in the [Dashboard Preview](#dashboard-preview) at the top.
 
---
 
## Sales Management Dashboard
 
The finished report has three pages:
 
1. **Overview dashboard** with headline KPIs and sales against budget
2. **Sales over time**, with visuals showing sales trends
3. **Sales per customer and per product**, with filters for detailed follow-up
The overview page is shown in the [Dashboard Preview](#dashboard-preview) at the top, and the detail pages are at the end of this README.
 
---
 
## More Screenshots
 
### Customer Details
 
![Customer Details](Customer%20Details.png)
 
### Product Details
 
![Product Details](Product%20Details.png)
 
-- 
This project was built by following the **AnalyzeWithAli** data analyst portfolio tutorial. The original project is at [analyzewithaliportfolio.wordpress.com](https://analyzewithaliportfolio.wordpress.com/home-2/).
 

-- Cleansed FACT_InternetSales Table --
SELECT
[ProductKey],
[OrderDateKey],
[DueDateKey],
[ShipDateKey],
[CustomerKey],
--. ,[Promotionkey]
--,[CurrencyKey]
--,[SalesTerritoryKey]
[SalesOrderNumber],
--[SalesOrderLineNumber],
--,[RevisionNumber]
--,[OrderQuantity],
--,[UnitPrice],
--,[ExtendedAmount]
--,[UnitPriceDisEountPct]
--,[DiscountAmount]
--,[ProductStandardCost]
--,[TotalProductCost]
[SalesAmount] -- ,[TaxAmt]
--,[Freight]
--,[CarrierTrackingNumber]
--,[CustomerPONumber]
--,[OrderDate]
--,[DueDate]
--,[ShipDate]
FROM 
[dbo].[FactInternetSales]
WHERE
LEFT (OrderDatekey, 4) >= YEAR(GETDATE()) -2 -- Ensures we always only bring two years of date from extraction.
ORDER BY
OrderDatekey ASC
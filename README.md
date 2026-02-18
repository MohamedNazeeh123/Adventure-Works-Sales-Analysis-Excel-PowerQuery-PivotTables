# Adventure-Works-Sales-Analysis-Excel-PowerQuery-PivotTables
## 🛠 Tech Stack
- Excel (Pivot Tables, Power Query)
- Data Analysis
## 🎯 Project Objective
Sales-Analysis-Excel-PowerQuery-PivotTables In this project, I utilized Excel and Power Query to perform an in-depth sales analysis. The goal was to analyze raw sales data, identify key performance metrics, and deliver actionable insights through a dynamic and interactive dashboard.

After downloading AdventureWorks dataset " .bak file " from Microsoft Learn (The Link is provided Finally in the last section "Data Sources") 1-) I imported the " .bak file " in SQL Server using SSMS. 2-) From Data view in Microsoft Excel, I imported the dataset from SQL Server to Microsoft Excel PowerQuery to apply some data transformation and cleaning to make dataset ready for analysis and visualization like:

-Clear unnecessary columns. -Handling NULL values. -Merging some columns together. -Joining Tables together using "Merge Queries" command. -Handling PK and FK between tables.
## 📂 Dataset Overview
The project uses **5 tables** from the AdventureWorks dataset:

1. **FactOrders**
   - Columns: SalesOrderID, CustomerID, TerritoryID, ShipMethodID, SubTotal, TaxAmt, Freight, TotalDue, SalesOrderDetailID, OrderQty, ProductID, LineTotal, OrderDateKey

2. **DimProducts**
   - Columns: ProductID, Product, Color, Category, SubCategory

3. **DimShipMethods**
   - Columns: ShipMethodID, ShipMethod

4. **DimTerritory**
   - Columns: TerritoryID, Territory, Group

5. **DimDate**
   - Columns: OrderDate, DateKey, Year, Month, Day, MonthName, OrderDate
   - 
<img width="1362" height="768" alt="data model" src="https://github.com/user-attachments/assets/7662c054-6cb3-4569-addc-3665964f0ef4" />

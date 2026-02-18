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
   - The Dimensional Data Model is STAR SCHEMA, as shown below: dataModel.png
<img width="1362" height="768" alt="data model" src="https://github.com/user-attachments/assets/7662c054-6cb3-4569-addc-3665964f0ef4" />
After data preparation in PowerQuery, I loaded data into Excel worksheet to start the analysis and visualization processes using Pivot Tables, as shown below:

Pivot Tables:

<img width="1366" height="728" alt="pivot table" src="https://github.com/user-attachments/assets/8b14e5d3-8a9d-4bcf-a758-af42ed7fd9b5" />
## 📈 Dashboard Analysis

The dashboard allows interactive filtering by:
- Order Date
- Year
- Product Category

### 📊 Main KPIs
- Total Sales
- Number of Orders
- Number of Customers
- Total Tax
- Total Freight

### 📊 Analytical Insights
- Orders distribution by Ship Method
- Top 5 best-selling Products
- Sales performance by Territory
- Monthly Sales Trend
- Orders and Sales breakdown by Category
FINAL Dashboard:

<img width="1327" height="440" alt="dashboard" src="https://github.com/user-attachments/assets/18d58c26-7c50-4b88-b92d-a246a555f2e9" />
data Source -> https://learn.microsoft.com/en-us/sql/samples/adventureworks-install-configure?view=sql-server-ver16&tabs=ssms

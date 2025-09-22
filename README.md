# Excel Interview Questions

### Basic & Data Handling 
1. What are the differences between Excel Tables vs Normal Ranges ? Why use tables in 
analytics?  
2. How do you remove duplicates from a dataset?  
3. How would you clean messy data (extra spaces, text to columns, inconsistent formats)?  
4. How do you handle missing values in Excel data?  
 
### Formulas & Functions  
5. Explain the difference between VLOOKUP, HLOOKUP, XLOOKUP, and INDEX -MATCH. 
Which one do you prefer and why?  
6. What is the difference between absolute, relative, and mixed cell referencing ? 
7. How would you use IF, IFS, and nested IFs  to categorize data?  
8. Can you explain the use of TEXT, LEFT, RIGHT, MID, FIND, LEN, TRIM, and 
### CONCATENATE (or TEXTJOIN)  in data cleaning?  
9. How does the MATCH function  work? Give an example with INDEX.  
10. When would you use OFFSET and INDIRECT functions ? 
 
### Analytical & Business Functions  
11. How do you calculate percentages, growth rates, and CAGR  in Excel?  
12. What is the difference between COUNT, COUNTA, COUNTBLANK, and 
COUNTIF/COUNTIFS ? 
13. How would you calculate running totals or cumulative sums  in Excel?  
14. What is the use of RANK, RANK.EQ, RANK.AVG  in analytics?  
15. How can you use SUMPRODUCT  for conditional calculations?  
 
### Pivot Tables & Reporting  
16. Explain how Pivot Tables help in data analysis.  
17. How do you create a Pivot Table that shows % contribution to total sales by region ? 
18. What are calculated fields and calculated items  in Pivot Tables?  
19. How can you create a dynamic dashboard  in Excel with slicers & pivot charts?  
20. How would you refresh Pivot Tables automatically when data updates?  
 

### Scenario -based Questions  
21. Suppose you have a dataset with Sales, Date, and Region . How would you find the top 
5 performing regions ? 
22. You have sales data for multiple years. How would you calculate Year-over-Year growth  
in Excel?  
23. How do you highlight the top 10% of sales reps  using conditional formatting?  
24. If you get two datasets (Orders & Customers) , how would you merge them in Excel?  
25. A dataset has outliers (e.g., one order with extremely high sales). How would you detect 
and handle it?  
 
### Advanced Excel (Optional, for Analyst Role)  
26. What are array formulas and dynamic arrays ? (Examples: FILTER, SORT, UNIQUE, 
SEQUENCE)  
27. How do you use Power Query  for data cleaning and transformation?  
28. What is the difference between Excel and Power Pivot / Data Model ? 
29. Can you write a formula to extract unique customers who purchased more than 3 
times? 
30. How would you handle large datasets (1M+ rows)  in Excel?  
 
 
 What are the differences between Excel Tables vs Normal Ranges? Why use tables in 
analytics?  
Normal Range  = Just a collection of cells.  
Excel Table  = Structured dataset with special features.  
  Key Differences:  
Feature  Normal Range  Excel Table  
Formatting  Needs manual formatting  Automatically formatted (banded 
rows, filter buttons)  
Dynamic Range  Fixed — if new rows are added, 
formulas/pivots won’t auto -update  Expands automatically when new 
rows/columns are added  
References  Uses cell references ( A2:C100) Uses structured references 
(Sales[Amount] ) 
Sorting/Filtering  Needs manual filter application  Built-in filter and sort buttons  

Feature  Normal Range  Excel Table  
Formulas  Must copy formulas down manually  Auto-fills formulas for the whole 
column  
Pivot Tables  Needs manual range selection  Auto-connects with table name 
(e.g., Sales) 
Readability  Hard to understand  Self-explanatory (column names 
used directly)  
  Why use Tables in Analytics?  
• Dynamic nature : When new data comes in, all charts, pivot tables, and formulas 
update automatically.  
• Clean structured references : Easy to read and maintain formulas.  
• Better data integrity : Prevents accidental formula breaks when rows are added.  
• Useful for dashboards : Tables + slicers work seamlessly.  
 Example : If your data range is converted to a table named Sales, you can write:  
=SUM(Sales[Amount])  
instead of  
=SUM(C2:C1000)  
 
 How do you remove duplicates from a dataset?  
  Methods:  
1. Using Remove Duplicates Tool (Quickest)  
o Select dataset → Go to Data → Remove Duplicates  
o Choose columns to check for duplicates (e.g., Customer ID + Date).  
o Click OK → Excel removes duplicate rows.  
2. Using Advanced Filter  
o Go to Data → Advanced  
o Select “Unique records only” → Copy to another location.  
3. Using Formulas (Flag Duplicates)  
o =COUNTIF($A$2:A2,A2)>1  → Returns TRUE if the value already appeared.  
o Use this to highlight or filter duplicate rows.  
4. Using Conditional Formatting  

o Home → Conditional Formatting → Highlight Duplicate Values  → visually spots 
duplicates.  
 Best Practice in Analytics : Don’t delete duplicates blindly. First, check why duplicates exist  
(could be data entry errors, or valid multiple records like repeat purchases).  
 
 How would you clean messy data (extra spaces, text to columns, inconsistent formats)?  
Messy data cleaning = Data Preprocessing Step  
  Common Techniques:  
1. Remove Extra Spaces  
o Use =TRIM(A2)  → removes unnecessary spaces (except single space between 
words).  
o Example: " Mukesh Bhai "  → "Mukesh Bhai" . 
2. Standardize Case  
o =PROPER(A2)  → Each Word Capitalized.  
o =UPPER(A2)  → All Uppercase.  
o =LOWER(A2)  → All Lowercase.  
3. Split Data (Text to Columns)  
o Go to Data → Text to Columns . 
o Example: "Mukesh,Bhai,India"  → separated into 3 columns.  
4. Fix Inconsistent Formats  
o Dates: Use =TEXT(A2,"dd -mm-yyyy") to standardize.  
o Numbers: Remove text formatting ( VALUE(A2) ). 
5. Remove Non -printable Characters  
o =CLEAN(A2)  → removes hidden characters like line breaks.  
6. Combine Columns  
o =CONCATENATE(A2," ",B2)  or =TEXTJOIN(" ",TRUE,A2:B2) . 
 Example:  
Original: " 25/12/2025 "  → TRIM + TEXT → "25-12-2025"  
 
 How do you handle missing values in Excel data?  
Missing values are common in real -world datasets.  
  Techniques:  

1. Identify Missing Values  
o Use =COUNTBLANK(A2:A100)  → Counts blanks.  
o Use filters or conditional formatting (highlight blanks).  
2. Handle Missing Values  
o Delete Rows  (only if very few and non -critical).  
o Replace with Zero  → =IF(A2="",0,A2)  
o Replace with Average/Median  → 
=IF(A2="",AVERAGE($A$2:$A$100),A2)  
o Forward Fill/Backward Fill  → Copy last known value to missing cell.  
o Use Interpolation  (if sequential data like time series).  
3. Flag Missing Data for Review  
o Add a helper column → =IF(A2="","Missing","OK")  
  Best Practice in Analytics:  
• If it’s customer data  (like missing phone number), mark as “Unknown” instead of 
deleting.  
• If it’s numeric data  (like missing sales amount), fill with mean/median or leave blank 
depending on analysis type.  
 Example:  
Dataset with missing sales values:  
Customer  Sales  
A 200  
B  
C 400  
→ Replace blank with average (300).  
Q5. Explain the difference between VLOOKUP, HLOOKUP, XLOOKUP, and INDEX -MATCH. 
Which one do you prefer and why?  
VLOOKUP (Vertical Lookup):  
• Searches for a value in the first column  of a table and returns a value from another 
column in the same row.  
• Syntax: =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])  
• Limitation: Cannot look to the left , requires sorted data for approximate matches, 
performance issues on large datasets.  
HLOOKUP (Horizontal Lookup):  

• Similar to VLOOKUP, but searches in the first row and returns a value from a different 
row.  
• Rarely used compared to VLOOKUP.  
INDEX-MATCH (Combination):  
• MATCH finds the position of a value.  
• INDEX returns the value at that position.  
• Together, =INDEX(return_range, MATCH(lookup_value, lookup_range, 0)) . 
• More flexible: can look left, dynamic column selection, faster with large datasets.  
XLOOKUP (Modern Excel):  
• Replaces VLOOKUP & HLOOKUP.  
• Syntax: =XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], 
[match_mode], [search_mode])  
• Can look left & right, handles errors, easier to read.  
 Preferred:  
• If available → XLOOKUP  (cleaner, powerful, handles errors).  
• If older Excel → INDEX-MATCH (flexible & robust).  
 
Q6. What is the difference between absolute, relative, and mixed cell referencing?  
When copying formulas in Excel, cell references behave differently:  
• Relative Reference ( A1) 
o Changes when copied across rows/columns.  
o Example: =B2+C2 → if copied down, becomes =B3+C3. 
• Absolute Reference ( $A$1) 
o Remains fixed, doesn’t change when copied.  
o Example: =$B$2+$C$2  → always refers to row 2, column B and C.  
• Mixed Reference ( $A1 or A$1) 
o Partially fixed (either column or row locked).  
o Example:  
▪ $A1: Column locked, row changes.  
▪ A$1: Row locked, column changes.  
 Used heavily in financial modeling, reporting templates, and lookup formulas . 
 

Q7. How would you use IF, IFS, and nested IFs to categorize data?  
IF Function:  
• Used for binary conditions (True/False).  
• Example: =IF(Sales>1000, "High", "Low") . 
Nested IF:  
• Multiple conditions using multiple IFs.  
• Example:  
• =IF(Sales>1000,"High", 
•     IF(Sales>500,"Medium" ,"Low")) 
IFS Function (Excel 2016+):  
• Cleaner way to write multiple IF conditions.  
• Example:  
• =IFS(Sales>1000,"High", 
•      Sales>500,"Medium" , 
•      Sales<=500,"Low") 
 Preferred:  Use IFS (readable, scalable). If Excel version doesn’t support it, use nested IF.  
 
Q8. Can you explain the use of TEXT, LEFT, RIGHT, MID, FIND, LEN, TRIM, and 
CONCATENATE (or TEXTJOIN) in data cleaning?  
• TEXT(value, format_text):  Format numbers/dates into text.  
o Example: =TEXT(TODAY(),"DD -MMM-YYYY") → 22-Sep-2025. 
• LEFT(text, num_chars):  Extracts characters from left.  
o Example: =LEFT("Analytics", 4)  → "Anal". 
• RIGHT(text, num_chars):  Extracts characters from right.  
o Example: =RIGHT("Analytics", 4)  → "tics". 
• MID(text, start_num, num_chars):  Extracts from middle.  
o Example: =MID("Analytics", 2, 4)  → "naly". 
• FIND(find_text, within_text):  Finds position of a character.  
o Example: =FIND("@","email@test.com")  → 6. 
• LEN(text):  Returns length of string.  
o Example: =LEN("Excel")  → 5. 

• TRIM(text):  Removes extra spaces (keeps single space).  
o Example: =TRIM(" Hello World ")  → "Hello World" . 
• CONCATENATE / TEXTJOIN:  Joins multiple strings.  
o =CONCATENATE("Excel"," ","Analytics")  → "Excel Analytics" . 
o =TEXTJOIN(", ", TRUE, A1:A3)  → "A, B, C". 
 These are key in data cleaning, parsing emails, formatting IDs, splitting text fields, etc.  
 
Q9. How does the MATCH function work? Give an example with INDEX.  
MATCH Function:  
• Returns the position of a value in a range.  
• Syntax: =MATCH(lookup_value, lookup_array, [match_type])  
o 0 = exact match  
o 1 = less than (array must be sorted ascending)  
o -1 = greater than (array must be sorted descending)  
Example:  
Dataset:  
A2:A6 = {"Apple","Banana" ,"Mango","Orange","Grapes" } 
Formula:  
=MATCH("Mango", A2:A6, 0)  → 3 (because Mango is the 3rd item).  
Using with INDEX:  
• INDEX returns a value from a position.  
• =INDEX(A2:A6, MATCH("Mango", A2:A6, 0))  → "Mango". 
 Together, INDEX -MATCH = more flexible alternative to VLOOKUP.  
 
Q10. When would you use OFFSET and INDIRECT functions?  
OFFSET(reference, rows, cols, [height], [width]):  
• Returns a reference to a cell range, offset from a starting point.  
• Example: =SUM(OFFSET(A1,0,0,5,1))  → sums A1:A5. 
• Used in dynamic ranges, rolling averages, dashboards . 
• Downside: Volatile (slows large workbooks).  
INDIRECT(ref_text, [a1]):  
• Returns a reference from a text string.  

• Example:  
o =INDIRECT("A"&5)  → returns value from A5. 
o =INDIRECT("Sheet2!B2")  → refers to cell B2 in Sheet2.  
• Useful for dynamic references, switching ranges based on input . 
• Downside: Also volatile, breaks if sheet names/columns change.  
 Use Case:  
• OFFSET → when you need moving/dynamic ranges (rolling 12 months sales).  
• INDIRECT  → when reference needs to be dynamic (user selects sheet name from 
dropdown).  
11. How do you calculate percentages, growth rates, and CAGR in Excel?  
  Percentages  
• Formula:  
• Percentage  = (Part / Total) * 100  
• Example:  
If Sales = 500  and Target = 1000 , then:  
=500/1000  → 0.5 → Format as % → 50%  
 
  Growth Rate  
• Formula for growth between two periods:  
• Growth % = (New Value – Old Value) / Old Value * 100  
• Example:  
Sales in 2024 = 1200, Sales in 2023 = 1000  
=(1200-1000)/1000  → 0.2 → 20% growth  
 
  CAGR (Compound Annual Growth Rate)  
• Formula:  
• CAGR = (Ending Value / Beginning Value) ^ ( 1 / Number of Years) - 1 
• Example:  
Beginning Sales (2020) = 500, Ending Sales (2024) = 1000 (4 years)  
=(1000/500)^(1/4) -1 → 0.1892 → 18.92% CAGR  
 CAGR is widely used in financial & business analytics  to measure long -term growth trends.  
 
12. What is the difference between COUNT, COUNTA, COUNTBLANK, and 
COUNTIF/COUNTIFS?  

Function  What it counts  Example  
COUNT  Only numeric values  =COUNT(A1:A10)  → counts numbers only  
COUNTA  All non-empty cells 
(numbers, text, dates, 
etc.)  =COUNTA(A1:A10)  
COUNTBLANK  Empty cells  =COUNTBLANK(A1:A10)  
COUNTIF  Cells meeting a single 
condition  =COUNTIF(A1:A10,">100")  → counts values >100  
COUNTIFS  Cells meeting multiple 
conditions  =COUNTIFS(A1:A10,">100",B1:B10,"East")  → counts 
sales >100 in East region  
 Usage in Analytics : 
• COUNT → number of transactions  
• COUNTA  → total entries (including names, IDs)  
• COUNTBLANK  → missing data check  
• COUNTIF/COUNTIFS  → data segmentation (e.g., sales above target by region)  
 
13. How would you calculate running totals or cumulative sums in Excel?  
Method 1: Simple Formula  
• Suppose you have Sales in Column B  (B2:B10).  
• Running total in C2: 
• =SUM($B$2:B2) 
• Drag down → each row shows cumulative sum up to that point.  
 
Method 2: Using Table (Structured Reference)  
• If your data is in a table named SalesData  with column [Sales]:  
• =SUM(SalesData[Sales]:[ @Sales]) 
• This automatically expands as new rows are added.  
 
Method 3: Using Pivot Table  
• Insert Pivot → Place Date in Rows, Sales in Values.  
• Right-click → Show Values As → Running Total In  → select Date.  

 Useful for tracking cumulative revenue, expenses, orders over time . 
 
14. What is the use of RANK, RANK.EQ, RANK.AVG in analytics?  
Functions  
• RANK / RANK.EQ : Returns the rank of a number in a dataset (equal values get same 
rank).  
• RANK.AVG : If ties exist, assigns the average rank . 
 
Example  
Sales Data: 500, 600, 600, 700  
Value  RANK.EQ  RANK.AVG  
700  1 1 
600  2 2.5 
600  2 2.5 
500  4 4 
 In analytics : 
• Rank products by sales volume  
• Rank students by scores  
• Rank regions by performance  
RANK.EQ  → good when you only care about order  
RANK.AVG  → good for fair ranking when ties exist  
 
15. How can you use SUMPRODUCT for conditional calculations?  
Basic Concept  
• SUMPRODUCT  multiplies arrays and sums the result.  
• It can also act as an alternative to COUNTIFS / SUMIFS . 
 
Example 1: Conditional Sum  
• Sales in B2:B10, Region in C2:C10  
• Find total sales in "East":  
• =SUMPRODUCT ((C2:C10="East")*(B2:B10)) 

 
Example 2: Multiple Conditions  
• Find sales in East region AND Product A : 
• =SUMPRODUCT ((C2:C10="East")*(D2:D10="Product A" )*(B2:B10)) 
 
Example 3: Conditional Count  
• Count how many sales >100 in East region:  
• =SUMPRODUCT ((C2:C10="East")*(B2:B10>100)) 
 
SUMPRODUCT in Analytics  is powerful for:  
• Weighted averages  
• Multi-condition sums/counts  
• Replacing array formulas  
 
16. Explain how Pivot Tables help in data analysis.  
Answer:  
Pivot Tables are one of the most powerful tools in Excel for data analytics. They help in 
summarizing, analyzing, and exploring large datasets  without writing complex formulas.  
• Summarization : Pivot Tables allow you to quickly calculate totals, averages, counts, 
percentages, max/min, etc.  
• Flexibility : You can rearrange rows, columns, and filters (drag & drop) to view data from 
multiple perspectives.  
• Aggregation : Large raw datasets (thousands or even millions of rows) can be 
summarized into meaningful insights.  
• Comparison : Easily compare sales by region, year, or product.  
• Drill Down : You can double -click a value to see the underlying raw data.  
• Dynamic Analysis : Pivot Tables can be connected with slicers and timelines for 
interactive reporting.  
 Example: If you have sales data with columns Date, Region, Product, Sales , you can create a 
Pivot Table to see total sales per region , average sales per product , or sales trend per year  — 
all without writing formulas.  
 
17. How do you create a Pivot Table that shows % contribution to total sales by region?  

Answer:  
Steps:  
1. Select your dataset (e.g., columns: Region, Sales ). 
2. Go to Insert → PivotTable  → Place in a new worksheet.  
3. In the Pivot Table Field List:  
o Drag Region into the Rows area.  
o Drag Sales into the Values area.  
o By default, it will show "Sum of Sales".  
4. Right-click on any value in the Pivot Table → Select Show Values As → % of Grand Total . 
5. Now each region’s sales will be shown as a percentage of the total sales . 
Example:  
If Total Sales = ₹1,000,000 and North = ₹250,000, then Pivot Table will show North = 25% . 
 
18. What are calculated fields and calculated items in Pivot Tables?  
Answer:  
• Calculated Field  
o A formula you create using existing fields in the dataset.  
o It adds a new field in the Pivot Table  that is not present in the source data.  
o Example: If you have Sales and Cost fields, you can create a Calculated Field:  
o Profit = Sales – Cost  
o Profit Margin = (Sales – Cost) / Sales  
o This appears as an additional column in the Pivot Table.  
• Calculated Item  
o Works within a field (row or column) to create a new item based on existing 
items.  
o Example: If you have Region field with North, South, East, West , you can create a 
Calculated Item : 
o North + East = Combined Sales of North & East  
o Useful when you want to group or create custom categories directly in the Pivot.  
Difference:  
• Calculated Field = Works on fields (columns in dataset).  
• Calculated Item = Works on items (row/column values inside Pivot Table).  

 
19. How can you create a dynamic dashboard in Excel with slicers & pivot charts?  
 Answer:  
Steps to create a dynamic Excel dashboard : 
1. Prepare Data : Clean and structure the dataset in tabular form.  
2. Insert Pivot Tables : Create multiple Pivot Tables for different metrics (e.g., Sales by 
Region, Sales Trend, Top Products).  
3. Add Pivot Charts : Convert Pivot Tables into charts (bar, line, pie, etc.).  
4. Insert Slicers : 
o Go to Insert → Slicer . 
o Add slicers for common filters like Region, Year, or Product. 
o Connect slicers to multiple Pivot Tables (via Report Connections).  
5. Arrange Layout : Place Pivot Charts & slicers neatly on one sheet.  
6. Format & Style : Use consistent colors, labels, and add KPIs (cards with total sales, 
profit, etc.).  
7. Make it Interactive : Users can click slicers (e.g., Region = North) and all charts update 
dynamically.  
Example:  
A dashboard could show Total Sales, Top 5 Products, Yearly Trend, Sales by Region , with 
slicers for Year and Region. 
 
20. How would you refresh Pivot Tables automatically when data updates?  
 Answer:  
There are multiple ways:  
1. Manual Refresh : Right-click on Pivot Table → Refresh. 
2. Refresh All : Go to Data → Refresh All  (updates all Pivot Tables in the workbook).  
3. Auto Refresh on File Open : 
o Go to Pivot Table Options → Data tab → Refresh data when opening the file . 
4. Auto Refresh via VBA Macro  (for dynamic automation):  
5. Private Sub Workbook_Open()  
6.     ThisWorkbook.RefreshAll  
7. End Sub  
o This will refresh all Pivot Tables automatically when the file is opened.  

8. Use Power Query Connection : If Pivot Table is based on Power Query, set it to auto -
refresh when opening the file.  
 Best Practice: For dashboards, enable Refresh All on Open  so users always see the latest 
data.  
21. Suppose you have a dataset with Sales, Date, and Region. How would you find the top 5 
performing regions?  
Answer:  
Step 1: Create a Pivot Table  
1. Select your dataset → Insert → PivotTable.  
2. Drag Region to Rows and Sales to Values → set to Sum of Sales . 
Step 2: Sort the Pivot Table  
1. Click the drop -down on Row Labels (Region).  
2. Choose Sort Largest to Smallest  by Sum of Sales.  
Step 3: Show Top 5  
1. In PivotTable → Row Labels → Value Filters → Top 10 → change to Top 5 by Sum of Sales . 
Alternative Formula Approach:  
• Use SUMIF + LARGE: 
=SUMIF(RegionRange, RegionName, SalesRange)  'Sum sales per region  
• Then use LARGE to find top 5 totals:  
=LARGE(SumSalesRange, 1)  'Top 1  
=LARGE(SumSalesRange, 2)  'Top 2  
 Result: You can either have a Pivot Table sorted or a formula -based top 5 list.  
 
22. You have sales data for multiple years. How would you calculate Year -over-Year (YoY) 
growth in Excel?  
Answer:  
Step 1: Organize Data  
• Ensure you have Year and Sales columns.  
Year  Sales  
2022  50000  
2023  60000  
Step 2: Apply YoY Growth Formula  

YoY Growth (%) = (Current Year Sales - Previous Year Sales) / Previous Year Sales * 100  
Example:  
If B3 = 2023 sales, B2 = 2022 sales  
=(B3-B2)/B2  
• Format as Percentage . 
Step 3: Fill Down  for all years.  
Alternative:  If data is in Pivot Table , you can use Show Values As → % Difference From → 
Previous Year . 
Result: Column showing YoY growth for each year.  
 
23. How do you highlight the top 10% of sales reps using conditional formatting?  
Answer:  
Step 1: Select the Sales Column  
• Highlight the range, e.g., C2:C100 . 
Step 2: Apply Conditional Formatting  
1. Home → Conditional Formatting → Top/Bottom Rules → Top 10%…  
2. Choose formatting (e.g., green fill).  
Alternative using Formula (for more control):  
=C2 >= PERCENTILE($C$2:$C$100, 0.9)  
• Apply this formula in Conditional Formatting → New Rule → Use a formula . 
 Result: Excel automatically highlights the top 10% of sales reps.  
 
24. If you get two datasets (Orders & Customers), how would you merge them in Excel?  
Answer:  
Step 1: Identify the common key  
• Usually CustomerID  or OrderID. 
Step 2: Use VLOOKUP or XLOOKUP (Excel 365)  
Example using VLOOKUP:  
=VLOOKUP(A2, Customers!$A$2:$D$1000, 2, FALSE)  
• A2 = CustomerID in Orders  
• Customers!A2:D1000  = Customers table  
• 2 = column index to fetch  

• FALSE = exact match  
Step 3: Alternative with Power Query (Recommended for large datasets)  
1. Data → Get & Transform → Get Data → From Table/Range  
2. Load Orders & Customers tables  
3. Merge Queries → select matching key → choose Left Join / Inner Join  
Result: Orders table now has customer details appended.  
 
25. A dataset has outliers (e.g., one order with extremely high sales). How would you detect 
and handle it?  
Answer:  
Step 1: Detect Outliers  
Method 1: Using Conditional Formatting  
• Select Sales column → Conditional Formatting → Color Scales / Top/Bottom Rules / 
Highlight values above threshold.  
Method 2: Using Z -Score  
Z = (Value - Mean) / Standard Deviation  
• Calculate mean and standard deviation of Sales:  
=AVERAGE(C2:C100)  
=STDEV.P(C2:C100)  
• Flag if Z > 3 or < -3. 
Method 3: Using IQR (Interquartile Range)  
1. Calculate Q1 and Q3  
=QUARTILE.INC(C2:C100,1)  'Q1  
=QUARTILE.INC(C2:C100,3)  'Q3  
2. Calculate IQR  
=Q3-Q1 
3. Determine outlier bounds  
Lower = Q1 - 1.5*IQR  
Upper = Q3 + 1.5*IQR  
• Any value outside Lower or Upper is an outlier.  
Step 2: Handle Outliers  
• Option 1: Remove them (if data entry error)  

• Option 2: Cap or replace with percentile values  
• Option 3: Analyze separately (if valid but extreme)  
 Result: Outliers are identified and treated appropriately without affecting analysis.  
 
26. What are array formulas and dynamic arrays? (Examples: FILTER, SORT, UNIQUE, 
SEQUENCE)  
Answer:  
• Array Formulas : 
These are formulas that can perform multiple calculations on one or more items in an 
array (range of cells) and return single or multiple results . 
o Traditional array formulas require pressing Ctrl+Shift+Enter . 
• Dynamic Arrays : 
Excel now has dynamic array functions  (Excel 365 / Excel 2021+) that spill results 
automatically into multiple cells  without Ctrl+Shift+Enter.  
Examples:  
1. UNIQUE – Extract unique values:  
2. =UNIQUE(A2:A10)  
Returns a list of unique values from the range A2:A10.  
3. FILTER – Filter data based on condition:  
4. =FILTER(A2:B20, B2:B20>100)  
Returns all rows where column B has values greater than 100.  
5. SORT – Sort data dynamically:  
6. =SORT(A2:A10, 1, 1)  
Sorts the range A2:A10 in ascending order.  
7. SEQUENCE  – Generate sequential numbers:  
8. =SEQUENCE(5,1,1,1)  
Returns 1,2,3,4,5 vertically.  
Key Point:  Dynamic arrays spill automatically  into adjacent cells, making analysis easier.  
 
27. How do you use Power Query for data cleaning and transformation?  
Answer:  
Power Query  is a self-service ETL tool  in Excel used to extract, transform, and load data  
(ETL). It is very useful for data cleaning, reshaping, and preparing data for analysis.  

Steps to use Power Query:  
1. Load Data:  
o Go to Data → Get Data → From File / Database / Web . 
o Import CSV, Excel, SQL, etc.  
2. Clean & Transform Data:  
o Remove unwanted columns / duplicates.  
o Change data types (text, number, date).  
o Split columns (Text to Columns).  
o Trim spaces and handle missing values.  
o Merge or append tables.  
3. Add Calculations:  
o Create custom columns using M formulas.  
o Group data for aggregation (sum, average).  
4. Load Data Back:  
o Click Close & Load  to load the cleaned/transformed data to Excel table or Data 
Model.  
Example:  
• Remove duplicates from a sales dataset:  
o In Power Query → Select column → Remove duplicates.  
• Filter sales greater than 1000:  
o Home → Filter Rows → Greater than 1000.  
Key Point:  Power Query makes repetitive data cleaning automated and repeatable . 
 
28. What is the difference between Excel and Power Pivot / Data Model?  
Feature  Excel  Power Pivot / Data Model  
Data Capacity  Limited (~1 million rows per 
sheet)  Can handle millions of rows efficiently  
Data Analysis  Basic formulas, PivotTables  Advanced calculations using DAX (Data 
Analysis Expressions)  
Relationships  Usually single table per 
sheet  Can create relationships between multiple 
tables  

Feature  Excel  Power Pivot / Data Model  
Data Storage  Stored in worksheet cells  Stored in a compressed Data Model  
Performance  Slower with large datasets  Optimized for large data, faster calculations  
Advanced 
Analytics  Limited to formulas and 
PivotTables  Advanced KPIs, Time Intelligence, complex 
aggregations  
Example Use Case:  
• Excel: PivotTable showing total sales by region.  
• Power Pivot: Multiple tables (Sales, Products, Customers), with relationships, advanced 
measures like Year-over-Year growth . 
 
29. Can you write a formula to extract unique customers who purchased more than 3 
times?  
Answer (using dynamic arrays in Excel 365/2021):  
Assume Customer Names in A2:A100  and Purchase Count in B2:B100 : 
=FILTER(UNIQUE(A2:A100), B2:B100>3)  
Explanation:  
• UNIQUE(A2:A100)  → Returns unique customer names.  
• FILTER(..., B2:B100>3)  → Returns only those whose purchase count > 3.  
Alternative using Pivot Table:  
• Insert Pivot Table → Rows: Customer → Values: Count of Purchases → Filter customers > 
3. 
 
30. How would you handle large datasets (1M+ rows) in Excel?  
Answer:  
Excel alone cannot efficiently handle millions of rows , so we need optimization:  
1. Use Power Query:  
o Load large datasets into Power Query → clean & filter before loading.  
2. Use Data Model / Power Pivot:  
o Load data into Data Model  instead of sheet.  
o Use DAX measures  for calculations instead of Excel formulas.  
3. Avoid Volatile Functions:  
o Functions like OFFSET, INDIRECT , NOW, TODAY slow down Excel.  

4. Use Tables & PivotTables:  
o Convert data into tables → use PivotTables for aggregation instead of formulas.  
5. Split Data / Use External Database:  
o Store in SQL Server / Access / CSV  → connect via Power Query instead of 
loading all rows in Excel.  
Key Point:  For analytics, Power Pivot + Power Query + PivotTables  is the preferred way to 
handle large datasets efficiently.  
 


<h1>Excel Interview Questions</h1>

<h2>Basic & Data Handling</h2>
<ol>
<li>What are the differences between Excel Tables vs Normal Ranges? Why use tables in analytics?</li>
<li>How do you remove duplicates from a dataset?</li>
<li>How would you clean messy data (extra spaces, text to columns, inconsistent formats)?</li>
<li>How do you handle missing values in Excel data?</li>
</ol>

<h2>Formulas & Functions</h2>
<ol start="5">
<li>Explain the difference between VLOOKUP, HLOOKUP, XLOOKUP, and INDEX-MATCH. Which one do you prefer and why?</li>
<li>What is the difference between absolute, relative, and mixed cell referencing?</li>
<li>How would you use IF, IFS, and nested IFs to categorize data?</li>
<li>Can you explain the use of TEXT, LEFT, RIGHT, MID, FIND, LEN, TRIM, and CONCATENATE (or TEXTJOIN) in data cleaning?</li>
<li>How does the MATCH function work? Give an example with INDEX.</li>
<li>When would you use OFFSET and INDIRECT functions?</li>
</ol>

<h2>Analytical & Business Functions</h2>
<ol start="11">
<li>How do you calculate percentages, growth rates, and CAGR in Excel?</li>
<li>What is the difference between COUNT, COUNTA, COUNTBLANK, and COUNTIF/COUNTIFS?</li>
<li>How would you calculate running totals or cumulative sums in Excel?</li>
<li>What is the use of RANK, RANK.EQ, RANK.AVG in analytics?</li>
<li>How can you use SUMPRODUCT for conditional calculations?</li>
</ol>

<h2>Pivot Tables & Reporting</h2>
<ol start="16">
<li>Explain how Pivot Tables help in data analysis.</li>
<li>How do you create a Pivot Table that shows % contribution to total sales by region?</li>
<li>What are calculated fields and calculated items in Pivot Tables?</li>
<li>How can you create a dynamic dashboard in Excel with slicers & pivot charts?</li>
<li>How would you refresh Pivot Tables automatically when data updates?</li>
</ol>

<h2>Scenario-based Questions</h2>
<ol start="21">
<li>Suppose you have a dataset with Sales, Date, and Region. How would you find the top 5 performing regions?</li>
<li>You have sales data for multiple years. How would you calculate Year-over-Year growth in Excel?</li>
<li>How do you highlight the top 10% of sales reps using conditional formatting?</li>
<li>If you get two datasets (Orders & Customers), how would you merge them in Excel?</li>
<li>A dataset has outliers (e.g., one order with extremely high sales). How would you detect and handle it?</li>
</ol>

<h2>Advanced Excel (Optional, for Analyst Role)</h2>
<ol start="26">
<li>What are array formulas and dynamic arrays? (Examples: FILTER, SORT, UNIQUE, SEQUENCE)</li>
<li>How do you use Power Query for data cleaning and transformation?</li>
<li>What is the difference between Excel and Power Pivot / Data Model?</li>
<li>Can you write a formula to extract unique customers who purchased more than 3 times?</li>
<li>How would you handle large datasets (1M+ rows) in Excel?</li>
</ol>

<h3>1. What are the differences between Excel Tables vs Normal Ranges?</h3>
<p><b>Normal Range =</b> Just a collection of cells.<br>
<b>Excel Table =</b> Structured dataset with special features.</p>

<table>
<tr><th>Feature</th><th>Normal Range</th><th>Excel Table</th></tr>
<tr><td>Formatting</td><td>Needs manual formatting</td><td>Auto formatted (banded rows, filters)</td></tr>
<tr><td>Dynamic Range</td><td>Fixed</td><td>Expands automatically</td></tr>
<tr><td>References</td><td>Cell refs (A2:C100)</td><td>Structured refs (Sales[Amount])</td></tr>
<tr><td>Sorting/Filtering</td><td>Manual</td><td>Built-in</td></tr>
<tr><td>Formulas</td><td>Copy down manually</td><td>Auto-fill entire column</td></tr>
<tr><td>Pivot Tables</td><td>Manual range selection</td><td>Connects with table name</td></tr>
<tr><td>Readability</td><td>Hard to interpret</td><td>Column names used directly</td></tr>
</table>

<ul>
<li><b>Dynamic:</b> Updates charts/pivots automatically.</li>
<li><b>Readable formulas:</b> Structured references.</li>
<li><b>Integrity:</b> Prevents broken formulas.</li>
<li><b>Dashboards:</b> Works with slicers.</li>
</ul>

<p><b>Example:</b><br>=SUM(Sales[Amount]) instead of =SUM(C2:C1000)</p>

<h3>2. How do you remove duplicates?</h3>
<ol>
<li>Remove Duplicates tool → Data → Remove Duplicates</li>
<li>Advanced Filter → Unique records only</li>
<li>Formula flag: <b>=COUNTIF($A$2:A2,A2)&gt;1</b></li>
<li>Conditional Formatting → Highlight Duplicates</li>
</ol>

<h3>3. Cleaning messy data</h3>
<ul>
<li>Remove spaces → <b>=TRIM(A2)</b></li>
<li>Standardize case → <b>=UPPER/LOWER/PROPER</b></li>
<li>Split columns → Text to Columns</li>
<li>Fix dates/numbers → <b>=TEXT</b>, <b>=VALUE</b></li>
<li>Remove non-printable → <b>=CLEAN(A2)</b></li>
<li>Combine → <b>=TEXTJOIN(" ",TRUE,A2:B2)</b></li>
</ul>

<h3>4. Handling missing values</h3>
<ol>
<li>Identify: <b>=COUNTBLANK(A2:A100)</b></li>
<li>Replace with 0, average, or forward-fill</li>
<li>Flag: <b>=IF(A2="","Missing","OK")</b></li>
</ol>

<table>
<tr><th>Customer</th><th>Sales</th></tr>
<tr><td>A</td><td>200</td></tr>
<tr><td>B</td><td></td></tr>
<tr><td>C</td><td>400</td></tr>
</table>
<p><b>→ Replace blank with average (300)</b></p>

<h3>5. VLOOKUP vs HLOOKUP vs INDEX-MATCH vs XLOOKUP</h3>
<ul>
<li><b>VLOOKUP:</b> Search first col → return other col.</li>
<li><b>HLOOKUP:</b> Search row → return from another row.</li>
<li><b>INDEX-MATCH:</b> Match finds position, Index returns value.</li>
<li><b>XLOOKUP:</b> Modern, can look left/right, handles errors.</li>
</ul>

<h3>6. Absolute vs Relative vs Mixed References</h3>
<ul>
<li><b>Relative (A1):</b> Changes when copied</li>
<li><b>Absolute ($A$1):</b> Fixed when copied</li>
<li><b>Mixed ($A1 or A$1):</b> One part locked</li>
</ul>

<h3>7. IF, IFS, Nested IF</h3>
<ul>
<li><b>IF:</b> <code>=IF(Sales&gt;1000,"High","Low")</code></li>
<li><b>Nested IF:</b> Multiple conditions</li>
<li><b>IFS:</b> Cleaner multi-condition</li>
</ul>

<h3>8. Text Functions</h3>
<ul>
<li><b>LEFT/RIGHT/MID:</b> Extract text</li>
<li><b>LEN:</b> Length</li>
<li><b>TRIM:</b> Remove spaces</li>
<li><b>FIND:</b> Find position</li>
<li><b>TEXTJOIN:</b> Combine text</li>
</ul>

<h3>9. MATCH + INDEX</h3>
<p><b>=MATCH("Mango",A2:A6,0)</b> → 3<br>
<b>=INDEX(A2:A6,MATCH("Mango",A2:A6,0))</b> → "Mango"</p>

<h3>10. OFFSET vs INDIRECT</h3>
<ul>
<li><b>OFFSET:</b> Dynamic ranges (rolling totals)</li>
<li><b>INDIRECT:</b> Build refs from text (dynamic sheet)</li>
</ul>

<h1>Excel Interview Questions</h1>

<h2>Basic & Data Handling</h2>
<ol>
<li>What are the differences between Excel Tables vs Normal Ranges? Why use tables in analytics?</li>
<li>How do you remove duplicates from a dataset?</li>
<li>How would you clean messy data (extra spaces, text to columns, inconsistent formats)?</li>
<li>How do you handle missing values in Excel data?</li>
</ol>

<h3>1. Excel Tables vs Normal Ranges</h3>
<p><b>Normal Range =</b> Just a collection of cells.<br>
<b>Excel Table =</b> Structured dataset with special features.</p>

<table>
<tr><th>Feature</th><th>Normal Range</th><th>Excel Table</th></tr>
<tr><td>Formatting</td><td>Manual</td><td>Auto formatted</td></tr>
<tr><td>Dynamic Range</td><td>Fixed</td><td>Expands automatically</td></tr>
<tr><td>References</td><td>A2:C100</td><td>Sales[Amount]</td></tr>
<tr><td>Sorting/Filtering</td><td>Manual</td><td>Built-in</td></tr>
<tr><td>Formulas</td><td>Copy manually</td><td>Auto-fill</td></tr>
<tr><td>Pivot Tables</td><td>Manual range</td><td>Connects by table name</td></tr>
<tr><td>Readability</td><td>Hard to interpret</td><td>Column names readable</td></tr>
</table>

<ul>
<li><b>Dynamic:</b> Charts/pivots auto-update</li>
<li><b>Readable:</b> Structured references</li>
<li><b>Integrity:</b> Prevents formula breaks</li>
<li><b>Dashboards:</b> Works with slicers</li>
</ul>

<p><b>Example:</b><br>=SUM(Sales[Amount]) instead of =SUM(C2:C1000)</p>

<h3>2. Removing Duplicates</h3>
<ol>
<li>Remove Duplicates → Data → Remove Duplicates</li>
<li>Advanced Filter → Unique records only</li>
<li>Formula flag: <b>=COUNTIF($A$2:A2,A2)&gt;1</b></li>
<li>Conditional Formatting → Highlight Duplicates</li>
</ol>
<p><b>Best Practice:</b> Check why duplicates exist before deleting.</p>

<h3>3. Cleaning Messy Data</h3>
<ul>
<li>Remove spaces: <b>=TRIM(A2)</b></li>
<li>Standardize case: <b>=UPPER/LOWER/PROPER</b></li>
<li>Split columns: Text to Columns</li>
<li>Fix formats: <b>=TEXT</b>, <b>=VALUE</b></li>
<li>Remove non-printable: <b>=CLEAN(A2)</b></li>
<li>Combine: <b>=TEXTJOIN(" ",TRUE,A2:B2)</b></li>
</ul>

<h3>4. Handling Missing Values</h3>
<ol>
<li>Identify: <b>=COUNTBLANK(A2:A100)</b></li>
<li>Handle: Replace with 0, average, forward fill, or interpolate</li>
<li>Flag: <b>=IF(A2="","Missing","OK")</b></li>
</ol>
<table>
<tr><th>Customer</th><th>Sales</th></tr>
<tr><td>A</td><td>200</td></tr>
<tr><td>B</td><td></td></tr>
<tr><td>C</td><td>400</td></tr>
</table>
<p><b>→ Replace blank with average (300)</b></p>

---

<h2>Formulas & Functions</h2>
<ol start="5">
<li>VLOOKUP vs HLOOKUP vs INDEX-MATCH vs XLOOKUP</li>
<li>Absolute vs Relative vs Mixed References</li>
<li>IF, IFS, Nested IF</li>
<li>Text Functions</li>
<li>MATCH with INDEX</li>
<li>OFFSET vs INDIRECT</li>
</ol>

<h3>5. Lookup Functions</h3>
<ul>
<li><b>VLOOKUP:</b> Vertical lookup, limited (no left lookup).</li>
<li><b>HLOOKUP:</b> Horizontal lookup, rarely used.</li>
<li><b>INDEX-MATCH:</b> Flexible, faster, looks both ways.</li>
<li><b>XLOOKUP:</b> Modern, supports left/right, errors handled.</li>
</ul>

<h3>6. Cell References</h3>
<ul>
<li><b>Relative (A1):</b> Changes when copied</li>
<li><b>Absolute ($A$1):</b> Fixed when copied</li>
<li><b>Mixed ($A1 or A$1):</b> Row or column locked</li>
</ul>

<h3>7. IF Functions</h3>
<ul>
<li><b>IF:</b> <code>=IF(Sales&gt;1000,"High","Low")</code></li>
<li><b>Nested IF:</b> Multiple conditions</li>
<li><b>IFS:</b> Cleaner syntax for multiple conditions</li>
</ul>

<h3>8. Text Functions</h3>
<ul>
<li><b>LEFT/RIGHT/MID:</b> Extract text</li>
<li><b>LEN:</b> Length</li>
<li><b>TRIM:</b> Remove spaces</li>
<li><b>FIND:</b> Find position</li>
<li><b>TEXT:</b> Format numbers/dates</li>
<li><b>TEXTJOIN:</b> Combine text</li>
</ul>

<h3>9. MATCH with INDEX</h3>
<p><b>=MATCH("Mango",A2:A6,0)</b> → 3<br>
<b>=INDEX(A2:A6,MATCH("Mango",A2:A6,0))</b> → "Mango"</p>

<h3>10. OFFSET vs INDIRECT</h3>
<ul>
<li><b>OFFSET:</b> Dynamic ranges (rolling totals)</li>
<li><b>INDIRECT:</b> Build references from text</li>
</ul>

---

<h2>Analytical & Business Functions</h2>

<h3>11. Percentages, Growth, CAGR</h3>
<ul>
<li><b>Percentage:</b> (Part/Total)*100</li>
<li><b>Growth:</b> (New–Old)/Old</li>
<li><b>CAGR:</b> (End/Start)^(1/Years) - 1</li>
</ul>

<h3>12. COUNT Functions</h3>
<table>
<tr><th>Function</th><th>Description</th><th>Example</th></tr>
<tr><td>COUNT</td><td>Numbers only</td><td>=COUNT(A1:A10)</td></tr>
<tr><td>COUNTA</td><td>Non-empty cells</td><td>=COUNTA(A1:A10)</td></tr>
<tr><td>COUNTBLANK</td><td>Empty cells</td><td>=COUNTBLANK(A1:A10)</td></tr>
<tr><td>COUNTIF</td><td>Single condition</td><td>=COUNTIF(A1:A10,">100")</td></tr>
<tr><td>COUNTIFS</td><td>Multiple conditions</td><td>=COUNTIFS(A1:A10,">100",B1:B10,"East")</td></tr>
</table>

<h3>13. Running Totals</h3>
<ul>
<li><b>Formula:</b> =SUM($B$2:B2)</li>
<li><b>Table:</b> Structured refs</li>
<li><b>Pivot:</b> Show Values As → Running Total</li>
</ul>

<h3>14. Ranking Functions</h3>
<table>
<tr><th>Value</th><th>RANK.EQ</th><th>RANK.AVG</th></tr>
<tr><td>700</td><td>1</td><td>1</td></tr>
<tr><td>600</td><td>2</td><td>2.5</td></tr>
<tr><td>600</td><td>2</td><td>2.5</td></tr>
<tr><td>500</td><td>4</td><td>4</td></tr>
</table>

<h3>15. SUMPRODUCT</h3>
<ul>
<li>Conditional sum: <code>=SUMPRODUCT((C2:C10="East")*(B2:B10))</code></li>
<li>Multi-condition: <code>=SUMPRODUCT((C2:C10="East")*(D2:D10="Product A")*(B2:B10))</code></li>
<li>Conditional count: <code>=SUMPRODUCT((C2:C10="East")*(B2:B10&gt;100))</code></li>
</ul>

---

<h2>Pivot Tables & Reporting</h2>

<h3>16. Pivot Tables</h3>
<ul>
<li>Summarization, comparison, aggregation</li>
<li>Drill down, slicers, dynamic dashboards</li>
</ul>

<h3>17. % Contribution by Region</h3>
<ul>
<li>Create Pivot → Region in Rows, Sales in Values</li>
<li>Right-click Sales → Show Values As → % of Grand Total</li>
</ul>

<h3>18. Calculated Field vs Item</h3>
<ul>
<li><b>Field:</b> New formula field (e.g., Profit=Sales–Cost)</li>
<li><b>Item:</b> New category within a field (e.g., North+East)</li>
</ul>

<h3>19. Dynamic Dashboard</h3>
<ul>
<li>Clean data → Pivot Tables → Pivot Charts</li>
<li>Add Slicers (Region, Year, Product)</li>
<li>Arrange & format</li>
</ul>

<h3>20. Auto Refresh</h3>
<ul>
<li>Manual → Refresh All</li>
<li>On file open → Pivot Options</li>
<li>VBA: <code>Workbook_Open → RefreshAll</code></li>
</ul>

---

<h2>Scenario-Based Questions</h2>

<h3>21. Top 5 Regions</h3>
<ul>
<li>Pivot → Sort Sales Desc → Value Filters → Top 5</li>
<li>Alternative: SUMIF + LARGE</li>
</ul>

<h3>22. YoY Growth</h3>
<ul>
<li>Formula: <code>=(Current–Previous)/Previous</code></li>
<li>Pivot → Show Values As → % Difference From → Previous Year</li>
</ul>

<h3>23. Highlight Top 10%</h3>
<ul>
<li>Conditional Formatting → Top 10%</li>
<li>Formula: <code>=C2&gt;=PERCENTILE($C$2:$C$100,0.9)</code></li>
</ul>

<h3>24. Merge Orders & Customers</h3>
<ul>
<li><b>VLOOKUP/XLOOKUP:</b> =VLOOKUP(A2,Customers!A:D,2,FALSE)</li>
<li><b>Power Query:</b> Merge Queries with CustomerID</li>
</ul>

<h3>25. Outliers</h3>
<ul>
<li>Detect with Conditional Formatting, Z-Score, or IQR</li>
<li>Handle by removing, capping, or analyzing separately</li>
</ul>

---

<h2>Advanced Excel</h2>

<h3>26. Array & Dynamic Arrays</h3>
<ul>
<li><b>UNIQUE:</b> =UNIQUE(A2:A10)</li>
<li><b>FILTER:</b> =FILTER(A2:B20,B2:B20&gt;100)</li>
<li><b>SORT:</b> =SORT(A2:A10,1,1)</li>
<li><b>SEQUENCE:</b> =SEQUENCE(5,1,1,1)</li>
</ul>

<h3>27. Power Query</h3>
<ul>
<li>Load → Clean (remove columns, fix types) → Transform (split, group) → Load back</li>
<li>Repeatable & automated ETL process</li>
</ul>

<h3>28. Excel vs Power Pivot</h3>
<table>
<tr><th>Feature</th><th>Excel</th><th>Power Pivot</th></tr>
<tr><td>Capacity</td><td>~1M rows</td><td>Millions</td></tr>
<tr><td>Analysis</td><td>Formulas, Pivots</td><td>DAX, advanced</td></tr>
<tr><td>Relationships</td><td>Single table</td><td>Multiple tables</td></tr>
<tr><td>Performance</td><td>Slower</td><td>Optimized</td></tr>
</table>

<h3>29. Unique Customers >3 Purchases</h3>
<p><b>=FILTER(UNIQUE(A2:A100),B2:B100&gt;3)</b></p>
<ul>
<li>Alternative: Pivot → Customer in Rows → Count Purchases → Filter >3</li>
</ul>

<h3>30. Handling Large Datasets</h3>
<ul>
<li>Use Power Query to clean before load</li>
<li>Use Data Model & DAX</li>
<li>Avoid volatile functions</li>
<li>Use Pivots instead of formulas</li>
<li>Store in SQL/Access & connect</li>
</ul>

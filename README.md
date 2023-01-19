# Office-Supply-Sales-Excel-Dashboard-Project-
Office Supply Sales (Project) Dashboard using Advance Excel Functions, Formulas, Pivot Tables, Graphs, Charts and Filters

<h3>Functions and Formulas (Quick Summary and Pivot tables):</h3>
- Sum 
- SumIF 
- COUNTNA 
- Max 
- Index-Match (Combined)
- Sort Function
- Remove Duplicate Function
- Custom Number Format


<h3>Quick Summary:</h3>
<h6>In the Quick Summary, we described the following **KEY FIGURES**  using Excel functions and formulas.</h6>

- Total Revenue

- Highest revenue of Region

- Highest representative Revenue

- Total No of Transactions

<h8>**Total Revenue** :  For Total revenue used (SUM) function for total column as SUM(Total) </h8><p></p>
<h8>**Highest revenue of Region** : </h8>
<p>In 1st step extracted region from region column from Data Sheet, Then remove the duplicates so that we get unique values for regions</p>
<p>In 2nd step calculated the SUM for each region using SUMIF with (Total) column as Sum range</p>
<p>In 3rd step found out Maximum value from step 2 using (Max ) function</p>
<p>In 4th step, found out respective region from unique regions for maximum value from step 3 using (Index-Match) Formula </p>

<h8>**Highest representative Revenue** :  Same steps repeated as of Highest Revenue of Region </h8><p></p>
<h8>**Total No of Transactions:** :  Used COUNTNA(Total) to count Total number of transactions </h8><p></p>

<h2>Graphs/Charts:</h2>
<h6>Created the following graphs visualizations for Visual Analysis:</h6>


-<h8>**Bar Chart**    : Top 3 Items Sold </h8><p></p>
-<h8>**Column Chart** : Sales By Representatives </h8><p></p>
-<h8>**Column Chart** : Sales By Region  </h8><p></p>
-<h8>**Line Chart**   : Sales Trend by Months </h8><p></p>
-<h8>**Column Chart** : Deal Count by Revenue  </h8>
<p>Created the custom group (buckets) of revene (Total) column by 500 gap in each year	
so we can see how much revenue we got in every bucket.</p>

<h3>Slicers:</h3>
<h6>Finally created the following slicers and then attached to each report to filter out data and make it Dynamic:</h6>

- Months Filter 
- Quarter Filter 
- Region Filter

<p align="center">
  <i>All charts are dynamic using Filters.</i>
</P>

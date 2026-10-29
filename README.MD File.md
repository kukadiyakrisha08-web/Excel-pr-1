# 📊 PR. 1 --- Fundamental Booster

### Advanced Excel Practical Project \| BCA

```{=html}
<p align="center">
```
`<strong>`{=html}From raw data to meaningful insights --- a
formula-driven Excel workbook.`</strong>`{=html}
```{=html}
</p>
```
![Project](https://img.shields.io/badge/Project-Fundamental%20Booster-17365D?style=for-the-badge)
![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Worksheets](https://img.shields.io/badge/Worksheets-5-087E8B?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)

> **Student / Project Prepared By:** Krisha Kukadiya\
> **Faculty Guide:** Prof. Girish Gondaliya\
> **Project Title:** PR. 1 --- Fundamental Booster\
> **Academic Area:** BCA --- Spreadsheet / Excel Practical\
> **Deliverable:** `PR1_Fundamental_Booster_5_Sheets.xlsx`

------------------------------------------------------------------------

## 📑 Table of Contents

1.  [Project Overview](#-project-overview)
2.  [Objectives](#-objectives)
3.  [Workbook Architecture](#-workbook-architecture)
4.  [Worksheet-by-Worksheet Details](#-worksheet-by-worksheet-details)
5.  [Function Coverage Matrix](#-function-coverage-matrix)
6.  [Key Calculations and Expected
    Outputs](#-key-calculations-and-expected-outputs)
7.  [How to Open and Use](#-how-to-open-and-use)
8.  [Validation and Quality
    Assurance](#-validation-and-quality-assurance)
9.  [Viva Preparation](#-viva-preparation)
10. [Limitations and Future
    Enhancements](#-limitations-and-future-enhancements)
11. [Conclusion and Credits](#-conclusion-and-credits)

------------------------------------------------------------------------

## 🎯 Project Overview

**Fundamental Booster** is a hands-on Microsoft Excel practical project
that demonstrates spreadsheet fundamentals, logical reasoning,
conditional aggregation, lookup techniques, text processing, date
calculations, mathematical rounding, and dynamic references.

The workbook contains **five interconnected learning areas** with
realistic sample datasets. It emphasizes **working formulas rather than
static answers**, enabling recalculation when input data changes.

**Core workflow**

``` text
Enter / Update Data
        │
        ▼
Apply Excel Formulas
        │
        ▼
Validate Calculations
        │
        ▼
Analyze Results & Lookups
        │
        ▼
Present Professional Workbook
```

## 🚀 Objectives

-   Demonstrate **relative, absolute, and mixed cell references**.
-   Apply **IF, nested IF, AND, and OR** for decision-making.
-   Perform **COUNTIFS, SUMIFS, and AVERAGEIFS** analysis.
-   Retrieve information with **VLOOKUP, XLOOKUP, XMATCH, INDEX, and
    MATCH**.
-   Work with **LEFT, FIND, UPPER, LOWER, and TEXT**.
-   Use **INDIRECT and OFFSET** for flexible cell references.
-   Analyze **dates, ages, service periods, and rounding operations**.
-   Present structured data using consistent headings, borders, dates,
    and currency formats.

## 🗂️ Workbook Architecture

  ------------------------------------------------------------------------
  \#                Worksheet         Primary Purpose    Data Volume
  ----------------- ----------------- ------------------ -----------------
  01                **Student Data**  Academic           15 student
                                      performance,       records
                                      grading, student   
                                      lookup             

  02                **Sales Data**    Sales, discounts,  15 sales
                                      regional/product   transactions
                                      analysis           

  03                **Product Data**  Inventory          10 products
                                      valuation,         
                                      eligibility,       
                                      product search     

  04                **Employee Data** Salary, age,       10 employees
                                      service, name      
                                      processing         

  05                **Formula         Dedicated          Multiple worked
                    Practice**        demonstrations of  examples
                                      reference and      
                                      utility functions  
  ------------------------------------------------------------------------

> **Note:** The fifth worksheet is a supplementary formula-demonstration
> sheet designed to consolidate practical topics; it should not be
> described as an explicitly mandated fifth worksheet unless confirmed
> by the faculty's original instructions.

## 📚 Worksheet-by-Worksheet Details

### 01 --- Student Data 🎓

**Columns:** Student ID, Full Name, Math Marks, Science Marks, Average,
Grade, Pass/Fail, Top Performer.

**What it demonstrates:** - `AVERAGE` computes the mean of Math and
Science marks. - Nested `IF` assigns **A / B / C / Fail** based on
average marks. - `IF(AND(...))` checks subject-level pass criteria and
top-performer status. - `COUNTIFS` counts students meeting multiple
conditions. - `AVERAGEIFS` calculates averages for specific grade/pass
groups. - `VLOOKUP` retrieves student details using Student ID. -
`FILTER` lists students whose average exceeds **80**.

**Representative formulas:**

``` excel
=AVERAGE(C5:D5)
=IF(E5>=80,"A",IF(E5>=60,"B",IF(E5>=40,"C","Fail")))
=IF(AND(C5>=40,D5>=40),"Pass","Fail")
=COUNTIFS(C5:C19,">60",D5:D19,">60")
=VLOOKUP(K16,$A$5:$H$19,2,FALSE)
=FILTER(A5:H19,E5:E19>80,"No Students")
```

### 02 --- Sales Data 💼

**Columns:** Sale ID, Date, Salesperson, Region, Product, Quantity, Unit
Price, Total Sales, Discount, Net Sales.

**Business logic:** - **Total Sales** = Quantity × Unit Price. -
**Discount** = 10% for total sales ≥ ₹50,000; 5% for ≥ ₹20,000;
otherwise 0%. - **Net Sales** = Total Sales − Discount. - `SUMIFS`
produces region-wise and product-wise net sales. - `INDEX` + `MATCH`
retrieves a salesperson, region, and net sale by Sale ID. - Date-based
`INDEX` + `MATCH` identifies the first sale in a chosen month. -
`OFFSET` calculates totals and averages across the last five
transactions in the fixed sample range.

**Representative formulas:**

``` excel
=F5*G5
=IF(H5>=50000,H5*10%,IF(H5>=20000,H5*5%,0))
=H5-I5
=SUMIFS($J$5:$J$19,$D$5:$D$19,L5)
=INDEX($C$5:$C$19,MATCH($M$13,$A$5:$A$19,0))
=SUM(OFFSET(J5,COUNT(J5:J19)-5,0,5,1))
```

### 03 --- Product Data 📦

**Columns:** Product Code, Product Name, Category, Unit Price, Stock
Quantity, Inventory Value, Discount Eligibility.

**What it demonstrates:** - Inventory valuation using **Unit Price ×
Stock Quantity**. - `IF` + `OR` eligibility: **Accessories** category
*or* stock quantity ≥ **50**. - `VLOOKUP` for product name and unit
price. - `XMATCH` to identify the position of a product code.

**Representative formulas:**

``` excel
=D5*E5
=IF(OR(C5="Accessories",E5>=50),"Eligible","Not Eligible")
=VLOOKUP(J5,$A$5:$G$14,4,FALSE)
=XMATCH(J11,$A$5:$A$14,0)
```

### 04 --- Employee Data 👥

**Columns:** Employee ID, Full Name, Department, Salary, Date of Birth,
Date of Joining, Age, Years of Service, First Name, Uppercase Name,
Lowercase Name.

**What it demonstrates:** - `DATEDIF` for age and completed years of
service. - `LEFT` + `FIND` to extract the first name. - `UPPER` and
`LOWER` for text transformation. - `XLOOKUP` for salary retrieval. -
`INDEX` + `MATCH` for employee name and department. - `YEAR`, `MONTH`,
and `TEXT` for joining-date analysis.

**Reference date:** **29 October 2026** is intentionally fixed to make
results reproducible.

**Representative formulas:**

``` excel
=DATEDIF(E5,$N$4,"Y")
=DATEDIF(F5,$N$4,"Y")
=LEFT(B5,FIND(" ",B5)-1)
=UPPER(B5)
=XLOOKUP(N8,$A$5:$A$14,$D$5:$D$14,"Not Found")
=TEXT(F5,"DD-MMM-YYYY")
```

### 05 --- Formula Practice 🧮

A structured demonstration area for:

  -------------------------------------------------------------------------------
  Concept                 Example Formula                 Example Output
  ----------------------- ------------------------------- -----------------------
  Relative Reference      `=A5*B5`                        `1000`

  Absolute Reference      `=E5*$F$4`                      `180`

  Mixed Reference         `=$A12*B$11`                    `10`

  INDIRECT                `=INDIRECT(F12)`                `5`

  OFFSET                  `=OFFSET(A17,2,0)`              `300`

  ROUND                   `=ROUND(D18,2)`                 `125.68`

  CEILING                 `=CEILING(D19,10)`              `130`

  FLOOR                   `=FLOOR(D20,10)`                `120`

  DATE                    `=DATE(2026,10,29)`             `29-Oct-2026`

  YEAR                    `=YEAR(B25)`                    `2026`

  TEXT                    `=TEXT(B25,"MMMM")`             `October`

  Dynamic Cross-Sheet     `=INDIRECT("'"&E24&"'!"&E25)`   `101`
  Reference                                               
  -------------------------------------------------------------------------------

## 🧩 Function Coverage Matrix

  Excel Feature                    Student   Sales   Product   Employee   Practice
  ------------------------------- --------- ------- --------- ---------- ----------
  IF / Nested IF / AND / OR          ✅       ✅       ✅        ---        ---
  COUNTIFS / AVERAGEIFS              ✅       ---      ---       ---        ---
  SUMIFS                             ---      ✅       ---       ---        ---
  VLOOKUP                            ✅       ---      ✅        ---        ---
  XLOOKUP                            ---      ---      ---        ✅        ---
  XMATCH                             ---      ---      ✅        ---        ---
  INDEX + MATCH                      ---      ✅       ---        ✅        ---
  FILTER                             ✅       ---      ---       ---        ---
  LEFT / FIND / UPPER / LOWER        ---      ---      ---        ✅        ---
  INDIRECT                           ---      ---      ---       ---         ✅
  OFFSET                             ---      ✅       ---       ---         ✅
  DATEDIF / YEAR / MONTH / TEXT      ---      ---      ---        ✅         ✅
  ROUND / CEILING / FLOOR            ---      ---      ---       ---         ✅
  Relative / Absolute / Mixed        ---      ---      ---       ---         ✅

## 📈 Key Calculations and Expected Outputs

The following checks use the workbook's included **sample data**.

  Check                                          Expected Result
  -------------------------------------------- -----------------
  Total students                                          **15**
  Students scoring above 60 in both subjects               **8**
  Grade A students                                         **6**
  Students passing both subjects                          **12**
  Students with average strictly above 80                  **4**
  Total net sales                                   **₹433,375**
  First February sale ID                                **S005**
  Sale S006 salesperson                                 **Diya**
  Sale S006 region                                     **South**
  Sale S006 net sales                                **₹42,750**
  Last five net sales total                         **₹160,500**
  Last five net sales average                        **₹32,100**
  Product P004 price                                 **₹12,000**
  Product P007 position                                    **7**
  Employee E005 salary                               **₹60,000**
  Employee E003 department                           **Finance**
  Formula Practice INDIRECT result                         **5**
  Formula Practice OFFSET result                         **300**
  Formula Practice date difference                  **301 days**

> Outputs may change if sample inputs, lookup criteria, or formulas are
> edited.

## 🖥️ How to Open and Use

**Recommended environment:** Microsoft Excel for Microsoft 365 or Excel
for the Web. Modern functions such as `FILTER`, `XLOOKUP`, and `XMATCH`
may not be supported in older versions, including Excel 2016.

1.  Download or open **`PR1_Fundamental_Booster_5_Sheets.xlsx`**.
2.  Open the workbook in **Excel for the Web** or a compatible desktop
    Excel version.
3.  Review each worksheet's input data and formula-generated columns.
4.  Click a formula cell and inspect its expression in the formula bar.
5.  Modify a sample input or lookup ID to observe the updated result.
6.  Confirm that date columns contain actual Excel dates rather than
    text.
7.  Save the final workbook as **`.xlsx`**, retaining formulas.

**Important:** When copying to another workbook, use regular paste to
preserve formulas. **Paste Values** removes the formulas and leaves only
their calculated results.

## 🧪 Validation and Quality Assurance

Use this checklist before submission:

-   [ ] All five worksheet names are correct.
-   [ ] Student grades and pass/fail classifications recalculate.
-   [ ] Student summary counts match the sample dataset.
-   [ ] Sales region totals and product totals both equal **₹433,375**.
-   [ ] Lookup IDs return the correct matching records.
-   [ ] `FILTER`, `XLOOKUP`, and `XMATCH` work in the selected Excel
    version.
-   [ ] Employee DOB and joining dates are recognized as dates.
-   [ ] Fixed reference date is **29-Oct-2026**.
-   [ ] `INDIRECT` and `OFFSET` return their expected sample outputs.
-   [ ] Currency, date formats, headings, and table borders are
    readable.
-   [ ] No unexpected `#N/A`, `#VALUE!`, `#REF!`, `#NAME?`, or `#SPILL!`
    errors remain.
-   [ ] Final `.xlsx` opens and formulas remain editable.

### Troubleshooting

  -----------------------------------------------------------------------
  Issue                   Likely Cause            Recommended Fix
  ----------------------- ----------------------- -----------------------
  `#N/A`                  Lookup ID absent or     Verify the search value
                          mismatched              and lookup range

  `#NAME?`                Function unsupported in Use Microsoft 365 /
                          older Excel             Excel for the Web

  `#VALUE!` in `DATEDIF`  Dates stored as text    Convert source cells
                                                  into real Excel dates

  `#SPILL!` with `FILTER` Output range is blocked Clear the cells where
                                                  results need to expand

  Wrong totals            Inconsistent ranges or  Inspect absolute
                          criteria                references and criteria
                                                  columns
  -----------------------------------------------------------------------

## 🎤 Viva Preparation

**Q1. What is a relative reference?**\
A cell reference that changes when a formula is copied, e.g., `A5*B5`.

**Q2. What is an absolute reference?**\
A fixed reference using `$`, e.g., `$F$4`, that does not move when
copied.

**Q3. What is a mixed reference?**\
A reference with only the column or row fixed, such as `$A12` or `B$11`.

**Q4. Difference between VLOOKUP and XLOOKUP?**\
`VLOOKUP` searches the first column of a table range and uses a column
index; `XLOOKUP` takes separate lookup and return arrays and supports
more flexible searches.

**Q5. What does COUNTIFS do?**\
Counts records satisfying **multiple conditions**.

**Q6. What is the difference between SUMIFS and AVERAGEIFS?**\
`SUMIFS` adds matching values; `AVERAGEIFS` computes the mean of
matching values.

**Q7. Why use INDEX with MATCH?**\
`MATCH` locates the position of a value; `INDEX` retrieves the
corresponding result.

**Q8. What does FILTER do?**\
Returns rows matching a condition and can spill results into multiple
cells.

**Q9. Difference between INDIRECT and OFFSET?**\
`INDIRECT` converts text into a reference; `OFFSET` returns a reference
at a specified displacement from a starting cell.

**Q10. What does DATEDIF with `"Y"` return?**\
The number of **completed years** between two dates.

**Q11. Why use a fixed reference date?**\
It ensures age and service calculations are reproducible during
assessment.

**Q12. What is the benefit of formulas over static values?**\
Formulas update calculations automatically when referenced inputs
change.

## ⚠️ Limitations and Future Enhancements

**Current scope** - Uses illustrative datasets for academic
demonstration, not real company records. - Sales rolling analysis uses a
fixed 15-row dataset; extending records may require formula range
updates. - Formula compatibility varies by Excel version. - No VBA
macros, external databases, or live API connections are included.

**Potential enhancements** - Excel Tables with structured references and
automatically expanding ranges. - Interactive dashboards, PivotTables,
and charts. - Data validation dropdowns for lookup IDs and categories. -
Conditional formatting to highlight low stock, failing students, or high
sales. - Power Query import workflows and refreshable datasets.

## 🏁 Conclusion and Credits

The **PR. 1 --- Fundamental Booster** project demonstrates how Excel
formulas can transform raw academic and business data into consistent,
searchable, and meaningful information. The five-sheet workbook provides
hands-on practice in **logical formulas, conditional calculations,
lookups, dynamic references, text functions, and date mathematics**.

### Project Credits

  Role                Name
  ------------------- -----------------------------------------
  **Prepared By**     **Krisha Kukadiya**
  **Faculty Guide**   **Prof. Girish Gondaliya**
  **Project**         **PR. 1 --- Fundamental Booster**
  **Platform**        **Microsoft Excel / Excel for the Web**

------------------------------------------------------------------------

```{=html}
<p align="center">
```
`<strong>`{=html}Made for learning • Built with formulas • Ready for
practical demonstration`</strong>`{=html}
```{=html}
</p>
```

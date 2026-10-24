# Fundamental Booster – Excel Functions Guide

> Documentation for the workbook **`Fundamental_Booster.xlsx`** (generated with Python).

## Objective

The objective of this workbook is to **build strong fundamentals in Excel formulas and functions**
through practical, real-life style datasets. Each task solves a small business question using
a specific function (or a combination of functions). By the end you will be able to:

- Calculate totals, percentages and grades using `SUM`, `AVERAGE`, `IF` (nested).
- Count and average data on conditions using `COUNTIFS` and `AVERAGEIF`.
- Extract matching records using `FILTER`.
- Look up values using `VLOOKUP`, `XLOOKUP`, and `INDEX` + `MATCH`.
- Combine conditions using `AND` / `OR` inside `IF`.
- Summarise data across two conditions using `SUMIFS`.

## Workbook Overview

| Sheet | Schema | Records | Columns | Purpose |
|---|---|---|---|---|
| Student Grade | Schema-1 | 20 students | Student ID, Name, Math, Science, English, Enrollment Date, Total, Grade, Percentage | Marks analysis and grading |
| Sales Data | Schema-2 | 20 sales | Sales ID, Product, Region, Salesperson, Amount, Date | Sales lookups, discounts, summaries |
| Employees Data | Schema-3 | 20 employees | Employee ID, Name, Department, Salary, Joining Date | Employee lookups |

## Task Index

| # | Task | Function(s) | Sheet |
|---|---|---|---|
| 1 | Total marks | `SUM` | Student Grade |
| 2 | Percentage | `SUM` | Student Grade |
| 3 | Grade | `IF`, `AVERAGE` | Student Grade |
| 4 | Count students scoring above 60% | `COUNTIFS` | Student Grade |
| 5 | Count students with Math above 30 | `COUNTIFS` | Student Grade |
| 6 | Average score of students above 60% | `AVERAGEIF` | Student Grade |
| 7 | Extract students scoring above 80% | `FILTER` | Student Grade |
| 8 | Fetch student name from ID | `VLOOKUP` | Student Grade |
| 9 | Above 80 in Maths and Science | `IF`, `AND` | Student Grade |
| 10 | Fetch product price from Sales ID | `VLOOKUP` | Sales Data |
| 11 | Total sales by region and product | `SUMIFS` | Sales Data |
| 12 | Discount amount | `IF` (nested) | Sales Data |
| 13 | Discount eligible / not eligible | `IF`, `AND`, `OR` | Sales Data |
| 14 | Sales value for a salesperson in a month | `INDEX`, `MATCH`, `MONTH` | Sales Data |
| 15 | Fetch employee salary from ID | `XLOOKUP` | Employees Data |
| 16 | Look up full employee details | Lookup | Employees Data |

---

## Sheet 1: Student Grade

### Q1. Calculate Total Marks

**Sheet:** `Student Grade`  |  **Cell(s):** `G4:G23`

**Formula used:**

```excel
=SUM(C4:E4)
```

**What it does:** Adds the marks of Math, Science and English for each student to get the total out of 300.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `=` | Starts every Excel formula. |
| `SUM(...)` | Adds all the numbers in the given range. |
| `C4:E4` | Range from column C (Math) to column E (English) in row 4. Copy down for each student. |

**Output:**

| Student ID | Name | Math | Science | English | Total |
|---|---|---|---|---|---|
| 1001 | Student 1 | 61 | 48 | 76 | 185 |
| 1002 | Student 2 | 78 | 76 | 70 | 224 |
| 1003 | Student 3 | 97 | 60 | 51 | 208 |
| 1004 | Student 4 | 74 | 55 | 46 | 175 |
| 1005 | Student 5 | 68 | 74 | 84 | 226 |

*Showing first 5 of 20 students.*

---

### Q2. Calculate Percentage

**Sheet:** `Student Grade`  |  **Cell(s):** `I4:I23`

**Formula used:**

```excel
=(SUM(C4:E4)/300)*100
```

**What it does:** Converts the total marks into a percentage. Maximum marks = 3 subjects x 100 = 300.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `SUM(C4:E4)` | Total marks of the three subjects. |
| `/300` | Divides by the maximum possible marks. |
| `*100` | Converts the fraction into a percentage. |

**Output:**

| Name | Total | Calculation | Percentage |
|---|---|---|---|
| Student 1 | 185 | (185/300)*100 | 61.67 |
| Student 2 | 224 | (224/300)*100 | 74.67 |
| Student 3 | 208 | (208/300)*100 | 69.33 |
| Student 4 | 175 | (175/300)*100 | 58.33 |
| Student 5 | 226 | (226/300)*100 | 75.33 |

*Showing first 5 of 20 students.*

---

### Q3. Assign Grade

**Sheet:** `Student Grade`  |  **Cell(s):** `H5:H23`

**Formula used:**

```excel
=IF(AVERAGE(C5:E5)>=80,"A",IF(AVERAGE(C5:E5)>=60,"B",IF(AVERAGE(C5:E5)>=45,"C",IF(AVERAGE(C5:E5)>=33,"D","F"))))
```

**What it does:** Finds the average of the three subjects and returns a grade using a nested IF (multiple conditions checked in order).

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `AVERAGE(C5:E5)` | Mean of Math, Science and English marks. |
| `IF(test, value_if_true, value_if_false)` | Checks a condition and returns one of two results. |
| Nested `IF` | The false part of one IF contains the next IF, so the conditions are tested top to bottom. |
| `>=80` -> A | Average 80 or more. |
| `>=60` -> B | Average 60 to 79. |
| `>=45` -> C | Average 45 to 59. |
| `>=33` -> D | Average 33 to 44. |
| Otherwise -> F | Below 33 (fail). |

**Output:**

| Name | Total | Percentage | Grade |
|---|---|---|---|
| Student 1 | 185 | 61.67 | 0 |
| Student 2 | 224 | 74.67 | B |
| Student 3 | 208 | 69.33 | B |
| Student 4 | 175 | 58.33 | C |
| Student 5 | 226 | 75.33 | B |
| Student 6 | 215 | 71.67 | B |
| Student 7 | 200 | 66.67 | B |
| Student 8 | 225 | 75 | B |
| Student 9 | 172 | 57.33 | C |
| Student 10 | 210 | 70 | B |
| Student 11 | 205 | 68.33 | B |
| Student 12 | 253 | 84.33 | A |
| Student 13 | 285 | 95 | A |
| Student 14 | 177 | 59 | C |
| Student 15 | 196 | 65.33 | B |
| Student 16 | 235 | 78.33 | B |
| Student 17 | 165 | 55 | C |
| Student 18 | 209 | 69.67 | B |
| Student 19 | 228 | 76 | B |
| Student 20 | 230 | 76.67 | B |

> **Note:** Cell `H4` (Student 1) contains `='Student Grade'!J5` instead of the grade formula, so it shows `0`. Student 1 scores 61.67%, so the correct grade is **B**. Replace H4 with the same IF formula used in H5:H23 to fix it.

---

### Q4. Count Students Scoring Above 60 (COUNTIFS)

**Sheet:** `Student Grade`  |  **Cell(s):** `Q3`

**Formula used:**

```excel
=COUNTIFS(I4:I23,">60")
```

**What it does:** Counts how many students have a percentage greater than 60.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `COUNTIFS(range, criteria)` | Counts the cells in a range that meet one or more conditions. |
| `I4:I23` | The Percentage column for all 20 students. |
| `">60"` | Criteria: value greater than 60. Operators go inside quotes. |

**Output:**

| Criteria | Range | Result |
|---|---|---|
| > 60 | I4:I23 | 16 |

---

### Q5. Count Students With Math Above 30 (COUNTIFS)

**Sheet:** `Student Grade`  |  **Cell(s):** `Z3`

**Formula used:**

```excel
=COUNTIFS(C3:C23,">30")
```

**What it does:** Counts how many students scored more than 30 in Math. The header cell (C3) is text, so COUNTIFS ignores it.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `COUNTIFS` | Conditional count. |
| `C3:C23` | Math column (header + 20 students). |
| `">30"` | Criteria: Math marks greater than 30. |

**Output:**

| Criteria | Range | Result |
|---|---|---|
| > 30 | C3:C23 | 20 |

> **Note:** This block carries the same title as Q4 in the sheet, but its formula checks Math marks, not percentage.

---

### Q6. Average Score of Students Above 60

**Sheet:** `Student Grade`  |  **Cell(s):** `Z5`

**Formula used:**

```excel
=AVERAGEIF(I4:I23,">60")
```

**What it does:** Calculates the average percentage of only those students whose percentage is above 60.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `AVERAGEIF(range, criteria)` | Averages the cells that meet a condition. |
| `I4:I23` | Percentage column - used for both the test and the average. |
| `">60"` | Only percentages greater than 60 are included. |

**Output:**

| Students Included | Criteria | Average |
|---|---|---|
| 16 | > 60 | 73.62 |

---

### Q7. Filter Students Scoring Above 80

**Sheet:** `Student Grade`  |  **Cell(s):** `L7:S9 (spill range)`

**Formula used:**

```excel
=FILTER(B3:I23, I3:I23>80)
```

**What it does:** Extracts every row of the table where the percentage is greater than 80. Results spill into the cells below automatically.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `FILTER(array, include)` | Returns only the rows of `array` for which `include` is TRUE (Excel 365/2021). |
| `B3:I23` | Data to return: Name to Percentage, with headers. |
| `I3:I23>80` | Test applied to each row's Percentage; TRUE rows are kept. |

**Output:**

| Name | Math | Science | English | Enrollment Date | Total | Grade | Percentage |
|---|---|---|---|---|---|---|---|
| Student 12 | 85 | 89 | 79 | 04-Jun-2023 | 253 | A | 84.33 |
| Student 13 | 98 | 90 | 97 | 01-Dec-2023 | 285 | A | 95 |

---

### Q8. Fetch Student Name Based on ID

**Sheet:** `Student Grade`  |  **Cell(s):** `Y10`

**Formula used:**

```excel
=VLOOKUP(X10,$A$4:$B$23,2,FALSE)
```

**What it does:** Searches the Student ID typed in X10 and returns the matching student name.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `VLOOKUP(lookup, table, col, match)` | Vertical lookup: finds a value in the first column of a table and returns a value from another column. |
| `X10` | Lookup value - Student ID (1001). |
| `$A$4:$B$23` | Table: ID in column A, Name in column B. `$` locks the range when copied. |
| `2` | Return the 2nd column of the table (Name). |
| `FALSE` | Exact match only. |

**Output:**

| Student ID (X10) | Name (Y10) |
|---|---|
| 1001 | Student 1 |

---

### Q9. Students Scoring Above 80 in Maths AND Science

**Sheet:** `Student Grade`  |  **Cell(s):** `Q14:Q33`

**Formula used:**

```excel
=IF(AND(C5>80,D5>80),"Above 80 Mark","No Above 80 Mark")
```

**What it does:** Flags a student only when BOTH Math and Science are above 80.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `IF` | Returns the first text if the test is TRUE, otherwise the second. |
| `AND(C5>80, D5>80)` | TRUE only when every condition inside is TRUE. |
| `C5>80` | Math greater than 80. |
| `D5>80` | Science greater than 80. |

**Output:**

| Student ID | Name | Math | Science | Result |
|---|---|---|---|---|
| 1001 | Student 1 | 61 | 48 | No Above 80 Mark |
| 1002 | Student 2 | 78 | 76 | No Above 80 Mark |
| 1003 | Student 3 | 97 | 60 | No Above 80 Mark |
| 1004 | Student 4 | 74 | 55 | No Above 80 Mark |
| 1005 | Student 5 | 68 | 74 | No Above 80 Mark |
| 1006 | Student 6 | 80 | 70 | No Above 80 Mark |
| 1007 | Student 7 | 75 | 47 | No Above 80 Mark |
| 1008 | Student 8 | 94 | 67 | No Above 80 Mark |
| 1009 | Student 9 | 41 | 75 | No Above 80 Mark |
| 1010 | Student 10 | 98 | 54 | No Above 80 Mark |
| 1011 | Student 11 | 73 | 79 | No Above 80 Mark |
| 1012 | Student 12 | 85 | 89 | Above 80 Mark |
| 1013 | Student 13 | 98 | 90 | Above 80 Mark |
| 1014 | Student 14 | 93 | 40 | No Above 80 Mark |
| 1015 | Student 15 | 70 | 51 | No Above 80 Mark |
| 1016 | Student 16 | 98 | 77 | No Above 80 Mark |
| 1017 | Student 17 | 66 | 59 | No Above 80 Mark |
| 1018 | Student 18 | 82 | 48 | No Above 80 Mark |
| 1019 | Student 19 | 90 | 78 | No Above 80 Mark |
| 1020 | Student 20 | 83 | 92 | Above 80 Mark |

> **Note:** Cell `Q14` contains a typo, `"Above 80 T17Mark"` (stray text `T17`). Its result is still `No Above 80 Mark`. Only Students 12, 13 and 20 qualify.

---

---

## Sheet 2: Sales Data

### Q10. Search Product Price From Sales Data

**Sheet:** `Sales Data`  |  **Cell(s):** `L6`

**Formula used:**

```excel
=VLOOKUP(K6,A4:F23,5,FALSE)
```

**What it does:** Takes a Sales ID (K6 = 2005) and returns the Amount recorded for that sale.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `VLOOKUP` | Vertical lookup in the first column of a table. |
| `K6` | Lookup value - Sales ID 2005. |
| `A4:F23` | Table from Sales ID (col A) to Date (col F). |
| `5` | 5th column of the table = Amount (column E). |
| `FALSE` | Exact match. |

**Output:**

| Sales ID (K6) | Price / Amount (L6) |
|---|---|
| 2005 | 43388 |

---

### Q11. Total Sales for a Specific Region and Product

**Sheet:** `Sales Data`  |  **Cell(s):** `S6:W9`

**Formula used:**

```excel
=SUMIFS(E4:E23,C4:C23,"South",B4:B23,"Keyboard")
```

**What it does:** Adds the Amount only for rows that match BOTH the region and the product, building a Region x Product summary matrix.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `SUMIFS(sum_range, range1, crit1, range2, crit2)` | Sums values that meet several conditions at once (AND logic). |
| `E4:E23` | Sum range - the Amount column. |
| `C4:C23, "South"` | Condition 1: Region equals South. |
| `B4:B23, "Keyboard"` | Condition 2: Product equals Keyboard. |
| Text case | SUMIFS is not case-sensitive, so `"SOUTH"` and `"South"` behave the same. |

**Output:**

| Region | Keyboard | Laptop | Monitor | Mouse | Printer |
|---|---|---|---|---|---|
| South | 42233 | 7283 | 0 | 23672 | 0 |
| East | 80349 | 31238 | 0 | 37335 | 25010 |
| North | 43388 | 0 | 77482 | 19757 | 0 |
| West | 0 | 0 | 38161 | 30436 | 51381 |

> **Note:** Some formulas in this block use shifted ranges (for example `E5:E24` instead of `E4:E23`). Results are correct here because the extra rows are empty, but keeping the same range (use `$E$4:$E$23`) in every cell is safer.

---

### Q12. Calculate Discount Amount

**Sheet:** `Sales Data`  |  **Cell(s):** `O11:O30`

**Formula used:**

```excel
=IF(E4>30000,E4*0.1,IF(E4>15000,E4*0.05,0))
```

**What it does:** Gives a tiered discount: 10% for sales above 30,000, 5% for sales above 15,000, otherwise no discount.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `IF(E4>30000, E4*0.1, ...)` | First test: amount above 30,000 -> 10% discount. |
| Inner `IF(E4>15000, E4*0.05, 0)` | If the first test fails, check above 15,000 -> 5% discount. |
| `0` | Amount 15,000 or less -> no discount. |

**Output:**

| Sales ID | Product | Amount | Discount |
|---|---|---|---|
| 2001 | Keyboard | 7048 | 0 |
| 2002 | Printer | 17428 | 871.4 |
| 2003 | Mouse | 23672 | 1183.6 |
| 2004 | Printer | 5264 | 0 |
| 2005 | Keyboard | 43388 | 4338.8 |
| 2006 | Mouse | 19757 | 987.85 |
| 2007 | Mouse | 32058 | 3205.8 |
| 2008 | Keyboard | 46429 | 4642.9 |
| 2009 | Printer | 19746 | 987.3 |
| 2010 | Keyboard | 35185 | 3518.5 |

*Showing first 10 of 20 rows.*

---

### Q13. Discount Eligible / Not Eligible

**Sheet:** `Sales Data`  |  **Cell(s):** `P11:P30`

**Formula used:**

```excel
=IF(AND(OR(J11="Keyboard",J11="Mouse",J11="Laptop"),M11>10000),"Discount Eligible","No Discount")
```

**What it does:** A sale is eligible only if the product is Keyboard, Mouse or Laptop AND the amount is above 10,000.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `OR(J11="Keyboard", J11="Mouse", J11="Laptop")` | TRUE if the product matches at least one of the three names. |
| `M11>10000` | Amount greater than 10,000. |
| `AND(OR(...), M11>10000)` | Both the product rule and the amount rule must be TRUE. |
| `IF(..., "Discount Eligible", "No Discount")` | Displays the final label. |

**Output:**

| Sales ID | Product | Amount | Eligibility |
|---|---|---|---|
| 2001 | Keyboard | 7048 | No Discount |
| 2002 | Printer | 17428 | No Discount |
| 2003 | Mouse | 23672 | Discount Eligible |
| 2004 | Printer | 5264 | No Discount |
| 2005 | Keyboard | 43388 | Discount Eligible |
| 2006 | Mouse | 19757 | Discount Eligible |
| 2007 | Mouse | 32058 | Discount Eligible |
| 2008 | Keyboard | 46429 | Discount Eligible |
| 2009 | Printer | 19746 | No Discount |
| 2010 | Keyboard | 35185 | Discount Eligible |

*Showing first 10 of 20 rows.*

---

### Q14. Sales Value for a Salesperson in a Specific Month

**Sheet:** `Sales Data`  |  **Cell(s):** `T15`

**Formula used:**

```excel
=INDEX($E$4:$E$23,MATCH(1,($D$4:$D$23=T13)*(MONTH($F$4:$F$23)=T14),0))
```

**What it does:** Finds the Amount for the salesperson in T13 whose sale happened in the month number in T14.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `INDEX(range, row)` | Returns the value at a given position in a range. |
| `MATCH(1, array, 0)` | Finds the position of the first 1 (TRUE) in the array; `0` = exact match. |
| `($D$4:$D$23=T13)` | Array of TRUE/FALSE: salesperson equals Person 1. |
| `MONTH($F$4:$F$23)=T14` | `MONTH` extracts the month number (1-12) from each date and compares it to 1. |
| `( ... ) * ( ... )` | Multiplying the two TRUE/FALSE arrays works as AND: result is 1 only when both are TRUE. |
| Array formula | Entered as an array formula (Ctrl+Shift+Enter in older Excel). |

**Output:**

| Sales Person (T13) | Month Number (T14) | Result (T15) |
|---|---|---|
| Person 1 | 1 | 7048 |

---

---

## Sheet 3: Employees Data

### Q15. Fetch Employee Salary From Employee ID

**Sheet:** `Employees Data`  |  **Cell(s):** `J5`

**Formula used:**

```excel
=XLOOKUP(J4,A4:A23,D4:D23)
```

**What it does:** Looks up the Employee ID in J4 and returns the matching salary.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `XLOOKUP(lookup, lookup_array, return_array)` | Modern lookup: finds a value in one range and returns the matching item from another range. |
| `J4` | Lookup value - Employee ID 3001. |
| `A4:A23` | Column searched - Employee ID. |
| `D4:D23` | Column returned - Salary. |
| Why use it | No column index number needed, looks in any direction, exact match by default. |

**Output:**

| Employee ID (J4) | Salary Amount (J5) |
|---|---|
| 3001 | 86766 |

---

### Q16. Look Up Employee Details

**Sheet:** `Employees Data`  |  **Cell(s):** `G10:K10`

**Formula used:**

```excel
=XLOOKUP(G10,A4:A23,B4:E23)   <- suggested formula
```

**What it does:** Shows the full record (Name, Department, Salary, Joining Date) for Employee ID 3005.

**Function meaning (part by part):**

| Part | Meaning |
|---|---|
| `G10` | Employee ID typed in the cell (3005). |
| `A4:A23` | Column searched for the ID. |
| `B4:E23` | Returns the whole row of Name to Joining Date, spilling across the columns. |
| Alternative | `=VLOOKUP($G10,$A$4:$E$23,2,FALSE)` for Name, changing `2` to `3`, `4`, `5` for the other columns. |

**Output:**

| Employee Id | Name | Department | Salary | Joining Date |
|---|---|---|---|---|
| 3005 | Employee5 | IT | 54149 | 27-Sep-2017 |

> **Note:** In the workbook these cells hold typed values, not formulas (the name is stored as `Employee5`, while the data table uses `Employee 5`). The formula above is the recommended way to make the result dynamic.

---

## Quick Reference: Function Summary

| Function | Purpose | Syntax |
|---|---|---|
| `SUM` | Adds numbers | `=SUM(range)` |
| `AVERAGE` | Mean of numbers | `=AVERAGE(range)` |
| `IF` | Returns a value based on a test | `=IF(test, true_value, false_value)` |
| `AND` / `OR` | Combine several tests | `=AND(test1, test2)` / `=OR(test1, test2)` |
| `COUNTIFS` | Counts cells meeting conditions | `=COUNTIFS(range, criteria)` |
| `AVERAGEIF` | Average of cells meeting a condition | `=AVERAGEIF(range, criteria)` |
| `SUMIFS` | Sums cells meeting several conditions | `=SUMIFS(sum_range, range1, crit1, ...)` |
| `FILTER` | Returns rows that match a condition | `=FILTER(array, include)` |
| `VLOOKUP` | Lookup in first column, return a column | `=VLOOKUP(value, table, col, FALSE)` |
| `XLOOKUP` | Flexible lookup with separate ranges | `=XLOOKUP(value, lookup_range, return_range)` |
| `INDEX` + `MATCH` | Position-based lookup, supports multiple criteria | `=INDEX(range, MATCH(value, range, 0))` |
| `MONTH` | Month number from a date | `=MONTH(date)` |

## Issues Found in the Workbook

| # | Location | Issue | Fix |
|---|---|---|---|
| 1 | Student Grade `H4` | Points to `J5` (empty), so Grade shows 0 | Use the same nested IF formula as H5:H23 |
| 2 | Student Grade `Q14` | Typo `"Above 80 T17Mark"` in result text | Change to `"Above 80 Mark"` |
| 3 | Student Grade `Z3` | Same title as Q4 but checks Math > 30 | Rename the label or change the formula |
| 4 | Sales Data `S7:V9` | Ranges drift (`E5:E24`, `E6:E25`...) | Use fixed `$E$4:$E$23`-style ranges |
| 5 | Employees Data `G10:K10` | Static typed values, not a formula | Use `XLOOKUP` as shown in Q16 |

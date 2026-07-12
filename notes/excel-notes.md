# Excel Notes — Functions, Formulas, and Dashboards

Personal learning notes covering core Excel skills: functions and formulas, conditional formatting, data validation, pivot tables/charts, lookups, macros, and dashboard design.

---

## 1. Introduction to Excel

**Why use Excel?**
- Cost-efficient option for working with data
- Great platform for performing mathematical calculations on large datasets
- Built-in features (searching, sorting, filtering) make it easy to work with data
- Lets you present data visually with charts, tables, and data bars
- Useful for reporting, accounting, and analysis
- Can support task lists, calendars, and goal-planning worksheets
- Provides data security — password protection and safe storage

**What is Excel?**
More than just a tool for storing data in tabular form — it's used for recording, analyzing, and visualizing data.

**Core topics covered:**
1. Functions & Formulas
2. Conditional Formatting
3. Data Validation
4. Pivot Tables & Pivot Charts

---

## 2. Functions & Formulas

- **Add:** `=SUM(D4:D7)` or shortcut `Alt + =` (automatically selects adjacent numeric cells, which you can expand or collapse)
- **Add with condition:** `=SUMIF(D11:D15, ">50")`
- **Fill down:** drag a formula down a column (or `Ctrl + D`); check the formula in the first cell before filling
- **Fill right:** drag a formula across a row (or `Ctrl + R`) — useful for checking totals horizontally

### Splitting Text

For consistent delimiters (e.g. splitting an email into first/last name):
```
Nancy.Smith@gmail.com
Andy.North@fabric.com
```
Use `Ctrl + E` (Flash Fill) — Excel detects the pattern from your first example and fills the rest.

For inconsistent or multiple delimiters (e.g. `Nancy,Smith,Contoso Ltd.`):
**Data → Text to Columns → Delimited → choose delimiter (e.g. comma) → Finish**

For manual extraction:
```excel
=LEFT(C56, FIND(" ", C56) - 1)
=RIGHT(C56, LEN(C56) - FIND(" ", C56))
```

### Transpose

Turns rows into columns (or vice versa).

**Method 1 — Paste Special:**
Select your data → `Ctrl + C` → select an empty cell → **Home → Paste Special → Transpose**

**Method 2 — Formula (array formula):**
```excel
=TRANSPOSE(C33:H34)
```
Select a range of the correct new dimensions (e.g. 6 columns × 2 rows → transposes to 2 columns × 6 rows), type the formula in the address bar, then confirm with `Ctrl + Shift + Enter`. The same formula populates every cell because it's an array formula (a single formula calculating across multiple cells).

---

## 3. Sorting and Filtering

- Select a column (e.g. Department) → **Home → Sort & Filter → A to Z** — sorts alphabetically, and the rest of the row data moves with it
- Sort largest to smallest for descending numeric order
- **Custom Sort:** choose specific columns, and sort by cell value, cell color, or conditional formatting icon
  - Sorting by date: choose the date field → sort oldest to newest
- **Filters** appear in column headers — use them to show only specific values, date ranges, or cell colors (e.g. "above average")
- Filters can be cleared with **Clear Filter**

---

## 4. Tables

Select your data range → **Insert → Table** (check "My table has headers" if applicable).

- Tables are a special collection of cells with built-in features: structured rows/columns and calculations
- **Total row:** select a cell in the total row, or use `Alt + =` on a column to get its total. You can also change the calculation type (e.g. average instead of sum) via the dropdown
- Converting a range to a table: select a cell in the range → `Ctrl + T`

---

## 5. Data Validation & Drop-Downs

Used to control what a user is allowed to enter into a cell.

**Creating a drop-down list:**
**Data → Data Validation → Allow: List → Source:** the range of allowed values (e.g. bakery, meat, produce)

- If invalid data is entered, an error message is displayed
- Helps ensure only valid, expected data gets entered

---

## 6. Importing Data

**Data → Get Data (Other Sources) → From Web** — paste a URL to pull data directly from a webpage
**Data → Get Data → From Database** — connect to an external database

---

## 7. Built-In Functions Reference

| Function | Purpose |
|---|---|
| `ISEVEN` | Checks if a number is even |
| `ISLOGICAL` | Checks if a value is a logical (TRUE/FALSE) value |
| `SUBTOTAL` | Aggregates filtered/visible data |
| `INT` | Rounds down to the nearest integer |
| `SUM` | Adds values |
| `AVERAGE` | Averages values |
| `TRUNC` | Truncates a number to a specified precision |
| `ABS` | Absolute value |
| `SQRT` | Square root |
| `REPT` | Repeats text a given number of times |
| `COUNT` | Counts numeric values |
| `MAX` | Returns the largest value |
| `NOW` | Current date and time |
| `TIME` | Constructs a time value |
| `TEXT` | Formats a value as text |
| `SUMIF` | Conditional sum |
| `TRIM` | Removes extra spaces |
| `LEN` | Length of a text string |

**Conditional aggregation:**
```excel
=SUBTOTAL(9, range)              ' Sum of visible/filtered cells (function code 9 = SUM)
=SUMIF(range, criteria, [sum_range])
=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)   ' multiple conditions
=COUNTIF(range, criteria)
=COUNTIFS(range1, criteria1, range2, criteria2, ...)
```

Example:
```excel
=COUNTIFS(C2:C44, "Gill", D2:D44, "Pencil")
```

Use `"<>Pencil"` as a criterion to **exclude** a value (e.g. everything except pencils).

---

## 8. Conditional Formatting

Formats cells automatically based on a pre-set condition, making it easier to spot important values.

**Common rule types:**
- Highlight Cell Rules
- Top/Bottom Rules
- Data Bars
- Color Scales

---

## 9. Pivot Tables & Pivot Charts

- Summarize large datasets with many rows and columns
- Let you group and aggregate data dynamically
- A **pivot chart** is the visual representation of a pivot table
- **Slicers** provide interactive, clickable filters for a pivot table/chart

---

## 10. Lookup Functions

### VLOOKUP
Searches for a value within a range and returns a corresponding value from another column.

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```
- **lookup_value:** what you're searching for
- **table_array:** the range to search within (often a named range)
- **col_index_num:** the column number (within the range) to return a value from
- **range_lookup:** `TRUE` = approximate/closest match, `FALSE` = exact match

The lookup value must be in the **leftmost** column of the range.

Wrap in `IFERROR` to handle no-match cases gracefully:
```excel
=IFERROR(VLOOKUP(...), "Not found")
```

### HLOOKUP
Same as VLOOKUP, but searches horizontally (left to right) — the top row is the one being searched.

### INDEX / MATCH
More flexible than VLOOKUP since it isn't restricted to the leftmost column.

```excel
=INDEX(array, row_num, [column_num])
=MATCH(lookup_value, lookup_array, [match_type])
```
- `match_type = 0` → exact match
- `match_type = 1` → approximate match

Often combined: `MATCH` finds the row/column position, and `INDEX` returns the value at that position.

### XLOOKUP (modern replacement for VLOOKUP)
- Can search in any direction (left or right of the lookup column), unlike VLOOKUP
- Defaults to an exact match — no need to type `FALSE`/`0`
- Doesn't break if columns are inserted or deleted
- Has built-in error handling — you can specify a custom value (e.g. `"Not Found"`) instead of needing `IFERROR`

**What is `_xlfn`?**
A prefix Excel adds to a function that doesn't exist in your current version. For example, if a friend uses a newer version of Excel (with a function like XLOOKUP) and sends you the file, but you're on an older version, you'll see `_xlfn.XLOOKUP` instead of it resolving normally.

---

## 11. Order of Precedence

| Symbol | Operator | Precedence |
|---|---|---|
| `^` | Exponentiation | 1 |
| `*` | Multiplication | 2 |
| `/` | Division | 2 |
| `+` | Addition | 3 |
| `-` | Subtraction | 3 |
| `&` | Concatenation | 4 |
| `=` | Equal to | 5 |
| `<` | Less than | 5 |
| `>` | Greater than | 5 |

---

## 12. Common Error Types

| Error | Example Formula | Cause |
|---|---|---|
| `#DIV/0!` | `=1/0` | Formula is dividing by zero |
| `#VALUE!` | `=B4 + "text"` | Formula includes an argument of the wrong type |
| `#REF!` | `=#REF!` | Formula refers to an invalid cell (usually deleted) |
| `#NAME?` | `=COUNTT(A3:A9)` | Formula uses a name/function Excel doesn't recognize |
| `#N/A` | `=VLOOKUP("Value", A1:A10, 2, FALSE)` | Formula refers to data that isn't available |
| `#NUM!` | `=SQRT(-1)` | Problem with the numeric value used in the formula |
| `#NULL!` | `=SUM(A1:A10 B1:B10)` | Formula tries to intersect two ranges that don't actually intersect |

---

## 13. Text Functions

- `TEXTJOIN` — joins multiple values with a delimiter
- `TEXTSPLIT` — splits text into multiple cells
- `RIGHT` / `LEFT` — extract characters from the right/left of a string
- `FIND` / `SEARCH` — locate a substring's position (`SEARCH` is not case-sensitive, `FIND` is)
- `MID` — extract characters from the middle of a string

## 14. Date & Time Functions

- `MONTH`, `DAY`, `YEAR` — extract date components
- `DATE` — construct a date value
- `TODAY` — current date
- `DATEDIF` — difference between two dates
- `HOUR`, `MINUTE`, `SECOND` — extract time components
- `TIME` — construct a time value
- `SEQUENCE` — generates a sequence of numbers

---

## 15. Descriptive Analytics

Available via the **Analysis ToolPak** add-in.

### Correlation
Measures the linear relationship (positive, negative, or none) between two variables (e.g. temperature and units sold).

```excel
=CORREL(range1, range2)
```

**ANOVA** — statistical test for comparing means across groups.

### Sampling
- **Periodic sampling:** samples at a fixed interval/frequency
- **Random sampling:** selects values at random (e.g. three random temperature readings)

### Regression
Estimates the relationship between two or more variables — one response/dependent variable and one or more predictor/independent variables. Can be used to make predictions.

- **X** = independent variable (predictor/explanatory)
- **Y** = dependent variable (response/target)
- A scatter plot helps visualize whether the relationship looks linear
- **R²** (coefficient of determination) shows how well the regression line fits the data

**Types of regression:**
- Simple linear regression (one predictor)
- Multiple linear regression (two or more predictors)
- Non-linear regression

**Related concepts to know:** R² vs. adjusted R², goodness of fit, standard error, residuals, ANOVA table, coefficients.

**Relevant functions:**
- `LINEST` — regression statistics
- Forecasting functions (e.g. `FORECAST.LINEAR`)

---

## 16. What-If Analysis

- **Goal Seek** — works backward from a target result to find the input value needed
- **Scenario Manager** — lets you save and compare multiple sets of input values

---

## 17. Macros & VBA

**What are macros?**
A recorded sequence of actions (mouse clicks and keystrokes) that you can replay as many times as needed — useful for repetitive daily/weekly/monthly tasks and reports.

**Adding a button to run a macro:**
**Developer tab → Insert → Form Controls → Button**

**Editing a macro:**
**Developer tab → Visual Basic → open the module → edit the code**

**VBA (Visual Basic for Applications)** is Excel's programming language for automating tasks via macros.
- Saves time on repetitive tasks
- Uses English-like statements
- Extends what Excel can do beyond built-in features

### VBA Basics

```vb
Sub ExampleMacro()
    ' Commands go here
End Sub
```

- `Dim` — declares a variable
- `MsgBox` — displays a message box
- `Range("A1").Value` — reads/sets a cell's value
- `Cells(6, 7)` — refers to row 6, column 7
- `Set example = Range("A1")` then `example.Value` — assigning an object reference

**Common properties/methods:**
```vb
Selection.Copy        ' Copy the current selection
Range("A1").Select    ' Select a cell
ActiveSheet.Paste      ' Paste into the active sheet
.Count                 ' Count of items
.Columns.Count         ' Number of columns
.Rows.Count            ' Number of rows
.Interior.ColorIndex   ' Cell fill color
.Interior.Pattern      ' Cell fill pattern
```

**If statement:**
```vb
Dim score As Integer, result As String
score = Range("A1").Value

If score >= 60 Then
    result = "Pass"
Else
    result = "Fail"
End If
```

---

## 18. Dashboards

A dashboard is a visual interface providing a high-level overview of key metrics relevant to a specific objective, using charts and graphs.

**Types of dashboards:**
- Strategic dashboards
- Analytical dashboards
- Operational dashboards

**Useful features:**
- Slicers and timelines (for interactive, clickable filtering)
- Built-in themes for styling — e.g. Office, Facet, Organic

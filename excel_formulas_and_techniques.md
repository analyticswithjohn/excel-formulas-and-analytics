# Excel Formulas and Techniques

This document outlines various Excel formulas and techniques from basic to advanced, including edge cases and practical applications. Each formula or technique is explained with examples, common pitfalls, and best practices.

## Table of Contents

1. [Lookup and Reference Functions](#lookup-and-reference-functions)
   - [VLOOKUP](#vlookup)
   - [HLOOKUP](#hlookup)
   - [INDEX-MATCH](#index-match)
   - [INDEX-MATCH-MATCH](#index-match-match)
   - [XLOOKUP](#xlookup)
   - [XMATCH](#xmatch)

2. [Text Functions](#text-functions)
   - [CONCATENATE and & Operator](#concatenate-and--operator)
   - [LEFT, RIGHT, MID](#left-right-mid)
   - [FIND, SEARCH](#find-search)
   - [SUBSTITUTE](#substitute)
   - [TEXTJOIN](#textjoin)

3. [Date and Time Functions](#date-and-time-functions)
   - [DATE, DATEVALUE](#date-datevalue)
   - [NETWORKDAYS, WORKDAY](#networkdays-workday)
   - [EDATE, EOMONTH](#edate-eomonth)
   - [DATEDIF](#datedif)

4. [Conditional Functions](#conditional-functions)
   - [IF](#if)
   - [IFS](#ifs)
   - [SWITCH](#switch)
   - [Nested IF Statements](#nested-if-statements)

5. [Statistical Functions](#statistical-functions)
   - [COUNTIF, COUNTIFS](#countif-countifs)
   - [SUMIF, SUMIFS](#sumif-sumifs)
   - [AVERAGEIF, AVERAGEIFS](#averageif-averageifs)

6. [Array Formulas](#array-formulas)
   - [Basic Array Formulas](#basic-array-formulas)
   - [Dynamic Arrays](#dynamic-arrays)
   - [FILTER, SORT, UNIQUE](#filter-sort-unique)

7. [Data Validation and Error Handling](#data-validation-and-error-handling)
   - [IFERROR, IFNA](#iferror-ifna)
   - [ISBLANK, ISERROR](#isblank-iserror)
   - [Data Validation Techniques](#data-validation-techniques)

8. [Power Functions](#power-functions)
   - [POWER QUERY](#power-query)
   - [POWER PIVOT](#power-pivot)
   - [DAX Basics](#dax-basics)

## Lookup and Reference Functions

### VLOOKUP

**Description:**
VLOOKUP (Vertical Lookup) is one of the most commonly used lookup functions in Excel. It searches for a value in the leftmost column of a table and returns a value in the same row from a column you specify.

**Syntax:**
```
VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

**Parameters:**
- `lookup_value`: The value to search for in the first column of the table
- `table_array`: The range of cells that contains the data
- `col_index_num`: The column number in the table from which to retrieve the value
- `range_lookup`: [Optional] TRUE for approximate match (default), FALSE for exact match

**Example Scenarios:**
1. **Basic Employee Lookup**: Finding an employee's salary based on their ID
2. **Product Information Retrieval**: Looking up product details based on product code
3. **Customer Data Management**: Retrieving customer information based on customer ID

**Edge Cases and Limitations:**
1. **Lookup Value Not Found**: Returns #N/A if the lookup value doesn't exist in the first column
2. **Case Sensitivity**: VLOOKUP is not case-sensitive
3. **Hidden Columns**: VLOOKUP counts hidden columns in the col_index_num
4. **Leftmost Column Requirement**: Can only look up values in the leftmost column of the table_array
5. **Duplicates**: Returns the first match if multiple matches exist

**Best Practices:**
1. Always use FALSE for exact matches unless you specifically need an approximate match
2. Sort your data if using approximate match (TRUE)
3. Consider using INDEX-MATCH for more flexibility
4. Use IFERROR to handle lookup failures gracefully

### HLOOKUP

**Description:**
HLOOKUP (Horizontal Lookup) works similarly to VLOOKUP but searches horizontally across the first row of a table instead of vertically down the first column.

**Syntax:**
```
HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
```

**Parameters:**
- `lookup_value`: The value to search for in the first row of the table
- `table_array`: The range of cells that contains the data
- `row_index_num`: The row number in the table from which to retrieve the value
- `range_lookup`: [Optional] TRUE for approximate match (default), FALSE for exact match

**Example Scenarios:**
1. **Monthly Data Analysis**: Looking up values for specific months in a year-based dataset
2. **Product Specifications**: Finding specific attributes of products listed horizontally
3. **Time-Series Data**: Retrieving data points at specific time intervals

**Edge Cases and Limitations:**
1. Similar limitations to VLOOKUP but in horizontal orientation
2. Less commonly used than VLOOKUP due to typical data organization in Excel

### INDEX-MATCH

**Description:**
INDEX-MATCH is a powerful combination of two functions that overcomes many limitations of VLOOKUP. INDEX returns a value at a specific position in a range, while MATCH finds the position of a value within a range.

**Syntax:**
```
INDEX(array, row_num, [column_num])
MATCH(lookup_value, lookup_array, [match_type])
```

Combined:
```
INDEX(return_range, MATCH(lookup_value, lookup_array, 0))
```

**Parameters:**
- `array`: The range of cells to return a value from
- `row_num`: The row position in the array
- `column_num`: [Optional] The column position in the array
- `lookup_value`: The value to find in the lookup_array
- `lookup_array`: The range of cells to search
- `match_type`: 1 for less than, 0 for exact match, -1 for greater than

**Example Scenarios:**
1. **Flexible Database Queries**: Looking up values when the lookup column is not the leftmost column
2. **Bidirectional Lookups**: Finding values based on criteria in any column
3. **Dynamic References**: Creating references that adjust when columns are inserted or deleted

**Advantages over VLOOKUP:**
1. Can look up values in any column, not just the leftmost
2. More efficient for large datasets
3. More flexible when columns are added or removed
4. Can perform right-to-left lookups

**Edge Cases and Limitations:**
1. More complex syntax than VLOOKUP
2. Returns #N/A if the lookup value is not found
3. Requires exact match (match_type=0) for most business applications

### INDEX-MATCH-MATCH

**Description:**
INDEX-MATCH-MATCH extends the INDEX-MATCH concept to perform two-dimensional lookups, allowing you to find values at the intersection of both row and column criteria.

**Syntax:**
```
INDEX(array, MATCH(row_lookup_value, row_lookup_array, 0), MATCH(column_lookup_value, column_lookup_array, 0))
```

**Example Scenarios:**
1. **Matrix Lookups**: Finding values at the intersection of two criteria (e.g., product and month)
2. **Financial Models**: Looking up values in complex financial tables
3. **Cross-Tabulated Data**: Retrieving specific data points from cross-tabulated reports

**Edge Cases and Considerations:**
1. Complexity increases with two MATCH functions
2. Both row and column criteria must exist for successful lookup
3. Excellent for dynamic dashboards where both row and column references may change

### XLOOKUP

**Description:**
XLOOKUP is a modern replacement for VLOOKUP, HLOOKUP, and INDEX-MATCH, introduced in Excel 365. It offers more flexibility and functionality in a single function.

**Syntax:**
```
XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

**Parameters:**
- `lookup_value`: The value to search for
- `lookup_array`: The range to search
- `return_array`: The range to return
- `if_not_found`: [Optional] Value to return if lookup_value is not found
- `match_mode`: [Optional] 0=exact match (default), -1=exact or next smaller, 1=exact or next larger, 2=wildcard match
- `search_mode`: [Optional] 1=first to last (default), -1=last to first, 2=binary ascending, -2=binary descending

**Example Scenarios:**
1. **Advanced Customer Lookups**: Finding customer data with fallback values if not found
2. **Reverse Lookups**: Searching from bottom to top of a dataset
3. **Range Lookups**: Finding values within specific ranges
4. **Wildcard Searches**: Looking up partial matches using wildcards

**Advantages over VLOOKUP and INDEX-MATCH:**
1. Can look up values to the left or right
2. Built-in error handling with if_not_found parameter
3. Can perform exact, approximate, and wildcard matches
4. Can search in any direction (first-to-last or last-to-first)
5. Supports array returns for multiple matches

**Edge Cases and Considerations:**
1. Only available in Excel 365 and later versions
2. More parameters to manage but more powerful
3. Can handle vertical and horizontal lookups in a single function

### XMATCH

**Description:**
XMATCH is the modern equivalent of MATCH, returning the relative position of an item in an array or range. It offers more flexibility than the traditional MATCH function.

**Syntax:**
```
XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
```

**Parameters:**
- `lookup_value`: The value to search for
- `lookup_array`: The range to search
- `match_mode`: [Optional] 0=exact match (default), -1=exact or next smaller, 1=exact or next larger, 2=wildcard match
- `search_mode`: [Optional] 1=first to last (default), -1=last to first, 2=binary ascending, -2=binary descending

**Example Scenarios:**
1. **Position Finding**: Determining the position of items in a list
2. **Dynamic Column References**: Finding column positions in dynamic tables
3. **Duplicate Handling**: Finding first or last occurrence of values

**Edge Cases and Considerations:**
1. Only available in Excel 365 and later versions
2. Returns #N/A if the value is not found
3. Can be combined with INDEX for powerful lookups

## Text Functions

### CONCATENATE and & Operator

**Description:**
These functions combine text from multiple cells or strings into one cell.

**Syntax:**
```
CONCATENATE(text1, [text2], ...)
```
Or using the & operator:
```
=text1 & text2 & text3
```

**Example Scenarios:**
1. **Name Formatting**: Combining first and last names
2. **Address Construction**: Building full addresses from components
3. **URL Building**: Creating dynamic URLs from base and parameters

**Edge Cases and Considerations:**
1. & operator is generally more flexible than CONCATENATE
2. Spaces must be explicitly included
3. Non-text values are automatically converted to text

### LEFT, RIGHT, MID

**Description:**
These functions extract a specific number of characters from a text string.

**Syntax:**
```
LEFT(text, [num_chars])
RIGHT(text, [num_chars])
MID(text, start_num, num_chars)
```

**Example Scenarios:**
1. **Data Cleaning**: Extracting specific portions of inconsistently formatted data
2. **Code Extraction**: Pulling product codes from longer strings
3. **Text Parsing**: Breaking down complex text into manageable components

**Edge Cases and Considerations:**
1. If num_chars is greater than the length of text, the entire text is returned
2. If start_num is greater than the length of text, MID returns an empty string
3. Combine with LEN for dynamic character counting

### FIND, SEARCH

**Description:**
These functions locate the position of one text string within another.

**Syntax:**
```
FIND(find_text, within_text, [start_num])
SEARCH(find_text, within_text, [start_num])
```

**Key Differences:**
- FIND is case-sensitive
- SEARCH allows wildcard characters (* and ?)

**Example Scenarios:**
1. **Data Extraction**: Finding the position of delimiters in strings
2. **Text Validation**: Checking if specific text exists within a cell
3. **Pattern Matching**: Locating patterns in text data

**Edge Cases and Considerations:**
1. Returns #VALUE! if the text is not found
2. Combine with MID for advanced text extraction

### SUBSTITUTE

**Description:**
SUBSTITUTE replaces specific text within a string.

**Syntax:**
```
SUBSTITUTE(text, old_text, new_text, [instance_num])
```

**Example Scenarios:**
1. **Data Standardization**: Replacing variations in text data
2. **Format Conversion**: Changing delimiters in strings
3. **Text Cleaning**: Removing or replacing unwanted characters

**Edge Cases and Considerations:**
1. Case-sensitive replacements
2. Use instance_num to replace only specific occurrences
3. For case-insensitive replacements, combine with UPPER or LOWER

### TEXTJOIN

**Description:**
TEXTJOIN combines text from multiple ranges with a specified delimiter.

**Syntax:**
```
TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
```

**Example Scenarios:**
1. **List Creation**: Creating comma-separated lists from cell ranges
2. **Advanced Concatenation**: Joining text with consistent delimiters
3. **Report Generation**: Building formatted text from multiple data points

**Edge Cases and Considerations:**
1. Only available in Excel 2019 and Office 365
2. Set ignore_empty to TRUE to skip empty cells
3. Can handle arrays and ranges directly

## Date and Time Functions

### DATE, DATEVALUE

**Description:**
These functions create or convert dates in Excel.

**Syntax:**
```
DATE(year, month, day)
DATEVALUE(date_text)
```

**Example Scenarios:**
1. **Date Construction**: Building dates from separate year, month, and day values
2. **Text-to-Date Conversion**: Converting text strings to proper Excel dates
3. **Dynamic Date Creation**: Creating dates based on calculations

**Edge Cases and Considerations:**
1. Excel stores dates as sequential numbers
2. Regional settings affect date interpretation
3. DATEVALUE requires text in a recognizable date format

### NETWORKDAYS, WORKDAY

**Description:**
These functions calculate working days, excluding weekends and optionally holidays.

**Syntax:**
```
NETWORKDAYS(start_date, end_date, [holidays])
WORKDAY(start_date, days, [holidays])
```

**Example Scenarios:**
1. **Project Planning**: Calculating project durations in working days
2. **Delivery Estimation**: Determining delivery dates excluding weekends and holidays
3. **Business Day Calculations**: Finding the number of business days between dates

**Edge Cases and Considerations:**
1. Holidays parameter is optional but important for accuracy
2. NETWORKDAYS counts the days between dates
3. WORKDAY calculates a date based on working days from a start date

### EDATE, EOMONTH

**Description:**
These functions calculate dates a specified number of months in the future or past.

**Syntax:**
```
EDATE(start_date, months)
EOMONTH(start_date, months)
```

**Example Scenarios:**
1. **Subscription Management**: Calculating renewal dates
2. **Financial Reporting**: Finding month-end dates for reporting periods
3. **Contract Management**: Determining contract milestone dates

**Edge Cases and Considerations:**
1. EDATE maintains the same day of the month when possible
2. EOMONTH always returns the last day of the target month
3. Negative values for months parameter move backward in time

### DATEDIF

**Description:**
DATEDIF calculates the difference between two dates in various units.

**Syntax:**
```
DATEDIF(start_date, end_date, unit)
```

**Units:**
- "Y": Years
- "M": Months
- "D": Days
- "YM": Months excluding years
- "YD": Days excluding years
- "MD": Days excluding months and years

**Example Scenarios:**
1. **Age Calculation**: Determining exact age in years, months, and days
2. **Service Duration**: Calculating length of service for employees
3. **Time Span Analysis**: Breaking down time periods into component parts

**Edge Cases and Considerations:**
1. Undocumented function in modern Excel versions
2. Different units provide different perspectives on the same time span
3. Combine multiple DATEDIF calls for comprehensive time difference analysis

## Conditional Functions

### IF

**Description:**
IF evaluates a condition and returns one value if true and another if false.

**Syntax:**
```
IF(logical_test, value_if_true, value_if_false)
```

**Example Scenarios:**
1. **Status Indicators**: Displaying "Completed" or "Pending" based on status values
2. **Performance Evaluation**: Categorizing performance metrics into bands
3. **Data Validation**: Flagging values that meet or fail specific criteria

**Edge Cases and Considerations:**
1. Can be nested up to 64 levels (though not recommended for readability)
2. Use IFS for multiple conditions
3. Can return calculations, text, or other formula results

### IFS

**Description:**
IFS evaluates multiple conditions and returns a value corresponding to the first TRUE condition.

**Syntax:**
```
IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)
```

**Example Scenarios:**
1. **Grading Systems**: Assigning letter grades based on numerical scores
2. **Pricing Tiers**: Determining pricing based on quantity brackets
3. **Status Classification**: Categorizing items based on multiple criteria

**Edge Cases and Considerations:**
1. Only available in Excel 2019 and Office 365
2. More readable alternative to nested IF statements
3. Returns #N/A if no conditions are met (consider adding a final TRUE condition)

### SWITCH

**Description:**
SWITCH compares an expression against a list of values and returns the result corresponding to the first match.

**Syntax:**
```
SWITCH(expression, value1, result1, [value2, result2], ..., [default])
```

**Example Scenarios:**
1. **Department Mapping**: Converting department codes to names
2. **Day of Week Handling**: Performing different calculations based on weekday
3. **Category Processing**: Applying different formulas based on product categories

**Edge Cases and Considerations:**
1. Only available in Excel 2019 and Office 365
2. More efficient than nested IF statements for equality comparisons
3. Optional default value if no matches are found

### Nested IF Statements

**Description:**
Multiple IF functions can be nested to evaluate more complex conditions.

**Syntax:**
```
IF(condition1, value1, IF(condition2, value2, IF(condition3, value3, value4)))
```

**Example Scenarios:**
1. **Complex Classification**: Categorizing data based on multiple sequential criteria
2. **Tiered Calculations**: Applying different formulas based on hierarchical conditions
3. **Decision Trees**: Implementing simple decision tree logic

**Edge Cases and Considerations:**
1. Becomes difficult to read and maintain beyond 3-4 levels
2. Consider IFS or SWITCH for better readability
3. Maximum of 64 nested functions in Excel

## Statistical Functions

### COUNTIF, COUNTIFS

**Description:**
These functions count cells that meet specified criteria.

**Syntax:**
```
COUNTIF(range, criteria)
COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

**Example Scenarios:**
1. **Inventory Analysis**: Counting products in specific categories
2. **Performance Metrics**: Counting occurrences of specific performance levels
3. **Data Profiling**: Analyzing data distribution across categories

**Edge Cases and Considerations:**
1. Criteria can include wildcards (* and ?) for partial matching
2. Text criteria must be enclosed in quotes
3. COUNTIFS allows multiple criteria across different ranges

### SUMIF, SUMIFS

**Description:**
These functions sum values that meet specified criteria.

**Syntax:**
```
SUMIF(range, criteria, [sum_range])
SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

**Example Scenarios:**
1. **Sales Analysis**: Summing sales for specific regions or products
2. **Budget Tracking**: Calculating expenses by category
3. **Performance Evaluation**: Summing metrics that meet quality thresholds

**Edge Cases and Considerations:**
1. If sum_range is omitted in SUMIF, the range is used for both criteria and summing
2. SUMIFS requires all criteria to be met (AND logic)
3. For OR logic, use multiple SUMIF functions and add the results

### AVERAGEIF, AVERAGEIFS

**Description:**
These functions calculate the average of values that meet specified criteria.

**Syntax:**
```
AVERAGEIF(range, criteria, [average_range])
AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

**Example Scenarios:**
1. **Performance Analysis**: Calculating average scores for specific groups
2. **Quality Control**: Finding average measurements for specific conditions
3. **Financial Analysis**: Determining average transaction values by category

**Edge Cases and Considerations:**
1. Ignores text and logical values in the average_range
2. Returns #DIV/0! if no cells meet the criteria
3. Similar structure to SUMIF/SUMIFS

## Array Formulas

### Basic Array Formulas

**Description:**
Array formulas perform multiple calculations on one or more items in an array.

**Traditional Syntax (Excel 2016 and earlier):**
```
{=FORMULA(array)}
```
(Entered with Ctrl+Shift+Enter)

**Example Scenarios:**
1. **Matrix Operations**: Performing calculations across rows and columns simultaneously
2. **Advanced Filtering**: Creating complex filters without helper columns
3. **Multi-Condition Analysis**: Analyzing data that meets multiple criteria

**Edge Cases and Considerations:**
1. Traditional array formulas require Ctrl+Shift+Enter
2. Can be computationally intensive for large ranges
3. Modern Excel uses dynamic arrays without CSE requirement

### Dynamic Arrays

**Description:**
Dynamic arrays automatically spill results into adjacent cells without using Ctrl+Shift+Enter.

**Example Functions:**
- SORT
- FILTER
- UNIQUE
- SEQUENCE
- RANDARRAY

**Example Scenarios:**
1. **Top N Analysis**: Displaying top performers without complex formulas
2. **Data Transformation**: Reshaping data for analysis or presentation
3. **Automated Reporting**: Creating dynamic report sections that adjust to data size

**Edge Cases and Considerations:**
1. Only available in Excel 365
2. Results spill automatically into adjacent cells
3. #SPILL! error occurs if the spill range is blocked by existing data

### FILTER, SORT, UNIQUE

**Description:**
These dynamic array functions filter, sort, and deduplicate data.

**Syntax:**
```
FILTER(array, include, [if_empty])
SORT(array, [sort_index], [sort_order], [by_col])
UNIQUE(array, [by_col], [exactly_once])
```

**Example Scenarios:**
1. **Dynamic Dashboards**: Creating filtered views of data that update automatically
2. **Data Cleaning**: Removing duplicates and organizing data
3. **Conditional Reporting**: Displaying only data that meets specific criteria

**Edge Cases and Considerations:**
1. Only available in Excel 365
2. FILTER returns if_empty value when no data meets criteria
3. UNIQUE can identify unique values or unique combinations of values

## Data Validation and Error Handling

### IFERROR, IFNA

**Description:**
These functions handle errors in formulas by providing alternative values.

**Syntax:**
```
IFERROR(value, value_if_error)
IFNA(value, value_if_na)
```

**Example Scenarios:**
1. **Lookup Error Handling**: Providing friendly messages for failed lookups
2. **Calculation Protection**: Preventing formula errors from breaking dashboards
3. **User Experience Improvement**: Displaying meaningful messages instead of errors

**Edge Cases and Considerations:**
1. IFERROR catches all errors, IFNA only catches #N/A errors
2. Can hide legitimate errors that should be investigated
3. Consider using more specific error handling when appropriate

### ISBLANK, ISERROR

**Description:**
These functions test for specific conditions in cells.

**Syntax:**
```
ISBLANK(value)
ISERROR(value)
```

**Example Scenarios:**
1. **Data Validation**: Checking for missing or erroneous data
2. **Conditional Formatting**: Highlighting cells with potential issues
3. **Process Control**: Preventing calculations on incomplete data

**Edge Cases and Considerations:**
1. ISBLANK returns TRUE only for truly empty cells, not for cells with ""
2. ISERROR returns TRUE for any error type
3. Can be combined with IF for conditional processing

### Data Validation Techniques

**Description:**
Excel's Data Validation feature restricts input to predefined criteria.

**Types:**
- List validation
- Date validation
- Number validation
- Custom formula validation

**Example Scenarios:**
1. **Form Controls**: Creating dropdown lists for standardized input
2. **Data Entry Rules**: Enforcing business rules during data entry
3. **Error Prevention**: Restricting inputs to valid ranges or formats

**Edge Cases and Considerations:**
1. Can be circumvented by pasting values
2. Consider using VBA for more robust validation
3. Custom formulas provide the most flexibility

## Power Functions

### POWER QUERY

**Description:**
Power Query (Get & Transform) is a powerful ETL tool for data preparation.

**Key Features:**
- Data connection and import
- Data transformation and cleaning
- Append and merge queries
- Parameterized queries

**Example Scenarios:**
1. **Data Integration**: Combining data from multiple sources
2. **Automated Reporting**: Creating refreshable report structures
3. **Data Cleaning**: Standardizing and transforming raw data

**Edge Cases and Considerations:**
1. Learning curve steeper than standard Excel functions
2. Excellent for repetitive data preparation tasks
3. Query steps are recorded and can be modified or reordered

### POWER PIVOT

**Description:**
Power Pivot enables data modeling and analysis with large datasets.

**Key Features:**
- Data modeling with relationships
- Calculated columns and measures
- KPI definition
- Hierarchies and perspectives

**Example Scenarios:**
1. **Business Intelligence**: Creating interactive dashboards with large datasets
2. **Financial Modeling**: Building complex financial models with multiple dimensions
3. **Sales Analysis**: Analyzing sales data across multiple dimensions

**Edge Cases and Considerations:**
1. Requires understanding of data modeling concepts
2. Uses DAX formula language
3. Enables working with much larger datasets than standard Excel

### DAX Basics

**Description:**
Data Analysis Expressions (DAX) is a formula language used in Power Pivot and Power BI.

**Key Functions:**
- CALCULATE
- FILTER
- RELATED
- SUMX, AVERAGEX
- Time intelligence functions

**Example Scenarios:**
1. **Advanced Calculations**: Creating complex business metrics
2. **Time Intelligence**: Analyzing trends over time periods
3. **Relationship-Based Analysis**: Leveraging table relationships for insights

**Edge Cases and Considerations:**
1. Different paradigm from standard Excel formulas
2. Context transition is a key concept to understand
3. Powerful for complex business logic implementation

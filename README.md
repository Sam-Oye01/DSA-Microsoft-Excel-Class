# PROJECT : Microsoft Excel Beginner class

## COURSE OUTLINE

### Introduction to Data Analysis
### Excel Functions 1 (Numbers)
### Excel Functions 2 (Texts)
### Excel Functions 3 (Conditionals)
### Excel Functions 4 (Dates)
### Excel functions 5 (LOOKUPS)
### Introduction to Pivot Table I
### Pivot Table II

## DAY ONE (21 - 04 - 2025) 

### Introduction to Data Analysis

Today, I started the journey of the Data Analysis with Digital Skill-Up Africa (DSA) with the introduction to Microsoft Excel. 

  "The capacity to learn is a gift; the ability to learn is a skill; the willingness to learn to learn is a choice" 

**Agenda**

- Foundations of Data
- Introduction to MS Excel
- Basic Excel functions
- Reports and dashboards in Excel

### Foundations of Data

**Understanding Data concept**

Data refers to a collection of facts, figures or information that can be analyzed and used to derive instghts and to make decisions.

**Data Types**

  - Qualitative data (Texts)
  - Quantitative data (Numbers)

**Data Terminologies**

  - Data points : A single piece of individual information within a data set (a single row or column).
  - Data sets : Combination of individual data points.
  - Data base : A repository where data is stored for efficient retrieval, management and manipulation.
  - Data cleaning : A process of removing unwanted data.

**Data Literacy pathway**

Data Generation - Data structure - Data storage - Data Analysis - Statistics - Data driven decision making.

**Categories of Data**

  - Structured (Tabular form)
  - Semi-structured (J-son - "Java Script Object notation" or Xml - "Extensible mark-up language")
  - Strucutured (media files)

**Sources of Data**

  - Primary Data : first-hand data.
  - Secondary Data : second-hand data (gotten from a third party).

**Data collection methods**

  - Statistical surveys (questioneers or interviews) for observational study
  - Experiments for experimental study.

#### DATA ANALYSIS : Exploring data to discover useful information, draw conclusions and support decision making.

The term that defines transformation of data to a structure that is ready for analysis is ETL.

  E - Extract : Connecting to data source either in cloud or premise.  
  T - Transform : Cleaning and manipulation of data into a proper structure.
  L - Load : Loading into an analysis tool like Power-BI.

**Data Analysis Life Cycle**

  - Ingestion - Ingesting data into tool for analysis.
  - Transformation - Data cleaning and manipulation.
  - Modelling - Connecting Data tables together.
  - Visualization - Pictorial representation of data using charts and illustrations.
  - Analysis - Exploratory Data Analysis (EDA).
  - Presentation of a report.

**Why are analysis carried out in organisations?**

  - To know what is working (Descriptive).
  - To know what is not working (Diagnostic).
  - To know the future course of event (Predictive/Prognostic).
  - To know what to focus on and proffer solutions (Prescriptive).

## DAY TWO (23 - 04 - 2025) 

### Excel Functions 1 (Numbers)

Today, we began hands-on practical classes using MS Excel. We learnt a few things about the features of the MS Excel platform.

### Introduction to MS Excel

MS Excel is a robust spreadsheet application created for organizing, calculating, and analyzing data. It provides tools for data entry, data manipulation and complex calculation.

- Workbook houses worksheets.
- The Ribbon Interface houses all the tools present in Excel
- The Tabs or Menu bar.
- Within each Tabs are different "Groups" that consists the various commands on Excel.
- Quick access tool bar helps pin important controls on Excel.
- Formula bar houses all the formula been used.
- Name box and formula bar are on the same row.
- Cells are the intercept of the rows and the columns on Excel.
- Scroll bar is at the extreme right on the worksheet.
- The max row on excel is 1048576.
- The max column is XFD.

Formatting in Excel (organizing data)

- Sort & Filtering.
- Embolding and labelling/heading
- Change of font (styling)
- Alignment
- Fill colour and tex colour
- Border
- Flash fill

**Introduction to Functions in Excel**

Functions are essential keywords that makes operations in Excel easy. Every function is excel always have arguements or parameters. The  Mandatory arguements are needed for the functions to work while the optional arguements defined by the square brackets are not necessarily needed for the functions to work.

Aggregation Functions : Sum, Average, Max, Min, Count, Large, Small.

Excel file for the class : [Download here](https://github.com/user-attachments/files/20906112/Excel.Functions.1.-.Numbers.xlsx)

**Note** : You must always start with an equal to sign whenever you want to begin any operation in Excel.

## DAY THREE (24 - 04 - 2025)

### Excel Functions 2 (Texts)

Today, we continued our Excel journey by focusing on some conditional (IF) functions. We also worked on some Data manipulation on qualitative data.  

**Note** : The difference between SumIF and SumIFS is that the latter can accomodate multiple conditions. You can use the function arguement box (ctrl A) to become aquainted with the meaning of each of the arguements in a function.

``` Microsoft Excel ```

Conditional (IF) functions

    SUMIF(range, criteria, [sum_range]) or SUMIFS
    AVERAGEIF(range, criteria, [average_range]) or AVERAGEIFS
    MINIFS
    MAXIFS
    COUNTIF(range, criteria)

Text Extraction

    LEFT(text, [num_chars])
    RIGHT(text,[num_chars])
    MID(text, start_num, num_chars)
    FIND(find_text, within_text, [start_num])
    SEARCH(find_text, within_text, [start_num])

**Data Cleaning** 

M - Missing values
I - Inconsistent values
D - Duplicate values
O - Outliers

**Inconsistent Data**

| Function | Inconsistency |
| :-------------: | :-------- |
| TRIM | Eliminate unnecessary spaces from Names |
| UPPER | Names written in Upper case |
| LOWER | Names written in Lower case |
| PROPER | Names properly entered |

    TRIM(text)
    UPPER(text)
    LOWER(text)
    PROPER(text)
    PROPER(TRIM(text))
    LEFT(text, FIND(" ", text, 1))

Excel file for the class : [Download here](https://github.com/user-attachments/files/20909715/Excel.Functions.2.-.Text-1.xlsx)
    
**Note:** Nesting - Putting a function inside another function or performing multiple functions in one. Use ctrl E for flash fill. Concatenation is the technical term for joining texts and you can use the concatenate function or the cells directly together with Ampersand (&).

## DAY FOUR (28 - 04 - 2025)

### Excel Functions 3 (Conditionals)

Today, we had a technical hitch in the transmission from the facilitator's end and another facilitator, by name Mr Femi, had to come in his stead to take us through the class. We took a more closer look on the conditional function (IF). We also looked at other Logical functions like AND & OR.

**IF function**

It allows you to test a condition, return a value if it is true and return another if it is false. It is a logical function that tests two or more conditions.

**OR function** 

It returns TRUE if any of its arguments evaluate to TRUE, and returns FALSE if all of its arguments evaluate to FALSE.

**AND function** 

It returns TRUE if all its arguments evaluate to TRUE, and returns FALSE if one or more arguments evaluate to FALSE

| Condition 1 | Condition 2 | AND | OR |
| :---------: | :---------: | :---: | :---: |
| True | True | True | True |
| True | False | False | True |
| False | True | False | True |
| False | False | False | False |

``` Microsoft Excel ```

    IF(logical_test, value_if_true, [value_if_false])
    OR(logical1, [logical2], ...)
    AND(logical1, [logical2], ...)
    
Excel file for the class : 1. [Download here](https://github.com/user-attachments/files/20910393/IF.Function.xlsx)
                           2. [Download here](https://github.com/user-attachments/files/20910420/Excel.Functions.3.-.Conditionals.xlsx)

## DAY FIVE (30 - 04 - 2025)

### Excel Functions 4(Dates)

Today, We looked at the different functions associated with Dates. The following functions were treated.

The TODAY function returns the serial number of the current date (shortcut - ctrl :)

The Now function returens the serial number of the current date and time. (shortcut - ctrl : ctrl shft ;)

The Text function converts a value to text in a specific number format

| Format | Text  |
| :---: | :----: |
| mmm | Aug |
| mmmm | August |
| ddd | Wed |
| dddd | Wednesday |

The IS functions checks the specified value and returns TRUE or FALSE depending on the outcome.

``` Microsoft Excel ```

    Today()
    Now()
    Year(serial number) brings back year in serial number
    Month(serial number) brings back month in serial number
    TEXT(value, format_text) - TEXT(E8, "mmmm") brings back the month in text format
    ISBLANK(value)
    ISNUMBER(value)
    ISTEXT(value)

**Note:** A volatile function is a function that updates itself as time progresses. Excel only recognises two things; a text and a number, and also behind the serial numbers of date there is a number that excel recognises which is counted from 01/01/1900. Always work with the data formatting control on the ribbon interface according to what is needed. 

Excel file for the class [Download here](https://github.com/user-attachments/files/20911208/Excel.Functions.4.-.Dates.xlsx)

## DAY SIX (01 - 05 - 2025)

### Excel functions 5(LOOKUPS)

Today, we went further to consider the LOOKUPS functions.

**VLOOKUP function - Vertical LOOKUP**

We use VLOOKUP when you need to find things in a table or a range by row.

In its simplest form, the VLOOKUP function says:

VLOOKUP(What you want to look up, where you want to look for it, the column number in the range containing the value to return, return an Approximate or Exact match – indicated as 1/TRUE, or 0/FALSE).

**XLOOKUP function - eXtended LOOKUP**

The XLOOKUP function searches a range or an array, and then returns the item corresponding to the first match it finds. If no match exists, then XLOOKUP can return the closest (approximate) match.

**Cell referencing in Excel**

- Relative referencing - Excel changes referencing accordigly as it moves.
- Absolute referencing - referencing does not change (shortcut : F4 or the dollar sign "$" before and after the cell column in the formula written).
- Column constant - column remains constant (shortcut : the dollar sign "$" before the cell column in the formula written).
- Row constant - row remains constant (shortcut : the dollar sign "$" after the cell column in the formula written).

| Referencing | Cell number |
| :--------: | :--------: |
| Relative | Y15 |
| Absolute | $Y$15 |
| Column constant | $Y15 |
| Row constant | Y$15 |

``` Microsoft Excel ```

    VLOOKUP (lookup_value, table_array, col_index_num, [range_lookup])
    XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]) 

Excel File for the class : [Download here}(https://github.com/user-attachments/files/20912580/Excel.Functions.5.-.LookUp.xlsx)

**Note** :  XLOOKUP function, an improved version of VLOOKUP that works in any direction and returns exact matches by default, making it easier and more convenient to use than VLOOKUP.

## DAY SEVEN (05 - 05 - 2025)

### Introduction to Pivot Table I

## DAY EIGHT (07 - 05 - 2025)

### Pivot Table II

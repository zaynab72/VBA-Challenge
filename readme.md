# Module-2-VBA
Module 2 Bootcamp Assignment - VBA Challenge
Course: Data Analytics Bootcamp 2024
Author: Murali C Veerabahu
Date: 11th Feb 2024

## Introduction
The files addressing the VBA Challenge for Module 2 has been uploaded to the GitHub Respository. All the necessary files have been uploaded to the repository and has been referenced in this 'ReadMe' file. The marking scheme for the assigment has been listed below and explation has been provided for each section. The VBA code for the project has been reproduced in this document for reference. 

## VBA Code Approach
The assigment calls for summary of an extensive stock dataset spanning over three years with each year in a seperate worksheet. The code helps create a summary table on each worksheet to give insight into the price changes and volumes for each stock.
The data was searched using 'FOR' loops and 'IF' conditionals. In class we used the condition '<>' 'Not equal', however I have used a '=' 'Equal' condtitional for the search. The code is divided into three parts:

- Part 1: Looping through the main table to create the summary table
- Part 2: Labelling and Formatting of summary table
- Part 3: Calculation of greatest values  

The excel file with the results was >100 MB and was not able to upload to GitHub, but the accompaying VBA code has been uploaded seperately as 'VBA Code.bas'. The screenshots of the three years of data has also been attached as '.png' files.

## Marking Scheme

### Retrieval of Data (20 points)
*The script loops through one year of stock data and reads/ stores all of the following values from each row:*
- *ticker symbol (5 points)*
- *volume of stock (5 points)*
- *open price (5 points)*
- *close price (5 points)*

Part 1 of the VBA code addresses this requirement by running through 'FOR NEXT' loops to summarise the data.

### Column Creation (10 points)
*On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:*

- *ticker symbol (2.5 points)*
- *total stock volume (2.5 points)*
- *yearly change ($) (2.5 points)*
- *percent change (2.5 points)*

The columns were created on the same worksheet. The data for this table was produced from Part 1 and the labells and formating was completed as Part 2.

### Conditional Formatting (20 points)
*Conditional formatting is applied correctly and appropriately to the yearly change column (10 points)*
*Conditional formatting is applied correctly and appropriately to the percent change column (10 points)*

The yearly change column was formatted with negative numbers having a red background and positive numbers being in green using a 'FOR NEXT' Loop and '.INTERIOR.COLORINDEX'. The values were also adjusted to two decimal places using '.NUMBERFORMAT'. To be able to see the data clearly the column widths were auto adjusted using '.COLUMNS.AUTOFIT'

Percent Change column was adjusted to two decimal places and percentage, again using '.NUMBERFORMAT'

### Calculated Values (15 points)
*All three of the following values are calculated correctly and displayed in the output:*

- *Greatest \% Increase (5 points)*
- *Greatest \% Decrease (5 points)*
- *Greatest Total Volume (5 points)*

This is done in Part 3 of the VBA code and again a 'FOR LOOP' is used to calculate all the three values.

### Looping Across Worksheet (20 points)
*The VBA script can run on all sheets successfully.*

To make the code run on all worksheets the 'FOP EACH .. IN' was used to rotate through all three years.

### GitHub/GitLab Submission (15 points)
*All three of the following are uploaded to GitHub/GitLab:*

- *Screenshots of the results (5 points)*
- *Separate VBA script files (5 points)*
- *README file (5 points)*

The screenshots of the three worksheets for the years 2018, 2019 and 2020 has been uploaded onto the GitHub repository as '.PNG' files.
The VBA script file 'VBA Code.bas' has been seperately uploaded. The Excel file with the main data was >100 MB in size and it was not permitted to by uploaded to GitHub. If required happy to provide the Excel file through other means.
The readme.md file has been uploaded to GitHub. 


## VBA Code for Assignment
```vb
'Module 2 Assignment - VBA Challenge
'-----------------------------------
'This code is split into 3 parts:
'Part 1: Looping through the main table to create the summary table
'Part 2: Labelling and Formatting of summary table
'Part 3: Calculation of greatest values


Sub ticker_summary()

For Each ws In Worksheets
    
    '---------------------------------
    'Part 1 :Creating of Summary Table
    '---------------------------------
    
    'Definition of Variables
    Dim counter As Integer
    Dim total_row, maxrow, minrow, maxvolrow As Long
    Dim ticker_symbol As String
    Dim total_stock As Variant
    Dim opening_year_price, closing_year_price, yearly_change, percentage_change  As Double
    Dim greatest_increase, greatest_decrease, greatest_total_volume As Double
        
    
    'Initial assignment of values to variables
    counter = 2
    total_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker_symbol = ws.Cells(2, 1).Value
    opening_year_price = ws.Cells(2, 3).Value
    
    'Looping through ticker symbol to create summary
    For x = 2 To total_row
                   
        If ws.Cells(x, 1) = ticker_symbol Then
            
            'Aggregates the volume of stock when the ticker symbol is the same
            total_stock = total_stock + ws.Cells(x, 7).Value
           
        Else
            'This would be at the change point when the ticker symbol does not match.
            'So outputs the calculated information to the summary table
            
            'Extraction of Ticker Symbol
            ws.Cells(counter, 9).Value = ticker_symbol
            ticker_symbol = ws.Cells(x, 1).Value
            
            'Output of calculated total stock
            ws.Cells(counter, 12).Value = total_stock
            total_stock = ws.Cells(x, 7).Value
            
            'Calculation and output of yearly change
            closing_year_price = ws.Cells(x - 1, 6).Value
            yearly_change = closing_year_price - opening_year_price
            ws.Cells(counter, 10).Value = yearly_change
            
            'Calculation and output of percentage change
            percentage_change = yearly_change / opening_year_price
            ws.Cells(counter, 11).Value = percentage_change
            
            'Reassignment of values for next run of loop
            opening_year_price = ws.Cells(x, 3).Value
            counter = counter + 1
        End If
         
    Next x
    
    '-------------------------------
    'Part 2: Labelling and formatting
    '-------------------------------
    
    'Column Labels for summary table
    ws.Range("i1") = "Ticker"
    ws.Range("p1") = "Ticker"
    ws.Range("j1") = "Yearly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("l1") = "Total Stock Volume"
    ws.Range("q1") = "Value"
    ws.Range("o2") = "Greatest % Increase"
    ws.Range("o3") = "Greatest % Decrease"
    ws.Range("o4") = "Greatest Total Volume"
    
    'Auto adjust column width for worksheet
    ws.Range("i:q").Columns.AutoFit
    
    'Formatting of yearly change column to 2 decimal places
    ws.Range("j:j").NumberFormat = "0.00"
    
    'Formatting to 2 decimal places and %
    ws.Range("k:k").NumberFormat = "0.00%"
    ws.Range("q2:q3").NumberFormat = "0.00%"
    
    'Colour Formatting of Yearly Change Column
    For x = 2 To counter - 1
        
        If ws.Cells(x, 10).Value < 0 Then
            'Negative values are red
            ws.Cells(x, 10).Interior.ColorIndex = 3
        Else
            'Positive values are green
            ws.Cells(x, 10).Interior.ColorIndex = 4
        End If
    
    Next x
    
    '----------------------------------------------
    'Part 3: Calculation & Output of Greatest Values
    '----------------------------------------------
    
    'Loops through summary table to find values
    For x = 2 To counter
    
        'Finding of greatest % increase
        If ws.Cells(x, 11).Value > greatest_increase Then
            greatest_increase = ws.Cells(x, 11).Value
            maxrow = x
        End If
        
        'Finding of greatest % decrease
        If ws.Cells(x, 11).Value < greatest_decrease Then
            greatest_decrease = ws.Cells(x, 11).Value
            minrow = x
        End If
        
        'Finding of greatest total volume
        If ws.Cells(x, 12).Value > greatest_total_volume Then
            greatest_total_volume = ws.Cells(x, 12).Value
            maxvolrow = x
        End If
        
    Next x
    
    'Output of greatest % increase - ticker symbol & value
    ws.Range("q2").Value = ws.Cells(maxrow, 11).Value
    ws.Range("p2").Value = ws.Cells(maxrow, 9).Value
    
    'Output of greatest % decrease - ticker symbol & value
    ws.Range("q3").Value = ws.Cells(minrow, 11).Value
    ws.Range("p3").Value = ws.Cells(minrow, 9).Value
       
    'Output of greatest total volume - ticker symbol & value
    ws.Range("q4").Value = ws.Cells(maxvolrow, 12).Value
    ws.Range("p4").Value = ws.Cells(maxvolrow, 9).Value

Next ws

End Sub
```


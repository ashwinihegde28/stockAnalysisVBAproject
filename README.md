# stockAnalysisVBAproject – VBA Scripting


## Overview of Project

The Project contains analysis of Stock data with VBA and refactoring the same to increase the efficiency. In Stock analysis Project, Steve wants to know differently Green Energy stocks have performed in two years 2017 and 2018 through their “yearly return” and is it worth of investing into them.
Through Visual Basic scripting the yearly increase/decrease in stock value for each ticker, the percentage change over the year, and the total volume for the year is calculated. The main goal is to achieve efficiency of the code through refactoring.
Refactoring is a key part of the coding process, and it makes the code more efficient — by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read without adding additional functionality. 

<br>



## Results
Post refactoring, code ran faster than it did in this module.
1.	Avoiding multiple loops reduces time to execute.
2.	When the subroutine is loaded a lot of logical code it can be divided into separate subroutine which can be called from the main subroutine as and when needed. This will not   only help in reusing the existing code instead of rewriting again in future, and cod looks neat too.
3.	Declaring and initializing is done in specified in a confined place so that they do not look distribute.
4.	All the above three was achieved without adding any additional functionality but with the existing logic only.


### Brief Analysis of the refactored code:
The entire code is divided into three subroutines to make it simpler
1.	ClearWorksheet() subroutine: 
 -  It clear the previous data on the "All Stocks Analysis" when clicked on cancel button or called from the any subroutine.
 ```
 'This Subroutine will clear the data in the worksheet
 
    Sub ClearWorksheet()
    
    Worksheets("All Stocks Analysis").Activate
        Cells.Clear    
        
    End Sub

 ```
<br>

2.	yearValAnalysis() subroutine: 
 -	Calls "ClearWorksheet" subroutine to delete existing data on the worksheet. 
 -	The input year is recorded and stored in the variable “yearValue”. Here we will do a small check for the input value is either 2018 or 2017 so that the code doesn’t land in runtime error for different values.
 -	 It passes this valid “yearValue” to the main subroutine “AllStocksAnalysisRefactored”.
	 ```
   	  'This subroutine will check for the correct year input and calls the subroutine allStocksAnalysis for the year given
      
      Sub yearValAnalysis()
            
               'Clear the prevous values from the Sheet
                ClearWorksheet
                
                Range("A1").Value = "All Stocks (" + yearValue + ")"
            
                'Display the text in the inputbox
                yearValue = InputBox("What year would you like to run the analysis on?")
                
                   'if the year is not entered than default year notify the user else runtime error occurs
                    If yearValue = "2018" Or yearValue = "2017" Then
                       AllStocksAnalysisRefactored (yearValue)
                    Else
                        MsgBox ("The year must be either 2018 0r 2017")
                    End If
            
        End Sub
 	 ``` 
<br>

3. formatAllStocksAnalysisTable() Subroutine:
 - This subroutine includes all formatting including number formating and cells formating.
 ```
  'formatAllStocksAnalysisTable function automate the formating our xlsm file
  
    Sub formatAllStocksAnalysisTable()
        
        'DEclare and set the value for the start and end rows for Tickers data.
        dataRowStart = 4
        dataRowEnd = 15
                
        'Selecting the workSheet for the formating
         Worksheets("All Stocks Analysis").Activate
         
        'Set the font columns from A3:C3 to BOLD
        Range("A3:C3").Font.FontStyle = "Bold"
        
        'Applying border for the contents
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        'Number formating for Total daily volume and the range columns
         Range("B4:B15").NumberFormat = "#,##0"
        
         Range("C4:C15").NumberFormat = "0.0%"
         
         Columns("B").AutoFit
        
        'iterate over the outcomes and If the value is negative(fails) than set the cell to red else green
         For i = dataRowStart To dataRowEnd
         
            If Cells(i, 3) > 0 Then
            
                Cells(i, 3).Interior.Color = vbGreen
            
            Else
        
                Cells(i, 3).Interior.Color = vbRed
            
            End If
        
        Next i
       
    End Sub
 ```

<br>

4.	AllStocksAnalysisRefactored(yearValue As String): This subroutine holds the main logic of the code that is refactored. This description includes the refactored code.For refactoring, few changes are incorporated as follows:

 -  The variables declaration and initiation are done at the beginning and then followed by the logic to make it more readable.

 -  This subrutine accepts the yearValue from subroutine yearValAnalysis() 
 
 -	Variable “tickerIndex” is ceated to hold the index of the ticker values and initialize it to 0.
~~~
 'variable to hold the index of each ticker and initialized to 0 before operating on it
        tickerIndex = 0        
~~~    
 - Declared three output arrays for storing Volume,Starting Price and Ending price for each ticker
~~~ 
'Declare Arrays to output of Volume,Starting Price and Ending price for the ticker
        Dim tickerVolumes(11) As Long
        Dim tickerStartingPrices(11) As Single
        Dim tickerEndingPrices(11) As Single
~~~
 - Initailise the array tickerVolumes to 0
~~~
'Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
        Next i
~~~
 - For each row 
  1. Increase volume for current ticker and add ticker volume for the current stock ticker
~~~
'Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
                

            'Increase volume for current ticker and add ticker volume for the current stock ticker
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
             
~~~
   2.  For calculating "tickerStartingPrices", the current row should be the first row with the selected tickerIndex.
~~~
'Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
             
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
             End If
~~~
   3.  For calculating "tickerEndingPrices" current row must be the last row of chosen ticker
~~~
 'check if the current row is the last row with the selected ticker and assign current ending price
            
             If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
             
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
             End If
~~~
   4.  The "tickerIndex" should be incremented when above operation for the chosen ticker is done and to start for the new ticker and thus the loop ends when all the tickers in the rows are targetted.
~~~
'If the next rows ticker does not match, increase the tickerIndex.
             If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerIndex = tickerIndex + 1
                
             End If
        
         Next i
~~~
 - Now iterate through each tickers counts one by one and set the Ticker name,Total Daily Volume, and Return
    ~~~
     ' Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
         For i = 0 To 11
            
            'Activate the worksheet where the output must be displayed
             Worksheets("All Stocks Analysis").Activate
             
             Cells(4 + i, 1).Value = tickers(i)
             Cells(4 + i, 2).Value = tickerVolumes(i)
             Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
             
             
          Next i
    ~~~

 - For formatting the code “formatAllStocksAnalysisTable” subfunction is called so that all the formatting code is placed in one place and not mixed with the actual logic. In future formatting enhancement codes for can be added there.
~~~ 
'For formatting call the subroutine "formatAllStocksAnalysisTable" 

formatAllStocksAnalysisTable
~~~
 - The output of both the codes for the year 2017 is showcased below and refractored code executes more faster than non refractered    


 - The output of both the codes for the year 2018 is showcased below. 
 ![VBA_Challenge_2017](https://github.com/ashwinihegde28/stockAnalysisVBAproject/blob/master/resourse/VBA_Challenge_2017.PNG) <br> Refactored VBA_Challenge_2017
 
 <br>
 
 ![VBA_Challenge_2017](https://github.com/ashwinihegde28/stockAnalysisVBAproject/blob/master/resourse/greenStoclAnalysis_2017.PNG) <br> GreenStockAnalysis_2017
 
 
 - From the Screenshots we conclude that the refractored code always takes lesser time to execute

![VBA_Challenge_2018](https://github.com/ashwinihegde28/stockAnalysisVBAproject/blob/master/resourse/VBA_Challenge_2018.PNG) <br> Refactored VBA_Challenge_2018
 
 <br>
 
 ![VBA_Challenge_2018](https://github.com/ashwinihegde28/stockAnalysisVBAproject/blob/master/resourse/greenStoclAnalysis_2018.PNG) <br> GreenStockAnalysis_2018


<br>


## Summary
### Advantages and disadvantages of refactoring code in general
#### Advantage of refactoring code
- Refactored code is always clean, readable and easy to understand. Any one with basic coding knowledge use it for future enhancement. Hence easy to add/modify or debug and fix the issues or build larger applications with it.
- It only alters the existing code base but do not alter the functionality so its safe to practice
- It helps to reuse the code without rewriting hence saves time.
- Increases the efficiency of 
 
#### Disadvantages of the refactored VBA script
- Refactoring code always takes longer time.
- Refracting code written without the knowledge of the project can alter the purpose of project.
- Even a small or simple mistake may render application unstable.
- Refactoring requires a thorough testing of entire functionality.
  <br>


 
### Advantages and disadvantages of the original and refactored VBA script
  #### Pros 
- Existing VBA code was simple, short and easier to understand hence the refactoring was not difficult.
- It already serves the purpose of automating the analysis.
- Cleaning the code was not at all time consuming.
- Since the code was already executed without errors and produced the correct output we need not have to test for its functionality.

 #### Cons
- When the code set is stable and the efficiency is sligtly different compared with other, refractering is not a must. This saves money,resources and time.

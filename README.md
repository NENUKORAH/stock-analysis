# Stock Analysis With Excel and VBA

## Overview of project

This project entails preparing a workbook for Steve to help him analyse a handfull of green energy stocks in addition to DAQO's new energy Corp stock.

Steve just graduated with a finance degree, he wants to help his parents who are passionate about alternative energy in making an investment decision.

### Purpose

The purpose of this challenge is to help Steve automate his analysis on the stock market. I have provided a template for him to analyse a list of alternative energy stocks for the year 2017 and 2018.

This challenge focuses on refactoring the Visual Basic for Applications (VBA) code which I have built for Steve in the excel workbook I provided to him earlier.

Refactoring entails improving the efficiency of a code by taking fewer steps, using less memory, or improving the logic of the code to make it easier for the future users to read.

The output of the refactored code will be the same as the first code. However, the run time of the query will improve considerably.

The updated code will allow Steve run his analysis faster with a click of a button. The same code can be used by Steve in analysing a larger set of data in the future.

This analytical tool will help Steve in making an investment decision quicker.

## Results

### Stock Performance 2017

The stock performance for the year 2017 accross the stocks is quite impressive. Only TERP stock had a negative return out of the 12 stocks analysed.

DQ, ENPH, FSLR, SEDG had a positive return of over 100% in 2017. This makes investment in DQ a viable option looking at the 2017 data.

The result shows that 92% of the stock analysed has a positive return.

Below is a screenshot of the stock performance for 2017;

![Stock_performance_2017png](https://user-images.githubusercontent.com/81701640/116452195-6f578980-a82b-11eb-85c1-2b9b29e1b40c.png)

### Stock Performance 2018

The stock performance for the year 2018 accross the stocks is significantly poor. Only the ENPH and RUN stock had a positive return out of the 12 stocks analysed.

DQ and JKS had a negative return of over 50% in 2018. The only viable investments using 2018 data is ENPH and RUN.

The result shows that 83% of the stock analysed had a negative return.

Below is a screenshot of the stock performance for 2018;

![Stock_performance_2018](https://user-images.githubusercontent.com/81701640/116452475-c5c4c800-a82b-11eb-924f-781bc8cc7450.png)

### Execution time of the Script

The result of the refactored code showed a significant improvement over the original code. The run time dropped by almost one minute for both years.

In the intial code provided to Steve, there was a code run time of 1.23 seconds and 1.21 seconds for 2017 and 2018 respectively.

Following the refactoring of the script, the code run time for 2017 and 2018 was 0.25 Seconds and 0.25 Seconds for 2017 and 2018 respectively.

The execution time for the original 2017 script is shown below;

![VBA_Challenge 2017 before Refactor](https://user-images.githubusercontent.com/81701640/116454035-89926700-a82d-11eb-85ec-b0ad8f4d52c8.png)

The execution time for the refactored 2017 script is shown below;

![VBA_Challenge 2017 after Refactor](https://user-images.githubusercontent.com/81701640/116454263-cb231200-a82d-11eb-9efd-840bb915125f.png)

The execution time for the original 2018 script is shown below;

![VBA_Challenge 2018 before Refactor](https://user-images.githubusercontent.com/81701640/116454334-dece7880-a82d-11eb-8960-808c233d841f.png)

The execution time for the refactored 2018 script is shown below;

![VBA_Challenge 2018 after Refactor](https://user-images.githubusercontent.com/81701640/116454393-f0b01b80-a82d-11eb-81b3-0f88a7333187.png)

### Code for the Analysis

I made some changes to the original code to arrive at the refactored code. 

The modified portion of the code is highlighted below;

```vbscript

'1a) Create a ticker Index
    'New variale for Tickerindex
    
    tickerindex = 0
       
    '1b) Create three output arrays
    'Output arrays defined
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
       
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'Initializing tickervolumes to zero
    
       For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    
    Next i
        
       ''2b) Loop over all the rows in the spreadsheet.
       'Loop through rows in the data
       
       For i = 2 To RowCount
              
              '3a) Increase volume for current ticker
              'Get total volume for current ticker
              
             tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
           
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
           If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
        
        End If

        '3c) check if the current row is the last row with the selected ticker
        'If  Then
        
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerEndingPrices(tickerindex) = Cells(i, 6).Value
         
         End If
            
           '3d Increase the tickerIndex.
            
         If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                tickerindex = tickerindex + 1
            
            End If
           
       Next i
          
   '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
       
       For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
   Next i

   ```

## Summary

1. What are the advantages or disadvantages of refactoring code?

### Advantages of Code Refactoring 

* One major advantage of code refactoring is that it improves the code efficiency by eliminating complex instructions or reducing the statements in the code
  into a smaller readable format.
* Another advantage of code refactoring is that it makes the code run faster. This saves time when running codes that takes time to process, refactoring the code will
  make the code run for a shorter peroid of time.
* Debugging becomes easier when a code is refactored.

### Disadvantages of Code Refactoring 

* The result might significantly defer when a code is wrongly refactored.
* Refactoring a code can be very time consuming.
* The process of improving the efficiency of an existing code might lead to an introduction of a bug in the code.

2. How do these pros and cons apply to refactoring the original VBA script?

  Firstly, the pros listed above all applies to the refactored VBA code used in this project. The run time improved significantly as seen in the screenshots shared in the analysis of result, 
  the code is more readable and understandable, debugging an issue in the refactored code is also easier compared to the original code.
  
  Secondly, there are lots of other advantages not listed above. The benefits of refactoring cannot be overemphasized as it helps in restructuring and improving the code quality without
  changing the result.

  Finally, for every advantage there is a disadvantage. The main applicable cons from the ones listed above is the fact that it is time consuming. I spent hours trying to understand what
  changes to make and also to make meaning of the refactored code. Though the result did not defer from the original code, this is a disadvantage that might arise if a code is not
  properly refactored. A bug can be introduced to a code if not properly refactored, this does not apply to this refactored code.

  **Nnaemeka Enukorah**
  
  **29th April, 2021**
 

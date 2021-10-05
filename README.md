# stocks-analysis
## Project Overview

The purpose of this project is to create an analysis of the entire stock market by refactoring code for Steve. Steve initially wanted to look into a dozen stocks as reference for his parents who are interested in making an investment. The original code worked well for Steve and made it easier for him to navigate the given stocks. However, Steve would like to continue to research more stocks for his parents and wants to have the dataset incorporate the entire stock matket.

## Results

### Year: 2017
After refactoring the script in VBA the stock performance was significantly better than the performance of 2018. The majority of stocks that year had a greater total volume which indicates how often a stock was shared. We can also look at the percentage of yearly returns to see if the price increased or decreased, and looking at the dataset we can conclude that the majority of the stocks in 2017 had an increase in price. The only stock that did not do well was the TERP stock which had dropped -7.2%.

![VBA_2017_Chart](https://user-images.githubusercontent.com/90741799/135965491-76a76ec4-0433-4198-afda-1fa3c4d31c2b.png)


### 2017 Code Time: 



Now looking into whether or not the code ran quicker compared to 2018 results should indicate how much the computer had to work to process the data. Seeing that the code ran in less than a second showed that it was easier to compute.


![2017_Code_Time](https://user-images.githubusercontent.com/90741799/136038644-89dd4a9d-4bbb-4b9c-8a19-d2a82fe8eabd.png)


### Year: 2018
VBA 2018 Chart: 
Comparing 2018's total volumes dataset to 2017's dataset, we can conclude that the volume had a significant drop in the results. The majority of the stocks had decreased by a few million which tells us that the stocks were not shared as often as it had been in 2017. ENPH and RUN were the only stocks that performed notably well even though in 2017, RUN only had 5.5% in returns.

![VBA_2018_Chart](https://user-images.githubusercontent.com/90741799/136038071-2c92a31a-f3db-4638-9923-f931705eae93.png)


VBA 2018 Time:

The 2018 VBA script took a little over 2 seconds to compute which indicates that the computer had to do more work in order to get the results.


![2018_Code_Time](https://user-images.githubusercontent.com/90741799/136038684-8d3a4918-5441-42d6-85fd-f6cee8544404.png)
 
Comparing the results of both years I would suggest RUN as the best option to invest in for Steve's parents. RUN was the only stock that performed well in both 2017 and 2018 while showing an increase in its returns. 
## Summary

### Advantages of Refactoring code:
Refactoring code could be an advantage by reducing the steps taken within the code which will make it less time consuming. Another advantage of refactoring code could be making the code more organized which could also make it easier for others to read. This provides a better use of whitespace as well as a way to help remember your last steps in the code because it's quite easy to forgot where you last left off. Adding comments is a good way to use whitespace for the purpose of making a code more readable.

        '1) Format the output sheet on All Stocks Analysis worksheet
          Worksheets("All Stocks Analysis").Activate
            Range("A1").Value = "All Stocks (2018)"

              'Headers in these specific cells A3,B3,C4
                    Cells(3, 1).Value = "Ticker"
                    Cells(3, 2).Value = "Total Daily Volume"
                    Cells(3, 3).Value = "Return"





### Disadvantages of Refactoring code:
The main disadvantage of refactoring code would be overlooking simple mistakes due to the function of the code not changing. This could also lead to it being a time consuming task because the mistake could be a missing text that was overlooked. For example, when writing the for loop of the ticker, the formula used the same word but one was without an s:  It took some time and another pair of eyes to point out why there was an error in the VBA script.


     "If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickersIndex) Then
    
    
      tickerEndingPrices(tickerIndex) = Cells(i, 6).Value". 
      
     'Fixed Variable 
    "If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    
    
      tickerEndingPrices(tickerIndex) = Cells(i, 6).Value".

### Advantages of the orginal & refactored VBA script:
The advantage of the original VBA script was the amount of stocks that were used. There were only twelve stocks that were used in the original script which made the coding process slightly easier. The orginal script always had...


    'Initialize array of all tickers
    Dim tickers(11) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    


The advantage of the refactored VBA Code was that there was a broader dataset but it took less time to analyse. Looping through the larger dataset could have been more difficult due to the data size but because of nested looping it much easier to execute.


### Disadvantages of the orginal & refactored VBA script:

The disadvantage of the original VBA script was the fact that it was solely focused on a small bit of information. Steve wanted to look for the best stock to invest in for his parents, if we only stuck with the original dozen there would have been less to compare to.

A disadvantage of the refactored VBA script was that there was an increase in information needed in order for VBA to specifically operate in the way we needed it too. The increase in volume and  looping is an example of taking an extra step to insure VBA included the additional information.

 
    '2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
     Next i
            
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
   
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
       If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If

            
            
       
        
        '3c) check if the current row is the last row with the selected ticker
         'If Then
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     
     
     '3d Increase the tickerIndex.
     
     
     tickerIndex = tickerIndex + 1
    
    End If
    Next i  
    

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i


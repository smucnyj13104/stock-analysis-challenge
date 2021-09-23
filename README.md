# stock-analysis

## Overview of Project

### Purpose

- The purpose of this project was edit the AllStocksAnalysisRefactored code in order to loop through the entire stock dataset and output the total daily volume and return for all stocks using more efficient code than what was written previously, as evidenced by faster completion times. The output of the new code should match the output of the old code.

## Results
- The old code was inefficient at determining stock returns and volumes, because it had to loop over the entire dataset multiple times in order to get the necessary information.
- The new code was therefore designed to only have to loop through the dataset once in order to get the necessary information.
- The refractored code is provided below, including comments.

    '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        '3a) Increase volume for current ticker

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
            '3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If

    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = (tickerEndingPrice(i) / tickerStartingPrice(i)) - 1
    Next i
    

- The new code successfully matched the output of the old code for both the year 2017 and the year 2018 as both the total daily volume and % return output was the same.
- The new code was substantially faster. The new code ran in 0.09375 seconds for the year 2017 (VBA_Challenge_2017.png?raw=true), and in .09375 seconds for the year 2018 (VBA_Challenge_2018.png?raw=true). The old code ran in .515625 seconds for the year 2017 (VBA_Challenge_oldcode_2017.png?raw=true) and in .546875 seconds for the year 2018 (VBA_Challenge_oldcode_2018.png?raw=true).
- What are two conclusions you can draw about the code results?
	1. The new code is 5 times faster than the old code because it only has to loop through the data once.
	2. The new code matches the results of the old code. Analyzing the stocks, it looks like 2017 was a better year than 2018 for these stocks overall as more of them were green in 2017 than 2018. In 2017 DQ performed the best (199.5% return), but in 2018 RUN performed the best (84% return).
	
## Summary
- Potential advantages of refractoring code are: 
	1) Improving its efficiency and therefore its run time.
	2) Improving its flexibility so it can run on more datasets.
	3) Improving how easy it is for other programmers to read and understand the code. 	
- Potential disadvantages of refractoring code are:
	1) It is not always easy to understand someone else's code. If the logic is not clear, refractoring the code might be difficult if not impossible, and the person doing the   	refracting might inadventently change the functionality of the code.
- As applied to the current project, it was clear that refractoring the code improved the efficiency of the old code, as the new code ran ~5 times faster. The only disadvantage I can see from refractoring it is if the previous programmer (the one who wrote the original code) saw the refractored code, they may not immediately understand how it works. With comments in the code, however, I would think it would be fairly easy to figure out the refractored code.

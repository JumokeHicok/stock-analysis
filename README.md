# Stock Analysis
### Overview of Project
The purpose of this project was to determine if it was possible to refactor our original code so that it would work faster for analyzing the entire stock market instead of just 12 stocks.
### Results
In order to make the code run faster I refactored the original code to not include the nested loop when going through the rows of data.  Instead of going through all of the rows multiple times (in this case 12 times, once for each ticker), I created an index variable and used that variable to store the data within the 3 output arrays.  This allowed the code to run through the lines just one time while capturing the data for all of the tickers at once.  

##### Original code:
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0      
        Sheets(yearValue).Activate
        
        For j = 2 To RowCount
           If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
           
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
        
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i


##### Refactored code:
    For i = 0 To 11
    ticker = tickers(i)
    tickerVolumes(tickerIndex) = 0
    Next i

    For i = 2 To RowCount   
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    Next i
The images below show that running the refactored code for both years is much faster than running the original code.

                            
##### Original output:
![2017 Original](/Resources/2017_Original.png)  ![2018 Original](/Resources/2018_Original.png)

##### Refactored output:
![2017 Refactored](/Resources/2017_Refactored.png)  ![2018 Refactored](/Resources/2018_Refactored.png)

### Summary

One advantage of refactoring code is that it can be an easier starting point than writing the code from scratch.  A disadvantage is that depending on how many comments are included with the code, it could be hard to follow someone else's original intent. In this case we did actually write the original code so the advantage and disadvantage did not apply.
  
    
    

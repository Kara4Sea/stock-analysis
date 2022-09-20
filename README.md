# Stock Analysis
An analysis of client stocks
## Overview of Project
### Purpose
The purpose of this analysis is to determine the performance of thirteen stocks in the client's portfolio. The total daily volume and return percentage are calculated to evaluate each stock.
## Results
### 2017 vs 2018 Performance
2017 saw far greater returns than 2018. All of the client's stocks in 2017 had positive returns, with DQ at the largest return of 199.4%. One stock saw a negative return in 2017, TERP at -7.2%. 

In 2018 all but two stocks had negative returns. DQ, which had the highest return in 2017, fell dramatically to -62.6%. TERP remained to reflect negative returns in 2018 with slight improvement with returns at -5.0%. ENHP and RUN were the only stocks with positive returns in 2018. RUN showed improvement with a return of 5.5% in 2017, which grew to 84.0% in 2018. While ENPH decreased from 129.5% in 2017 to 81.9% in 2018.

![2017_All Stocks](https://user-images.githubusercontent.com/110419577/191360506-d74dbaef-3af6-40ee-81cd-22c85ebfd71b.png)![2018_All Stocks](https://user-images.githubusercontent.com/110419577/191360536-7abb8885-b566-4350-9a9c-07724016e17b.png)

### Execution Times - Original vs. Refactored Script
The execution time in the refactored script was far improved when pulling data for both 2017 and 2018. I believe the reason for this improvement is due to reducing the number of loops required for the same data to be output.

#### Original Script
In the original script, we used the first for loop to loop through all the tickers by using the following code. 

`For i = 0 To 11`

    `ticker = tickers(i)`

Then in the nested loop, we used If, Then to get the total volume for the current ticker.

`For j = 2 To RowCount`
           
           `If Cells(j, 1).Value = ticker Then`

            `totalVolume = totalVolume + Cells(j, 8).Value`

           `End If`
           
We used the same nested loop to establish two more If, Then to get the starting price and ending price for the current tracker.

            `If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then`

            `startingPrice = Cells(j, 6).Value`

           ` End If`
           
           
           `If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then`

          `endingPrice = Cells(j, 6).Value`
          
            `End If`

      ` Next j`

#### Refactored Script
We eliminated the needs for a second loop by creating a variable for tickerIndex, and creating three output arrays for tickerVolumes, tickerStartingPrices and tickerEndingPrices.

` Dim tickerIndex As Long
    
    tickerIndex = 0

    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single`
    
 We then established a loop to run through all of the arrays in the original ticker variable. and used If Then to iterate throug the tickerIndex to pull for each assinged number in the array starting with 0 for AY. 
 
 ` For i = 0 To 11
    
    tickerVolumes(i) = 0
    
    Next i

   For i = 2 To RowCount
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value`
        
 Then to determine tickerStartingPrices, we checked if the current row was the first row of the selected tickerIndex.
 
 ` If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then

        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
         End If`
         
 Similarly to determine tickerEndingPrices, we checked if the current row was the last row with the selected ticker.
 
 `If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value`
            
We also increased the tickerIndex to allow it to run through all in the array, before ending the If.
            
            `tickerIndex = tickerIndex + 1
    
         End If`
         
By reducing the number of loops information for 2017 improved from 0.8554688 seconds with the original script to 0.1640625 seconds with the refactored script. For 2018, we improved from with the original script to 0.125 seconds for the refactored script.

* Refactored:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/110419577/191367023-cea72877-b552-4989-96a5-bb3b3722fe48.png)![VBA_Challenge_2018](https://user-images.githubusercontent.com/110419577/191367043-4f9f7d95-85eb-4ab7-9ffa-1250e7e4bcca.png)

* Original:

![VBA_Challenge_2017_original script](https://user-images.githubusercontent.com/110419577/191367087-380953d9-115a-467b-b749-1981d6edd4bd.png)![VBA_Challenge_2018-original script](https://user-images.githubusercontent.com/110419577/191367096-542cd949-8f7c-4f2b-b33e-dab1f428b370.png)

## Summary

### Advantages vs. Disadvantages of Refactoring Code

### Pros and Cons applied to refactoring VBA script







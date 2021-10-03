# VBA Stock Analysis

## Overview
In this project i have been creating stock analysis scripts in VBA to review data on 12 Green energy stocks. My goal is to create a script that can quickly review the stock data producing the annual volume for the stocks along with the return after one year. The code should also be able to format the results nicely so it is easily read. 

## Stock Analysis
The results from the first script created are listed below. 

![image](https://github.com/JonLev03-hub/Stock-analysis/blob/main/2017%20Stock%20analysis.png)
![image](https://github.com/JonLev03-hub/Stock-analysis/blob/main/2018%20Stock%20analysis.png)

These images show the annual volume and return for the stock data that was reviewed. As you would notice all the stocks in this list performed much better in 2017 than 2018. This was likely caused by the conditions of the stock market or the conditions of the green energy sector itself. When looking at the results from this analysis you must understand that these stocks are in develouping industries so price changes can be spontanious. One thing you might notice from the results is that in 2018 -- A year when the green energy has majority negitive returns -- stocks ENPH and RUN maintained a positive return, and a considerable large one at that. This could lead some to believe that they will have more stability in the market overall. Another trend that should be noted is that the stocks that perform well tend to have higher volumes. 

## Better analysis performance
When first creating the script to compute this data it was for 12 stocks total, and since it was for such a small sample the performance of the script didnt mean much, but to perform analysis on larger datasets I decided that it was important to speed up the processes. To do this I refined the code to cut out unnecesary processes. In the first version it would run through the data 12 times, and each time it read the data it would collect the information about one stock, but since the data was nicely formatted in alphabetical order it was easy to adjust that to look through the data one time by telling the computer when the ticker switches in the data save the information into a new position in an array. The method of doing this is shown below.

###### Example 
      ''1a) Create a ticker Index and output arrays
    tickerindex = 0
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrice(11) As Single
    Dim tickerEndingPrice(11) As Single
    
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
    
        ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        ''3a) Increase volume for current ticker
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8)
        
        ''3b) Check if the current row is the first row with the selected tickerIndex.
        If (Cells(i, 1) <> Cells(i - 1, 1)) Then
        tickerStartingPrice(tickerindex) = Cells(i, 6)
        
        ''3c) check if the current row is the last row with the selected ticker
        ElseIf (Cells(i, 1) <> Cells(i + 1, 1)) Then
        tickerEndingPrice(tickerindex) = Cells(i, 6)
        tickerindex = tickerindex + 1
        
        End If
    Next i
    
    
This cut the processing time in almost a tenth because it only went through the data 1 time instead of 12. You will see that the first version performed the analysis in approximately .6 seconds, and the second version performed it in approximately .07 seconds. Below I have the results and performance of the second version on the 2017 and 2018 datasets. 
![image](https://github.com/JonLev03-hub/Stock-analysis/blob/main/2017%20Refined%20Stock%20Analysis.png)
![image](https://github.com/JonLev03-hub/Stock-analysis/blob/main/2018%20Refined%20Stock%20Analysis.png)

## Potenital Deeper analysis 
When looking at the results from the first analysis performed on the stocks you are able to find two stocks that may be more likley to perform well and there was a correlation found between stocks with high volumes and positive returns; but its definitely possible to do deeper analysis and find more correlations in this dataset. For further analysis i would consider looking into the correlation between the 2017 volume and the 2018 performance, and you may also be able to find substantial correlations for the stability of the stocks by looking not solely at the yearly return but how quickly or slowly they gained that return. I hypothesize that you would be able to find correlations between the speed of growth and consistency of growth. 

#Stock_Analysis with VBA

## Overview of Project:

In this project, a friend of mine Steven needs my help to analyze stocks data from green energy companies so that he can assist his parents whether investing on one or more of these stocks are reasonable. 
In order to establish a quick and user friendly tool that he might want to use it in future as well, I decided to utilize my VBA skills create a macro that will help us to determine which of these stocks are worthful. 

## Purpose

In this project, I have created an analysis by utilizing elements of VBA such as Arrays, Loops, Conditionals and assigning variables with correct data types. After executing the macro, I have seen that it takes around 0.7-0.8 seconds which seem quite long time for a small data set. Therefore, I decided to dive into refactoring to code and see if it can get any faster. 

## Results 

### Stock Performance by Year 

In 2017 " DQ " return was the highest with 199% and 35,796,200 total daily volume, however in 2018 return was -62% and 107,873,900 total daily volume. This would help Steve to understand having high volume doesn't necessarily mean high return. 

On the other hand, in 2017 "ENPH" return was 129.5% with 221,772,100 total daily volume and in 2018 it was 81.9% with 607,473,500 total daily volume. 

<p align="center"><img width="778" alt="Screen Shot 2022-02-05 at 5 47 47 PM" src="https://user-images.githubusercontent.com/98676400/152662805-6b55682e-feb9-4109-bd74-77467825c9d1.png"><p align="center">Image 1- ENPH and DQ Performance Compared by Years </p></p>

Thus, we can conclude that investing on "ENPH" for his family would be a better choice. I will also use following outcomes from my Macro to explain my analyze to Steve. 
<p align="center">
<img width="344" alt="2017_Performance" src="https://user-images.githubusercontent.com/98676400/152662725-cde1ceaf-3a61-4ca4-8908-3c4d2907aa67.png">

<img width="343" alt="2018_Performance" src="https://user-images.githubusercontent.com/98676400/152662730-d71cbd19-a70a-4213-bffd-59a27b151c0c.png">
  <p align="center">Image 2- All Stocks Performance Compared by Years </p>
</p>

### Performance Comparison between Initial and Refactored Macro

In order to speed up the initial macro, I needed to find a better way than using nested for loop so that I created a variable called "tickerIndex" which will be able to access correct index across the four different arrays as following :

* tickers 
* tickerVolumes
* tickerStartingPrices
* tickerEndingPrices

By using "tickerIndex" variable the array would run :

* tickers(tickerIndex)= ("AY", ... ,"VSLR")
* tickerVolumes(tickerIndex)=("totalVolumes_AY", ... , "totalVolume_VSLR")

Refactoring Code as follow :

Dim tickerIndex As Single
    tickerIndex = 0

    '5b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '6a) Initialize ticker volumes to zero
        
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    '6b) loop over all the rows
    
    For i = 2 To RowCount
    
        '7a) Increase volume for current ticker
       
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '7b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '7c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '7d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i

After Executing Initial and Refactoring codes, run time shows significan decerease  : 





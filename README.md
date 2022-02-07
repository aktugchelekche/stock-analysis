# Stock_Analysis with VBA

## Overview of Project:

In this project, a friend of mine Steven needs my help to analyze stocks data from green energy companies so that he can assist his parents whether investing on one or more of these stocks are reasonable. 
In order to establish a quick and user friendly tool that he might want to use it in future as well, I decided to utilize my VBA skills create a macro that will help us to determine which of these stocks are worthful. 

## Purpose

In this project, I have created an analysis by utilizing elements of VBA such as Arrays, Loops, Conditionals and assigning variables with correct data types. After executing the macro, I have seen that it takes around 0.7-0.8 seconds which seem quite long time for a small data set. Therefore, I decided to dive into refactoring to code and see if it can get any faster. 


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

####  Run-Time Of  Initial Code For 2017 and 2018

<p align ="center">
<img width="300"  src="https://user-images.githubusercontent.com/98676400/152704832-2f71dc35-8b5e-41b1-88db-ff6bfc056b02.png">

<img width="300"  src="https://user-images.githubusercontent.com/98676400/152705236-3a29f86a-3b5b-43e3-91e3-87cb5cae5300.PNG">
                                                                                                                             

</p>                                                                                                                         

####  Run-Time Of ReFactored Code For 2017 and 2018

<p align ="center">
  <img width="300"  src="https://user-images.githubusercontent.com/98676400/152705178-4a9ff081-3e56-44e3-8369-eae054473a16.png">
                                                                                                           
  <img width="300"  src="https://user-images.githubusercontent.com/98676400/152719246-64321706-513d-4c54-a469-1b34a1eaffe0.png">
                                                                                                                                                </p>   
                                                                                                                           
Upon completlation of execution, there is a significant decrease on runtime between the initial and refactored code. 

## Results

### Advantages and Disadvantages of Refactoring code 

###### Advantages
Nowadays, artificial intelligence and machine learning are two imperative fields that leading the entire world in any industry. Industries are always looking for the most efficient and quick application to implement on their system. We can create a algorithm with a code that working properly but speed and efficiency can be improved always. While we want to make improvement, need to reduce the amount of time as well. Refactoring an initial code and make improvement would help us to achieve both of our goals which are time and efficiency . 

###### Disdvantages     

As it was mentioned on advantages, we are seeking for redecing amount of time, however, we might face with a code that understading and refactoring it might take much longer time than creating a fresh new code. Therefore before making decision, we have to look for cons and pros than do the math if refoctraing is the path we want to take .

###### Advantages and Disdvantages original and refactored VBA script 

I am going to refactor my two statements above and connect to this part. The advantages of usiung YearValue code then refactoring it was saved me fair amounth time. There are too many lines that take while to type such as the "tickers" array. On the other hand, I spent quite time to understand and find out the logic behind of the "tickerIndex". I didnt remove my nested for loop at the beginnng of refactoring and I could not figure out why my code was not working. After consulting with " google " and classmate I understood that the whole idea of utilizing that variable is to remove nested loop and use arrays in lieu of it. 
Spending that much time to understand the reasoning of refactoring code would worth it if this code will be used for long term but if this is a code for a short term plan, I would reconsider using initial one rather than refactored . 

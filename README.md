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

<table >
<tr>
<th>
<table  >
<caption>2017</caption>
  <tr>
    <th>Stock</th>
    <th>Total Daily Volume</th>
    <th>Return</th>
  </tr>
  <tr>
    <td>DQ</td>
    <td>199%</td>
    <td>35,796,200 </td>
  </tr>
  <tr>
    <td>ENPH</td>
    <td>129.5%</td>
    <td>221,772,100 </td>
  </tr>
</table>
</th>
<th>
<table  >
<caption>2018</caption>
  <tr>
    <th>Stock</th>
    <th>Total Daily Volume</th>
    <th>Return</th>
  </tr>
  <tr>
    <td>DQ</td>
    <td>-62%</td>
    <td>107,873,900 </td>
  </tr>
  <tr>
    <td>ENPH</td>
    <td>89.1%</td>
    <td>607,473,50 </td>
  </tr>
</table>
</th>
</tr>
</table>

Thus, we can conclude that investing on "ENPH" for his family would be a better choice. I will also use following outcomes from my Macro to explain my analyze to Steve. 


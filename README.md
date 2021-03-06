# Stock Analysis Report for Steve

## Overview of Project
This analysis satisfies the 2nd Challenge in the Data Analytics Bootcamp. For this challenge, Steve has asked for help in reviewing stock data.
Steve's parents have been interested in and are ready to invest in a company called DAQO New Energy Corp with Stock Ticker "DQ".
Steve wants to do further research into a variety of environmentally focused company's stock results so that he can point his parents in the right direction.
This project compares 2017 and 2018 volumes, starting and ending values for 12 different companies. 
With this information a ratio is calculated and results for each year are shown in the output table All Stocks Analysis.

## Analysis
An export of stock data that includes dates, opening/closing values, volume, etc. was reviewed using Microsoft Excel.
Visual Basic for Applications (VBA) code was written to review 3012 data points for each year. 
The equation (ending price / starting price - 1) was used to determine the percent return.
DQ had a good year in 2017 with 199.4% return and 35,796,200 shares traded. The stock started at $19.85 January 3, 2017 and ended December 30, 2017 at $59.44.
In 2018 the opposite is true. In January the stock started at $62.57, moving up a little bit from December 2017; but ended at $23.40. 
This ends the year with a -62.6% return with 107,873,900 shares traded.

#### 2017 Analysis-Results

![2017 Analysis-Results](https://github.com/summerstime/stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)

#### 2018 Analysis-Results

![2018 Analysis-Results](https://github.com/summerstime/stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

### Better Stock?
The results of other stocks, as shown above, were reviewed to determine which company may be a better investment for Steve's parents. 
Almost all stocks had negative returns in 2018, except the following:
* Enphase Energy Inc (ENPH) has shown to be less volatile from 2017 to 2018 with returns of 129.5% and 81.9%, respectively. Their shares traded each year are 221,772,100 for 2017 and 607,473,500 for 2018.
* Sunrun Inc (RUN) had a increase from 2017 of 5.5% (267,681,300 shares) to 2018 of 84.0% (502,757,100 shares). This could be due to more consumer interest in solar energy.

## The Code  
Development of the VBA code to determine the returns of the 12 companies included nested loops and arrays, which became a challenge. 
There were two approaches taken with both utilizing loops; but one had variables for addition and the other approach utilized separate arrays for volume, start price, and end price.
One advantage of the original code was that it was easy to follow/read, the disadvantage was that it was slow.

#### Original Code

![Original Code](https://github.com/summerstime/stock-Analysis/blob/main/Resources/Original_Code_For_Loops.png)

#### Refractored Code

![Refactored Code](https://github.com/summerstime/stock-Analysis/blob/main/Resources/Refactored_Code_For_Loops.png)

### Time
Time was evaluated between the two approaches with the array method being much faster in its processing. 
The elapsed time in processing one iteration of the code was decreased from approximately 1.40 seconds to approximately 0.25 seconds
This is an advantage when using arrays, even though the code can become more complex. This decrease in processing time will be even more noticeable when evaluating larger datasets.

#### 2017 Time Comparison

![2017 Time Comparison](https://github.com/summerstime/stock-Analysis/blob/main/Resources/VBA_Challenge_2017_B4-After.png)

#### 2018 Time Comparison

![2018 Time Comparison](https://github.com/summerstime/stock-Analysis/blob/main/Resources/VBA_Challenge_2018_B4-After.png)

### Refactoring Code
The array method was a refactoring of the original code, which focused on For Loops completing full cycles i.e. 3012 inner loop cycles and 12 main loop cycles. 
The refactored code determined the last row of the particular stock name and moved to the next stock. This decreased the amount of time spent processing the data.
By using the inner For Loop to increment the ticker array, the outer For Loop was not used as much as the original code required.
The array method was a great advantage and the refactoring of the original code, while somewhat challenging, created much better time results.
Refactoring code, in general, can be challenging to keep track of what has been changed and what needs to be changed. 
The disadvantage is that time is spent understanding the orignal code in order to make the correction in the first place. 
It may better to rewrite whole sections depending on the amount of change needed. 

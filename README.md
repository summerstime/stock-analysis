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

![2017 Analysis-Results](https://github.com/summerstime/stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png))
![2018 Analysis-Results](https://github.com/summerstime/stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png))

## Results
The results of other stocks were reviewed to determine which company may be a better investment for Steve's parents. 
It appears that Enphase Energy Inc (ENPH) has shown to be more consistent from 2017 to 2018 with returns of 129.5% and 81.9%, respectively.
The shares traded each year are 221,772,100 for 2017 and 607,473,500 for 2018.

### Challenges 
Development of the VBA code to determine the returns of the 12 companies included nested loops and arrays, which became a challenge. 
There were two approaches taken with one utilizing loops that added to a static variable and the next approach utilized separate arrays for volume, start price, and end price.
Time was evaluated between the two approaches with the array method being much faster in its processing. 
Time spent in processing one iteration of the code was decreased from approximately 1.40 seconds to approximately 0.25 seconds
That is a dramatic decrease in processing time which will be much more noticeable for larger datasets.

![2017 Time Comparison](https://github.com/summerstime/stock-Analysis/blob/main/Resources/VBA_Challenge_2017_B4-After.png))
![2018 Time Comparison](https://github.com/summerstime/stock-Analysis/blob/main/Resources/VBA_Challenge_2018_B4-After.png))



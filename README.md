# Stock Analysis using VBA

## Overview of Project

### Purpose
Steve, a recent finance graduate, is assisting his parents with their investments. They are fond of green energy and have decided to invest all of their money into a single company, DAQO New Energy Corp, based on their ticker symbol, DQ. Steve is interested in helping make an informed decision and diversify his parents portfolio by analyzing a few different green energy stocks based on the prior years' performances.

## Results

### Analysis & Performance
Twelve different green energy stocks were evaluated for the years of 2017 and 2018. The analysis consisted of running a VBA script gathering the Total Daily Volumes and calculating the Stock Return for each stock.
* Total Daily Volume is the collective traded shares during the given year displayed
* Return is the performance based on the stork price at the end of the year compared to the beginning of the year.

The results were added to a table to better compare which stocks performed the best.   

In order to better evaluate the stock performance the positive returning stock are highlighted in green and the negative returning stocks are highlighted in red.

**2017 Green Stock Performance**

<img src="https://github.com/mcgibbenyd1/stock-analysis/blob/main/Green_Stocks%20original%20code%202017.png" width="650" height="340" />

**2018 Green Stock Performance**

<img src="https://github.com/mcgibbenyd1/stock-analysis/blob/main/Green_Stocks%20original%20code%202018.png" width="650" height="340" />

Based on the results, 2017 showed positive returns for all the stocks chosen besides the stock identified by TERP. The ticker symbol DQ performed the best with a return of almost 200%. In 2018, the results were not as favorable with only two stocks with positive returns, ENPH and RUN.

The VBA script used to evaluate the raw data for the stocks was able to perform the analysis within the time indicated within the pop-up dialog boxs next to the tables of each years print out.
* Each year (2017 and 2018) had 3012 rows of data, a row for the performance of each stock for each day of the year the stock market was open. 

#### Refactored Script To Optimize Output time

The original script looped through the rows of data each time evaluating if the ticker in a given row was the current ticker being evaluated. If the current ticker was found then the Volumes were summed up and if not then the evaluation moved to the next row. If it was the first occurance of the ticker being found then the starting price was taken, and if it was the last occurance of the ticker appearing then the ending price was taken in order to calculate the Return value on the table.

The script was able to be refactored to help optimize the output time it takes for the code to evaluate the given dataset.

Rather than evaluating every row each time the nested For loop indexed, the three variables being reviewed (Volumes, Starting Price, Ending Price) were changed to an array. This way the data set only needed to be ran through one time due to being able to eliminate the nested For loop. Now, each time the ticker is found the values can be saved to a the array index associated to the ticker. 
 
By refactoring the codes as such, the output time was able to be greatly reduced. As can be seen in the pop up dialog boxes. Additionally, you can see that the script was just as successful at evaluating the dataset with the same results being found but with a much quicker processing time. 

**2017 Refactored Stock Performance**

<img src="https://github.com/mcgibbenyd1/stock-analysis/blob/main/VBA_Challenge_2017.png" width="650" height="340" />

**2018 Refactored Stock Performance**

<img src="https://github.com/mcgibbenyd1/stock-analysis/blob/main/VBA_Challenge_2018.png" width="650" height="340" />

### Summary: Refactoring
1) What are the advantages or disadvantages of refactoring code?
  * Advantages of refactoring code can include more efficient code that runs much faster as well as the possible to better equip the code to handle much broader datasets
  * Disadvantages can include increased testing to verifiy that the new simplicity or complexity does not introduce new bugs or errors. 

2) How do these pros and cons apply to refactoring the original VBA script?
  * The new refactored code was able to reduce the processing time by almost a factor of 10 by reducing the number of time the dataset needed to be scanned within the For loops. 
  * Changing the variables for Volumes, Starting Price, and Ending Price to an array made it possible to eliminate the nested For loop
  * The added array usage increased the complexity slightly that required updating some of the inputs for the variables in order to reference to correct ticker which caused to multiple testing attempts to verify the same output as the original analysis.  
  * Additionally the code could handle many more stocks with the refactoring by including a loop to create an array of unique tickers within the dataset so the code was not limited to the dozen stocks reviewed. 

# AN ANALYSIS OF GREEN STOCKS 
_Analyzing Green Stocks Data From Two Periods To Revail Trends_



## Overview of the Project
Refactoring involves the act of modifying an exisiting code structure to make it efficient in execution and easy to read. The purpose of the project was to refactor the "All Stocks analysis" VBA script so it can run at a more faster rate than before hence this was an attempt to reduce the runtime of the code.


## Results

### Original Code Analysis

The original code (All Stocks Analysis) was created to provide the Total daily volume and the Percentage returns of all green stocks from 2017 and 2018.
The code was developed by creating an array of tickers and asigning a value (0) to the "total volume" variable. A nested For Loop was then created where the iterator; "i" read through the array of tickers as against the second iterator "j"  which read through all rows in the worksheets. Conditionals or If - Then statements were added to generate and populate the totalvolume, startingPrice and endingPrice variables for each ticker.

_**Example of code used to generate starting prices per ticker_**

**If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then**

  **startingPrice = Cells(j, 6).Value**
  
Finally, Columns "B" and "C" on the active worksheet was filled with the total daily volumes and returns for all the stocks respectively and based on the output of the total volume, startingprice and endingPrice variables.

_**code to output data per ticker_**

**Cells(4 + i, 1).Value = ticke**r

**Cells(4 + i, 2).Value = totalVolume**

**Cells(4 + i, 3).Value = endingPrice / startingPrice - 1**

The results of the analysis showed that stocks performed well in 2017 than 2018. In 2017 the only stock that ill perfomed was "TERP" at a return of -7.2% whiles all but "ENPH" and "RUN" had negative returns in 2018.

### Refactored Code Analysis

Though the initial code worked in generating the information for which it was created, it had to be refactored to increase its efficiency.
 The internal structure of the code was modified by the introduction of an array for totalVolumes, startingPrices and endingPrices. 
A tickerIndex variable was created to aid loop through all the arrays than loop through the entire worksheet to generate the outputs. 
The final results of running this code indicated a slight change in the runtime of the code for both years.
The runtimes for both 2017 and 2018 analysis was 2.41 seconds and 2.42 seconds respectively when the original code was executed. 
However the runtime improved by 2.36 seconds and 2.21 seconds when the refactored code was run.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/100079292/157147552-1f178e76-a1d4-403d-bbb9-c8e1ce6ab918.PNG)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/100079292/157147569-1fa3dfc4-122d-4784-b8c6-d3372c508d41.PNG)


## Summary
- The advantages of refactoring code includes fewer bugs, faster execution and easy to read whiles the time involved in refactoring and the tendency of making mistakes serve as its cons.

- Much time was invested in restructing the script though the runtime of the original VBA script was improved after refactoring. 



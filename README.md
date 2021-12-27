# Stock Analysis for Steve
## Overview of Project
Steve's parents are passionate about green energy as they think there should be more alternative energy sources than fossil fuel. They would like to invest in a renewable energy company named "DAQO New Energy Corp". Since they don't know much about stocks and their trends, they want their son, who has recently graduated with a Finance degee, to do some stock analyses for them to avoid accidents and errors.

### Purpose
The purpose of the analyses is to write a VBA script to find total daily volume traded of 12 different stocks and their annual returns in two consecutive years, 2017 and 2018.

Another purpose is to refactor the code to make it faster and take up less memory so that Steve can expand the dataset to include the entire stock market over the last few years and  help his parents decide better about investing in stocks.

## Analysis and Results
### Analyses of stock performance
In general, the 12 stocks considered showed stronger performance in 2017 compared to 2018. DQ, SEDG, ENPH and FSLR all grew twice or more in value over 2017. SPWR and FSLR had exceptional total daily volume traded at more than half a billion for each.

Only two stocks saw positive growth in 2018, RUN and ENPH. The stocks which experienced relatively small decrease at in value less than 10% were VSLR, TERP, AY, SEDG. Total total daily volume traded for ENPH, SPWR and RUN were half a billion or more.

Considering both years, ENPH had kept a steady growth over the period. It should be noted that growth was defined based on value on first and last date and does **NOT** consider the variability in the intermediate dates.

### Performance of original and refactored codes
**Original code**

The original VBA script titled "AllStockAnalysis" looped thorugh all rows in the data within another loop for all twelve tickers to calculate total daily volume traded and annual returns for the stocks for a year.

The results and elapsed time for this script for each of the two years are shown below:

![Stock_Analysis_2017](https://github.com/Nusratnimme/stock-analysis/blob/main/Stock_Analysis_2017.png)

![Stock_Analysis_2018](https://github.com/Nusratnimme/stock-analysis/blob/main/Stock_Analysis_2018.png)

As can be seen, the script took more or less **1.5 seconds** to run.

**Refactored code**

The refactored script titled "AllStockAnalysisRefactored" gets rid of the nested loop structure. First, a ticker volume array of length 12 is created and then a loop initiates the array by assigning 0 as values. Next, the script it iterates through all rows in the data only once to calculate total daily volume traded and annual returns for the stocks for a year.

This is done by changing the ticker to the next one if next row to current has a different ticker. This works because the original data is ordered by ticker and dates.

The refactored script runs much faster than the original script. The results and elapsed time for the refactored script for each of the two years are shown below:

![VBA_Challenge_2017](https://github.com/Nusratnimme/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA_Challenge_2018](https://github.com/Nusratnimme/stock-analysis/blob/main/VBA_Challenge_2018.png)

It took **less than 0.2 seconds** for the script to run.

## Summary
### Advantages of refactoring
Refactoring of a code means making the existing code more efficient without addding any new features or functionalities.

The advantages of refactoring are:
* It makes code cleaner which in turn makes it easier for future users to read.
* It makes the code use less memory by employing fewer steps which results in the code running faster.
* It helps in finding and fixing bugs.
* It makes the code more organized.

### Disadvantages of refactoring
* It takes time to refactor a code, especially if it's complex and lengthy.
* If assumptions are made about the data in refactoring a code to make it run faster, it may yield error when those assumptions are not held.
* If refactoring is achieved by employing complex logic, it may be difficult to understand for other users. 

### Pros and cons of refactored stock analysis script
**Pros**
* The refactored code was significantly faster.
* It used less memory.
* It was cleaner and more organized.

**Cons**
* The refactoring was a time-consuming process.
* The refactoring assumes the data is ordered by tickers and dates. Even if a signle row is moved from its position, the code will yield inaccurate results.
* It may be difficult for another user to understand the logic of calculating ticker volumes without looping over all 12 tickers.

### Pros and cons of original stock analysis script
**Pros**
* The original script was detailed.
* It did not assume that the rows were ordered by tickers.
* The nested loop logic was more intuitive.

**Cons**
* The original script looked messy.
* It used more memory.
* It used nested loops which took longer time to run.

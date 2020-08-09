# Stock Analysis
An analysis of stock data utilizing VBA

## Overview of Project
During this analysis, we sought to create an easy to run analysis of all stocks in 2017 and 2018.  There was a particular emphasis on knowing the performance of the DQ stock, but this was later rolled up into an analysis of all stocks.

### Purpose 
The goal of this analysis to analyze the performance of a particular stock.  We did this by looking specifically at the annualized return based on the starting price and the closing price of the stock.  This will provide a year-over-year comparison that allows us to look at the performance of a particular stock through the previous/current year(s) and future years to come.  

## Results
###
** Code Comparison **
When running the 2017 and 2018 analysis with our original code it took .71 seconds.  When we run the same data set with our refactored code the time to run the code drops too .14 and .24 respectively.  The code run time may appear to be marginal at first glance but as the data set continues to grow and/or more stocks are added this improved speed and refactored script will make run time substantially faster.  The key difference between the two codes is the steps required to analyze the data set.  Our first code operates within a nested loop whereas our refactored code is processing through an index that references and array.  There was no need to use a nested loop for this refactored code.  

** Original Code **
![](Resources\VBA_Module_Code.PNG)

** Refactored Code **
![](Resources\VBA_Challenge_Code.PNG)

** Run Time on Refactored Code (2017 & 2018) **
![](Resources\VBA_Challenge_2017.png)

![](Resources\VBA_Challenge_2018.png)

** Stock Comparison **
In comparing the performance of the stocks within our dataset we can clearly see that 2017 was a much better year than 2018.  Stocks in 2017 generated an average return of 67.3% with only one of the stocks showing a small loss.  Stocks in 2018, however, did not fare as well as the previous year with an average loss of -8.5%.  Based on the data for the last 2 years stock TERP has not performed well but did show less of a loss in 2018 which could mean momentum towards a positive return.  ENPH and RUN were two stocks that performed exceptionally well over the last 2 years and could continue to show the same results based on the strong returns in 2018.

## Summary

**- What are the advantages or disadvantages of refactoring code?**

    **Advantages**
    * Easier to read
    * Improved effeciency
    * Organized

    **Disadvantages**
    * Injecting bugs when none existed before
    * Can be difficult to discern what is what without comments
    * Net result may not be worth the time spent

**- How do these pros and cons apply to refactoring the original VBA script?**

    The original code was pretty well organized but did lack some efficiency with the nested loop and not using an index.  The refactored code is very easy to read by allowing the code to reference a named index/array we can discern more quickly what it is we are trying to accomplish within the code.  The argument could be made that the time it took to refactor this code was not worth the time it saved in efficiency, however, if the dataset were to exponentially grow then we could assume that the time saved on performance justified the refactoring.  I injected plenty of bugs while refactoring and it took time to resolve each one.   

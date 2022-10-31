# Stocks_Analysis

## Overview of Project
   
   Analysis of stock data to make predictions on the best possible return investment.

### Purpose
   
   The overall purpose is to calculate total daily volume and yearly return for each stock in the data for any year. The total daily volume and yearly return are parameters used to predict whether a stock is worth investing in or not. 
  
  More specifically, Excel and Visual Basic will be used to analyze, calculate and automate the predictions on which stock will be the better investment. By writing macros to calculate the total daily volume and the yearly return for each stock we can automate the process. By summing up the daily volume for each stock, the yearly volume is calculated and a rough idea of the value of the stock. It's suggested that the higher the volume the more often the stock is traded and therefore the price is a better predictor of the value of the stock. By calculating the yearly return for each stock we can measure how each stock is performing. The yearly return shows the percentage increase or decrease in price from the beginning of the year to the end of the year. 
  
  To improve performance analysis on larger datasets, the original code and a refactored version will calculate how long the macro/code takes to execute and output. The code with the shortest run-time will be used to improve performance.


 
## Results

Stock performance for 2017 and 2018.


<img src = "https://github.com/cjstreet/stocks_analysis/blob/main/Resources/VB_2017_Output.png" width ="392" height ="242" align ="left">
<img src = "https://github.com/cjstreet/stocks_analysis/blob/main/Resources/VB_Output.png" width ="492" height ="242" align ="left"> <br />
<br /><br /><br /><br /><br /><br /><br /><br /><br /> <br />
Based on the total daily volume and yearly return, 2017 was a more profitable year for the stocks. For 2017, all stocks with the exception of the TERP (-7.2%) stock were profitable. The total daily volume and yearly return were not necessary correlated. SPWR had the highest total daily volume with roughly 800,000,000 but only a 23% return. DQ had the highest return for 2017 of approximately 200%, but only 36,000,000 in total daily value. So total daily value may not necessary an be a good indicator of the profitability of the stock. 

For the year 2018, all but two stocks (RUN & ENPH) reported a yearly loss of return. DQ reported the highest loss with -63%. Again, the highest total daily volume did not necessary indicate a positive return. SPWR had one of the highest total daily volumes ($5m) but their yearly return was -45%. The two stocks that faired well in 2018 were ENPH with 82% return and RUN with 84%. RUN was the only stock to improve from 5.5% in 2017 to 84% in 2018. 

In conclusion, 2017 was a profitable year for the majority of stocks and 2018 was the opposite with the majority of stocks reporting a loss. 
<br /><br />
### Original Code vs. Refactored Code<br /> <br />
<img src = "https://github.com/cjstreet/stocks_analysis/blob/main/Resources/VBChallenge_2018_Original.png" width ="325" height ="250" align ="left">
<img src = "https://github.com/cjstreet/stocks_analysis/blob/main/Resources/VBChallenge_2018_Refactored.png" width ="325" height ="250" align ="left"> <br /> 
<img src = "https://github.com/cjstreet/stocks_analysis/blob/main/Resources/VB_Code.png" width ="1100" height ="650" align ="left"> <br />
<br /><br /><br /><br /><br /><br /><br /><br /><br /> <br />

The above timed scripts tests show that the original code for 2018 ran in .445 seconds and the refactored code ran in .086 seconds. The refactored code ran 20x faster the orginal code. By refactoring the code we improved the performance of the code. Most likely this was done by using less memory (using arrays) and/or performing fewer operations (reducing the number of times the outer loop ran), the logic was similar to the original.

## Summary

### What are the advantages or disadvantages of refactoring code?
Advantages of refactoring are faster run times, simplifying the code, maintainability, and making the code easier to update.
Disadvantages of refactoring are a programmer could introduce a logic error into the code and it would have to be tested more.

### How do these pros and cons apply to refactoring the original VBA script?
Overall refactoring the original script by using arrays and running fewer operations (reducing the number of times we entered the outer loop) the script faster and easier to understand.


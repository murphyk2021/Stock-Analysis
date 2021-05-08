# Module 2: Stock Analysis Using VBA
- - -
## Overview of Project
In this module students were given a dataset which included information about how often 12 different stocks were traded and how much they cost on each day over the years of 2017 and 2018.  We created a macro using VBA to automatically read through the 3,013 rows of data in a year and extrapolate the first and last closing price of the stock DAQO, specifically, to determine the percent return for the year.  We also counted the total number of times it had been traded during the year.  This exercise introduced us to the basic logic behind *for loops* and *conditionals* in addition to introducing cell formatting through VBA.  We then added buttons to run our coded macros within a worksheet for user ease.

The next activity we asked the program to look through all of the data and calculate the yearly percent return and total number of trades for each of the 12 different stocks.  This required us to create an array, or list, of our different stocks and expand our knowledge of for loops by creating *nested for loops*-loops inside of loops.  

Finally, in our challenge excercise we were asked to rewrite our code in such a way that it would run quicker and more efficiently.  One way to achieve that goal is to reduce the number of times that we are looping through each row of data.
- - -
## Results
Below is the code we created as we worked through the module.  You can see there are two loops in this code-one inside of the other.
![Nested Loop Code](https://github.com/murphyk2021/Stock-Analysis/blob/da980120bd370356cf578fb15c14ba30e84f1fca/module_VBA_Code.png)
The first/outer loop will go through each of our tickers and apply the conditions.  The second/inner loop will go through our dataset and apply the information to each of the tickers.

This is what will happen when we apply these instructions to the **first** ticker in our array *(i=0).*
  - Define the string value for the variables:  **ticker** = “AY” (from our previously defined array)
  - Define the variable **TotalVolume** = 0'
  - Enter the nested loop (Run through each of the 3013 rows and apply the following to each:)
    - If the row contains the **“AY”** then redefine the **tickerVolume** as a sum of the last tickerVolume(“0”) and the value in the Volume column of our original data *(if not, do nothing)*
    - If the row contains the **“AY”**, but the row above it does not, then set the value of the first closing price (variable: startingPrice) *(if not, do nothing)*
    - If the row contains the **"AY”**  but the row below it does not, then set the value for the last closing price (variable: endingPrice) *(if not, do nothing)*
  - Exit nested loop
  - Record the values for ticker, totalVolume, and Return in the worksheet
  - Run Same sequence on the next ticker in our array.

When we use this code, it means we are asking the program to read our 3013 data points for each ticker.  That means it has to read **36,156** cells!
Use an Array + Not nested formula means we only read the 3013 data points once.



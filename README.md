# Module 2: Stock Analysis Using VBA
- - -
## Overview of Project
In this module students were given a dataset which included information about how often 12 different stocks were traded and how much they cost on each day over the years of 2017 and 2018.  We created a macro using VBA to automatically read through the 3,013 rows of data in a year and extrapolate the first and last closing price of the stock DAQO, specifically, to determine the percent return for the year.  We also counted the total number of times it had been traded during the year.  This exercise introduced us to the basic logic behind **for loops** and **conditionals** in addition to introducing cell formatting through VBA.  We then added buttons to run our coded macros within a worksheet for user ease.

The next activity we asked the program to look through all of the data and calculate the yearly percent return and total number of trades for each of the 12 different stocks.  This required us to create an array, or list, of our different stocks and expand our knowledge of for loops by creating **nested for loops**-loops inside of loops. 

Finally, in our challenge excercise we were asked to rewrite our code in such a way that it would run quicker and more efficiently.  One way to achieve that goal is to reduce the number of times that we are looping through each row of data.
- - -
## Results
Below is the code we created as we worked through the module.  You can see there are two loops in this code-one inside of the other.
![Nested Loop Code](https://github.com/murphyk2021/Stock-Analysis/blob/da980120bd370356cf578fb15c14ba30e84f1fca/module_VBA_Code.png)
The first/outer/red loop will go through each of our **tickers** and apply the conditions.  The second/inner/purple loop will go through **each row in our dataset** and apply the information to each of the tickers.

This is what will happen when we apply these instructions to the **first** ticker in our array *(i=0).*
  - Define the string value for the variables:  **ticker** = “AY” (from our previously defined array)
  - Define the variable **TotalVolume** = 0'
  - Enter the nested loop (Run through each of the 3013 rows and apply the following to each:)
    - If the row contains the **“AY”** then redefine the **totalVolume** as a sum of the last totalVolume(“0”) and the value in the Volume column of our original data *(if not, do nothing)*
    - If the row contains the **“AY”**, but the row above it does not, then set the value of the first closing price (variable: startingPrice) *(if not, do nothing)*
    - If the row contains the **"AY”**  but the row below it does not, then set the value for the last closing price (variable: endingPrice) *(if not, do nothing)*
  - Exit nested loop
  - Record the values for ticker, totalVolume, and Return in the designated output worksheet
  - Run Same sequence on the next ticker in our array.

In the refactored code below, we have created three more arrays--one for each of the three values we wanted to collect from the dataset for each ticker: **totalVolume, startingPrice, and endingPrice**
![Three consecutive loops](https://github.com/murphyk2021/Stock-Analysis/blob/98db6244f6178bebac633ee859d80b318e913679/Challenge_VBA_Code.png)
- In the **first for loop** we simply set the value for our tickerVolume to 0.  This ensures that we get an accurate count when the second loop is complete.  

- The **second for loop** is very similar to what we wrote in our first code. However, instead of recording the values on our output worksheet and running through the loop again for the next stock, we are simply storing the values for each stock in an array which we can reference later!

- The **third for loop** records the information from our 4 arrays into the output worksheet.

[See source code](https://github.com/murphyk2021/Stock-Analysis/blob/b9dfdcf7c048faeb6bc90a4c79fb693442429637/VBA_Challenge.xlsm)
- - -
When we compare the two strategies, both gave us the same output data.  With the additional formatting, it is easy to see that most of these stocks performed better in 2017 than they did in 2018!  Additionally, when we use a timer to measure how long it took for each code to run there was a marked difference between the two.

**All Stocks 2017**
![Comparison of run times for 2017](https://github.com/murphyk2021/Stock-Analysis/blob/1a81f29ec5d2fa9cc8a7a81263bb9fb787ec91e9/Resources/VBA_Challenge_2017comp.png)


**All Stocks 2018**
![Comparison of run times for 2018](https://github.com/murphyk2021/Stock-Analysis/blob/1a81f29ec5d2fa9cc8a7a81263bb9fb787ec91e9/Resources/VBA_Challenge_2018comp.png)
- - -
## Summary
It is not unusual for our first attempt at creating or achieving something new to be a little unrefined.  However, once we have taken that first step towards success we can remove redundant or unnecessary steps and streamline the process.  This also applies to writing codes.  It may take some time to do, but will save us time and resources later!  

Although the first code we created achieved the task we set out to do, it took longer to run because we were asking the program to read our 3,013 data points for each of our 12 stocks.  That means it has to read **36,156** cells!  In contrast, when we use the refactored code containing three separate loops which ran one after the other the program  only had to read through the 3,013 data points once.  If our dataset were to remain small, this may not be that big of a deal.  However, if our dataset were to grow, this could potentially become an issue!  




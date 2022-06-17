**# Stock-Analysis**

Running analysis on a couple stocks with VBA to determine the percentage of yearly return for each stock from the dataset.

**## Overview**

In this project, I will be creating a code in order to loop through the stock dataset in order to determine the yearly return for each stock to draw the conclusion of whether or not the stock should be invested in by our client, Steve, and his parents. Additionally, I will be refactoring the original code built so that this VBA script successfully runs faster than the original. By creating a flexible macro for running multiple stocks, this code can be reused in the future to be able to run on a different set of stocks. The purpose of this code is to be able to read how the stocks performed at a glance with proper formatting and legibility. 

**## Results**

To make my original code more efficient, I had to properly set arrays for my variables for my code to run through more quickly. Looping my arrays to a ticker Index is what, in my opinion, made the code so efficient. Instead of going through each ticker at a time like my original code had, my refactored code was able to run through all the tickers simultaneously by setting the tickerIndex. 

### Run-time for each yearValue
Below are the run times with my original code
![green_stocks_2017](https://user-images.githubusercontent.com/105755095/174363576-eae46d4f-1d43-4deb-a144-7ff972993474.png)
![green_stocks_2018](https://user-images.githubusercontent.com/105755095/174363940-125a7b4a-cb89-4f84-b57f-6fef4df6e3f2.png)

After refactoring the code, I was able to cut down the time it took to process the code by nearly 110%.
![VBA_Challenge_2017](https://user-images.githubusercontent.com/105755095/174364554-42f7114f-3618-45e6-b06d-899b3a409c92.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/105755095/174364905-bc7ec1e4-601c-4963-8ee3-ba426d315a8c.png)

Based on these run-times, the code now runs almost 1 second faster than the original code, making this refactored code much more efficient. 

**## Summary**

Now, with a click of a button, Steve is able to run the code to analyze a list of stocks and their returns with the benefit of running for any year. 

### The Advantages of Refactoring
-	The code is now much more efficient and faster.
-	The code works the same way so there is no surprise for the end user.
-	Really learn what all parts to the code means
-	Reduces the redundancies of the original code

### The Disadvantages of Refactoring
-	You may risk damaging the whole code and creating more errors
-	May take a lot of time and proves little significant advantages
-	May take a while to understand someone else’s code in order to refactor it

### Application to Refactoring Original VBA Script

These pros apply to refactoring this VBA script as this method did make the code much faster, which was our goal. Through the refactoring process, it really forced me to learn what each step of the original code meant in order to effectively implement the changes in which I thought would make the loops more efficient. Another pro was that I was able to reduce the redundancies of the code looping through each ticker by creating that tickerIndex. The only disadvantage I felt like I experienced through this process was the time it took me to figure out the proper syntax of the code to make it actually run with no errors. Many trials and errors were made to finally produce the end refactored results.  

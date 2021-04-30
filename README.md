# stock-analysis

## Overview of Project

### Purpose

Steve is analyzing stock market data for his parents to help them in choosing stocks that will perform well and lead to a high overall return. They are specifically looking in the renewal energy market, and Steve is trying to find alternatives to the DAQO stock they currently have as there are better options to produce a higher return on their investment. Steve wants to use the workbook we have already given him and expand its capacity to be able to include the entire stock market in the analysis. We will need to optimize the workbook using refactoring so it can more efficiently handle large amounts of data. 

## Results

### 2018 vs 2017 Stock Performance

Overall stock performance based on a percentage return was much lower in 2018 than 2017. Majority of stocks in 2018 (12/14 stocks chosen) had a negative return, where as only one stock in 2017 had a negative return. However, in 2018 there was a higher daily volume when looking at all stocks together. 

### Efficiency & Optimization of Original vs Refactored Script

After refactoring the VBA script we found that speed of the macro dramatically reduced. The macro went from taking 0.93 seconds to 0.22 seconds for 2018 and 1.03 seconds to 0.21 seconds for 2017. See below for example comparison screenshots. 

Refactored 2018
![image_name](VBA_Challenge_2018.png)

Refactored 2017
![image_name](VBA_Challenge_2017.png)

Original 2018
![image_name](VBA_Challenge_2018_Original.png)

Original 2017
![image_name](VBA_Challenge_2017_Original.png)


## Summary

### Advantages & Disadvantages to refactoring code

- What are the advantages or disadvantages of refactoring code?

  - Advantages:
     -  One of the main advantages is that refactoring improves efficiency, so the macros will run faster. Another advantage is that it makes the code easier to read and understand. This becomes very useful if many other people will be looking at your code in the future. May reduce debugging time if changes were made.  
  
  - Disadvantages:
    - One disadvantage to refactoring code could be that it takes extra time to develop the new refactored code on top of the time it took you to produce the first macro, so this could be a hit to productivity in the short term. Additionally, a disadvantage could be that members of your team got used to the macro before it was refactored and don't know how to interpret it after refactoring. 

- How do these pros and cons apply to refactoring the original VBA script? 

  - In this case we saw that the pros to refactoring the original script were that macro runtime was greatly reduced. This will also allow the workbook to be easily expanded to include the whole stock market. One disadvantage was that we had to take extra time to build on an already working macro. 

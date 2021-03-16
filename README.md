# Stock-Analysis

## Overview of Project

### Purpose:
Perform data analysis on green energy stock for our client Steve. As well as, create a user-friendly macro for Steve to view the data set. 

## Analysis:

### Analysis of stocks based on returns
- A VBA macro was generated to run through collected data for all green energy stocks from 2017 and 2018. A nested “For” Loops was scripted to analyze the total volume for each stock and determine the return value based on the ending price and starting price. The For loop used a series of “Arrays”, which relied on the ticker value (see image below) to track the total volume as well as the starting price and ending price for individual stocks. The macro then prints the results in a pseudo table format on “All Stock Analysis” worksheet of the Excel workbook. 

![](https://github.com/Apollo619/Stock-Analysis/blob/main/Resources/Arrays.PNG)

### Formatting Data Set
- VBA formatting was scripted at the end of the code and applied to the data using another “For” loop on the pseudo table data set as a point of reference (see image below). Once the macro was complete and working properly it was assigned to a convenient button that allows the user to reference which year they would like to analyze.  A second button was also set up to clear previous data set quickly. 

![](https://github.com/Apollo619/Stock-Analysis/blob/main/Resources/Formatting.PNG)

## Results:
1.	2017 was a good year for green energy companies (see “All Stocks (2017)”), with one exception being “TERP” who reported a negative return value, all stocks 			netted a positive return. “DQ” showed the biggest return with nearly 200%. Unfortunately, 2018 painted a different picture (see “All Stocks (2018)”). Nearly all 		stocks reported negative returns except for “ENPH” and “RUN”.
![](https://github.com/Apollo619/Stock-Analysis/blob/main/Resources/All_Stock_2017_Return.PNG)        ![](https://github.com/Apollo619/Stock-Analysis/blob/main/Resources/All_Stock_2018_Return.PNG)
2.	Based on the results, it is recommended that Steve advise his parents to invest in “ENPH” and/or “RUN” green energy as they are the only two companies to report 		positive returns over the past two years. 
3.	Refactoring of the code lead to in improved performance of the VBA macro compared to performance of the original code. 
![](https://github.com/Apollo619/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017_OriginalCode.png)     ![](https://github.com/Apollo619/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)
![](https://github.com/Apollo619/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018_OriginalCode.png)   ![](https://github.com/Apollo619/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)
4.	A larger data set ranging back further over a longer period would allow for a more thorough analysis of green energy stock. 

## Summary:

### Advantages and Disadvantages of refactoring code
- One of the many advantages of refactoring code is the ability to increase the efficiency of the VBA macro, which increases the performance speed of the code. Also, different technique such as nested loops, arrays, and conditional formatting can allow the code to encompass more of the user’s want/needs from the data set.  
- Subsequently a disadvantage of refactoring the code is, you are dependent on the previous coders “comments” to understand what they are wanting to accomplish. Without these comments it can be time consuming to read and interpret the code. Another disadvantage can be when a coder changes the structure of the code it can cause the program to lose track of its variables or it will get lost in an infinite loop, depending on the circumstances. 

### Advantages and Disadvantages of refactoring original VBA script
Pro: the refactoring of the code improved performance speed, as seen in the above timer images *All Stocks (2017)* and *All Stocks (2018)*.

Con: adding a variable to the array caused miss match issues and other run time errors that needed to be debugged.  

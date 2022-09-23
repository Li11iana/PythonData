# Stock Analysis
## Overview of Project
This project reviews stock data in Excel format from several green energy companies. Using VBA to automate stock analysis eliminating common repetitives calculations and providing clear and consistant results, so the data can be easily reviewed by finance advisors and their clients.

### Purpose
Facilitate financial analysis of stocks in Excel files by automating the analysis of each stock's yearly total volume and returns. Additionaly the analysis result's formatting will be integrated as part of the macro to ensure the results are inuser-friendly.

Given the nature of stock analysis, a larger number of companies and transacctions might be necessary in the future, therefore the initial code (see Green_Stocks_Analysis.xlsm) will be refactored to improve its efficiency. 


## Results
The initial version of the VBA code to automate stock analysis focus on simple readable code that would locate appropriate data sheet to and in the Analysis sheet calculate Total Volume, starting price, ending price and returns for each of the tickers define in an array. This method worked but consumed noticeable time as shown below:
###Time to calculate 2017's analysis
![2017.1](https://github.com/Li11iana/stock-analysis/blob/main/Resources/2017.1.png)
###Time to calculate 2018's analysis
![2018.1](https://github.com/Li11iana/stock-analysis/blob/main/Resources/2018.1.png)

In the refactored version of the code all calculations were made and held in arrays until they were retrive to populate the analysis table. 
###Refactired time to calculate 2017's analysis
![2017.2](https://github.com/Li11iana/stock-analysis/blob/main/Resources/2017.2.png)
###Refactored time to calculate 2018's analysis
![2018.2](https://github.com/Li11iana/stock-analysis/blob/main/Resources/2018.2.png)

The decrease in calculated time 

*The analysis is well described with screenshots and code (4 pt).

## Summary
### Refactoring


**Why Developers Refactor Source Code: A Mining-based Study.**
Mauro Pezzè (Ed.). 2020. Continuous Special Section: AI and SE. ACM Trans. Softw. Eng. Methodol. 29, 4 (October 2020).
https://dl.acm.org/doi/10.1145/3408302

### Original code vs refactored code
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).



There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

# Stock Analysis
## Overview of Project
This project reviews stock data in Excel format from several green energy companies. Using VBA to automate stock analysis eliminates common repetitive calculations and provides clear and consistent results, so the data can be easily reviewed by finance advisors and their clients.

### Purpose
Facilitate financial analysis of stocks in Excel files by automating the analysis of each stock's yearly total volume and returns. Additionally, the analysis result's formatting will be integrated as part of the macro to ensure the results are user-friendly. Given the nature of stock analysis, a larger number of companies and transactions might be necessary in the future, therefore the initial code (see Green_Stocks_Analysis.xlsm) will be refactored to improve its efficiency.

## Results
The initial version of the VBA code to automate stock analysis focuses on simple readable code that would locate appropriate data sheets to and in the “Analysis” sheet calculate Total Volume, starting price, ending price and returns for each of the tickers defined in an array. This method worked but consumed noticeable time as shown below:
### Time to calculate 2017's analysis.
![2017.1](https://github.com/Li11iana/stock-analysis/blob/main/Resources/2017.1.png)
### Time to calculate 2018's analysis.
![2018.1](https://github.com/Li11iana/stock-analysis/blob/main/Resources/2018.1.png)

It took 0.5488281 seconds to analyze how 11 different tickers traded in 2017 and 0.5625 seconds to assess those same tickers during 2018. Even though this average time of 0.556 seconds is not long enough to be frustrating to the user, it’s still noticeable, and for the limited number of elements that are being evaluated in this case it suggests code can be improved. In the refactored version of the code all calculations were made and held in arrays until they were retrieved to populate the analysis table. The new refactored calculating time is shown below:

### Refactored time to calculate 2017's analysis.
![2017.2](https://github.com/Li11iana/stock-analysis/blob/main/Resources/2017.2.png)

### Refactored time to calculate 2018's analysis.
![2018.2](https://github.com/Li11iana/stock-analysis/blob/main/Resources/2018.2.png)

The time it took to analyze the data for both 2017 and 2018 using the refactored code was 0.0703125 seconds in both cases. The decrease in calculated time (rounded up) went from 0.556 seconds to 0.070 seconds, taking only 12% of the time initially required to execute the stock analysis. Therefore, the refactored code is more efficient and will work better if higher volumes of data were to be processed. It is important to consider that if further changes are necessary the source code now interlaces 3 different arrays, so evaluating more tickets must be considered and modified in every array.


## Summary
### Refactoring
Refactoring is described as the process of altering the internal structure of the code without changing the functionality or behavior. It is especially important in the industry as a way of providing product consistency to the end user but improving its maintainability and extensibility. Refactoring is a common practice in that benefits reduce operational costs as well as to adapt to an increase in users.

### Original code vs refactored code
In the study: _The buggy side of code refactoring: understanding the relationship between refactorings and bugs_, the authors acknowledge that even when refactoring is done to improve a working code and reduce the risk of errors, in a significantly number of cases the refactoring introduces new "bugs" to the source code. An important factor to determine the success of the refactoring is the complexity of both the source code and the intent of the refactoring, it is vital to have a clear goal of the structural improvement that is going to take place and avoid taking major structural changes in the disguise of a refactoring project.

## Sources for the Summary

Isabella Ferreira, Eduardo Fernandes, Diego Cedrim, Anderson Uchôa, Ana Carla Bibiano, Alessandro Garcia, João Lucas Correia, Filipe Santos, Gabriel Nunes, Caio Barbosa, Baldoino Fonseca, and Rafael de Mello. 2018. The buggy side of code refactoring: understanding the relationship between refactorings and bugs. In Proceedings of the 40th International Conference on Software Engineering: Companion Proceedings (ICSE '18). Association for Computing Machinery, New York, NY, USA, 406–407. https://doi.org/10.1145/3183440.3195030
Code Refactoring: Meaning, Purpose, Benefits.Published online March 28, 2022.https://sumatosoft.com/blog/code-refactoring-meaning-purpose-benefits

# Stock Analysis using VBA

## Overview of Project

The purpose of this project is to develop a complex code using Visual Basic for Applications in order to help the user to simplify the analysis around selected stocks and their performance in 2017 and 2018, with that, the user will be capable to offer alternatives to diversify funds and maximize the returns for his clients. 

The original code developed for this purpose was working well for a certain number of stocks, but the concern is that it might not work as well for thousands of stocks. Then in this challenge we have the task to improve the quality of the code using the refactoring technique. 

>*Refactoring is a disciplined technique for restructuring an existing body of code, altering its internal structure without changing its external behavior*  -Martin Fowler (Father of Code Smell)

## Results
Prior to start the optimization of our code, it was needed to make a copy of the workbook that was developed in Module2 and attached the “refactored subroutine” using Visual Studio to open the file and copy the scrip to VBA. Then using the instructions detailed in this Challenge, I was able to refactor the code and run the macro faster than the original code (≈ 80% of efficiency) as you can see below: 

Year 2017

![Year2017 comparison](https://user-images.githubusercontent.com/106939511/175434641-7ae01224-42b1-4842-aad5-030c39ce436b.png)


Year 2018

![Year2018 comparison](https://user-images.githubusercontent.com/106939511/175434655-19815b97-557d-4fd4-92ab-3ae9ea19131f.png)

To see the scrip please visit: [GitHub pages]:(https://github.com/AAGA85/stocks-analysis/blob/aa1616c6be9a5d7c3498cc87ac978e46c5044004/VBA_Challenge.xlsm)

## Summary

### Advantages and disadvantages of refactoring code
Refactoring gives us some advantages like having a clean and more understandable code then that makes easier to find a mistake or makes future changes much simpler, also help us to run our programs faster and with that, we improve the efficiency and performance of our application.  Nevertheless, refactoring should not be use always because sometimes it could take more time and then more money to develop a code and the efficiency or the performance do not have a substantial improvement, then it is important to know the expectations of the user (or customers) and their resources (time, developers, money) to find a balance between improvement and functionality.

### Advantages and disadvantages of refactoring the Stock Analysis code
If we compare the *AllStockAnalysis* subroutine vs the *AllStockAnalysisRefactored* subroutine, basically the refactoring technique allowed us to turn two loops into one (see image below) and have an impact in the time and structure complexity (faster algorithm).  As we learned in lesson 2.3.2 the nested loops could be a bad practice because if we apply several for loops in a code, we could get lost trying to find a mistake in our coding or if other person tries to make a change in our code it will be more difficult to understand the structure.

![nested loops](https://user-images.githubusercontent.com/106939511/175435222-5182b58d-db1a-4ebe-9a97-943f98a6b66e.PNG)

But as I mentioned above refactoring isn’t a good idea always. In our case of study, the improvement was that our code ran 80% more faster than the original one but to be honest both macros stayed into the <2 sec that means there wasn’t any substantial advantage for the final user but we as a programmers take us extra time to develop this “new program” with the same performance as the original.

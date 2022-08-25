# Stock Analysis using VBA
## Overview of Project
In this project we analyzed the stock data of 2017 and 2018. We helped Steve to find out stocks with best returns in renewable energy as to help his parents want to diversify their portfolio and have better returns. Initially we have an excel file that represents the data in the raw form. It has different Columns like ticker, date, open, close, and volume etc. The first one represents the ticker for each stock (as traded on stock exchange). Others represent values for opening, closing, high and low prices for the day, and lastly volume for a stock on a specific date.

Firstly, as Steve’s parents have invested in **DQ**, we calculated the data for **DQ**. After that, we calculated the total volume for all stocks during a particular year. By doing this can see how a stock traded during a year. After The yearly return of stocks is the percentage increase or decrease in price from beginning of the year to the end of the year. To make our results more valuable we add a button which can give us same comparison with one click. We added colors through conditional formatting to make our results easily understandable. We used a time function to measure the code performance.

## Results
As we can clearly see from images that 2017 a good year for stocks as we can mostly green color which represents the increase in returns. The **sedg** performed best this year and **terp** is the one whose value decreased during this year.  For year 2018 was not good for many stocks. It shows red for most of the stocks value. During this year only two stocks performed well.  As for results we can clearly see that **enph** is the best stock as its value increased during both years.

![results2017](https://user-images.githubusercontent.com/111101038/186575106-d2e887e5-0057-490e-b3a6-f5eb9a4473c1.PNG)

![2018results](https://user-images.githubusercontent.com/111101038/186575148-ea44bcdf-f77e-475c-bef2-975a0e402415.PNG)


The run time images show the timely comparison between two same data files. AS we can see that original **VBA** script takes more time than refactored script. To reduce execution time, we replaced the nested for loop in old code with three separates for loops. Our results show approximately 1 minute difference between both execution times. 
![2017](https://user-images.githubusercontent.com/111101038/186575186-69b4817d-60e3-41ac-9dc7-20036c84ea09.PNG)
![2018](https://user-images.githubusercontent.com/111101038/186575209-838bb0d6-3ec1-430a-b6b0-f7ce920fc052.PNG)


## Summary

### Advantage and Disadvantage

The advantage of refactoring the code is reduction in execution time.
The disadvantage of refactoring the code is creating new arrays which is hard to understand.

### Pro and Cons

By the less execution time we can get results faster. The only con that we can see is complex code that is harder to understand.

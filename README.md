# Stock-Analysis

## Overview of Project
For this project, we created a workbook for Steve to analyze stocks from 2017 and 2018 to research what stocks his parents should invest in. The scope of the original project was to analyze just a few dozen stocks, but we refactored the original code to be more efficient. With the refactored code, Steve has the capacity to thousands of stocks if he wanted to do more in-depth research later.

Click [here](https://github.com/liviblocker/stock-analysis/blob/master/VBA_Challenge.xlsm) to see the full analysis.

### Results
Steve's parents are interested in the DQ stock and they want to better understand if that is the right investment. Through this analysis, we can see that while DQ surged in 2017 with an annual return of 195%, it dropped 63% in 2018. Based on these results, Steve's parents may want to invest in another stock.

Though it is important to note the overall performance of stocks in 2017 and 2018. In general, 2017 was an excellent year for the collection of stocks that Steve chose to analyze. The yearly return in almost all stocks increased with the biggest drop being only 8%. It is notable that DQ's return increased 195%, higher than any other stock that year. In contrast, 2018 was a terrible year. Almost all the stocks in the dataset dropped save ENPH and RUN. DQ dropped the most with 63%. While I would still recommend Steve's parent invest in a different stock, it's helpful to contextualize the drop in a larger economic context.

I would recommend Steve's parents invest in ENPH - the only stock in the dataset to increase it's yearly return over both 2017 and 2018.

As mentioned in the overview, we refactored the code to increase the efficiency. In the original code, 2017 and 2018 both ran at 0.625 seconds:

![OriginalCode_2017](https://github.com/liviblocker/stock-analysis/blob/master/OriginalCode_2017.png)
![OriginalCode_2018](https://github.com/liviblocker/stock-analysis/blob/master/OriginalCode_2018.png)

The refactored code ran much faster at only 0.109375:

![VBA_Challenge_2017](https://github.com/liviblocker/stock-analysis/blob/master/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/liviblocker/stock-analysis/blob/master/VBA_Challenge_2018.png)

The refactored code streamlined the original code making it clear and easier to understand. The major difference between the original code and the refactored code was the use of nested loops. While the original code use that structure, it became clear that it was an unnecessary in the refactored code - instead using two For Loops:

![ForLoops](https://github.com/liviblocker/stock-analysis/blob/master/ForLoop.png)

## Summary
<b>-What are the advantages or disadvantages of refactoring code?</b>

There are multiple advantages of refactoring code. In this project, refactoring the code meant making the analysis run fast (as shown in the results above). Refactoring is also a teaching tool - you become more familiar with the code and logic by breaking it down and making it neater and easier to read.

Unfortunately, refactoring can also feel repetative and time-consuming, especially when working with smaller datasets in which the efficiency of the run time is negligible. 

<b>-How do these pros and cons apply to refactoring the original VBA script?</b>

Refactoring was deeply helpful to me to better learn the logic of VBA and how make code cleaner and easier to understand. More than the faster run time of the analysis, the process of breaking down and approaching the code from a new perspective was a deeply educational experience. Because I had been working through the logic, I had the ability to identify an Overflow error and correct it by assigning the tickerStartingPrices and tickerEndingPrices variables as Integer instead of Single. Had it not been for the refactorign process, I would not have been able to adjust and make that correction.

It wasn't perfect. The process was time-consuming and, frankly, frustrating, especially for code that runs only .5 second faster.

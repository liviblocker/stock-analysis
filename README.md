# stock-analysis

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

## Summary
<b>-What are the advantages or disadvantages of refactoring code?</b>

There are multiple advantages of refactoring code. Refactoring code means the making it more efficient, as such refactoring code means taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

<b>-How do these pros and cons apply to refactoring the original VBA script?</b>



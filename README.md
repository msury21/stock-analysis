# VBA of Wall Street

## Overview of Project
The intention of this project is to create a code using VBA that can analyze stock data and highlight whether a certain stock is performing well or not.

### Refactoring
An additional component to this project has been an attempt to refactor the code in order to have it run faster, leading to a comparison of run times between 2018 in the original script and 2018 in the refactored script.

## Results

![2017_Stock_Analysis.png](/Resources/2017_Stock_Analysis.png)
![2018_Stock_Analysis.png](/Resources/2018_Stock_Analysis.png)

In 2017, barring TERP, every company had a positive return, with some companies having a nearly 200% return. SPWR, JKS, FSLR, CSIQ, and AY all experienced a decrease in total daily volume in 2018. Additionally, in contrast to positive returns for nearly every company in 2017, 2018 had only 2 companies with positive returns: ENPH (81.9%) and RUN (84.0%). Both of those companies had a major growth spurt with regards to their total daily volume over 2018; ENPH's increasing by 385,701,400 and RUN's increasing by 235,075,800.

### Refactoring

![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)
![VBA_Challenge_2018_Refactored.png](/Resources/VBA_Challenge_2018_Refactored.png)

The refactored code was able to analyze the over 8 times faster than the original code. The main difference in the codes for the two macros is that the original code involves a nested loop while the refactored code has only single loops. In the original code, the macro would go through the tickers in a loop, and for each ticker, it would go through each row of the dataset and check whether the current ticker was present before altering the total volume, starting price, or ending price. In the refactored code, several arrays were set so that while scanning every row of the dataset, each ticker's information would update. Thus, in the refactored code, the entire dataset is examined once as opposed to the original code looping over the dataset 12 times - once for each ticker.

## Summary

In conclusion for the stock performances, ENPH and RUN appear to be the best companies to purchase stock from. With regard to refactoring, in this case, refactoring the code worked out well; the code ran much faster and produced the same results. In general, refactoring code can lead to higher efficiency, but at the cost of wracking your brains trying to a) figure out how to make it more efficient b) figure out how to make it work. I was originally stumped by the presence of only one loop in the refactored code; I was convinced that I had to make another loop in order to prevent the row-scanning-loop from going infinitely or having too many entries for the size of the arrays I created. However, once I figured out that the loop would end without adding anything else, the refactored code became my favorite for how sleek it felt compared to the original code.

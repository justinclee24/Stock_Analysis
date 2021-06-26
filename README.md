# Stock Analysis
Creating VBA analysis for Steve's portfolio

## Overview
The purpose of this analysis is to provide Steve with easy-to-read knowledge on the performance of his stock portfolio. Knowing the return and volume of his stocks at a simple glance, it will make his job much easier as to which stocks to further pursue and which to eliminate from his portfoli.

## Results
### 2017
Regarding performance for the stocks in 2017, the difference the refactored script made was enormous (see images below). 

Original

<img width="352" alt="Original 2017" src="https://user-images.githubusercontent.com/85330159/123522400-406a6200-d67a-11eb-9856-bb4875e00a1f.png">

Refactored

<img width="354" alt="Refactored 2017" src="https://user-images.githubusercontent.com/85330159/123522402-46f8d980-d67a-11eb-8600-5c01444c08ea.png">

In the refactored script, we were tasked to create an array to use for the output table. This is most of what changed the look of the table from mostly red to mostly green in the Return column. This would also explain the significant reduction in run time from the first original script to the refactored version (see images below).

<img width="418" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85330159/123522496-d1413d80-d67a-11eb-8378-f1869e7b189e.png">

<img width="419" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/85330159/123522499-d605f180-d67a-11eb-99c6-b398fb0f395d.png">

### 2018
Dealing with year 2018, the original and refactored script create similar tables of volume and return for each ticker. However, the refactored version, like in 2017, significantly reduced the run time for the code. This also has to do with the addition of creating an array for our tickers, rather than the method used in the original coding (see above run time images). 

## Summary
### Advantages/Disadvantages of Refactoring
One advantage of refactoring the code is that is becomes more organized and efficient. It essentially alters the structure of the code, thus resulting in faster run times. However, potential disadvantages could be that it could create more opportunity for bugs to be within your coding script that may not be caught easily. 

### Pros/Cons of Applying Refactoring to Original VBA Script
The pros of applying refactoring to this VBA script is that it definitley decreases our run time for Steve, and seems to paint a better picture, especially in 2017, of which tickers performed well enough for Steve to keep in his portfolio. Potential cons however, as discussed earlier, is that with the potential for bugs that may be missed, the table formatting might be misrepresenting Steve's stocks in some way. 

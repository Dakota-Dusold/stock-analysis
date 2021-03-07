# stock-analysis
## Overview of the Project
    We have been asked to rebuild a workbook for our client Steve to analyze more stocks than previously requested. This refactoring of the code will improve upon our current macro which analyzes stocks and gives us informations such as their total daily volume for the year and their return percentages for the year. If Steve wants to run an analysis on the entire stock market for his parents, we will need to save him some time! We can do this by creeating indexes and variables to simplify the code.
## Results
### Refactored Code
This code was refactored using:

    -The creation of a tickerIndex variable that was set = 0 
    -Next we used output arrays for the tickerIndex to access for our outputs
    -Then we used for loops to loop through our spreadsheets using our tickerIndex
    'Finally we used if-then statements to find our starting and ending prices and to increase tickerIndex


### 2017
![](https://github.com/Dakota-Dusold/stock-analysis/blob/main/Resources/Stocks2017.PNG)
![](https://github.com/Dakota-Dusold/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
    
   - In 2017, the majority of Steve's selected green energy stocks performed really well. Half of these stocks had a return of 50% or more! Theese look like promising investmeents so far.
    
   - The All Stocks Analysis Refactored Macro ran almost half a second faster for 2017 (.4453125 seconds to be exact). This will save Steve a lot of time when he uses this macro to analyze thousands of stocks as opposed to 12 stocks.

### 2018
![](https://github.com/Dakota-Dusold/stock-analysis/blob/main/Resources/Stocks2018.PNG)
![](https://github.com/Dakota-Dusold/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)
   
   - In 2018, only two of Steve's selected green energy stocks had a return. In addition, most of these stocks had less total daily volume meaning they were traded less. It was not a good year for these stocks, however, Steve should consider looking into the two stocks that had returns in both years: ENPH and RUN.
    
   - The All Stocks Analysis Refactored Macro also ran almost half a second faster for 2018 (.4550781 seconds to be exact). This proves consistency in our macro when running it across different data sets.

## Summary Statement
    1. The Advantages of Refactoring Code:
        - Makes code more readable
        - Speeds up code run time
        - "It changes the way a developer thinks about the implementation when not refactoring." 
        ![Link to source!](https://aip.scitation.org/doi/abs/10.1063/1.3516393?journalCode=apc) 
       The Disadvantages of Refactoring Code:
        - It is time-consuming
        - May not actually make the code run time faster
        - The time spent refactoring needs to be worth the time that will be saved by the faster code 

    2. The Pros of Refactoring Original VBA Code:
        - It helped me work through what we learned in this module
        - The new code is more readable and runs faster as proven by the timestamps 
       The Cons of Refactoring Original VBA Code:
        - At one point I got so confused that I deleted everything and started from scratch

# Stock Analysis

This analysis reviews market activity from 2017 and 2018 stocks.
----

[VBA Data & Visual Sources](https://github.com/emilymcdaniel/stock-analysis/blob/master/VBA_Challenge.xlsm) 

----
## Overview
This code analyzes stock history in 2017 and 2018. The ticker (stock identifier), transaction date, prices at time of opening, daily high, daily low, closing, and adjusted close; as well as the volume traded were recorded. Using this information, we aggregated the Total Daily Volume (total number of stocks traded in each year) and Return (earliest price vs the latest closing price) of each ticker.


## Results: 
The refactored code took 0.21 seconds to run the 2017 data, and 0.16 seconds to run the 2018 data.

Refactored 2018 Result:

![Refactored 2018 Result](https://github.com/emilymcdaniel/stock-analysis/blob/master/Resources/VBA_Challenge_2018_refactored.PNG)

Original 2018 Result:

![Original 2018 Result](https://github.com/emilymcdaniel/stock-analysis/blob/master/Resources/VBA_Challenge_2018_original.PNG?raw=true)

Refactored 2017 Result:

![Refactored 2017 Result](https://github.com/emilymcdaniel/stock-analysis/blob/master/Resources/VBA_Challenge_2017_refactored.PNG?raw=true)

Original 2017 Result:

![Original 2017 Result](https://github.com/emilymcdaniel/stock-analysis/blob/master/Resources/VBA_Challenge_2017_original.PNG?raw=true)

## Summary
### What are the advantages or disadvantages of refactoring code?
Advantages of using refactored code include: 
- Faster processing: Fewer lines that can be easily interpreted allows the program to quickly interpret instructions.
- Easier to debug: When writing code, it's easier to correct errors when there is less code to review.
- Broader application: Refactored code can be more readily applied to a larger dataset.

Disadvantages of refactored code include:
- Debugging irregularities: Any data that doesn't run will need to be identified and corrected, and potentially require unique code.
- Overlooking errors: It is possible probelmatic rows may be skipped without throwing an error, meaning data gaps would go unnoticed.

### How do these pros and cons apply to refactoring the original VBA script?
Comparing the original to the refactored VBA, notice the following:
- The original includes user-friendly buttons to activate distinct macros.
- The original is unable to run multiple macros simulateously.
- The refactored is not user-friendly, but once "run" is selected, all macros are applied.
- The refactored version applies logic and formatting macros simultaneously.

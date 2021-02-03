# Stock Analysis

This analysis reviews market activity from 2017 and 2018 stocks.
----

## Overview
This code analyzes stock success in 2017 and 2018. The ticker (stock identifier), transaction date, prices at time of opening, daily high, daily low, closing, and adjusted close; as well as the volume traded were recorded. Using this information, we aggregated the Total Daily Volume (total number of stocks traded in each year) and Return (earliest price vs the latest closing price) of each ticker.


## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
The refactored code took () seconds to run the 2017 data, and () seconds to run the 2018 data.

## Summary
### What are the advantages or disadvantages of refactoring code?
Advantages of using refactored code includes: 
- Faster processing: Fewer lines that can be easily interpreted allows the program to quickly interpret instructions.
- Easier to debug: When writing code, it's easier to correct errors when there is less code to review.
- Broader application: Refactored code can be more readily applied to a larger dataset because individual code blocks can process information unchanged.

Disadvantages of refactored code includes:
- Irregular data bugs. Any data that may not run in the program will need to be identified and corrected in its source program, and potentially require unique code.
- Overlooking errors. It is possible that programs may skip data lines without throwing an error, meaning data gaps would go unnoticed.

### How do these pros and cons apply to refactoring the original VBA script?
In the original version, we needed to add buttons to activate distinct macros; the refactored version automatically runs all functions so that buttons were not needed. 
Additionally, it was challenging to run multiple macros simultaneously; the refactored version easily applied logic and formatting macros to the data.
That said, the factored version must be run from the script itself.

# Module 02 Challenge - Refactoring VBA Code

## Overview of Project

Refactoring of VBA Stock Analysis Code Previously Written

### Purpose

The purpose of this project was to Refactor a set of Visual Basic for Applications (VBA) Excel
Subroutines previously written to perform analysis on a set of trading volume
and price observations to see if they could be made to execute
more quickly.

The initial set of data contained one Dozen unique
Stock Ticker Symbols upon which to perform analyses, and did not take
a burdensome amount of computing time to complete. However, if the
original subroutine(s) were presented with a hypothetically much
larger set of data to compute, or a larger number of unique Stock Ticker Symbols
to consider, the code in its original form could
have potentially taken much longer to execute.

As an exercise, the original code was Refactored in a manner to
complete all calculations and analyses by looping through the
dataset only one time to completion. In this way, as the dataset
gets larger, all other things being equal, compute time would grow
linearly with the size of the data set, rather than exponentially
in the case of the original code version utilizing multiple loop
iterations.

## Results

Upon successful completion of the refactoring, both versions of the code
were tested for runtime performance for both Year 2017 and Year 2018.

The results of these tests are summarized below in

Table 1: [Seconds, rounded to 4 decimal places]
|          |2017    |2018     |
|----------|--------|---------|
|Original  |0.7344  |1.0508   |
|Refactored|0.1250  |0.1680   |

Even on a small dataset comprizing observations for 12 unique values over
approximately 3,000 rows, it can readily be seen that the Refactored code performs
about 6 times faster than the Original code. On larger datasets, one
can imagine that this could quickly add up to much more significant
runtimes in actual use. Since no matter how large the dataset grows,
the Refactored code only runs the loop one time, the runtime will grow
linearly according to the size of the dataset; however, the Original code
adds an additional loop iteration for each new unique value under consideration,
and thus the runtime could increase exponentially as the size of the dataset grows if
more and more Stock Ticker Symbol values were taken into account.

## Summary

### Advantages of Refactoring Code

In this case, a definite advantage of Refactoring the code
is the resulting increase in runtime efficiency.


### Disadvantages of Refactoring Code

A disadvantage of the Refactored Code, as it is currently
implemented, is in terms of the order of the values contained within
the dataset. In order for the correct expected results to be output,
all the values contained within the dataset need to be sorted prior
to running the analysis so that the Stock Ticker Symbols and corresponding
values are in ascending alphabetical order, corresponding to the order
of the Stock Ticker Symbols contained within the `tickers` Array. If the blocks of
values, or individual values themselves were presented out of order, a result
would be returned other than the correct expected one.

Despite longer runtimes, the Original version of the code avoids this problem by
considering each row of data in isolation and checking for a Stock Ticker Symbol match.
In this way, the Original Code would output the correct expected result for each
Stock Ticker Symbol even if the dataset under consideration was un-sorted.

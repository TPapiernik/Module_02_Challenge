# Module 02 Challenge - Refactoring VBA Code

## Overview of Project

Refactoring of VBA Stock Analysis Code Previously Written

### Purpose

The purpose of this project was to Refactor a set of Visual Basic for Applications (VBA) Excel
Subroutines previously written to perform analysis on a set of trading volume
and price observations to see if they could be made to execute
more quickly.

The initial set of data contained approximately 3,000 observations of one Dozen unique
Stock Ticker Symbols upon which to perform analyses, and did not take
a burdensome amount of computing time to complete using the Original version of the code. However, if the
original subroutine(s) were presented with a hypothetically much
larger set of data to compute, or a larger number of unique Stock Ticker Symbols
to consider, the code in its Original form could
have potentially taken much longer to execute.

As an exercise, the Original code was Refactored in a manner to
complete all calculations and analyses by looping through the
dataset only one time to completion. In this way, as the dataset
gets larger, all other things being equal, compute time would grow
linearly with the size of the data set, rather than exponentially
in the case of the original code version utilizing multiple loop
iterations.

## Results

Upon successful completion of the Refactoring, both versions of the code
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
runtimes in actual use.

Since no matter how large the dataset grows,
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

A disadvantage of the Rafactored code is in usage of available
memory. The Original code uses only one Array to contain the
Stock Ticker Symbols, while the calculated results are output
once all the rows have been looped through for each iteration.

In contrast, the Refactored code uses 4 Arrays, 3 of which are
used to temporarily store the calculated results until all the rows
have been looped through. Then the Arrays themselves are looped
through to output the results.

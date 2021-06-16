# Module 02 Challenge - Refactoring VBA Code

## Overview of Project

Refactoring of VBA Code Previously Written

### Purpose

The purpose of this project was to refactor a set of Visual Basic Excel
Subroutines previously written to see if they could be made to execute
more quickly. The initial set of data contained one Dozen unique
Stock Ticker Symbols upon which to perform analyses, and did not take
a burdensome amount of computing time to complete. However, if the
original subroutine(s) were presented with a hypothetically much
larger set of data to compute, the code in its original form could
have potentially taken much longer to execute.

As an exercise, the original code was refactored in a manner to
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



## Summary

### Advantages of Refactoring Code


### Disadvantages of Refactoring Code

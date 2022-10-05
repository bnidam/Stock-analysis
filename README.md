# Stock-analysis
Refactoring VBA code for efficiencies with larger volumes of data

## Overview of Project
### Purpose
Refactor (or edit) the VBA code developed during the Module lessons to loop through the data one time to collect the same information and then measure the run time of the refactored code compared to the original code

### Description
For clarity, I will be referring to the code developed while doing the Module as the "original code" and the refactored code as the "refactored code". 

## Refactoring Process
From a high level view, both sets of code produce the same results: a list of stock tickers with their total daily volume for a given year and the percent return for that year. The percentage return is column is conditionally formatted to quickly show positive and negative returns by color - green or red, respectively.

In the original code, the process identifies the first ticker names, collects the volume numbers, collects the starting and ending prices, calculates the return, and puts that information into the designated cells for the first ticker name, moves to the next ticker name, and so on through the list of twelve ticker names in the data. With this method, the code moves through the data twelve times.

In the refactored code, the same identification, collection, calculation, and output is accomplished by moving through the data only one time, addressing each of the stocks simultaneously.

The two sets of code are included at the bottom of this file as OriginalCode and Refactored code.

## Results
Both sets of code include a timer for the processing time that shows in a message box when the run is complete. Below are two images for each code (original first, then refactored) showing the processing time. Note that the stock analysis values are identical in all of the images, meaning that the original code and the refactored code produce the same results.
Original code:
![Run Time for 2018 using Original Code 1](https://github.com/bnidam/Stock-analysis/blob/main/Resources/2018RunTime_AllStocksAnalysis.png)
![Run Time for 2018 using Original Code 2](https://github.com/bnidam/Stock-analysis/blob/main/Resources/2018RunTime_AllStocksAnalysis2.png)

Refactored code:
![Run Time for 2018 using Refactored Code 1](https://github.com/bnidam/Stock-analysis/blob/main/Resources/2018RunTime_AllStocksAnalysisRefactored.png)
![Run Time for 2018 using Refactored Code 2](https://github.com/bnidam/Stock-analysis/blob/main/Resources/2018RunTime_AllStocksAnalysisRefactored2.png)

Below is a table showing the above run times, along with the percentage change.
![Table of 2018 run times and % change](https://github.com/bnidam/Stock-analysis/blob/main/Resources/RunTimesComp%25Change.png)

The refactored code runs much faster than the original code - 85% faster.  

## Summary
This exercise shows the potential benefits of refactoring code, illustrating that correct code - code that produces a correct, desired outcome - could be refactored to produce better results, such as faster run times.

There are many scenarios were refactoring code could be beneficial:
 - Perhaps a project was developed in sections or pieces, and the code was written in independent sections and then pasted together. Reviewing and refactoring the code could produce better results.
 - Perhaps a project has been in use for a long time. Refactoring the code with newer, more optimizing code could decrease storage requirements or allow for a more interactive experience.
 - Perhaps the volume of data being processed has increased significantly. Refactoring the code could decrease storage requirements and decrease the run time.

 Refactoring, however, may not always be practical. The cost of reviewing and doing the refactoring should be taken into account, along with the scale of potential benefits. For example, saving 10 seconds of run time for a project that takes 100 minutes to run probably doesn't justify hiring a new developer to do that work. Each scenario should be evaluated individually.

## Code
Below is the original code section that is changed during refactoring.
![original code](https://github.com/bnidam/Stock-analysis/blob/main/Resources/OriginalCode.png)

Below is the refactored code.
![refactored code](https://github.com/bnidam/Stock-analysis/blob/main/Resources/RefactoredCode.png)







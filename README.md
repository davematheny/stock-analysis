# stock-analysis
# Kickstarting with Excel
UNC Boot Camp assignment Module 2 Challenge

# Overview of Project
1. Take the excel stock data and VBA module provided and create an automated Macro.   
2. Fill out the VBA code to refactor the prior subroutine.   
3. Use variables,loops and subloops.
4. Use an input box to take in data and message box to record time.  
5. Perform analysis with the color coded data.   
 
 
# Purpose
In this challenge, I edited, or refactored, the Module 2 solution code to loop through all the data one time in order to collect the same information that I did in this module.  Then
I determined whether the refactored changes made an impact on the run time.  The new code with Click event button, loops etc should now provide a quicker way to do the analysis.

# Results
## Stock performance between 2017 and 2018
The total daily volume from 2017 to 2018 significantly decreased except for ENPH and RUN.  Basically ENPH and RUN were good investments the other stocks did poorly.  Year after year
ENPH and RUN went up, particularly ENPH.  See below 
![Graph 1. Results](resources/VBA_Challenge_2017.png)
![Graph 2. Results](resources/VBA_Challenge_2018.png)

## Execution times of the original script and the refactored script
Below are the run times before the refactor
![Graph 3. Results](resources/Old Code VBA_Challenge_2017.png)
![Graph 4. Results](resources/Old Code VBA_Challenge_2018.png)

And here are the times for code after the refactor
![Graph 3. Results](resources/VBA_Challenge_2017.png)
![Graph 4. Results](resources/VBA_Challenge_2018.png)

We can clearly see that the refactor cut down on the time significantly, almost over half a second for both years.


# Summary
## What are the advantages of refactoring code?
The advantages are you could come across a quicker or cleaner way of the coding the solution.  I think its always good to brainstorm or have second set of eyes on a problem.
I think no matter how good you think you are at programming its always good to keep an open mind, there is always something new that you dont know.

## What are the disadvantages of refactoring code?
The biggest disadvantage I see to refactoring code is if its not broken dont fix it.  From my own experience I think its good to keep a similar architecture.  I dont think its a good
idea to do something different for the sake of doing it different.  Keep things simple and similar, so that when your up at 4am fixing a production issue you dont have to keep jumping
around.  Also any time you change something your going to have to test the hell out of it.

## How do these pros and cons apply to refactoring the original VBA script?
The advantage was we were able to cut down the run time significantly.  To me this was worth the change, it had a real impact on the end user.  On the other hand it needed to be tested,
in this case it wasnt much work.  Also the code did get more fancy so you would also need to make sure your team could handle maintaining more complicated code.

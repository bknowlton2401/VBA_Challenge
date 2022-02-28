# VBA_Challenge

Stock market analysis to compare year to year returns.

## Overview of Project: 

The purpose of this analysis is to automate an Excel spreadsheet using VBA code.  The spreadsheet contained the comparison of opening and closing stock prices of 12 different companies from 2017 to 2018. The goal was to refactor the orginal code, to make it run more efficiently, cleanly and faster than previously. I added a button, to make it user friendly for any future clients that would like to see each year and to clear the values.   

## Results: 

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

To begin, I copied the working code and began working on the ticker Index, output arrays, created a loop to initialize the tickerVolume to zero and to loop over all the rows in the speadsheet.  

![image1](https://user-images.githubusercontent.com/96890065/155920480-9b57573a-21ad-4276-a3bf-c36c18d0073c.PNG)

A problem I encountered was not assigning the variable to calling the each index (i) to set the array to zero.  Another problem I encountered was to give each dimension the correct size (12). 

Next, I worked on increasing the volume for the current ticker and ran conditinal statements to check that the first row, current row and last row contained the same ticker symbol. 

![image2](https://user-images.githubusercontent.com/96890065/155920627-19ce4c45-061b-47cb-a3d7-8fb6dcc8cad6.PNG)

A problem I encountered here was mis-typing key words.  For example, I mispelled index as indes.  

Finally, I was able to have the code read through all the arrays to give an output for the Ticker, Total Volume and Return for the indicated year. 

![image3](https://user-images.githubusercontent.com/96890065/155920738-479dee9a-04c5-471c-bc2d-6e1c2767540f.PNG)

Due to the previous misspellings, my code kept encountering an error on the last line.  Once I fixed the misspellings the code ran fine. 



## Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

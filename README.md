# VBA of Wallstreet

## Overview of Project
This project analyzes stock returns of energy companies, to be used as a resource by client Steve.  

### Purpose
 Steve will be using
this data as a means to inform the financial advice he provides to his clients regarding their invesment portfolio.

## Results

##### Comparison of stock performance between 2017 and 2018
The 2017 returns for the stocks evaluated were predominantly prosperous, with only one stock closing the year with a
negative return. The opposite is true of the 2018 returns for most stocks. A majority of the reviewed stocks closed 2018
at a loss, with only ENPH and RUN tickers producing gains at the year's close. Looking more closely at these two stocks,
ENPH did experience a significant loss of its returns from 2017 to 2018, its 129.5% gain in 2017 was reduced to 81.9 
the following year. Only RUN consistently produced gains. 2017 may have led to only a 5.5% return.
However, 2018's gains increased to 84%. Based on the data, RUN may prove to be a more reliable investment.

##### Comparison of execution times between original script and refactored script
The execution time for the year 2017 of the Refactored script reduced the execution time of the original script by 85%.
The original script runtime of 2017 was 1.875 seconds, reduced to 0.28125 seconds in the refactored script.

The execution time for the year 2018 of the Refactored script reduced the execution time of the original script by 85%.
The original script runtime of 2018 was 1.828125 seconds, reduced to 0.2734375 seconds in the refactored script.

The above percentages were calculated by finding the difference between each year's execution times, dividing that
difference by the original execution time and multiplying by 100 to get a percentage.

It is abundantly clear that refactoring improves the efficiency of the code by greatly reducing the execution times.
Despite the time invested it required to refactor, I believe it is a sensible choice to allocate a portion of resources
and personnel to refactoring code, depending on the code's purpose. 

## Summary 

##### What are the advantages and disadvantages of refactoring code?
Some advantages of refactoring code include reduction of code repetition. Repeatedly copying code with an error can make
debugging arduous, as the error will also be repeated and will need to be addressed in each line it is present in order
to complete debugging. This, in turn, would make debugging more time consuming. Refactored code also increases its human
readability and as a result makes the data more accessible to a wider audience, one that may be less experience in
reading code.

Some disadvantages of refactoring code include that it is a time consuming process. It occurs after the code is written,
as an afterthought. It typically only occurs budget and time allowing, as it is otherwise nonessential to the code.
Refactoring also decouples elements within the code and, as a result, can make the code more complex.
 
##### How do these pros and cons apply to refactoring the original VBA script?
Refactoring the original VBA code is an important skill to learn, but it is one that is not likely to be used in all
work situations. The code may be more visually accessible, but dedicating time to refactoring is still a matter of
allocating resources to a task that is not essential for the effectiveness of the code.
###### stock_analysis
An Analysis on Green Energy Stock.

![Wall Street Picture](https://user-images.githubusercontent.com/93900628/144791424-8e4b48e5-2006-4e82-896c-bdb8c6905fe2.jpeg)

# Stock Analysis

## Overview of Project

### Purpose:
Steve's parents are interested in investing in Green Energy Stock, they are looking at putting all their money in to DAQO New Energy Corp. (ticker DQ), they have not done any research on DQ and have asked Steve to help them with their decision. Steve has data for 12 Green Energy Stocks for the years 2017 and 2018. Steve is going to use this data to determine the Total Daily Volume and Yearly Return, based on this he can give his parents an informed recomendation.

### Results:

Definations (referenced from Module 2.2.1)

 - Total Daily volume is sum all the total number of shares traded throughout the day in the year; it measures how actively a stock is traded. 
 - Yearly return is the percentage difference in price from the beginning of the year to the end of the year.

#### DQ Analysis (2018):
The data from 2018 was used to determine the Total Daily volume which was almost 108M as well as the yearly return of -63%. Steve decided to do an analysis on all 12 of the stock for 2017 and 2018 since the 2018 return on DQ was so poor.

#### All Stock Analysis (2017-2018):

The analysis on the 2017 and 2018 yearly return shows that only EPNH and RUN performed better in 2018, all other stock had poor performanses in 2018 vs. 2017. (see screenshots for results)

###### All Stock Analysis (2017)
![Screen Shot 2021-12-05 at 11 11 21 PM](https://user-images.githubusercontent.com/93900628/144792190-964502c4-5e73-424f-b5f2-fb831ed9a635.png) 

###### All Stock Analysis (2018)
![Screen Shot 2021-12-05 at 11 11 34 PM](https://user-images.githubusercontent.com/93900628/144792210-ee0a0c52-760c-4c6d-b119-2d57d6afac31.png)

###### Refactored Code:
Steve has determined that he could use this workbook to analyse data for future years and has asked to refactor the VBA code so that it runs faster.

The orginal code took almost 0.30 seconds to run for 2017 and 0.296 seconds for 2018. (see screenshots below)

![Screen Shot 2021-12-05 at 9 41 04 PM](https://user-images.githubusercontent.com/93900628/144794000-1794d49c-ee61-4df3-a0d2-824e6b8861d8.png)![Screen Shot 2021-12-05 at 9 42 02 PM](https://user-images.githubusercontent.com/93900628/144794048-34788b67-bd96-4512-a40a-abfa180835b0.png)

After the code was refactored it took 0.085 seconds to run for 2017 and .078 seconds to run for 2018.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/93900628/144794524-b2a681c2-e56d-4446-ba80-ad4604d091bc.png)![VBA_Challenge_2018](https://user-images.githubusercontent.com/93900628/144794540-98062c34-18fb-48c9-9560-5f3ec23ecda1.png)

#####

### Summary:
##### What are the advantages or disadvantages of refactoring code?
- Advantages of refactoring code is that it improves readability, the code is cleaner and much easier to comprehend. When working with big sets of data, you can save a considerable amount of time as the code runs faster.

-  Disadvantages of refactoring code if there is a lack of understanding of what the code does you risk creating bugs, refactoring is also very time-consuming.

##### How do these pros and cons apply to refactoring the original VBA script?
- Pros: The time to run the code has decreased by almost 300%, this is useful when using this workbook to analyze massive amounts of data.

- Cons: It was time-consuming to debug the modified code. Debugging would be challenging and time-consuming if there were thousands of lines of code. 

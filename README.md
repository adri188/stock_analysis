# stock_analysis
## Background

Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

## Results 

### Stock Performance 

#### 2017
<img width="232" alt="Screen Shot 2022-04-16 at 3 35 31 PM" src="https://user-images.githubusercontent.com/102937320/163693316-e0130e1d-0512-433e-995c-f10f2ac98d2f.png">

#### 2018
<img width="232" alt="Screen Shot 2022-04-16 at 3 33 59 PM" src="https://user-images.githubusercontent.com/102937320/163693280-b7e2388e-466c-4033-bed0-b98bfde7439f.png">

From 2017 to 2018 this Stocks had quite a turn around, with the exception of  RUN that showed growth. Only RUN and ENPH maintained a positive return rate . 
DQ trippled in volume however had quite a negative turnaround from ~200% return in 2017 to -62.6% 

### Code Performance

#### 2017

As seen on the images below Refactoring the original code made it real fast , only took 0.08 to run which is significantly more efficient. Original code was taking aroun 0.28 secs. When analyzing bigger data sets this will make a big difference. 


![VBA_Challenge_2017](https://raw.githubusercontent.com/adri188/stock_analysis/864dcfa55c7a07805932b52d470aaa5c9e3274ca/Resources/VBA_Challenge_2017.png)

#### 2018

![VBA_Challenge_2018](https://raw.githubusercontent.com/adri188/stock_analysis/864dcfa55c7a07805932b52d470aaa5c9e3274ca/Resources/VBA_Challenge_2018.png)


## Summary
' advantage and disadvantages of refactoring code and How do these pros and cons apply to refactoring the original VBA script?
ReFactoring the code gives you the opportunity to go over ways to always improve the way you programm, also allows for avoiding wasting time trying to start everything from scratch when a similar solution is already developed and available. However when doing so one must make sure the code is well tagged and structured, and that you are able to follow a structure and a debugging/testing strategy that allows for the process to really be value added rather than consuming more time and effort. 
Regarding the Challenge Example 

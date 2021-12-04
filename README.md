# Stock Analysis - Challenge 2 
## Overview of Project 

* In this analysis we are using VBA to analyze Green Energy Stock Data to help our client make a clear decision when choosing which  stock to invest in. In our previous work we compared stocks from years 2017 and 2018, however, the purpose of this project is to create code that can be used for thousands of different stocks so we can easily analyze which stocks provide the best returns. To do this, we need to refactor the current code we have so that we can utilize it across all stocks and not just a couple. 
* By refactoring, we can improve the logic, use less memory, and take fewer steps, allowing the analysis to run at a faster pace. This is useful, especially when dealing with stock data with thousands of sheets. 

## Results of Project

* In order to gauge, how efficient our refactoring of the code is, we can compare the time it takes to run each code. The faster the code runs, assuming that it holds the same functions, the more efficient the code. Below we can see the comparison of our code before and after refactoring by comparing Stock Data from 2017 & 2018. 

### 2017 Efficiency: Before & After Refactoring

* Before Refactoring: The code ran in 1.18 seconds
<img width="262" alt="2017 - Time WITHOUT refactoring" src="https://user-images.githubusercontent.com/94571150/144718377-d4b0b1ff-d3c2-471d-8b34-9a38b2a26231.png">

* After Refactoring: The code ran in 0.2 seconds
<img width="265" alt="VBA_Challenge_2017 png " src="https://user-images.githubusercontent.com/94571150/144718417-0f22da3e-24a8-4f95-bbac-dd3789b0cf7d.png">

* In 2017, there was a difference of roughly 1 full second to run this code. As a result, the refactoring was successful as it provided the same function in less time using fewer steps. 

### 2018 Efficiency: Before & After Refactoring

* Before Refactoring: The code ran in 0.98 seconds
<img width="264" alt="2018 - Time WITHOUT refactoring" src="https://user-images.githubusercontent.com/94571150/144718524-87135d68-b923-47b2-9467-899a44a0d1f3.png">

* After Refactoring: The code ran in .18 seconds
<img width="261" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/94571150/144718598-548e59a5-27d8-40bf-b397-832f72f1960b.png">

* In 2018, there was a difference of roughly .8 seconds to run this code. As a result, the refactoring was successful as it provided the same function in less time using fewer steps.  

## Summary

### Pros & Cons: Refactoring Code

* By definition, refactoring code allows for code to run using less memory, fewer steps, and improved logic. In this way, refactoring code holds many advantages as it can provide a cleaner and more concise script and can run more efficiently and in less time. When dealing with a large amount of data in the real world, timing is a key factor to having accurate data and meeting deadlines. Having refactored code that provides the same functions but more efficiently will prove beneficial for effective analysis. That being said, refactoring code when can also 

* generalized easily understood

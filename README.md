# VBA of Wallstreet

## Overview of Project

### Purpose
In this project, we were tasked with finding a way to create a piece of looping code that would search through a large data set with a list  of stock tickers and their respective starting and closing prices for 2 different years: 2017 and 2018. While we could have used cell formulas to accomplish this, we wanted to make a tool that would allow for more data to be added to the original sheets and make sure that our processes still worked no matter how much data was added. Throughout the course of the project, we built on to our VBA code so that it became more efficient, comprehensive, and readable. Using For loops, we were able to have our code pull from the relevant rows of data for each ticker and then display the total grossed amount and also the return on investment for either year. From there, we refactored our code to make it run more efficiently so that when we add more data to the set, it will continue to run smoothly and without issue. 

## Analysis and Challenges

### Results

#### What the data shows us.

In 2017, almost all of the stocks created a positive return making it seem like all of the options of green stocks seem relatively safe (except for TERP). However, after running our analysis on 2018 we found that the opposite to be true. All but two stock options made negative returns in 2018 (ENPH and RUN). Looking at the data itself (without taking into account outside factors such as marketing and public perception on "green" businesses), it shows that the majority of these green stocks are relatively volatile and more due diligence should be performed to get a more complete look at the risk or reward of these 12 green stocks moving forward.


##### Charts Showing Returns for 2018 and 2017


![2017_Chart](https://user-images.githubusercontent.com/85508764/122953525-658e6600-d344-11eb-8551-71c2f15de701.png)&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  ![2018_Chart](https://user-images.githubusercontent.com/85508764/122953731-78a13600-d344-11eb-8ad2-18f475a306ee.png)

#### Runtime of Code Before and After Refactoring

As already discussed, our code was running for almost a whole second. Which, in today's world of high-powered machines in our pockets is very slow compared to what most people are accustomed to using on a daily basis. Through the use of a variable index (tickerIndex) we were able to have our code run more efficiently because this allowed all of our loops ran at once rather than having it go through each ticker one by one.

#### Original Code

![2017_Not_Refactored](https://user-images.githubusercontent.com/85508764/122947429-c23b5200-d33f-11eb-8433-2c9fea3a3507.png)&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  ![2018_Not_Refactored](https://user-images.githubusercontent.com/85508764/122955586-d8e4a780-d345-11eb-8a91-e15ebeae48a6.png) 


#### Refactored Code

![2017_Refactored](https://user-images.githubusercontent.com/85508764/122956179-61fbde80-d346-11eb-808f-6ee58a85514f.png)&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ![2018_Refactored](https://user-images.githubusercontent.com/85508764/122956100-53152c00-d346-11eb-96c7-496f151fe966.png)



### Summary

## What Are the Advantages or Disadvantages of Refactoring code

- Advantages of Refactoring Code

Refactoring our code made it so that we can easily add more data to our workbook and have our code be more modular. Also, as discussed previously it allows for our program to run around 90% faster than it originally did! This will be important for our client who will be presenting our workbook to clients of his and shows another level of sophistication for our VBA macro.

- Disadvantages of Refactoring Code

The main disadvantage that I could see with refactoring code is that it takes more time away from development of other macros and other ways to analyze the data that was given in the workbook.  

- How do these pros and cons apply to refactoring the original VBA script?

# stock-analysis

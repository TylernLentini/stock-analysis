# Stock-Analysis

## VBA on Wall Street

### Overview of Project

### Background

    Unless you're experienced, investing in the stock market will be daunting. Overtime, stocks rise and fall, at times with predictable or at least historically precedented market trends but sometimes stocks act unpredictably. Unpredictability breeds uncertainty and that can scare new investors away from trying their hands at investing. 
    The answer to that uncertainty is data. The numbers don't lie. While investing in the stock market can be daunting, tracking stock performance and making educated investment decisions doesn't have to be. 
    
    In this lesson, Steve tasked us to help him research the stock, DQ, track it's performance and create code that will help his parents choose healthy investments in the Green Energy sector. One way to track a stock's peformance is to calculate it's yearly return. The yearly return is the percentage incerase or decrease in price from beginning of the year to the end of the year. Then we compare yearly return of the stock over the course of several years to determine if the stock is a healthy investment for Steve's parents. 

### Purpose

    In the challenge, I will refactor the code created to help Steve with his analysis of stocks from 2017-2018 in module 2. We will loop through the data one time to collect that same data from the module and determine if the refactoring of our code has made the VBA  script run faster. Lastly, I will compose a written analysis to explain my findings. 

### Results

### Compare the stock performance between 2017-2018

 DQ saw a 199.4% increase in 2017! That would be incredible for Steve's parents if they had invested at the end o 2016. However, DQ saw a 62.6% decrease in 2018. A decrease is unfavorble but this results in a 136.8% net positive return over two years. That's a great investment short term. We would need to execute our analysis over a greater time sample and determine what type of investment Steves' parents are looking to make to determine definitely if DQ is a good fit for them as investors. Additionally, the data begins to give us a glimpse at how renewable energy is performing as a trend. The vast majority of stocks in this list saw a positive return in 2017 and a decrease in yearly return in 2018. 2018 was a bad year for renewables. TERP and JKS saw a net negative return from 2017-2018. TERP decreased year after year while JKS' saw a loss of yearly return in 2018 that outweighed it's 53.9% growth in 2017. That being said, there are many factors that influence market trends. I tend to look at something like renewables as a longterm investment because factors like poliy change and corporation holdouts can influence the price of these stocks in the short-term. If Steve's parents believe in renewables and in DQ more research should be done. 

### Compare execution times of the scripts 

My execution times ran only slightly faster after refactoring the code. I conferred with classmates who had much faster code after refactoring and our changes matched. I also consulted a TA about my concern. He said it's not a satisying answer but there's nothing wrong with the code. The code does everything the module asks. Sometimes code run times differ depending on machine or how the code is executed; with a button, from developer etc. 

He was right, that wasn't a satisfying answer. SO I googled, and googled, and added with statements and reduced lines and added semicolons and reduced indexes to single letters. It's all there and honestly it was good practice but I never got my run times to be significantly faster, just *slightly*. 


## Summary

### What are the advantages or disadvantages of refactoring code?

I do see the advantages of refactoring and optimizing your code. It's a good way to make sure your logic is correct and your code can be read and **understood** by other developers. 
This makes collaboration a whole lot easier. Code that is unclear or not using the most efficient logic can create more problems down the road. It will save time overall to refactor your code as you go and while you are fresh on the current objective.

### How do these pros and cons apply to refactoring the original VBA script?
The disadvantage I experienced is only a very minimal decrease in runtime though I spent hours refactoring. I do think the resulting code is cleaner and
more simplified but I can also see times when a quick answer will be necessary.
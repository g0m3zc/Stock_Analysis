# Stock_Analysis

## Overview & Purpose of Project
In this analysis, I reviewed stock returns for alternative energy companies in the years of 2017 and 2018.  The stock of interest was DQ for this project.  I did expand the analysis to include other options in this category of stocks.  

I was able to use VBA in Excel to conduct this analysis.  Through VBA I calculated the total daily volume for the year as well as the % return for the full year.

## Results
DQ increased in total daily volume more than two fold in 2018 when compared to 2017.  However, their return went down in the same year.


![All_Stocks_(2017)](https://user-images.githubusercontent.com/106936638/175470666-d533b83c-8497-444d-a245-6e03657908dd.png)

TERP was the only company to have a negative return in 2017.

<br/><br/>

![All_Stocks_(2018)](https://user-images.githubusercontent.com/106936638/175470692-ba8e404c-bdbe-4df6-b8bc-b6de9033dd5c.png)

Overall, the alternative energy companies performed much better in 2017 than they did in 2018.  Only ENPH & RUN had positive returns in both years.
The company with the largest total daily volume in both years was FSLR.  

## Summary
Through this analysis I compared the efficiency (speed) of using refactored code to non-refactored code.  When analyzing the two years of data, the refactored code performed at a fraction of the time.

Challenges of refactoring code includes the initial investment of time to rework code that already works as is.  It also challenges one to think creatively and perhaps use other resources or methods then initially attempted. 

The benefits of refactoring code are especially worthwhile when the code will be used for multiple datasets, over time, or is a repeated task that needs to be completed.  It is also an investment in supporting any future users with the understanding of the code and its purpose.

In this example, refactoring the code to leverage loops and if/then statements made it possible to conduct the same analysis over a list of different tickers.  It also allowed for increases in speed.

###### Initial Code Length
![Initial 2017](https://user-images.githubusercontent.com/106936638/175471806-5f18d890-0157-445c-98c4-857bd9ea879a.PNG)
![Initial 2018](https://user-images.githubusercontent.com/106936638/175471809-d1ecf09a-dfd7-49f9-946c-c306df8ad181.PNG)

###### Refactored Code Length
![VBA_Challenge_2017](https://user-images.githubusercontent.com/106936638/175471819-f635c456-e1cf-4077-be9c-761e7254e8b7.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/106936638/175471825-df749543-1841-4f41-9774-9653ad3ba728.PNG)

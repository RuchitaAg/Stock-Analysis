# Stock Analysis with VBA


## Overview of Project

### Purpose

As part of this analysis I am helping a friend Steve, a fresh finance graduate, who plans to start his stock consulting business. His parents want to be his first client and they would like to invest all their money in **DAQO New Energy Corp.**, which is a supplier to solar panel manufacturing companies, as they are passionate about the future of Green Energy/Alternate Energy. 

Since, they haven't done due-dilligence and market study to check performance of DAQO stock in the past, Steve woud like my assist in making the best choice for his parents and suggest a stock that help them maximize their ROI. 

For this purpose, I conducted a deepdive and developed a VBA code to automate the analysis and share the findings to optimise returns.


## Analysis and Findings

### Analysis:

As part of the review I studied the outcomes of DAQO stock first (_Ticker DQ_) and realised that in the year 2018, the stock had dismal performance of &#x1F53B;63%. The analysis also suggested that the traded volume for DQ for the whole year was 104 Mn.

Since, this doesn't appear be such a great investment, I expanded the excercise to include all other stocks and studied their performance in the Year 2018. 

This revealed that 2018 wasn't a great year overall as almost all stocks, barring 2 (ENPH and RUN), were bleeding red. This data also highlights that DQ was one of the least traded stock for some reason.

Refer below snaphot of 2018:

![image](https://user-images.githubusercontent.com/102870991/164873183-4bd96cb6-53b6-443d-82b7-77d843d7f0cc.png)

 
I wanted to be sure which stocks are consistently performing so I made the code dynamic to summarise analysis for any year at a click of a button. Refer below summary for analysis of all stocks in 2017 for their volumes and returns:

![image](https://user-images.githubusercontent.com/102870991/164873189-6881c5bd-437d-4d0d-ba8a-a2c6228b2375.png)
 

### Findings:
The analysis for 2017 suggested that while DQ was the least traded stock even in 2017, it generated maximum returns. Standalone this is excellent news for DQ investors but when comparing with 2018, ENPH turns out to be a clear winner as even in the year of down-turn, it not just continued to generate positive returns but also showed growth of 174% in the total volumes.

RUN showed a big swing in the return from mid-single digit return in 2017 to highest return of 84% in 2018. But this kind of volatility doesn't make it the safest bet. So I would recommend to either keep it on a watch list or make a small investment to diversify out invetment. 

Although, DQ was one of the few stocks showing triple digit growth in transactions in 2018,its recent negative returns do not make it the best bet. I suggest monitoring it's performance for another year before putting stakes in it.


![image](https://user-images.githubusercontent.com/102870991/164873803-6a0d8e7e-b674-4037-abf5-7f26d562efd7.png)


### Analysis of Code Performance


The code runs fast and it is seen that the overall time taken to conduct the analysis for 2018 and 2017 with refactored code is around 2 seconds as seen below:

![image](https://user-images.githubusercontent.com/102870991/164873203-f844d6fa-8591-4e7b-81b3-02b778c907b1.png)

 
![image](https://user-images.githubusercontent.com/102870991/164873215-996d5513-965f-46d3-8b0f-e9cbc374eab6.png)


##### Benefits of Refactoring of a VBA code -

Code refactoring means restructuring your existing code, in a way that improves the internal structure but doesn't change its external behavior i.e. the code formatting and naming the variables, defining the steps etc. without changing the final output. This helps in making the code efficient for the end-user.


### Refactoring of Stock Analysis code -


The All stock analysis code used multiple variables like i, j etc. for looping through therows or columns on different sheets. I refactored the code to use tickerindex as the variable and initialised that as 'Long' datatype. 
The refactored code also includes step # with well defined headers making the code more readable. The code does all that and still gives the same output.
The total run-time of the code reduced by almost 50%.

Run time for original code is as below:

 ![image](https://user-images.githubusercontent.com/102870991/164873769-117e3775-731e-4e04-abd0-5cec6a77cdea.png)

![image](https://user-images.githubusercontent.com/102870991/164873780-2f1da18c-7c6b-46a1-a942-4cc54e52d36b.png)


 

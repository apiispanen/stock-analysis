# Stock Analysis with VBA

## Overview of Project
The purpose of the project is to analyze green stock data for 12 stocks to aid in Steve's analysis. We will be looking at total volumes of trades during two years, 2017 and 2018. Using Visual Basic for Applications, we can investigate the effectiveness of Steve's portfolio over the 2 year span.

### Background
The dataset here is for green stocks, which are environmentally friendly stocks. Per our module, we have picked 12 stocks to analyze specifically, in aid of Steve's project. 

## Results
Based on the data we compiled, we can see an obvious trend between the years - generally speaking, the returns have flip flopped between years. In 2017, the returns on all stocks but one (TERP) were positive.
![Results from our 2017 Macro]("Resources/All_stocks_2017.png")

 Contrasting that to the following year, 2018, all stocks but 2 (ENPH & RUN) are all negative, indicating a down year for Steve's portfolio.
![Results from our 2018 Macro]("Resources/All_stocks_2018.png")



## Challenges
Using VBA to analyze stock information has many benefits. With its benefits, however, there are significant setbacks.  

### Using VBA
Visual Basic for Applications is a useful tool for analyzing spreadsheet data in bulk. When compared with other bulk data analytics tools such as Python, it may be limited in scope. Since the code works under the operating system, it can be problematic when a code turns buggy, especially when working with loops. However, for working with Microsoft Excel files, it can be efficient and simple to get started with. 

### Refactoring Code

Refactoring code that was taken from the module in VBA can be quite difficult. For one, refactoring existing code and changing data structures require significant debugging and error handling. However, there are both advantages and disadvantages to doing so:

- Advantages
Refactoring code is extremely useful for optimizing existing codes. For one, we simplified the data structures by creating arrays for the volumes and starting and ending prices. It can also give a developer a head start on a problem, utilizing existing code and changing it internally to optimize the amount of processing it requires, the simplicity of the code, or to make it less buggy. 

- Disadvantages 
Refactoring existing code can also make waste. For one, it can cause unecessary extra steps to be made, as usually refactored code is not made in tandem with its original. Instead, it is typically made by another developer, or by the same person after a substantial period of time has passed. This results in a lack of full understanding of the code, whereas starting from scratch may help to enhance their understanding of the code they are working with. 


## Conclusions
Steve's portfolio has seen a downward trend from 2017 to 2018. Only one stock truly prevailed, and that was "RUN" - which went from a positive 5.5% return in 2017 to a positive 84.0% return. These were both positive growth rates, indicating growth in the company's equity positions by a substantial margin. <br> Overall, there was a significant decline in returns in 2018 when compared to 2017, which would be advisable to err on the side of caution when selecting next year's portfolio. To summarize our findings, Steve should consider a new portfolio for the future years.
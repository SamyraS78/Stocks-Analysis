# Stocks-Analysis
Overview of Project: Explain the purpose of this analysis.
This analysis aimed to compare the return rates for the stock portfolio under review, specifically to review the performance from 2017 to 2018. To identify these trends VBA in excel was utilized to develop a code that could analyze performance. To initialize our analysis a subroutine was created in VBA to identify the steps that would need to be run to compare our data for analysis. 
This subroutine was the outline for the code that would be developed to carry out the necessary equations or calculations to determine performance. Once it was understood what needed to be measured, the data was organized and structured for the output of our findings. 

 
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
2017 Analysis vs 2018 Analysis 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/102603966/167313946-3ccfbb6d-5ede-492b-ab1b-6a91417a121d.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/102603966/167313952-71095cb4-c8a0-4a99-a792-6fb5fd2ed7b6.png)

After reviewing the results of the “All Stocks Analysis” it appears that DQ has not been trading as high as it was in 2017. This can indicate several issues, however, based on the overall results there were two stocks that showed promise of sustained performance. ENPH and RUN were able to maintain positive returns over the last two years, based on this information it may be safe to forecast a third year of positive performance in 2019. 
 
       
   
 
All Stocks Analysis vs. Refactored All Stocks Analysis
The “All Stocks Analysis” subroutine was developed to calculate the performance of the stocks by year, at the bottom of the code was the dimension for the timer and message box function. Noticing this structure appears to contribute to the extended length of time it took to calculate the performance time. 

All Stocks Analysis Code
 
![Module 2 timer code](https://user-images.githubusercontent.com/102603966/167314009-d91cc3ec-b979-4263-bfda-231ccc64ab48.png)

Refactored Stocks Analysis Code
 

 
![All Stocks Top](https://user-images.githubusercontent.com/102603966/167313987-2f29c4d5-e698-43e9-8011-b90db60ab8b0.png)


![All Stocks Bottom](https://user-images.githubusercontent.com/102603966/167313995-6373388b-4b43-4202-a9c7-fbc4c3f73cee.png)



“All Stocks Analysis” 2017 Timer vs. “Refactored All Stocks Analysis” 2017 Timer
 
  ![VBA_Challenge_2017_timer](https://user-images.githubusercontent.com/102603966/167314017-b7dceb9f-6117-4a3c-ad07-746a0c25c411.png)

 
 ![VBA_Challenge_2017_timer_AllStocksAnalysis](https://user-images.githubusercontent.com/102603966/167314020-862b58b4-209a-4a41-9df3-ba3b229a8d2a.png)


“All Stocks Analysis” 2018 Timer vs. “Refactored All Stocks Analysis” 2018 Timer
 
 ![VBA_Challenge_2018_Timer_AllStocksAnalysis](https://user-images.githubusercontent.com/102603966/167314023-04cd9871-220d-46a6-9382-7d20200be27b.png)

 
 ![VBA_Challenge_2018_Timer](https://user-images.githubusercontent.com/102603966/167314025-3886ecbb-45c7-48ca-b56c-a5ed28d83236.png)


Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
The advantages of the refactoring script are many, while learning how to edit VBA is complicated, it can be accomplished. Specifically, nesting the for loops within themselves can save a lot of time when steps need to be repeated across large data sets. In the project example, the ticker was to loop over 11 times from the top to bottom row, while starting new rows once the steps were completed and needed to be repeated. I feel that once this step was understandable it can be applied to save a lot of time and yield faster results more efficiently. 
How do these pros and cons apply to refactoring the original VBA script?
Specifically, once the VBA code was refactored it reduced the time to calculate the total ticker volume and returns from minutes to seconds. This is an advantage when clients call with unexpected questions, or if a manager needs to know the rate of return for a meeting within the hour. 

# VBA Challenge

### Overview

​	The data sets were provided with the intent of facilitating the client's ability to perform analysis on a number of different stocks and their performance over a 2 year span. A macro was created in VBA to output basic calculations about a number of stock options, allowing for intuitive analysis. An interface was created to allow the user to perform these basic functions with the click of a button.  

### Results

​	Below is a comparison of the completed analysis of all stocks by year. The "Total Daily Volume", or the number of trades of the stock in a given day, is summed over the course of the year. In the "Return" column, the stock's price at the end of the year is divided by the price at the beginning of the year, and converted to show percentage growth or loss. This indicates how much return on investment a stock in a given corporation will provide the owner, with positive (green) values indicating increased value and negative (red) indicating losses.

<img src="C:\Users\soude\Desktop\Data Analytics Bootcamp\Module 2  - VBA of Wall Street\Resources\All_Stocks_Comparison.png" style="zoom: 80%;" />

​	The conditional formatting here allows the differences between stock performance in 2017 and 2018 for the selected companies to be displayed easily and clearly. In 2017, only one ticker ("TERP") shows a negative return, while in 2018 most of the selected tickers showed a negative return. "DQ", the main stock of interest has a spectacular year with a nearly 200% return in 2017. However, in 2018, that return drops to -63%. While the nature of the stock market is a risky one, a change of this magnitude likely indicates major changes or setbacks in the business which it is unlikely to fully recover from in the next year. Other options with less drastic changes likely indicate a less risky stock purchase/portfolio. 

​	Code refactoring was a major part of this project. The initial analysis was written using a nested for loop - an iterative process within which multiple additional iterative processes are contained. An example of the code is shown below: 

```visual basic
 
    'initiate ticker loop and totalVolume
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
       
            'activate data worksheet
            Worksheets(yearValue).Activate
            
            'initiate nexted row loop
            For j = 2 To RowCount
            
                'find total volume for current ticker
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                
                End If
```

While nested for loops are an effective way to complete the project, they do come at a price - speed. 

![](C:\Users\soude\Desktop\Data Analytics Bootcamp\Module 2  - VBA of Wall Street\Resources\Timer1_Comparison.png)

In the loop above, the computer runs through *each record*  of the data set once for each possible ticker category (or *i*). In this case, there are only 12 categories and 3013 rows of data. However, in a worst case scenario, each record could be unique, meaning that the computer would run through the values n<sup>2</sup> (3013 <sup>2</sup>) times. 

 

​	To reduce the time and resources consumed, we refactored the code into an array (snippet below).

```visual basic
'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
   
 'Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```

Creating the index and using the array method allows us to reduce this time because the computer now needs to access each data row only once. By using pre-sorted data, we are essentially designating a zone, and once the computer recognizes that it has left the "AY" zone, it doesn't have to go check back through those rows to confirm that there are no other values in that range. This cuts down the time resources required to n, rather than n<sup>2</sup>. 



![](C:\Users\soude\Desktop\Data Analytics Bootcamp\Module 2  - VBA of Wall Street\Resources\Timer2_Comparison.png)



### Summary

​	Refactoring code is a generally beneficial process which can change the way that code is implemented, how resources are allotted,  or the design of the project. The goal with refactoring is typically to improve readability, reduce complexity (and thus increase speed), and streamline the code, making it easier to maintain or extend. However, refactoring code has its own challenges. For example, an improved process may not always be clear and coming up with new solutions can be time/money consuming. In addition, if the code breaks while refactoring only a portion of it, locating the root issue and debugging can be more difficult. 

​	Refactoring the original VBA script here provided a lot of benefits. the main being decreased logarithmic complexity (and thus, quicker results). By removing the nested loop, the code is also more readable and simple looking. 


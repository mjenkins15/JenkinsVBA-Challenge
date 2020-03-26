# JenkinsVBA-Challenge
Stock Market Analysis- The VBA of Wall Street

In this assignment, we used VBA scripting to analyze real stock market data and are looking to find the yearly change from opening price at the beginning of the year to closing price as the end of the year, the percent change from opening price at the beginning of th eyeat to the closing price at the end of the year, and the total stock volume of the stock. 

Files provided:
[Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.
[Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.


After setting up our initial variables, we need to create a way to keep a running total of the volume of stocks with the same ticker symbol. Even though this is listed as one of the last things to do, it should be done at the beginning of writing the code, after variables (Dim) have been set. Below is the code used to keep a running total of the total volume of stock in each ticker. 

'Set an initial variable for holding the total stock volume per ticker
Dim ticker_total As Double
    ticker_total = 0
    
In order to determine what the yearly change from the beginning of the year to the end of the year was, we need to create a For loop to loop through all of the rows given in our dataset. (For i = 2 To lastrow). We need to set the ticker name, make sure to continously add to our running total of stock volume within each ticker, and print those results in the summary table. It is important to reset the ticker total to zero. This ensures that the total stock volume will reset each time a new ticker symbol is found and will only total the stock volume in each specific ticker. 
    
    


# stock_analysis


### Overview of the project
In the project - VBA_challenge, I refactored our initial code which means modifying the code so that it runs faster than it did before. It does not mean creating the code again but making it smaller and effecient.
#### Analysis 

Changed the module name to **AllStocksAnalysisRefactored** and added the timer

    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
Create a header row

    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

Initialize array of all tickers
    
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
    
Activate data worksheet


    Worksheets(yearValue).Activate
    
Get the number of rows to loop over


    RowCount = Cells(Rows.Count, "A").End(xlUp).Row


###### Step 1 

a) I created the tickerIndex variable = 0.
          
           tickerIndex = 0

b) Created the output arrays as follows:
- tickerVolumes as **_long_** because the long data type have a range of whole numbers between -2 billion to 2 billion.
- tickerStarting prices and tickerEnding prices as **_Single_** because single data type is **_32 bits_** so it means it takes less memory than **_Double_** data type which stores value in **_64 bits_**.
             
       Dim tickerVolumes(12) As Long

       Dim tickerStartingPrices(12) As Single

       Dim tickerEndingPrices(12) As Single
       
       
###### Step 2 

a) Created the loops for totalVolumes from 0.

          For i = 0 To 11

          tickerVolumes(i) = 0
          
          Next i

b) Now the loop starts from 2nd row to the last row of the data set.

          For i = 2 To RowCount
       
###### Step 3 

a) Now we add the next value in tickerVolume in the column **H** as the loop is doing its job.

    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

b) Checking if the current row is the first row using the _if conditionals_.

      If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <>              tickers(tickerIndex) Then
    
       tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
    End If

c) Similarly using _if conditionals_ to check if the current row is the last row.


    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <>              tickers(tickerIndex) Then
     
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
     End If

d) Increasing the tickerIndex

    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
            tickerIndex = tickerIndex + 1
            
        End If
        
        
###### Step 4

Using loops to update the Ticker, Total daily Volume and Return.

For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = tickers(i)
    
    Cells(4 + i, 2).Value = tickerVolumes(i)
    
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i


Formatting the text and cells in the totalVolume column
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.FontStyle = "Bold"
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    
    Range("C4:C15").NumberFormat = "0.0%"
    
    Columns("B").AutoFit

    dataRowStart = 4
    
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 Ending the timer using 
 
 
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

     End Sub


![2017 run time - 0.109375 seconds](stock_analysis/VBA_challenge_2017.png)



![2018 run time - 0.101562 seconds](stock_analysis/VBA_challenge_2018.png)

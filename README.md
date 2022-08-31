# Module-2-Challenge-Refactor-VBA-Code-and-Measure-Performance

## Overview of Project

The purpose of this project was to refactor a Microsoft Excel VBA code to gather year 2017 and 2018 stock information and determine whether the stocks are worthy enough for investing. Original code was developed and tested for this requirement. However, the goal of this effort was to modify the original code to make the code more efficient and reduce run time. 

## The Data

The data includes two sheets with stock information of 12 different stocks. The stock information contains a ticker value, the date of the information, the opening price, highest and lowest price, closing and adjusted closing price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock for a particular year.

## Results

### Analysis
Before refactoring the code, the original code that was developed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet was copied to a new file with .VBA extension. Next, the steps were listed out to set the structure for the refactoring. The instructions and the code written in the file are shown below.

    Sub AllStocksAnalysisRefactored()
      Dim startTime As Single
      Dim endTime  As Single

      yearValue = InputBox("What year would you like to run the analysis on?")

      startTime = Timer
    
      'Format the output sheet on All Stocks Analysis worksheet
      Worksheets("All Stocks Analysis").Activate
    
      Range("A1").Value = "All Stocks (" + yearValue + ")"
    
      'Create a header row
      Cells(3, 1).Value = "Ticker"
      Cells(3, 2).Value = "Total Daily Volume"
      Cells(3, 3).Value = "Return"

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
    
      'Activate data worksheet
      Worksheets(yearValue).Activate
    
      'Get the number of rows to loop over
      RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
      '1a) Create a ticker Index
      tickerIndex = 0
    
      '1b) Create three output arrays
      Dim tickerVolumes(12) As Long
      Dim tickerStartingPrices(12) As Single
      Dim tickerEndingPrices(12) As Single
    
      ''2a) Create a for loop to initialize the tickerVolumes to zero.
      For i = 0 To 11
          tickerVolumes(i) = 0
          tickerStartingPrices(i) = 0
          tickerEndingPrices(i) = 0
      Next i
        
      ''2b) Loop over all the rows in the spreadsheet.
      For i = 2 To RowCount
    
          '3a) Increase volume for current ticker
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
          '3b) Check if the current row is the first row with the selected tickerIndex.
          'If  Then
        
        
          'End If
        
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
              tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
          End If
            
          '3c) check if the current row is the last row with the selected ticker
          'If the next rowÃ¢â‚¬â„¢s ticker doesnÃ¢â‚¬â„¢t match, increase the tickerIndex.
          'If  Then
        
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
              tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
          End If

            '3d Increase the tickerIndex.
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
               tickerIndex = tickerIndex + 1
          End If
            
          'End If
        
          
      Next i
    
      '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
      For i = 0 To 11
        
          Worksheets("All Stocks Analysis").Activate
          Cells(4 + i, 1).Value = tickers(i)
          Cells(4 + i, 2).Value = tickerVolumes(i)
          Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
      Next i
    
      'Formatting
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
 
      endTime = Timer
      MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub


The technique used in refactoring the code is simple but extremely efficient. Instead of reading the spreadsheet record by record from top to bottom for each iteration while processing the data in a loop, arrays are defined in the code to store the information in memory and used in the code when required. This reducess the processing substantially time while dealing with a large volume of data. 

For instance, the original code took 0.6 seconds to process 2018 stock datadata and .5 seconds to process 2017 data. Where as after refactoring the code, it took .08 seconds to process 2018 data and .07 seconds to process 2017 data.

### Original Code
![image](https://user-images.githubusercontent.com/31812730/187798085-4242d5c7-bc85-4194-8666-c26ade4601d3.png)

![image](https://user-images.githubusercontent.com/31812730/187798673-99e29782-c5ea-4c1d-bb85-05947fb70c20.png)

### Code After Refactoring 
![image](https://user-images.githubusercontent.com/31812730/187797788-db5ec34e-5adc-479b-a0be-6759811ab53d.png)
![image](https://user-images.githubusercontent.com/31812730/187799043-4b1cdc3e-313e-4907-9a49-82056ed1df1a.png)

## Summary

### Advantages and Disadvantages of Refactoring Code

Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. However, we do not always have the luxury to refactor our code due to disadvantages. These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.

### Advantages and Disadvantages of the Original and Refactored VBA Script

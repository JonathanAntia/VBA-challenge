
'Instructions: Create a script that will loop through all the stocks for one year and output ticker symbol, yearly change, percent change, and stock volume
              'Extract the ticker with the greates percent increase, another with the greatest percent decrease, and the one with the greatest total volume
              'Modify the script to run on multiple WorkSheets at once
        
'1. Add headers to columns I, J, K, L for Ticker, Yearly Change, Percent Change, Total Stock Volume
'2. Dim a LastRow variable to account for the last row of every sheet
'3. Write a For loop that will evaluate column A to see if Cells(row,1)<>Cells(row+1,1).
    '3.1 Initialize the loop at row = 2 and run to LastRow
'4. Record the ticker name everytime the condition of the For loop is met
    '4.1 Add a numTicker counter to move to the next cell in column I as tickers are added to the summary
    '4.2 Add a sumStockVolume counter to record the sum of stock volume per ticker
        '4.2.1 Reset the sumStockVolume counter for every ticker
'5. Write another For loop to calculate the yearly change and percent change per ticker
    '5.1 Initialize the loop at row = 2
    '5.2 Record the opening value for the first ticker as a variable that will later be modified inside the loop
    '5.3 Add a nested loop with conditional formatting to show positive yearly change in green and negative in red
    '5.4 Add a nested loop to account for instances in which the openValue is zero. This will prevent bugs due to division by zero
'6. Add headers Ticker and Value for a new table in column O and P
'7. Define the last row for the summary table, using column I
    '7.1 Use the Application.WorkSheetsFunction.Max(Range) to calculate the max in the Percent Change Column, same for the min, and add them to column P
    '7.2 Find the tickers associated with each value using a for loop and add them to column O
'8. Add code to make the Sub VBAStock() Module run on every WorkSheet at once

Sub ApplyMacroToAllWorkSheets()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call VBAStocks
    Next
    Application.ScreenUpdating = True
End Sub

Sub VBAStocks()

'Insert headers for summary table in columns I, J, K, L
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change ($)"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Variables:
Dim LastRow As Long, row As Long, numTicker As Long
Dim sumStockVolume As Variant
Dim openValue As Double, closeValue As Double
Dim TickerSymbol As String

'calculate the last row
LastRow = Cells(Rows.Count, 1).End(xlUp).row

'Initialize all the counters at zero
numTicker = 0
sumStockVolume = 0

'this loop records all the different ticker symbols, sums the volume of each stock, and adds them to the summary table
For row = 2 To LastRow
    sumStockVolume = sumStockVolume + Cells(row, 7).Value
        If (Cells(row, 1).Value <> Cells(row + 1, 1).Value) Then
        numTicker = numTicker + 1 'increase this counter by one so that the next ticker is added the to next row in the summary table
        TickerSymbol = Cells(row, 1).Value
        Range("I" & numTicker + 1).Value = TickerSymbol
        Range("L" & numTicker + 1).Value = sumStockVolume
        sumStockVolume = 0 'reset the sumStockVolume counter everytime the ticker changes
        End If
Next row

numTicker = 0 'reset the numTicker counter
openValue = Cells(2, 3).Value 'set open value as the opening value for the first ticker in the table

'this loop calculates the yearly change for each ticker symbol
For row = 2 To LastRow
    If (Cells(row, 1).Value <> Cells(row + 1, 1).Value) Then
    numTicker = numTicker + 1
    closeValue = Cells(row, 6).Value
    Range("J" & numTicker + 1).Value = closeValue - openValue 'this calculates the yearly change and adds it to the summary table
    
        If Range("J" & numTicker + 1).Value >= 0 Then 'this adds conditional formatting to each cell in column J
        Range("J" & numTicker + 1).Interior.ColorIndex = 4
        Else
        Range("J" & numTicker + 1).Interior.ColorIndex = 3
        End If
        
        If (openValue = 0) Then
        Range("K" & numTicker + 1).Value = closeValue
        Range("K" & numTicker + 1).NumberFormat = "0.00%"
        openValue = Cells(row + 1, 3)
        Else
        Range("K" & numTicker + 1).Value = ((closeValue - openValue) / openValue) 'this calculates the fraction of change per ticker
        Range("K" & numTicker + 1).NumberFormat = "0.00%"
        openValue = Cells(row + 1, 3)
        End If
    End If
Next row

'insert headers for new table in column O and P
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("N1:N4").Select
Selection.Columns.AutoFit

'define the last row of the summary table
LastRow = Cells(2, 9).End(xlDown).row

'find the greatest percent increase and greatest percent decrease
Range("P2").Value = Application.WorksheetFunction.max(Range(Cells(2, 11), Cells(LastRow, 11)))
Range("P3").Value = Application.WorksheetFunction.Min(Range(Cells(2, 11), Cells(LastRow, 11)))
Range("P2:P3").NumberFormat = ("0.00%")

'find the greatest total volume
Range("P4").Value = Application.WorksheetFunction.max(Range(Cells(2, 12), Cells(LastRow, 12)))
Range("P1:P4").Select
Selection.Columns.AutoFit

'find the tickers associated with the greatest percent increase, decrease, and total volume
For row = 2 To LastRow
    If (Cells(row, 11).Value = Range("P2").Value) Then
    Range("O2").Value = Cells(row, 9).Value
    End If
    If (Cells(row, 11).Value = Range("P3").Value) Then
    Range("O3").Value = Cells(row, 9).Value
    End If
    If (Cells(row, 12).Value = Range("P4").Value) Then
    Range("O4").Value = Cells(row, 9).Value
    End If
Next row

End Sub


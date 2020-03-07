Attribute VB_Name = "Module1"
Sub MultipleYearStockData()
        'Setup all the column headers
        Call SetupHeaders
        'Create a list of Tickers
        Call Ticker
        'Calculate Yearly & Percentage Change
        Call YearlyAndPercentageChange
        'Calculate Total Stock volume
        Call TotalStockVolume
    
End Sub

Sub SetupHeaders()

'Clear Rows and setup Row Header
    Columns("I:Q").Select
    Selection.Clear
'Ticker
    Range("I1").Value = "Ticker"
    Range("I1").Columns.AutoFit
'Yearly Change
    Range("J1").Value = "Yearly Change"
    Range("J1").Columns.AutoFit
'Percent Change
    Range("K1").Value = "Percent Change"
    Range("K1").Columns.AutoFit
    'Setup Cell format to percentage to 2 decimal
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"
    'unselect the column
    Cells(1, 1).Select
'Total Stock Volume
Range("L1").Value = "Total Stock Volume"
Range("L1").Columns.AutoFit

End Sub

Sub Ticker()

'Setup variable
Dim ticker_column, summary_row As Integer
Dim last_row As Long

'Ticker is column 1 on the each sheet
ticker_column = 1
'Row 1 has the column labels so Row 2 is where summary information will start
summary_row = 2
'Figure out how many rows have data in them for FOR_LOOP
last_row = Cells(Rows.Count, 1).End(xlUp).Row

'Since tickers are sorted alphabetically, using For loop to find variance between rows and recording the ticker in Summary column "I"
For i = 2 To last_row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'copy ticker value into Column I, summary row starting at 2
            Range("I" & summary_row) = Cells(i, 1).Value
            'Increment the summary row number to record next ticker
            summary_row = summary_row + 1
        End If
Next i

End Sub

Sub YearlyAndPercentageChange()

'Setup variable
Dim yearly_change_column, summary_row As Integer
Dim opening_price, closing_price As Double
Dim last_row As Long

'Row 1 has the column labels so Row 2 is where summary information will start
summary_row = 2
'Figure out how many rows have data in them for FOR_LOOP
last_row = Cells(Rows.Count, 1).End(xlUp).Row

'Opening and Closing Price are both positive value so to reset, set it to -1.
'Use both variable to grab starting and ending values to do calculation
opening_price = -1
closing_price = -1


'Since tickers are sorted, using For loop to find when value changes and then record opening and closing price
For i = 2 To last_row
        'if the opening_price is reset, pick the first row of the tickers opening price for the day as opening price
        If opening_price = -1 Then
            'grab value from column 3 and set opening price
            opening_price = Cells(i, 3).Value
        End If
        
        'check the next row to see if all ticker rows have been accounted for
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'if ticker changes, record the closing value from column 6 as last closing price
            closing_price = Cells(i, 6).Value
            'set the value of  yearly change in Column J based on closing - opening price calc
            Range("J" & summary_row) = closing_price - opening_price
            
            'manage the scenario where opening price is 0 as it throws exceptions
            If opening_price <> 0 Then
                'Column K has the percentage of change
                Range("K" & summary_row) = (closing_price - opening_price) / opening_price
                
                'if value is less than 0 then mark it as red other wise mark it as green
                If Range("K" & summary_row) < 0 Then
                    Range("K" & summary_row).Interior.Color = RGB(255, 0, 0)
                Else
                    Range("K" & summary_row).Interior.Color = RGB(50, 205, 50)
                End If
            Else
                'if opening price was 0, then change is 100%
                Range("K" & summary_row) = 100
            End If
            'reset opening and closing price
            opening_price = -1
            closing_price = -1
            'increment the summary row
            summary_row = summary_row + 1
        End If
Next i


End Sub

Sub TotalStockVolume()

'setup variables
Dim ticker_column, summary_row As Integer
Dim last_row As Long
Dim total_stock_volume As Double


'Row 1 has the column labels so Row 2 is where summary information will start
summary_row = 2
'Figure out how many rows have data in them for FOR_LOOP
last_row = Cells(Rows.Count, 1).End(xlUp).Row
'stock volume initialized
total_stock_volume = 0

'Since tickers are sorted, using For loop to find variance and record ticker
For i = 2 To last_row
        'keep adding up the stock volume until the ticker changes with value from column 7
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        'when ticker is about to change, record the stock volume into in the cell and reset the stock volume and increment the summary row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Range("L" & summary_row) = total_stock_volume
            summary_row = summary_row + 1
            total_stock_volume = 0
        End If
Next i
End Sub

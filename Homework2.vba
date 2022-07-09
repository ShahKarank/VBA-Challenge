VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_market()
'Create a loop to cycle through the worksheets in the workbook
    'Set a variable to cycle through the worksheets
    Dim ws As Worksheet
    
    'Start loop
    For Each ws In Worksheets
    
    'Create column labels for the summary table
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Keep track of location for each ticker symbol in the summary table
    Dim TickerRow As Long
    TickerRow = 2
    
    'Set variable to hold the ticker symbol
    Dim ticker As String
    
    'Set variable to hold year open price
    Dim open_price As Double
    open_price = 0
    
    'Set variable to hold year close price
    Dim close_price As Double
    close_price = 0
    
    'Set variable to hold the change in price for the year
    Dim price_change As Double
    price_change = 0
    
    'Set variable to hold the percent change in price for the year
    Dim price_change_percent As Double
    price_change_percent = 0
    
    'Set variable to hold total volume of stock traded
    Dim Total_stock_volume As Double
    Total_stock_volume = 0
    
    'Set variable for total rows to loop through
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop to search through ticker symbols
    For i = 2 To lastrow
    
    'Grab open price
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    open_price = ws.Cells(i, 3).Value
    End If
    
    'Determine the total stock volume
    Total_stock_volume = Total_stock_volume + ws.Cells(i, 7)
    
    'Determine when the ticker symbol changes
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    
    'Move ticker to the summary table
    ws.Cells(TickerRow, 9).Value = ws.Cells(i, 1).Value
    
    'Move the total stock volume to the table
    ws.Cells(TickerRow, 12).Value = Total_stock_volume
    
    'Grab Close price
    close_price = ws.Cells(i, 6).Value
    
    'Calculate the change in Price by subtracting close price from open
    price_change = close_price - open_price
    ws.Cells(TickerRow, 10).Value = price_change
    
    'Conditional formatting to highlight positive change as green and negative as red
    If price_change >= 0 Then
    ws.Cells(TickerRow, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(TickerRow, 10).Interior.ColorIndex = 3
    End If
    
    'If open or close price are zero then determine percent change as zero or else determine the percent change and display it in 0.00% format as shown below.
    If open_price = 0 And close_price = 0 Then
    price_change_percent = 0
    ws.Cells(TickerRow, 11).Value = price_change_percent
    ws.Cells(TickerRow, 11).NumberFormat = "0.00%"
    ElseIf open_price = 0 Then
    
    'Sometimes open price starts as 0 and then increases which means its a new addition so we need to specify as such
    Dim new_stock_percent As String
    new_stock_percent = "New Stock"
    ws.Cells(TickerRow, 11).Value = price_change_percent
    Else
    price_change_percent = price_change / open_price
    ws.Cells(TickerRow, 11).Value = price_change_percent
    ws.Cells(TickerRow, 11).NumberFormat = "0.00%"
    End If
    
    'Add 1 to move on to the next row
    TickerRow = TickerRow + 1

    
    'Reset the values once it moves on to the next row
    open_price = 0
    close_price = 0
    price_change = 0
    price_change_percent = 0
    Total_stock_volume = 0
    End If
    Next i
    
    'Creating best and worst performing charts table
    'Titles
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Assign Lastrow as we did previously to count the rows in the table
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Assign Variables to hold the best stock, worst stock and the stock with most volume
    Dim high_stock As String
    Dim high_value As Double
    Dim low_value As Double
    Dim low_stock As String
    Dim best_vol As String
    Dim best_value As Double
    
    'Start off the high value,low value and best vol value as first respectively
    high_value = ws.Cells(2, 11).Value
    low_value = ws.Cells(2, 11).Value
    best_value = ws.Cells(2, 12).Value
    
    'Start the loop that scans through the table
    For j = 2 To lastrow
    
    'Determine the best performer
    If ws.Cells(10, 11).Value > high_value Then
    high_value = ws.Cells(10, 11).Value
    high_stock = ws.Cells(10, 9).Value
    End If
    
    'Determine the worst performer
    If ws.Cells(10, 11).Value < low_value Then
    low_value = ws.Cells(10, 11).Value
    low_stock = ws.Cells(10, 9).Value
    End If
    
    'Determine highest volume stock
    If ws.Cells(10, 12).Value > best_value Then
    best_value = ws.Cells(10, 12).Value
    best_vol = ws.Cells(10, 9).Value
    End If
    
    Next j
    
    'Move the high value, low value and best volume to the table we dedicated
    ws.Range("P2").Value = high_stock
    ws.Range("Q2").Value = high_value
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P3").Value = low_stock
    ws.Range("Q3").Value = low_value
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P4").Value = best_vol
    ws.Range("Q4").Value = best_value
    
Next ws
    

    
    
End Sub



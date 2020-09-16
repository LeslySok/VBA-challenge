Sub StockDataAnalysis()
    'Loop to add headers to all worsheets
        For Each WS In Worksheets
        WS.Activate

    'Setting column names to cells
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Changes"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        Range("O2") = "Greatest % Increase"
        Range("O3") = "Greatest % Decrease"
        Range("O4") = "Greatest Total Volume"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        
    'Formatting column names to Autofit into Cells
        Range("J1:L1").EntireColumn.AutoFit
        Range("O2:O4").EntireColumn.AutoFit
 
    'Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

    'Declaring all Variable
        Dim Ticker_Name As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Volume As Double
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim MaxIncrease As Double
        Dim MaxDecrease As Double
        Dim MaxTotal As Double
        Dim MaxTic As String
        Dim MinTic As String
        Dim MaxValueTic As String
        Dim Row As Double
        Dim Column As Integer
        Dim i As Long

        Volume = 0
        Row = 2
        Column = 1
     
    'Assign Open price
        Open_Price = Cells(2, Column + 2).Value
    
    'Loop through all tickers
        For i = 2 To LastRow
    
    'Find unique ticker and new trigger
        If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
    
    'Assign Ticker name
        Ticker_Name = Cells(i, Column).Value
        Cells(Row, Column + 8).Value = Ticker_Name
    
    'Assign Close Price
        Close_Price = Cells(i, Column + 5).Value

    'Assign Yearly Change
        Yearly_Change = Close_Price - Open_Price
        Cells(Row, Column + 9).Value = Yearly_Change
    
    'Calculate Percent Change
        If (Open_Price = 0 And Close_Price = 0) Then
        Percent_Change = 0
            ElseIf (Open_Price = 0 And Close_Price <> 0) Then
            Percent_Change = 1
                Else
                Percent_Change = Yearly_Change / Open_Price
                Cells(Row, Column + 10).Value = Percent_Change
                Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If

    'Add Total Ticker Volume to Stock Volume
        Volume = Volume + Cells(i, Column + 6).Value
        Cells(Row, Column + 11).Value = Volume
    
    'Adding to summary table
        Row = Row + 1
    
    'Reset Open Price
        Open_Price = Cells(i + 1, Column + 2)
    
    'Reset Volume Total
        Volume = 0
    
    'If duplicate ticker
        Else
            Volume = Volume + Cells(i, Column + 6).Value
            End If
            
            Next i
        
    'Determine the Last Row of Yearly Change per WS
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row

    'Assign Cell Colors to Yearly Change
        For j = 2 To YCLastRow
        If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
        Cells(j, Column + 9).Interior.ColorIndex = 4
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
                End If
                Next j
        
    'Set Greatest % Increase, % Decrease, and Total Volume
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
    
    'Search each row and return tikcer's greatest increase value
        For K = 2 To YCLastRow
            If Cells(K, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, Column + 15).Value = Cells(K, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(K, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(K, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, Column + 15).Value = Cells(K, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(K, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(K, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, Column + 15).Value = Cells(K, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(K, Column + 11).Value
            End If
        Next K
        
    Next WS
        
End Sub




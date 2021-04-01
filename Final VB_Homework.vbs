Attribute VB_Name = "Module1"
Sub GroupstockData()

'Worksheet Loop
For Each ws In Worksheets

'Set an Initial variable for the brand name
Dim Ticker As String
'Set an initial vaiable for holding the total change per ticker
Dim Ticker_total As Double
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Stock_Total As Double
Ticker_Change = 0
Stock_Total = 0

'Keep the track of the total change
Summary_Table_Row = 2

LastR = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'Main Header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "% Change"
    ws.Cells(1, 12).Value = "Stock Total Volume"

For i = 2 To LastR

    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
       
       Opening_Price = ws.Cells(i, 3).Value
       Ticker = ws.Cells(i, 1).Value
       Stock_Total = Stock_Total + ws.Cells(i, 7).Value
    
            ElseIf (ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value) And (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
            Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            ElseIf (ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value) And (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                    Closing_Price = ws.Cells(i, 6).Value
                    Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                    Yearly_Change = Closing_Price - Opening_Price
                    
                                If ws.Cells(Summary_Table_Row, 10).Value < 0 Then
                                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                                        Else
                                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                                        
        
        ws.Cells(Summary_Table_Row, 9).Value = Ticker
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        If Opening_Price > 0 Then
        ws.Range("K" & Summary_Table_Row).Value = (Closing_Price - Opening_Price) / Opening_Price
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        Else
        ws.Range("K" & Summary_Table_Row).Value = 0
        End If
             
        ws.Range("L" & Summary_Table_Row).Value = Stock_Total
        Summary_Table_Row = Summary_Table_Row + 1
    
        Stock_Total = 0
        
            End If

End If

Next i

'Label
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

Dim j As Long
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total_Volume As Double
Dim Grt_Inc, Grt_Dec, Grt_Vol As String

LastR2 = ws.Cells(Rows.Count, "J").End(xlUp).Row

For j = 2 To LastR2

    If ws.Cells(j, 11).Value > Greatest_Increase Then
        Greatest_Increase = ws.Cells(j, 11).Value
        Grt_Inc = ws.Cells(j, 9).Value
        If ws.Cells(j, 12).Value > Greatest_Total_Volume Then
        Greatest_Total_Volume = ws.Cells(j, 12).Value
        Grt_Vol = ws.Cells(j, 9).Value
        End If
        
        ElseIf ws.Cells(j, 11).Value < Greatest_Decrease Then
        Greatest_Decrease = ws.Cells(j, 11).Value
        Grt_Dec = ws.Cells(j, 9).Value
        If ws.Cells(j, 12).Value > Greatest_Total_Volume Then
        Greatest_Total_Volume = ws.Cells(j, 12).Value
        Grt_Vol = ws.Cells(j, 9).Value
        End If
        
        ElseIf ws.Cells(j, 12).Value > Greatest_Total_Volume Then
        Greatest_Total_Volume = ws.Cells(j, 12).Value
        Grt_Vol = ws.Cells(j, 9).Value
        
        
        End If
        Next j
        
ws.Cells(2, 16).Value = Greatest_Increase
ws.Cells(3, 16).Value = Greatest_Decrease
ws.Cells(4, 16).Value = Greatest_Total_Volume

ws.Cells(2, 15).Value = Grt_Inc
ws.Cells(3, 15).Value = Grt_Dec
ws.Cells(4, 15).Value = Grt_Vol

ws.Range("P2:P3").NumberFormat = "0.00%"

Next ws

End Sub

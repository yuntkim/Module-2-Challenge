Sub StockData()

For Each ws In Worksheets

Dim Stock_Name As String
Dim Yearly_Change As Double
Dim Year_Open As Double
Dim Year_Close As Double
Dim Summary_Row As Integer
Dim TotalVolume As Double
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total_Volume As Double


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
TotalVolume = 0
i = 1 
Year_Open = ws.Cells(i + 1, 3).Value
Year_Close = 0
PercentChange = 0
Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_Total_Volume = 0


ws.Range("J1") = "Ticker"
ws.Range("K1") = "Yearly Change"
ws.Range("L1") = "Percent Change"
ws.Range("M1") = "Total Stock Volume"
ws.Range("J1:M1").Interior.ColorIndex = 36
ws.Range("J1:M1").Font.Bold = True

ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("O2:O4").Interior.ColorIndex = 35
ws.Range("O2:O4").Font.Bold = True
  
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"
ws.Range("P1:Q1").Interior.ColorIndex = 36
ws.Range("P1:Q1").Font.Bold = True
  
Summary_Row = 2
For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Stock_Name = ws.Cells(i, 1).Value
    
    Stock_Ticker = ws.Cells(i, 1).Value
    
    Year_Close = ws.Cells(i, 6).Value

    Yearly_Change = Year_Close - Year_Open

    PercentChange = Round((Year_Close - Year_Open) / Year_Open * 100, 2) & "%"

    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    

    ws.Range("J" & Summary_Row).Value = Stock_Name
    ws.Range("K" & Summary_Row).Value = Yearly_Change
    ws.Range("L" & Summary_Row).Value = PercentChange
    ws.Range("M" & Summary_Row).Value = TotalVolume
    Summary_Row = Summary_Row + 1


    Year_Open = ws.Cells(i + 1, 3).Value
    Stock_Total = 0
    TotalVolume = 0
    Year_Close = 0
    PercentChange = 0
    
  Else

    TotalVolume = TotalVolume + ws.Cells(i, 7).Value

  End If

  Next i
  
    For i = 2 To LastRow
        If ws.Cells(i, 11).Value > 0 Then
            Yearly_Change = ws.Cells(i, 11).Value
            ws.Cells(i, 11).Interior.ColorIndex = 4
            
        ElseIf ws.Cells(i, 11).Value < 0 Then
            Yearly_Change = ws.Cells(i, 11).Value
            ws.Cells(i, 11).Interior.ColorIndex = 3
        
  End If
    
  Next

    For i = 2 To LastRow
        If (ws.Cells(i, 12).Value) > Greatest_Increase Then
                ws.Cells(2, 17) = ws.Cells(i, 12).Value * 100 & "%"
                Greatest_Increase = ws.Cells(i, 12).Value
                ws.Cells(2, 16) = ws.Cells(i, 10).Value
                
        ElseIf (ws.Cells(i, 12).Value) < Greatest_Decrease Then
                ws.Cells(3, 17) = ws.Cells(i, 12).Value * 100 & "%"
                Greatest_Decrease = ws.Cells(i, 12).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 10).Value

        ElseIf (ws.Cells(i, 13).Value) > Greatest_Total_Volume Then
                ws.Cells(4, 17) = ws.Cells(i, 13).Value
                Greatest_Total_Volume = ws.Cells(i, 13).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
                
        Else
        End If
        
        Next i
        Next ws
    
        
End Sub

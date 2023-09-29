# vba_challenge
Sub VBA_Challenge()

    For Each ws In Worksheets

    Dim Ticker As String
    Dim Date2 As String
    Dim Open2 As String
    Dim Close2 As String
    
    Dim Ticker2 As Double
    Ticker2 = 0
    
    Dim TotalStockVolume As Double
    TotalStockVolume = 2
    
    Dim YearlyChange As Double
    YearlyChange = 2
    
    Dim PercentChange As Double
    PercentChange = 2
    
    Dim Value As Double
    Value = 2
    
    Dim Start As Double
    Start = 2
    
    Dim WorksheetName As String
    
    
    For i = 2 To 753001
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker = ws.Cells(i, 1).Value
    
    Ticker2 = Ticker2 + ws.Cells(i, 7).Value
    
    ws.Range("J" & TotalStockVolume).Value = Ticker
    
    ws.Range("M" & TotalStockVolume).Value = Ticker2
    
    TotalStockVolume = TotalStockVolume + 1
    
    Ticker2 = 0
    
    YearlyChange = ws.Cells(i, 6) - ws.Cells(Start, 3)
    
    ws.Range("K" & TotalStockVolume).Value = YearlyChange
    

    Else
    
    Ticker2 = Ticker2 + ws.Cells(i, 7).Value
    
    End If
    
    If YearlyChange < 0 Then
    ws.Cells(TotalStockVolume, 11).Interior.ColorIndex = 3
    
    ElseIf YearlyChange > 0 Then
    ws.Cells(TotalStockVolume, 11).Interior.ColorIndex = 4
    
    End If
    
    PercentChange = YearlyChange / ws.Cells(Start, 3)
    
    ws.Range("L" & TotalStockVolume).Value = PercentChange

    
    Next i
    
    Next ws
    
    
End Sub

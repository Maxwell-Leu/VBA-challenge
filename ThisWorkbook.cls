VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub StockData()
    Dim i As Long
    Dim x As Integer
    Dim ws As Worksheet
    Dim currentTicker As String
    Dim uniqueTickers As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim startRange As Long
    Dim endRange As Long
    Dim dataRow As Long
    

    For Each ws In Worksheets

        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("j1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        currentTicker = ws.Cells(2, 1).Value
        startRange = 2
        dataRow = 2
        
        For i = 2 To LastRow + 1
        
            If currentTicker = ws.Cells(i, 1).Value Then
                
            Else
            
                ws.Cells(dataRow, 9).Value = currentTicker
                currentTicker = ws.Cells(i, 1).Value
                uniqueTickers = uniqueTickers + 1
                
                endRange = i - 1
                openPrice = ws.Cells(startRange, 3).Value
                closePrice = ws.Cells(endRange, 6).Value
                
                quarterlyChange = closePrice - openPrice
                ws.Cells(dataRow, 10).Value = quarterlyChange
                If quarterlyChange < 0 Then
                
                    ws.Cells(dataRow, 10).Interior.ColorIndex = 3
                
                ElseIf quarterlyChange > 0 Then
                
                    ws.Cells(dataRow, 10).Interior.ColorIndex = 4
                
                Else
                
                    ws.Cells(dataRow, 10).Interior.ColorIndex = 0
                
                End If
                
                percentChange = quarterlyChange / openPrice
                ws.Cells(dataRow, 11).Value = percentChange
                ws.Cells(dataRow, 11).NumberFormat = "0.00%"
                
                ws.Cells(dataRow, 12).Value = WorksheetFunction.Sum(ws.Range("G" & startRange, "G" & endRange))
                startRange = i
                dataRow = dataRow + 1
                
            End If
        
        Next i
        Dim greatestIncrease As Double
        greatestIncrease = 0
        Dim increaseTicker As String
        Dim greatestDecrease As Double
        greatestDecrease = 0
        Dim decreaseTicker As String
        Dim greatestVolume As String
        Dim volumeTicker As String
        
        
        ws.Range("n2").Value = "Greatest % Increase"
        ws.Range("n3").Value = "Greatest % Decrease"
        ws.Range("n4").Value = "Greatest Total Volume"
        ws.Range("o1").Value = "Ticker"
        ws.Range("p1").Value = "Value"
        
        Dim largestFound As String
        largestFound = 1
        
        For x = 2 To dataRow - 1
            
            Dim comparedPercent As Double
            comparedPercent = ws.Cells(x, 11).Value
            If 0 < comparedPercent Then
                
                If comparedPercent > greatestIncrease Then
                    greatestIncrease = comparedPercent
                    increaseTicker = ws.Cells(x, 9).Value
                End If
                
            Else
                If comparedPercent < greatestDecrease Then
                    greatestDecrease = comparedPercent
                    decreaseTicker = ws.Cells(x, 9).Value
                End If
            End If
            
            If ws.Range("p4").Value < ws.Cells(x, 12).Value Then
                ws.Range("p4").Value = ws.Cells(x, 12).Value
                volumeTicker = ws.Cells(x, 9).Value
                
            End If
            
        Next x
        
        
        'greatestVolume = WorksheetFunction.Max(ws.Range("L2", "L1501"))
        
        
        ws.Range("o2").Value = increaseTicker
        ws.Range("p2").Value = greatestIncrease
        ws.Range("p2", "P3").NumberFormat = "0.00%"
        ws.Range("o3").Value = decreaseTicker
        ws.Range("p3").Value = greatestDecrease
        ws.Range("o4").Value = volumeTicker
        'ws.Range("p4").Value = greatestVolume
    Next ws



End Sub




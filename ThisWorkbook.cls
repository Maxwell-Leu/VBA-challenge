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
    
    'Variable for data that gets what is requested
    Dim dataRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim startRange As Long
    Dim endRange As Long
    
    'Data asked for
    Dim currentTicker As String
    Dim quarterlyChange As Double
    Dim percentChange As Double
    
    'Looping for each sheet in workbook
    For Each ws In Worksheets
        
        'Gets last row
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("j1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Preload data for logic to work
        currentTicker = ws.Cells(2, 1).Value
        startRange = 2
        dataRow = 2
        
        'Iterating on data
        For i = 2 To LastRow + 1
            'Checking if the next ticker is the same as the past one
            If currentTicker = ws.Cells(i, 1).Value Then
                'Could have done a while loop looking back
            Else
                'Populating the ticker row once we ran out of the current one
                ws.Cells(dataRow, 9).Value = currentTicker
                currentTicker = ws.Cells(i, 1).Value
                uniqueTickers = uniqueTickers + 1
                
                'Setting the range for math to avoid overflow
                endRange = i - 1
                openPrice = ws.Cells(startRange, 3).Value
                closePrice = ws.Cells(endRange, 6).Value
                
                'Setting value and fill based on it
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
                
                'Using excel to add the numbers together to avoid overflow
                ws.Cells(dataRow, 12).Value = WorksheetFunction.Sum(ws.Range("G" & startRange, "G" & endRange))
                
                'Move range forward to where the new ticker is
                startRange = i
                
                'move data onto the next row to prevent overwriting
                dataRow = dataRow + 1
                
            End If
        
        Next i
        
        'Variables to store the 6 needed values and logic variables
        Dim greatestIncrease As Double
        greatestIncrease = 0
        Dim increaseTicker As String
        Dim greatestDecrease As Double
        greatestDecrease = 0
        Dim decreaseTicker As String
        Dim greatestVolume As String
        Dim volumeTicker As String
        
        'Info
        ws.Range("n2").Value = "Greatest % Increase"
        ws.Range("n3").Value = "Greatest % Decrease"
        ws.Range("n4").Value = "Greatest Total Volume"
        ws.Range("o1").Value = "Ticker"
        ws.Range("p1").Value = "Value"
        
        
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
                
                'compare cell values because of overflow
            If ws.Range("p4").Value < ws.Cells(x, 12).Value Then
                ws.Range("p4").Value = ws.Cells(x, 12).Value
                volumeTicker = ws.Cells(x, 9).Value
                
            End If
            
        Next x
        
        'tried as a better way to find the max but could figure out how to get the row for the ticker
        'greatestVolume = WorksheetFunction.Max(ws.Range("L2", "L1501"))
        
        'Populating data and formating
        ws.Range("o2").Value = increaseTicker
        ws.Range("p2").Value = greatestIncrease
        ws.Range("p2", "P3").NumberFormat = "0.00%"
        ws.Range("o3").Value = decreaseTicker
        ws.Range("p3").Value = greatestDecrease
        ws.Range("o4").Value = volumeTicker
        'ws.Range("p4").Value = greatestVolume
    Next ws

End Sub





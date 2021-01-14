Attribute VB_Name = "Module2"
Sub bonus():

    Dim lastrow As Long
    Dim maxValue As Double
    Dim minValue As Double
    Dim greatestTotalVolume As Double
    Dim i As Long
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim maxTickerName As String
    Dim minTickerName As String
    Dim volTickerName As String
    
    For Each ws In Worksheets
    
        currentSheet = ws.Name
        Debug.Print ("Working on sheet " & currentSheet)
        
        'find last row of summary table
        lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
                
        'set initial values for max, min, and greatest total
        maxValue = 0
        minValue = 0
        greatestTotalVolume = 0
        
        'iterate through stock summary table to find max and min percent change values
        For i = 2 To lastrow
            
            'access the current cell percentage change
            percentChange = ws.Cells(i, 11).Value
            
            'find the max value by comparing percentChange to stored max value
            If percentChange > maxValue Then
                
                'set maxValue to percentChange if true
                maxValue = percentChange
                
                'record ticker name of max value
                maxTickerName = ws.Cells(i, 9).Value
                
            End If
            
            'find the min value by comparing percentChange to stored min value
            If percentChange < minValue Then
                
                'set minValue to percentChange if true
                minValue = percentChange
                
                'record ticker name of min value
                minTickerName = ws.Cells(i, 9).Value
                
            End If
                   
        Next i
        
        'iterate through stock summary table to find greatest total volume value
        For i = 2 To lastrow
        
            'access the current cell total volume
            totalVolume = ws.Cells(i, 12).Value
            
            If totalVolume > greatestTotalVolume Then
            
                'set greatestTotalVolume to totalVolume if true
                greatestTotalVolume = totalVolume
                
                'record ticker name of greatest total volume
                volTickerName = ws.Cells(i, 9).Value
                
            End If
            
        Next i
        
        'create summary table
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        'write values to summary table
        ws.Cells(2, 15).Value = maxTickerName
        ws.Cells(2, 16).Value = Format(maxValue, "0.00%")
        ws.Cells(3, 15).Value = minTickerName
        ws.Cells(3, 16).Value = Format(minValue, "0.00%")
        ws.Cells(4, 15).Value = volTickerName
        ws.Cells(4, 16).Value = greatestTotalVolume
        
        ws.Columns("I:P").AutoFit
        
    Next ws
    
End Sub

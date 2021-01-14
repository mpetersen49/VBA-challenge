Attribute VB_Name = "Module1"
Sub stockSummary():

    Dim i As Long
    Dim lastrow As Long
    Dim nextTicker As String
    Dim currentTicker As String
    Dim summaryRow As Long
    Dim volumeTotal As Double
    Dim counter As Long
    Dim openValue As Double
    Dim closeValue As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    For Each ws In Worksheets
        
        currentSheet = ws.Name
        Debug.Print ("Working on sheet " & currentSheet)
        
        'find last row of dataset
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'write summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
    
        'begin for loop to iterate through stock data
        summaryRow = 2
        counter = 0
        volumeTotal = 0
        
        For i = 2 To lastrow
        
            'check if next ticker value is equal to current ticker value
            nextTicker = ws.Cells(i + 1, 1).Value
            currentTicker = ws.Cells(i, 1).Value
            
            If nextTicker <> currentTicker Then
                
                'calculate summary values
                openValue = ws.Cells(i - counter, 3).Value 'use counter variable to find 1st row of dataset
                closeValue = ws.Cells(i, 6).Value
                yearlyChange = closeValue - openValue
                
                'test if openValue is 0 to avoid div/0 error when calculating percentChange
                If openValue = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openValue
                End If
                
                volumeTotal = volumeTotal + Cells(i, 7).Value
                
                'write summary values to cells
                ws.Range("I" & summaryRow).Value = currentTicker
                ws.Range("J" & summaryRow).Value = yearlyChange
                ws.Range("K" & summaryRow).Value = Format(percentChange, "0.00%")
                ws.Range("L" & summaryRow).Value = volumeTotal
                
                'check if yearlyChange is positive or negative and format cell color
                If yearlyChange >= 0 Then
                    
                    ws.Range("J" & summaryRow).Interior.ColorIndex = 4 'green
                
                Else
                    
                    ws.Range("J" & summaryRow).Interior.ColorIndex = 3 'red
                
                End If
                
                'reset summaryRow, counter and volumeTotal
                summaryRow = summaryRow + 1
                counter = 0
                volumeTotal = 0
                
            Else
                'count number of cells with same ticker name
                counter = counter + 1
                
                'sum volume with same ticker name
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
                
            End If
                
        Next i
        
    Next ws
    
End Sub

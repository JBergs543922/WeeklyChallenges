Sub StockEvaluation()
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    'variable declaration
    
    'original table
    Dim ticker As String
    Dim openP As Double
    Dim closeP As Double
    Dim vol As Long
    'new table
    Dim currentEntry As String
    Dim currentEntryV As Double
    Dim stockNum As Long
    'new table2
    Dim greatestPerIn As Double
    Dim greatestPerDec As Double
    Dim greatestVol As Double
    Dim curEntry As Double
    Dim prevEntry As Double
    Dim curEntryV As Double
    Dim prevEntryV As Double
    'variable initialization
    ticker = " "
    vol = 0
    openP = 0
    closeP = 0
    currentEntry = " "
    currentEntryV = 0
    foundStock = 0
    stockNum = 0
    
    'setup table1
    ws.Range("H1:H10000").Interior.ColorIndex = 15
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Columns("I:M").EntireColumn.AutoFit
        
        
        'conditional logic
        'itterate through the original table
        For i = 2 To 800000 'anything over 100k results in a crash
            currentEntry = ws.Cells(i, 1).Value
            currentEntryV = ws.Cells(i, 7).Value
            
            'yearly change calculation - save first opening price
            If ws.Cells(i - 1, 1).Value <> currentEntry Then
                openP = ws.Cells(i, 3).Value
            End If
            
            'yearly change calculation - save final closing price
            If ws.Cells(i + 1, 1).Value <> currentEntry Then
                closeP = ws.Cells(i, 6).Value
                ws.Cells(stockNum + 1, 10).Value = closeP - openP
                'color grade
                If ws.Cells(stockNum + 1, 10).Value > 0 Then
                ws.Cells(stockNum + 1, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(stockNum + 1, 10).Interior.ColorIndex = 3
                End If
                'percent calculation and format
                ws.Cells(stockNum + 1, 11).Value = ((closeP - openP) / closeP)
                ws.Cells(stockNum + 1, 11).NumberFormat = "0.00%"
            End If
            'logic control
            foundStock = 0
            
            'check to see if any entry matches
            For j = 1 To stockNum

                'we found a stock that matches, in our save list
                If currentEntry = ws.Cells(j + 1, 9).Value Then
                    
                    'add vol
                    ws.Cells(j + 1, 12).Value = ws.Cells(j + 1, 12).Value + currentEntryV
            
                    'logic control
                    foundStock = 1
                End If
                
            Next j

            'we did not find a stock that matches, in our save list
            If (foundStock = 0) Then
    
            'ticker
               ws.Cells(stockNum + 2, 9).Value = currentEntry
                
            'yearly calculation
                
            
            'vol
                ws.Cells(stockNum + 2, 12).Value = ws.Cells(stockNum + 2, 12).Value + currentEntryV
                
                stockNum = stockNum + 1
             End If
            
            'logic control
            itNum = itNum + 1
        Next i
        
        'set up table2
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Columns("N:Q").EntireColumn.AutoFit
        
        'Finding greatest % incerease, % decrease, and greatest volume
        'greatest % increase
        greatestPerIn = -1000
        For i = 2 To stockNum
            If ws.Cells(i, 11).Value > greatestPerIn Then
            greatestPerIn = ws.Cells(i, 11).Value
            ticker = ws.Cells(i, 9).Value
            End If
        Next i
        'output found value
        ws.Cells(2, 16).Value = ticker
        ws.Cells(2, 17).Value = greatestPerIn
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        'greatest % decrease
        greatestPerDec = 1000
        For i = 2 To stockNum
            If ws.Cells(i, 11).Value < greatestPerDec Then
            greatestPerDec = ws.Cells(i, 11).Value
            ticker = ws.Cells(i, 9).Value
            End If
        Next i
        'output found value
        ws.Cells(3, 16).Value = ticker
        ws.Cells(3, 17).Value = greatestPerDec
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        'greatestVol
        greatestVol = 0
        For i = 2 To stockNum
            If ws.Cells(i, 12).Value > greatestVol Then
            greatestVol = ws.Cells(i, 12).Value
            ticker = ws.Cells(i, 9).Value
            End If
        Next i
        'output found value
        ws.Cells(4, 16).Value = ticker
        ws.Cells(4, 17).Value = greatestVol
        
        Next ws

End Sub
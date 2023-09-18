Attribute VB_Name = "Module1"
Sub tickerLoop()

    'go through each worksheet and end with Next ws
    For Each ws In Worksheets
    
    'now i gotta determine all the variables
    
        'the variable that holds the ticker name within the column
        Dim tickerName As String
        
        'now set the volume of the tickers starting at 0
        Dim tickerVolume As Double
        tickerVolume = 0
        
        'make ticker counter and start it on the second row/looks at this column
        Dim tickCount As Long
        tickCount = 2
        
        'need the ending of the table marked
        Dim lastRow As Long
        
        
        'placing all the headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'set initial price at row2col3
        Dim openPrice As Double
        openPrice = ws.Cells(2, 3).Value
        
        'set rest
        Dim closePrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double

        'counting the rows in first column
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'now to loop trough the rows looking at ticker names
        'basic i for index
        For i = 2 To lastRow
        'BIG NOTE, THIS IS ALL OF ENDING THE LOOP AND WHERE TO BRING THE INFO
        'IF next value ISN'T EQUAL TO current value THEN get info and print
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'consolidate the name, volume and prices
                tickerName = ws.Cells(i, 1).Value
            
                'add all the volumes
                tickerVolume = tickerVolume + ws.Cells(i, 7).Value
            
                'put the info into summary table, name, volume, prices
                ws.Range("i" & tickCount).Value = tickerName
                ws.Range("l" & tickCount).Value = tickerVolume
                'closePrice at index col6
                closePrice = ws.Cells(i, 6).Value
                'yearly change is closePrice - openPrice
                yearlyChange = (closePrice - openPrice)
                'now place the year change
                ws.Range("j" & tickCount).Value = yearlyChange
                
                'be sure to check if opening is 0 then percent has to be 0
                If openPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openPrice
                End If
                
            'now place the percentChange into the K kolumn and change it to % format
              ws.Range("k" & tickCount).Value = percentChange
              ws.Range("k" & tickCount).NumberFormat = "0.00%"
            
            'reset the tickCount then add 1 to make it 1 further where we just indexed
              tickCount = tickCount + 1
            
            'volume resets as well
              tickerVolume = 0
            
            'resetting opening price (i+1) = going one further than were we initially started
              openPrice = ws.Cells(i + 1, 3).Value
            
            Else
                'WE GO ON AS NORMAL JUST ADDING VOLUME SINCE THE NEW AND OLD VALUE ARE =
                tickerVolume = tickerVolume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
    
    
    'time for coloring of the end results column
    lastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To lastRowSummary
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
        'tutor mentioned using min max functions from excel make this easy
        For i = 2 To lastRowSummary
           'finding max percent
           If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("k2:k" & lastRowSummary)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            'min percent
           ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRowSummary)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            'now max volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("l2:l" & lastRowSummary)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                
            End If
            
        Next i
     

    Next ws

End Sub


Sub stonks():

    ' variables
    Dim openprice As Double
    Dim closeprice As Double
    Dim pricechange As Double
    Dim volume As LongLong
    volume = 0
    
    'check to see if its the first row of that ticker symbol
    Dim firstTickerRow As Boolean
    firstTickerRow = True
    
    'for each ws loop
    For j = 1 To ActiveWorkbook.Worksheets.Count
    
    	'keep track of what row summary is on
    	Dim summaryRow As Integer
    	summaryRow = 2
    
        'add summary table headers
        Sheets(j).Cells(1, 9).Value = "Ticker"
        Sheets(j).Cells(1, 10).Value = "Yearly Change"
        Sheets(j).Cells(1, 11).Value = "Percent Change"
        Sheets(j).Cells(1, 12).Value = "Total Stock Volume"
    
        'find last row of ws
       lastRow = Sheets(j).Cells(Rows.Count, 1).End(xlUp).Row
       
       'loop through data
       For i = 2 To lastRow
       
            'grab initial open price if its a new ticker
            If firstTickerRow = True Then
                openprice = Sheets(j).Cells(i, 3).Value
                firstTickerRow = False
            End If
            
            'check to see if next row is a differnt ticker sybmol
            If Sheets(j).Cells(i + 1, 1).Value <> Sheets(j).Cells(i, 1).Value Then
                
                'add ticker symbol to summary table
                Sheets(j).Cells(summaryRow, 9).Value = Sheets(j).Cells(i, 1).Value
                
                'add final volume to total
                volume = volume + Sheets(j).Cells(i, 7).Value
                
                'add final volume of ticker to table
                Sheets(j).Cells(summaryRow, 12).Value = volume
                
                'grab closing price
                closeprice = Sheets(j).Cells(i, 6).Value
                
                'calculate yearly change for ticker and add to table
                pricechange = closeprice - openprice
                Sheets(j).Cells(summaryRow, 10).Value = pricechange
                
                'conditional format yearly change positive change as green cell and negative change as red cell
                If pricechange > 0 Then
                    Sheets(j).Cells(summaryRow, 10).Interior.ColorIndex = 4
                
                ElseIf pricechange < 0 Then
                    Sheets(j).Cells(summaryRow, 10).Interior.ColorIndex = 3
                
                End If
                
                'calculate percent change and add to table
                If openprice <> 0 Then
                    Sheets(j).Cells(summaryRow, 11).Value = pricechange / openprice
                
                Else
                    Sheets(j).Cells(summaryRow, 11).Value = 0
                    
                End If
                
                'format percent change
                Sheets(j).Cells(summaryRow, 11).Style = "percent"
                
                'moving onto new ticker, clear out data
                summaryRow = summaryRow + 1
                firstTickerRow = True
                volume = 0
            
            'if it's the same
            Else
                
                'add vol to total stock volume
                volume = volume + Sheets(j).Cells(i, 7).Value
            
            End If
            
       
       Next i
        
    
    
    Next j

End Sub


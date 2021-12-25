Sub Ticker_final()

    Dim biggestInc As Double
    Dim biggestDec As Double
    Dim biggestTot As LongLong
    Dim biggestIncTicker As String
    Dim biggestDecTicker As String
    Dim biggestTotTicker As String
    
    biggestInc = 0
    biggestDec = 0
    biggestTot = 0
    
For Each ws In Worksheets

    'Declare a new variable Year so can do the macro at annual level. Don't know how many years are in each file.
    Dim Year As String
    
    'Declare ticker and year counters as Integer to support ticker and year array creation
    Dim tickerCnt As Integer
    Dim yearCnt As Long
    Dim symbolCnt As Long
    
    'Declare perChg as Double to display the percent change from opening price at the beginning of a given year to the closing price at the end of that year
    Dim perChg As Double
    Dim totVol As LongLong
                
    'Initialize symbolCnt, tickerCnt and yearCnt to populate beginning in 2nd row to begin populating array
    tickerCnt = 2
    yearCnt = 2
    symbolCnt = 0
    totVol = 0
        
    'Declare iterationCnt as new variable based on row count function. Use Long variable type to account for large number of records
    Dim itrCnt As Long
    
    itrCnt = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

    
    'Test accurate iteration cnt derived from iterationCtn
    ws.Cells(1, 10).Value = "Iteration Count"
    ws.Cells(2, 10).Value = itrCnt
    
    ws.Cells(1, 14).Value = "Table Count"

    
  'Populate header row for new data elements and new table with consolidated view sought
    ws.Cells(1, 8).Value = "Year"
    ws.Cells(1, 9).Value = "Symbol Count"
    ws.Cells(1, 12).Value = "Unique Year"
    ws.Cells(1, 16).Value = "Unique Ticker"
    ws.Cells(1, 17).Value = "Year Open Price"
    ws.Cells(1, 18).Value = "Year Close Price"
    ws.Cells(1, 19).Value = "Percent Change"
    ws.Cells(1, 20).Value = "Total Volume"
    ws.Cells(1, 21).Value = "Yearch Change"
    ws.Cells(1, 24).Value = "Ticker"
    ws.Cells(1, 25).Value = "Value"
    ws.Cells(2, 23).Value = "Greatest % Increase"
    ws.Cells(3, 23).Value = "Greatest % Decrease"
    ws.Cells(4, 23).Value = "Greatest Total Volume"
    
    For i = 2 To itrCnt
        'Populate year value into 8th
        Year = Left(ws.Cells(i, 2).Value, 4)
        ws.Cells(i, 8).Value = Year
        
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            symbolCnt = symbolCnt + 1
            ws.Cells(i, 9).Value = symbolCnt
                       
        Else
            symbolCnt = symbolCnt + 1
            ws.Cells(i, 9).Value = symbolCnt
            symbolCnt = 0
            
        End If
        
        
        'Create list of unique ticker symbols
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(tickerCnt, 16).Value = ws.Cells(i, 1).Value
            ws.Cells(tickerCnt, 18).Value = ws.Cells(i, 6).Value
            totVol = totVol + ws.Cells(i, 7).Value
            ws.Cells(tickerCnt, 20).Value = totVol

            'to have the ticker counter move to the next row add 1
            tickerCnt = tickerCnt + 1
            totVol = 0
            
            
            
     
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And symbolCnt = 1 Then
            ws.Cells(tickerCnt, 17).Value = ws.Cells(i, 3).Value
            totVol = totVol + ws.Cells(i, 7).Value
           
        Else
            totVol = totVol + ws.Cells(i, 7).Value
        End If
        
    Next i
    
    For i = 2 To itrCnt
        'Create list of array of unique years
        If ws.Cells(i + 1, 8).Value <> ws.Cells(i, 8).Value Then
            ws.Cells(yearCnt, 12).Value = ws.Cells(i, 8).Value
         
            'to have the year counter move to the next row add 1
            yearCnt = yearCnt + 1
        
        End If
    Next i
    
    
    Dim tblCnt As Long
    Dim greatestInc As Double
    Dim greatestDec As Double
    Dim greatestTot As LongLong
    Dim greatestIncTicker As String
    Dim greatestDecTicker As String
    Dim greatestTotTicker As String
    Dim yearly_chg As Double
    
    greatestInc = 0
    greatestDec = 0
    greatestTot = 0
    
    tblCnt = ws.Cells(Rows.Count, 16).End(xlUp).Row
    ws.Cells(2, 14).Value = tblCnt
    
    For k = 2 To tblCnt
    
          If ws.Cells(k, 17) <> 0 Then
           perChg = ((ws.Cells(k, 18).Value / ws.Cells(k, 17).Value) - 1)
           yearly_change = ws.Cells(k, 18).Value - ws.Cells(k, 17).Value
           ws.Cells(k, 19).Value = perChg
           ws.Cells(k, 21).Value = yearly_change
           
          End If
            
           'format cell color based on whether positive or negative change yearend close Vs year start open
           If ws.Cells(k, 19).Value > 0 Then
                ws.Cells(k, 19).Interior.ColorIndex = 4
                ws.Cells(k, 19).NumberFormat = "0.00%"
                ws.Cells(k, 20).NumberFormat = "#,##0"
                
           ElseIf ws.Cells(k, 19).Value < 0 Then
                ws.Cells(k, 19).Interior.ColorIndex = 3
                ws.Cells(k, 19).NumberFormat = "0.00%"
                ws.Cells(k, 20).NumberFormat = "#,##0"
        
           End If
           
           'format cell color based on whether positive or negative change yearend close Vs year start open
           If ws.Cells(k, 21).Value > 0 Then
                ws.Cells(k, 21).Interior.ColorIndex = 4
                ws.Cells(k, 21).NumberFormat = "0.00%"
                ws.Cells(k, 21).NumberFormat = "#,##0.00"
                
           ElseIf ws.Cells(k, 21).Value < 0 Then
                ws.Cells(k, 21).Interior.ColorIndex = 3
                ws.Cells(k, 21).NumberFormat = "0.00%"
                ws.Cells(k, 21).NumberFormat = "#,##0.00"
        
           End If
           
        'Greatest percent increase
            If ws.Cells(k, 19).Value > greatestInc Then
                greatestInc = ws.Cells(k, 19).Value
                ws.Cells(2, 24).Value = ws.Cells(k, 16).Value
                ws.Cells(2, 25).Value = ws.Cells(k, 19).Value
                greatestIncTicker = ws.Cells(k, 16).Value

            End If
            
         'Greatest percent decrease
            If ws.Cells(k, 19).Value < greatestDec Then
                greatestDec = ws.Cells(k, 19).Value
                ws.Cells(3, 24).Value = ws.Cells(k, 16).Value
                ws.Cells(3, 25).Value = ws.Cells(k, 19).Value
                greatestDecTicker = ws.Cells(k, 16).Value

            End If
         'Greatest total volume
            If ws.Cells(k, 20).Value > greatestTot Then
                greatestTot = ws.Cells(k, 20).Value
                ws.Cells(4, 24).Value = ws.Cells(k, 16).Value
                ws.Cells(4, 25).Value = ws.Cells(k, 20).Value
                greatestTotTicker = ws.Cells(k, 16).Value

            End If
                
    
    Next k
    
        'Biggest percent increase
            If greatestInc > biggestInc Then
                biggestInc = greatestInc
                biggestIncTicker = greatestIncTicker
            End If
            
        'Biggest percentalrigh decrease
            If greatestDec < biggestDec Then
                biggestDec = greatestDec
                biggestDecTicker = greatestDecTicker
            End If
        'Biggest total volume
            If greatestTot > biggestTot Then
                biggestTot = greatestTot
                biggestTotTicker = greatestTotTicker

            End If


Next ws

    ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Big Movers"
    'move created sheet to be first sheet
    Sheets("Big Movers").Move Before:=Sheets(1)
    ' Specify the location of the combined sheet
    Set big_movers = Worksheets("Big Movers")
    
    'Format cells for biggest movers
    big_movers.Range("C2:C3").NumberFormat = "0.00%"
    big_movers.Cells(4, 3).NumberFormat = "#,##0"
    
    big_movers.Cells(2, 1).Value = "Biggest % Increase"
    big_movers.Cells(3, 1).Value = "Biggest % Decrease"
    big_movers.Cells(4, 1).Value = "Biggest Total Volume"
    big_movers.Cells(1, 2).Value = "Ticker"
    big_movers.Cells(1, 3).Value = "Value"
    big_movers.Cells(2, 2).Value = biggestIncTicker
    big_movers.Cells(3, 2).Value = biggestDecTicker
    big_movers.Cells(4, 2).Value = biggestTotTicker

  
                
   
    big_movers.Cells(2, 3).Value = biggestInc
    big_movers.Cells(3, 3).Value = biggestDec
    big_movers.Cells(4, 3).Value = biggestTot
        
End Sub


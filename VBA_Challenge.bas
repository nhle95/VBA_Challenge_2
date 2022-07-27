Attribute VB_Name = "Module1"
Sub VBA_of_Wall_Street()
        'loop through all worksheets
        For Each ws In Worksheets
        
               
        'get the last row column 1
        lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'get worksheet name
        worksheetname = ws.Name
        
        'create new cells
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'set ticker count
        tickercount = 2
        
        'set start row for stock price
        stock_open_price = 2
        
        'set start value for stock volume
        total_vol = 0
        
        'start the loop
        For i = 2 To lastrow1
            
            'add all vol together first, then if ticker change, print result and reset value
            total_vol = total_vol + ws.Cells(i, 7).Value
            'check if ticker name changes
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                
                'write ticker in column 9
                ws.Cells(tickercount, 9).Value = ws.Cells(i, 1).Value
                
                'calculate the yearly change and print result in column 10.take column 6 - column 3
                 yearly_change = ws.Cells(i, 6).Value - ws.Cells(stock_open_price, 3).Value
                 'MsgBox (yearly_change)
                 ws.Cells(tickercount, 10).Value = yearly_change
                
                'formatting
                    If ws.Cells(tickercount, 10).Value < 0 Then
                    
                    'set background to red
                    ws.Cells(tickercount, 10).Interior.ColorIndex = 3
                    
                    'set background to green
                    Else
                    ws.Cells(tickercount, 10).Interior.ColorIndex = 4
                                
                    End If
                              
                'to calculate the percentage change and print result in column 11
                    If (ws.Cells(tickercount, 10).Value <> 0) Then
                    ws.Cells(tickercount, 11).Value = yearly_change / ws.Cells(stock_open_price, 3).Value
                    ws.Cells(tickercount, 11).Value = Format(ws.Cells(tickercount, 11), "Percent")
                
                    Else
                    ws.Cells(tickercount, 11).Value = 0
                    ws.Cells(tickercount, 11).Value = Format(ws.Cells(tickercount, 11), "Percent")
                    End If
                       
                'print result in column 12
                ws.Cells(tickercount, 12).Value = total_vol
                
                'reset the stock volume total
                total_vol = 0
                               
                'to move to open price of next ticker
                               
                tickercount = tickercount + 1
                stock_open_price = i + 1
                'MsgBox (i)
                'MsgBox (stock_open_price)
                     
            End If
                    
        Next i
                
                'find min, max and print result
                lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
                'MsgBox (lastrow2)
                
                ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
                ws.Cells(3, 17).Value = Format(Application.WorksheetFunction.Min(ws.Range("K:K")), "percent")
                ws.Cells(2, 17).Value = Format(Application.WorksheetFunction.Max(ws.Range("K:K")), "percent")
                
                For bonus = 2 To lastrow2
                                
                If (ws.Cells(bonus, 12).Value = ws.Cells(4, 17).Value) Then
                ws.Cells(4, 16) = ws.Cells(bonus, 9)
                End If
                
                If (ws.Cells(bonus, 11).Value = ws.Cells(3, 17).Value) Then
                ws.Cells(3, 16) = ws.Cells(bonus, 9)
                End If
            
                 If (ws.Cells(bonus, 11).Value = ws.Cells(2, 17).Value) Then
                ws.Cells(2, 16) = ws.Cells(bonus, 9)
                End If
                
                
                Next bonus
                
       
            
           
                    Next ws
                    
         

      End Sub
      
      


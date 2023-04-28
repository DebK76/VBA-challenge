Sub calc_all_results()
     
    'Dim stock_volume As Long
    
    ticker_check = ""
    open_price = 0
    closed_price = 0
    stock_volume = 0
    lastrow = Range("A1").End(xlDown).Row
    StartDate = "20180102"
    EndDate = "20181231"
    pasteRowCounter = 2
    'MsgBox lastrow
    For i = 2 To lastrow
        
        ' if current ticker is not equal to last
        If ticker_check <> Cells(i, 1).Value Then
            
            'update value of ticker_checker
            ticker_check = Cells(i, 1).Value
            
            'check open_price
            If Cells(i, 2).Value = StartDate Then
                open_price = Cells(i, 3).Value
            End If
            
            'check closed date
            If Cells(i, 2).Value = EndDate Then
                closed_price = Cells(i, 6).Value
            End If
                
            'update volume
            stock_volume = stock_volume + Cells(i, 7).Value
                
            'check if you have open and closed price
            
            If open_price <> 0 And closed_price <> 0 Then
                'ticker
                Cells(pasteRowCounter, 9).Value = ticker_check
                'year change
                Cells(pasteRowCounter, 10).Value = closed_price - open_price
                
                'format cell interior
                If Cells(pasteRowCounter, 10).Value < 0 Then
                    Cells(pasteRowCounter, 10).Interior.ColorIndex = 3
                Else
                    Cells(pasteRowCounter, 10).Interior.ColorIndex = 4
                End If
                
                'percent change
                Cells(pasteRowCounter, 11).Value = (closed_price / open_price) - 1
                'volume
                Cells(pasteRowCounter, 12) = stock_volume
                
                open_price = 0
                stock_volume = 0
                closed_price = 0
                pasteRowCounter = pasteRowCounter + 1
            End If
        
        Else
        
            'check open_price
            If Cells(i, 2).Value = StartDate Then
                open_price = Cells(i, 3).Value
            End If
            
            'check closed date
            If Cells(i, 2).Value = EndDate Then
                closed_price = Cells(i, 6).Value
            End If
                
            'update volume
            stock_volume = stock_volume + Cells(i, 7).Value
                
            'check if you have open and closed price
            
            If open_price <> 0 And closed_price <> 0 Then
                'ticker
                Cells(pasteRowCounter, 9).Value = ticker_check
                'year change
                Cells(pasteRowCounter, 10).Value = closed_price - open_price
                
                'format cell interior
                If Cells(pasteRowCounter, 10).Value < 0 Then
                    Cells(pasteRowCounter, 10).Interior.ColorIndex = 3
                Else
                    Cells(pasteRowCounter, 10).Interior.ColorIndex = 4
                End If
                
                'percent change
                Cells(pasteRowCounter, 11).Value = (closed_price / open_price) - 1
                'volume
                Cells(pasteRowCounter, 12) = stock_volume
                
                open_price = 0
                stock_volume = 0
                closed_price = 0
                pasteRowCounter = pasteRowCounter + 1
            End If
        
        End If
     
    Next i

End Sub


Sub calc_greatest_increase()

    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

'change to decimal, might be a bit off

    Cells(2, 17).NumberFormat = "0.0000"
    Cells(3, 17).NumberFormat = "0.0000"
    
    tickrow = 3001
    
    Cells(2, 17).Value = WorksheetFunction.Max(Range("K2:K" & tickrow))
    Cells(3, 17).Value = WorksheetFunction.Min(Range("K2:K" & tickrow))
    Cells(4, 17).Value = WorksheetFunction.Max(Range("L2:L" & tickrow))
    
    'find the ticker
   
    Set Rang = Range("K:K").Find(Cells(2, 17).Value)
    Cells(2, 16).Value = Cells(Rang.Row, 9)
    
    Set Rang = Range("K:K").Find(Cells(3, 17).Value)
    Cells(3, 16).Value = Cells(Rang.Row, 9)
    
    Set Rang = Range("L:L").Find(Cells(4, 17).Value)
    Cells(4, 16).Value = Cells(Rang.Row, 9)
    
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    
    Range("K2:K" & tickrow).NumberFormat = "0.00%"
    
        'Columns Autofit
        
    Range("J1:R1").EntireColumn.AutoFit
        
    
End Sub
    
    Sub MyMacro()
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Range("A1").Select
    Next ws
End Sub


Sub MyMacro()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Perform your desired actions on each worksheet here
        ' For example, to select cell A1 on each worksheet:
        ws.Range("A1").Select
    Next ws
End Sub



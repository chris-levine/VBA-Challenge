Attribute VB_Name = "Module1"
Sub StockData():

For Each ws In Worksheets

    Dim i As Long
    Dim j As Long
    Dim StockCounter As Double
    Dim PercentChange As Double
    Dim Increase As Double
    Dim Decrease As Double
    Dim Volume As Double
    
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest Percent Increase"
    ws.Cells(3, 14).Value = "Greatest Percent Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    StockCounter = 2
    
    j = 2
    
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ws.Cells(StockCounter, 9).Value = ws.Cells(i, 1).Value
            
            ws.Cells(StockCounter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            
            ws.Cells(StockCounter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
                If ws.Cells(StockCounter, 10).Value > 0 Then
                
                   ws.Cells(StockCounter, 10).Interior.ColorIndex = 4
                
                Else
                
                    ws.Cells(StockCounter, 10).Interior.ColorIndex = 3
                
                End If
                
                If ws.Cells(j, 3).Value <> 0 Then
                
                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                
                ws.Cells(StockCounter, 11).Value = Format(PercentChange, "Percent")
                
                End If
                
            
            StockCounter = StockCounter + 1
            
            j = i + 1
            
            End If
        
    Next i
    
Increase = ws.Cells(2, 11).Value
Decrease = ws.Cells(2, 11).Value
Volume = ws.Cells(2, 12).Value

    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        If ws.Cells(i, 11).Value > Increase Then
        
        Increase = ws.Cells(i, 11).Value
        ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 16).Value = Increase
        
        End If
        
        If ws.Cells(i, 11).Value < Decrease Then
        
        Decrease = ws.Cells(i, 11).Value
        ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 16).Value = Decrease
        
        End If
        
        If ws.Cells(i, 12).Value > Volume Then
        
        Volume = ws.Cells(i, 12).Value
        ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 16).Value = Volume
        
        End If
    
        ws.Cells(2, 16).Value = Format(Increase, "Percent")
        ws.Cells(3, 16).Value = Format(Decrease, "Percent")
        
    Next i
       
    
Next ws

End Sub

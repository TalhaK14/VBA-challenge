Sub Stock_Data()
Dim ws As Worksheet

For Each ws In Worksheets
    
Dim Ticker As String
    
Dim Open_Price As Double
    
Dim Close_Price As Double
    
Dim Yearly_Change As Double
    
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
    
Dim Percent_Change As Double
    
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
    
Dim Initial_Price As Long
Initial_Price = 2

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker = ws.Cells(i, 1).Value
                
            Open_Price = ws.Cells(Initial_Price, 3).Value
                
            Close_Price = ws.Cells(i, 6).Value
            
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            Yearly_Change = Close_Price - Open_Price
            ws.Cells(i, 10).Value = Yearly_Change
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
            If Open_Price = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = Yearly_Change / Open_Price
            End If
            
    ws.Range("I" & Summary_Table_Row).Value = Ticker
        
    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        
    ws.Range("K" & Summary_Table_Row) = Format(Percent_Change, "Percent")
      
    ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
            If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            End If
        
    Summary_Table_Row = Summary_Table_Row + 1
        
    Total_Stock_Volume = 0
        
    Initial_Price = (i + 1)
        
        Else

        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        End If
    
    Next i
    
    ws.Range("A:M").Columns.AutoFit

Next ws

End Sub



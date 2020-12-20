Attribute VB_Name = "Module1"
Sub Stock_Market()
    For Each ws In Worksheets
    
    Dim Ticker As String
    Dim SummaryTableRow As Long
    Dim yearly_change As Double
        yearly_change = 0
    Dim percent_change As Double
        percent_change = 0
    Dim stock_vol As Double
        stock_vol = 0
    Dim Op As Double
    Dim Closed As Double
    SummaryTableRow = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Op = ws.Cells(2, 3).Value
    
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
            Closed = ws.Cells(i, 6).Value
            Ticker = ws.Cells(i, 1).Value
            yearly_change = Closed - Op
                If Op <> 0 Then
                        percent_change = ((Closed - Op) / Op)
                Else
                percent_change = 0
                End If
            stock_vol = stock_vol + ws.Cells(i, 7).Value
            
                    ws.Range("I" & SummaryTableRow).Value = Ticker
                    ws.Range("J" & SummaryTableRow).Value = yearly_change
                    ws.Range("K" & SummaryTableRow).Value = percent_change
                    ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                    ws.Range("L" & SummaryTableRow).Value = stock_vol
                    
            If ws.Range("J" & SummaryTableRow).Value < 0 Then
                     ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            Else
                     ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            End If
            
                SummaryTableRow = SummaryTableRow + 1
                yearly_change = 0
                percent_change = 0
                stock_vol = 0
                Op = 0
                Closed = 0
                Op = ws.Cells(i + 1, 3).Value
            Else
    
           ' yearly_change = yearly_change + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
        ' percent_change = percent_change + (((Closed - Op) / Op) * 100)
            
            stock_vol = stock_vol + ws.Cells(i, 7).Value
        End If
      
        Next i
        
        Next ws
End Sub




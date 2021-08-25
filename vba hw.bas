Attribute VB_Name = "Module1"
Sub Stock()

For Each ws In Worksheets
    
    
    ' naming the cells
    ws.Cells(1, 10) = "ticker"
    ws.Cells(1, 11) = "yearly change"
    ws.Cells(1, 12) = "Percent change"
    ws.Cells(1, 13) = "Total Stock Volume"
    Dim Summary_index As Long
    Summary_index = 2
    Dim total_stock_vol As Single
    total_stock_vol = 0
    Dim yearly_change As Double
    yearly_change = 0
    Dim start_index As Double
    start_index = 2
    Dim percent_yearly_change As Double
    percent_yearly_change = 0
    
    
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            If total_stock_vol = 0 Then
                ws.Cells(Summary_index, 10).Value = 0
                ws.Cells(Summary_index, 13).Value = 0
                ws.Cells(Summary_index, 11).Value = 0
                ws.Cells(Summary_index, 12).Value = 0
            Else
                ws.Cells(Summary_index, 10).Value = ws.Cells(i, 1).Value
                ws.Cells(Summary_index, 13).Value = total_stock_vol
                yearly_change = (ws.Cells(i, 6).Value - ws.Cells(start_index, 3).Value)
                ws.Cells(Summary_index, 11).Value = yearly_change
                If ws.Cells(start_index, 3).Value = 0 Then
                    percent_yearly_change = (yearly_change / 1) * 100
                Else
                    percent_yearly_change = (yearly_change / ws.Cells(start_index, 3).Value) * 100
                End If
                ws.Cells(Summary_index, 12).Value = percent_yearly_change
                start_index = i + 1
                total_stock_vol = 0
                Summary_index = Summary_index + 1
            End If
            
            
        End If
    Next i




Next ws



End Sub


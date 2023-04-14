Sub Stock()

    ' Define dims
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double

    For Each ws In Worksheets

        ' Count number of records in the file ie last row and last column
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' Get Last row for the Part2
        LastRowPart2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        ' Process Header row
        ws.Cells(1, 9).Value = Ticker
        ws.Cells(1, 10).Value = Yearly Change
        ws.Cells(1, 11).Value = Percent Change
        ws.Cells(1, 12).Value = Total Stock Volume
        ws.Cells(1, 16).Value = Ticker
        ws.Cells(1, 17).Value = Value
        ws.Cells(2, 15).Value = Greatest % Increase
        ws.Cells(3, 15).Value = Greatest % Decrease
        ws.Cells(4, 15).Value = Greatest Total Volume
        ws.Range(ws.Cells(1, 9), ws.Cells(1, 12)).EntireColumn.AutoFit
        ws.Range(ws.Cells(1, 16), ws.Cells(1, 17)).EntireColumn.AutoFit
        ws.Range(ws.Cells(2, 15), ws.Cells(4, 15)).EntireColumn.AutoFit
        row_tracker = 2
        open_price = ws.Cells(2, 3).Value
        total_volume = 0
    
        ' Main loop that goes through the entire data

        For i = 2 To LastRow
        
            total_volume = total_volume + ws.Cells(i, 7).Value
        
            If ws.Cells(i, 1).Value  ws.Cells(i + 1, 1).Value Then
            
                ws.Cells(row_tracker, 9).Value = ws.Cells(i, 1).Value
                close_price = ws.Cells(i, 6).Value
                yearly_change = close_price - open_price
                ws.Cells(row_tracker, 10).Value = yearly_change
                If ws.Cells(row_tracker, 10).Value  0 Then
                    ws.Cells(row_tracker, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(row_tracker, 10).Value  0 Then
                    ws.Cells(row_tracker, 10).Interior.ColorIndex = 3
                End If

                If open_price  0 And Not IsNull(open_price) Then
                    percent_change = yearly_change  open_price
                Else
                    percent_change = open_price
                End If
                
                ws.Cells(row_tracker, 11).Value = percent_change
                ws.Cells(row_tracker, 11).NumberFormat = 0.00%
                
                ws.Cells(row_tracker, 12).Value = total_volume
                total_volume = 0
                row_tracker = row_tracker + 1
                open_price = ws.Cells(i + 1, 3).Value
            
            End If
        Next i
        
    Next ws
    MsgBox (Completed)
End Sub

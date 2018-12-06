'Homework - VBS code (Moderate)

Sub totalStockVol()
    volume_subtotal = 0
    current_summary_row = 2
    last_row = Cells(Rows.Count, 1).End(xlUp).Row

    temp_year_close = 0
    temp_year_open = 0
    yearly_change = 0
    percent_change = 0

    for row = 2 to last_row   
        previous_stock = Cells(row - 1, 1).Value
        current_stock = Cells(row,1).Value
        next_stock = Cells(row + 1, 1).Value
        current_volume = Cells(row,7).Value


        If current_stock = next_stock And current_stock = previous_stock Then 
            volume_subtotal = volume_subtotal + current_volume
        
        ElseIf current_stock = next_stock And current_stock <> previous_stock Then
        'need to keep subtotal going 
            volume_subtotal = volume_subtotal + current_volume
        'and need to record the first_open
            temp_year_open = Cells(row,3).Value
        
        Else
            volume_subtotal = volume_subtotal + current_volume
            Cells(current_summary_row,9).Value = current_stock
            Cells(current_summary_row,12).Value = volume_subtotal
        ' 'need to record the last close & 
            temp_year_close = Cells(row,6).Value
        ' 'calculate yearly change & 
            yearly_change = temp_year_close - temp_year_open
        ' 'caluculate %  change 
            percent_change = yearly_change/temp_year_open * 100
        ' 'add yearly change to new current_row_summary,10
            Cells(current_summary_row,10).Value = temp_year_close
        ' 'add % change to new current_summary_row,11
            Cells(current_summary_row,11).Value = percent_change
        ' 'reset the row for summary & reset the subtotal
            current_summary_row = current_summary_row + 1
            volume_subtotal = 0
        End If
    Next row
End sub


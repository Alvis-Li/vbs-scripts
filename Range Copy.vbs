Sub Copy()
    Dim r_row
    Dim Range_column
    Dim column_1
    Dim column_2
    Dim increase
    For Each r In Worksheets("Sheet1").Range("A2:H999").Rows
        r_row = r.Row - 2
        increase = 8 * (r_row \ 2)
        Range_column = "B"
        column_1 = "C"
        column_2 = "F"

        If r_row Mod 2 = 1 Then
            Range_column = "H"
            column_1 = "I"
            column_2 = "L"
        End If
    
        Worksheets("Sheet2").Range("B2:F8").Copy _
            Destination:=Worksheets("Sheet3").Range(Range_column + CStr(2 + increase))
        Rows(CStr(4 + 8 * r_row) + ":" + CStr(8 + 8 * r_row)).RowHeight = 21

        Range(column_1 + CStr(4 + increase)).Value = r.Cells(, 3).Value
        
        Range(column_2 + CStr(4 + increase)).Value = r.Cells(, 2).Value
        
        Range(column_1 + CStr(5 + increase)).Value = r.Cells(, 4).Value

        Range(column_2 + CStr(5 + increase)).Value = r.Cells(, 6).Value
        
        Range(column_1 + CStr(6 + increase)).Value = r.Cells(, 5).Value
        
        Range(column_1 + CStr(7 + increase)).Value = r.Cells(, 7).Value
        
        Range(column_2 + CStr(6 + increase)).Value = r.Cells(, 8).Value
    Next

End Sub

Sub SortTablesbyDeadline()

'Sorts the Deliverables Table by oldest to newest deadline
Dim tbl As ListObject
Set tbl = Sheet2.ListObjects("Table24")
Dim sortcolumn As Range
Set sortcolumn = Range("Table24[DEADLINE]")
With tbl.Sort
   .SortFields.Clear
   .SortFields.Add Key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
   .Header = xlYes
   .Apply
End With

'Sorts the Tests Table by oldest to newest deadline
Dim tbl2 As ListObject
Set tbl2 = Sheet2.ListObjects("Table1")
Dim sortcolumn2 As Range
Set sortcolumn2 = Range("Table24[DEADLINE]")
With tbl2.Sort
   .SortFields.Clear
   .SortFields.Add Key:=sortcolumn2, SortOn:=xlSortOnValues, Order:=xlAscending
   .Header = xlYes
   .Apply
End With

'Sorts the Deliverables Worksheet by oldest to newest deadline
    ActiveWorkbook.Worksheets("Deliverables").Sort.SortFields.Add2 Key:=Range( _
        "G2:G5"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Deliverables").Sort
        .SetRange Range("A1:G5")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Sorts the Tests Worksheet by oldest to newest deadline
    ActiveWorkbook.Worksheets("Tests").Sort.SortFields.Add2 Key:=Range("H2:H3"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Tests").Sort
        .SetRange Range("A1:H3")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

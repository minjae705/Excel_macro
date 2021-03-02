Sub Sorting()
    Dim rAll As Range
    Set rAll = Sheets("격리자현황").Range("A3:T150")
    rAll.Sort Range("H2"), xlDescending, Header:=xlNo

    ActiveWorkbook.Worksheets("격리자현황").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("격리자현황").Sort.SortFields.Add(rAll.Columns(2), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(146, 208, 80)
    ActiveWorkbook.Worksheets("격리자현황").Sort.SortFields.Add(rAll.Columns(2), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 0)
    
    With ActiveWorkbook.Worksheets("격리자현황").Sort
        .SetRange rAll
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

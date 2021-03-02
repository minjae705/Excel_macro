Sub Haejejamove()
  Dim rAll As Range
  Dim iIdx As Integer
  Dim mIdx As Integer
  Dim j As Integer
  Dim i As Integer
  Dim wValue '격리해제여부
  Dim yValue '격리해제일자
  
  Set rAll = Sheets("격리자현황").Range("A3:Q300")

  '자료채우기
  i = 0
  '자료 목록
  iIdx = rAll.Rows.Count
  mIdx = Sheets("해제자현황").Cells(Rows.Count, 3).End(xlUp).Row
  
  For j = 0 To iIdx - 1
    wValue = rAll.Cells(rAll.Row + j - 2, 15)
    yValue = rAll.Cells(rAll.Row + j - 2, 16)
     If (wValue = "O") Then
        Sheets("격리자현황").Rows(j + 3).Copy Sheets("해제자현황").Rows(mIdx + 1 + i)
        Sheets("격리자현황").Rows(j + 3).EntireRow.Delete
        
        i = i + 1
        j = j - 1
     End If
    
  Next j
End Sub

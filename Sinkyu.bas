Sub Sinkyu()
  Dim rAll As Range
  Dim iIdx As Integer
  Dim j As Integer
  'Static i As Integer
  Dim vValue '검사결과
  Dim xValue '신규
  
  Set rAll = Sheets("격리자현황").Range("A3:Q150")
    
  '자료채우기
  'i = 0
  '자료 목록
  iIdx = rAll.Rows.Count
  
  For j = 0 To iIdx - 1
     xValue = rAll.Cells(rAll.Row + j - 2, 2)
     If (xValue = "O") Then
        '연번
        Sheets("보고서양식").Cells(9 + i, 1) = i + 1
        Sheets("보고서양식").Cells(9 + i, 1).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 1).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '기관명
        Sheets("보고서양식").Cells(9 + i, 2) = rAll.Cells(rAll.Row + j - 2, 3)
        Sheets("보고서양식").Cells(9 + i, 2).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 2).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 2).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '직급
        Sheets("보고서양식").Cells(9 + i, 3) = rAll.Cells(rAll.Row + j - 2, 5)
        Sheets("보고서양식").Cells(9 + i, 3).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 3).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '성명
        Sheets("보고서양식").Cells(9 + i, 4) = rAll.Cells(rAll.Row + j - 2, 6)
        Sheets("보고서양식").Cells(9 + i, 4).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 4).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 4).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 4).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 4).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '담당업무
        Sheets("보고서양식").Cells(9 + i, 5) = rAll.Cells(rAll.Row + j - 2, 7)
        Sheets("보고서양식").Cells(9 + i, 5).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 5).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '시작일
        Sheets("보고서양식").Cells(9 + i, 6) = rAll.Cells(rAll.Row + j - 2, 8)
        Sheets("보고서양식").Cells(9 + i, 6).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 6).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 6).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 6).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 6).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '종료일
        Sheets("보고서양식").Cells(9 + i, 7) = rAll.Cells(rAll.Row + j - 2, 9)
        Sheets("보고서양식").Cells(9 + i, 7).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 7).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 7).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 7).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '격리장소
        Sheets("보고서양식").Cells(9 + i, 8) = rAll.Cells(rAll.Row + j - 2, 10)
        Sheets("보고서양식").Cells(9 + i, 8).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 8).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 8).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 8).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 8).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '사유
        Sheets("보고서양식").Range(Cells(9 + i, 9), Cells(9 + i, 10)).Merge
        Sheets("보고서양식").Cells(9 + i, 9) = rAll.Cells(rAll.Row + j - 2, 12)
        Sheets("보고서양식").Cells(9 + i, 9).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 9).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 10).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 9).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 10).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 9).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 10).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '비고
        Sheets("보고서양식").Cells(9 + i, 11) = rAll.Cells(rAll.Row + j - 2, 13)
        Sheets("보고서양식").Cells(9 + i, 11).Interior.ColorIndex = 6
        Sheets("보고서양식").Cells(9 + i, 11).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 11).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 11).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + i, 11).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        i = i + 1
     End If
    
  Next j
End Sub

Sub Haeje()
  Dim rAll As Range
  Dim iIdx As Integer
  Dim j As Integer
  Dim k As Integer
  Dim wValue '해제여부
  Dim yValue '격리해제일자
  
  Set rAll = Sheets("격리자현황").Range("A3:Q150")
    
  '자료채우기
  k = 0
  '자료 목록
  iIdx = rAll.Rows.Count
  
  For j = 0 To iIdx - 1
    wValue = rAll.Cells(rAll.Row + j - 2, 15)
    yValue = rAll.Cells(rAll.Row + j - 2, 16)
     If (wValue = "O") Then
        '연번
        Sheets("보고서양식").Cells(9 + k, 12) = k + 1
        Sheets("보고서양식").Cells(9 + k, 12).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 12).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 12).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 12).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '기관명
        Sheets("보고서양식").Cells(9 + k, 13) = rAll.Cells(rAll.Row + j - 2, 3)
        Sheets("보고서양식").Cells(9 + k, 13).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 13).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 13).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 13).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '직급
        Sheets("보고서양식").Cells(9 + k, 14) = rAll.Cells(rAll.Row + j - 2, 5)
        Sheets("보고서양식").Cells(9 + k, 14).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 14).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 14).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 14).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '성명
        Sheets("보고서양식").Cells(9 + k, 15) = rAll.Cells(rAll.Row + j - 2, 6)
        Sheets("보고서양식").Cells(9 + k, 15).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 15).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 15).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 15).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '담당업무
        Sheets("보고서양식").Cells(9 + k, 16) = rAll.Cells(rAll.Row + j - 2, 7)
        Sheets("보고서양식").Cells(9 + k, 16).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 16).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 16).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 16).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '시작일
        Sheets("보고서양식").Cells(9 + k, 17) = rAll.Cells(rAll.Row + j - 2, 8)
        Sheets("보고서양식").Cells(9 + k, 17).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 17).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 17).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 17).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '종료일
        Sheets("보고서양식").Cells(9 + k, 18) = rAll.Cells(rAll.Row + j - 2, 9)
        Sheets("보고서양식").Cells(9 + k, 18).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 18).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 18).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 18).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '격리장소
        Sheets("보고서양식").Cells(9 + k, 19) = rAll.Cells(rAll.Row + j - 2, 10)
        Sheets("보고서양식").Cells(9 + k, 19).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 19).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 19).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 19).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '사유
        Sheets("보고서양식").Range(Cells(9 + k, 20), Cells(9 + k, 21)).Merge
        Sheets("보고서양식").Cells(9 + k, 20) = rAll.Cells(rAll.Row + j - 2, 12)
        Sheets("보고서양식").Cells(9 + k, 20).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 21).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 20).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 21).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 20).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 21).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '비고
        'Sheets("보고서양식").Cells(9 + k, 22) = rAll.Cells(rAll.Row + j - 2, 13)
        Sheets("보고서양식").Cells(9 + k, 22).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 22).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 22).Borders(xlEdgeTop).LineStyle = xlContinuous
        Sheets("보고서양식").Cells(9 + k, 22).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        k = k + 1
     End If
    
  Next j
End Sub

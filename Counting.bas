Sub Counting()
  Dim rAll As Range
  Dim iIdx As Integer
  Dim j As Integer
  Dim i As Integer
  Dim k As Integer
  
  Dim o_o As Integer '1급+1급
  Dim o_t As Integer '1급+2급
  Dim t_t As Integer '2급+2급
  Dim o_g As Integer '1급+구급교육
  
  Dim sob1 '소방서1
  Dim bu1 '부서1
  Dim te1 '팀1
  Dim sod1 '소대1
  Dim da1 '대원1
  Dim ja1 '자격1
  
  Dim sob2 '소방서2
  Dim bu2 '부서2
  Dim te2 '팀2
  Dim sod2 '소대2
  Dim da2 '대원2
  Dim ja2 '자격2
  
  Set rAll = Sheets("구급대 자격현황(수정)").Range("A4:N2000")

  i = -3
  j = -2
  k = 53
  o_o = 0
  o_t = 0
  t_t = 0
  o_g = 0
  
  iIdx = rAll.Rows.Count
  
  Do While (k < 77)
    sob1 = rAll.Cells(rAll.Row + i, 2)
    bu1 = rAll.Cells(rAll.Row + i, 3)
    te1 = rAll.Cells(rAll.Row + i, 4)
    sod1 = rAll.Cells(rAll.Row + i, 5)
    da1 = rAll.Cells(rAll.Row + i, 6)
    ja1 = rAll.Cells(rAll.Row + i, 12)
    
    sob2 = rAll.Cells(rAll.Row + j, 2)
    bu2 = rAll.Cells(rAll.Row + j, 3)
    te2 = rAll.Cells(rAll.Row + j, 4)
    sod2 = rAll.Cells(rAll.Row + j, 5)
    da2 = rAll.Cells(rAll.Row + j, 6)
    ja2 = rAll.Cells(rAll.Row + j, 12)
    
     'If ((sob1 = sob2) And (bu1 = bu2) And (te1 = te2) And (sod1 = sod2) And (da1 = "구급대원1") And (da2 = "구급대원2")) Then
        If ((ja1 = "1급" Or ja1 = "1급(간호사)" Or ja1 = "2급(간호사)" Or ja1 = "간호사") And (ja2 = "1급" Or ja2 = "1급(간호사)" Or ja2 = "2급(간호사)" Or ja2 = "간호사")) Then
            o_o = o_o + 1
        ElseIf (ja1 = "구급교육" Or ja2 = "구급교육") Then
            o_g = o_g + 1
        ElseIf (ja1 = "2급" And ja2 = "2급") Then
            t_t = t_t + 1
        Else
            o_t = o_t + 1
        End If
     'End If

    If (sob1 <> rAll.Cells(rAll.Row + i + 3, 2)) Then
        Sheets("구급대 자격현황(수정)").Cells(k, 20) = o_o
        Sheets("구급대 자격현황(수정)").Cells(k, 22) = o_t
        Sheets("구급대 자격현황(수정)").Cells(k, 24) = t_t
        Sheets("구급대 자격현황(수정)").Cells(k, 26) = o_g
        k = k + 1
        o_o = 0
        o_t = 0
        t_t = 0
        o_g = 0
    End If
    
    i = i + 3
    j = j + 3
  Loop
End Sub

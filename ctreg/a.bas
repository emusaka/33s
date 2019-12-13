Function ctreg(rtyu As String, tyui As String) As Long
    '最終行を返す(ctrl+end次行がhiddenの場合の対処版)
    ctreg = Workbooks(rtyu).Sheets(tyui).Range("A1").SpecialCells(xlLastCell).Row()
    ctreg = ctreg + 1
    Do Until Workbooks(rtyu).Sheets(tyui).Cells(ctreg, 1).EntireRow.Hidden = False
        ctreg = ctreg + 1
    Loop  'ctrl+endの次行がhiddenだったら、hiddenされた最終行を返す。
    ctreg = ctreg - 1
End Function

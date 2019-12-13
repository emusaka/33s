Function ctreg(rtyu As String, tyui As String) As Long
    ctreg = Workbooks(rtyu).Sheets(tyui).Range("A1").SpecialCells(xlLastCell).Row()
    ctreg = ctreg + 1
    Do Until Workbooks(rtyu).Sheets(tyui).Cells(ctreg, 1).EntireRow.Hidden = False
        ctreg = ctreg + 1
    Loop
    ctreg = ctreg - 1
End Function

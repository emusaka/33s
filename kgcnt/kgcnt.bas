Function kgcnt(cef As String, kgr As String) As Long
    '文字列に含まれる区切り文字の数を返す。
    'cef:文字列、kgr:区切り文字
    Dim cunt As Long
    cunt = 0
    If kgr <> "" Then
        Do Until InStr(bb + 1, cef, kgr) = 0
            cc = InStr(bb + 1, cef, kgr)
            cunt = cunt + 1
            bb = cc
        Loop
    End If
    kgcnt = cunt
End Function

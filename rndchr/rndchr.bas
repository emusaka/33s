Function rndchr(suei As String) As String
    Dim aa As Long, bb As Integer, san As Integer
    '最小値～最大値の範囲で乱数の整数を取得  変数二文字化　86_014k
    Randomize
    san = Int(4 * Rnd + 1)  '1~4が取得範囲
    If suei = "suu" Then  '数字指名
        rndchr = LTrim(Str(Int(7 * Rnd + 3)))  '0+3~6+3→3~9が取得範囲(0~2は取得除外)
    ElseIf suei = "syou" Then '英小文字指名
        aa = 0
        Do Until aa = 1
            bb = Int((26 - 1 + 1) * Rnd + 1)
            If InStr(1, "aeiucklosvwxz", Chr(bb + 97 - 1)) = 0 Then aa = 1 '取得除外リスト
        Loop
        rndchr = Chr(bb + 97 - 1)
    ElseIf suei = "dai" Then '英大文字指名
        aa = 0
        Do Until aa = 1
            bb = Int((26 - 1 + 1) * Rnd + 1)
            If InStr(1, "ABCEIKOSUVWXZ", Chr(bb + 65 - 1)) = 0 Then aa = 1 '取得除外リスト
        Loop
        rndchr = Chr(bb + 65 - 1)
    'ここから無指名
    ElseIf san = 1 Then '数字
        rndchr = LTrim(Str(Int(7 * Rnd + 3)))  '0+3~6+3→3~9が取得範囲(0~2は取得除外)
    ElseIf san >= 2 And san <= 3 Then '英小文字
        aa = 0
        Do Until aa = 1
            bb = Int((26 - 1 + 1) * Rnd + 1)
            If InStr(1, "aeiucklosvwxz", Chr(bb + 97 - 1)) = 0 Then aa = 1 '取得除外リスト
        Loop
        rndchr = Chr(bb + 97 - 1)
    Else  '英大文字
        aa = 0
        Do Until aa = 1
            bb = Int((26 - 1 + 1) * Rnd + 1)
            If InStr(1, "ABCEIKOSUVWXZ", Chr(bb + 65 - 1)) = 0 Then aa = 1 '取得除外リスト
        Loop
        rndchr = Chr(bb + 65 - 1)
    End If
End Function

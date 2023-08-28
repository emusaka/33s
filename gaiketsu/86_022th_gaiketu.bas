Attribute VB_Name = "gaiketu"
Public flag As Boolean 'UserForm1連携のため要public
Public dw As String, fmt As String 'UserForm2連携のため要public
Const hrkt As Long = 0  'strconv,phonetic関連 phonetic無効(新)　'この行、消さない
'Const hrkt As Long = 16  'strconv,phonetic関連 phonetic有効(従来)　'この行、消さない
Dim cnt As Long, k As Long, xFlag As Boolean, shog As String
Dim bfn As String, shn As String, bfshn As Worksheet, twbsh As Worksheet, twt As Worksheet
Dim dd1 As Long, dd2 As Long, gg2 As Long, gg1 As Long, mghz As Long, mg2 As Long
Dim twn As String, sr(8) As Long
'ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー
Function kgcnt(cef As String, kgr As String) As Long
    '文字列に含まれる区切り文字の数を返す
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
'ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー◇ー
Function ctreg(rtyu As String, tyui As String) As Long
    '最終行を返す(ctrl+end次行がhiddenの場合の対処版)
    ctreg = Workbooks(rtyu).Worksheets(tyui).Range("A1").SpecialCells(xlLastCell).Row()
    ctreg = ctreg + 1
    Do Until Workbooks(rtyu).Sheets(tyui).Cells(ctreg, 1).EntireRow.Hidden = False
        ctreg = ctreg + 1
    Loop  'ctrl+endの次行がhiddenだったら、hiddenされた最終行を返す。
    ctreg = ctreg - 1
End Function
'ー◇ー◇ー以上、元pubikoued
Sub GetRGBValue() '◇共有対象外
    Dim ii As Long
    g1g = Selection.Row    '選択開始行
    g2g = Selection.Rows(Selection.Rows.Count).Row   '選択終了行
    d1d = Selection.Column    '選択開始列
    d2d = Selection.Columns(Selection.Columns.Count).Column    '選択終了列
    
    For ii = 1 To 21
        haro = Cells(ii, d1d).Font.Color
        Cells(ii, d1d).Font.Color = haro
    Next
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function syutoku() As String

    If Left(Application.OperatingSystem, 9) = "Macintosh" Then
        syutoku = MacScript("set userName to short user name of (system info)" & vbNewLine & "return userName")
    Else
        syutoku = Environ("Username")
    End If

End Function
'Sub syutoku2()
'    MsgBox Environ("Username")
'    MsgBox Application.UserName
'    MsgBox Application.OperatingSystem
'
'
'End Sub
Sub kinkyu()  'これがある事で強制終了、データ消滅を阻止できているので、これは消さない。
    On Error GoTo myError
    bfn = ActiveWorkbook.Name 'bfn,shnはパブリック
    shn = ActiveSheet.Name
    '↓この2行でエラーが起こる
    ThisWorkbook.Activate
    Sheets("▲集計_雛形").Activate
    DoEvents
    Workbooks(bfn).Activate          'こちらへ（いかが？）
    Sheets(shn).Activate
    DoEvents
    Exit Sub
myError:
    MsgBox "エラーです。終わります。"
    End
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub kyosydou()
    '共通（外結、初回、複写）共通の初動をまとめる。
    kinkyu

    Dim ii As Long, nm As Variant
    Dim xsheet As Worksheet 'xWsheet→xsheet

    Application.CutCopyMode = False
    DoEvents
    gg1 = Selection.Row    '選択開始行
    gg2 = Selection.Rows(Selection.Rows.Count).Row   '選択終了行
    dd1 = Selection.Column    '選択開始列
    dd2 = Selection.Columns(Selection.Columns.Count).Column    '選択終了列
    
    If dd1 = 0 Then MsgBox "dd1がゼロですね"
    bfn = ActiveWorkbook.Name 'bfn,shnはパブリック
    shn = ActiveSheet.Name
    Set bfshn = Workbooks(bfn).Worksheets(shn) '30ｓ63より。当シート・当ファイル
    
    twn = ThisWorkbook.Name 'マクロファイル名そのもの(?.xlsm)
    Set twbsh = Workbooks(twn).Worksheets("▲集計_雛形") '30ｓ74より
    
    nm = Array("", "対象ｼｰﾄ名", "当：突合列", "対：突合列", "対：ｵｰﾙ1列", "対：ｶｳﾝﾄ列", "対：加算列･他", "当：転載列", "対：転載列", "対：実質加算列", "当：実質転載列") '30s63,array化
    If IsError(Application.Match(nm(1), Range(bfshn.Cells(1, 2), bfshn.Cells(200, 2)), 0)) Then Call oshimai("", bfn, shn, 4, 2, "（処理中止）「" & nm(1) & "」が見つかりません")
    For ii = 1 To 8
        sr(ii) = WorksheetFunction.Match(nm(ii), Range(bfshn.Cells(1, 2), bfshn.Cells(200, 2)), 0)
        If sr(0) < sr(ii) Then sr(0) = sr(ii)
    Next
    sr(0) = sr(8) + 1  '下側から移設
    
    If IsError(Application.Match("。", Range(bfshn.Cells(1, 1), bfshn.Cells(1, 5000)), 0)) Then
        Call oshimai("", bfn, shn, 1, 0, "「" & shn & "」シート右上に「。」がありません。入れて下さい。")
    Else
        mghz = Application.Match("。", Range(bfshn.Cells(1, 1), bfshn.Cells(1, 5000)), 0)
    End If
    shog = "log_" & syutoku() & "_" & Format(Date, "yyyymm")
    
    'ログシート有無chk、30s82、初回のみ→ここに移設
    For Each xsheet In ThisWorkbook.Sheets
        If xsheet.Name = shog Then xFlag = True 'boolean型の初期値はfalse
    Next xsheet
    If xFlag = True Then ' 該当のシートがある場合の処理
        '（何もしない）
    Else ' 該当のシートがない場合の処理 '
        Workbooks(twn).Activate '30s83
        Worksheets.Add
        ActiveSheet.Name = shog
        nm = Array("", "項目名", "項番", "log", "date", "timestamp", "メモ", "to", "最右列", "from9") '30s83,array化
        For ii = 1 To 9
            Workbooks(twn).Sheets(shog).Cells(1, ii).Value = nm(ii)
        Next
        Workbooks(bfn).Activate          'こちらへ（いかが？）
        bfshn.Select
    End If
    
    xFlag = False
    For Each xsheet In ThisWorkbook.Sheets     '転記有無chk、30s82e
        If xsheet.Name = "高速シート_" & syutoku() Then xFlag = True
    Next xsheet

    If xFlag = True Then ' 該当のシートがある場合の処理
        Set twt = Workbooks(twn).Worksheets("高速シート_" & syutoku()) '30s82f
        twt.Cells.Clear
    
        ThisWorkbook.Activate
        twbsh.Select '30s86_017q　変なエラー解消用(フォントをテーマの色にすると再起動する件)
                                  '↑100%解消している訳ではない
        Sheets("高速シート_" & syutoku()).Activate
        Rows("6:6").Select
        Selection.Delete Shift:=xlUp
        DoEvents
    
        Range(Cells(2, 3), Cells(2, 3)).Select
        Workbooks(bfn).Activate       '86_017h
        bfshn.Select
        DoEvents
    Else ' 該当のシートがない場合の処理 '
        Worksheets.Add
        ActiveSheet.Name = "高速シート_" & syutoku()
        Set twt = Workbooks(twn).Worksheets("高速シート_" & syutoku()) '30s82f
        Workbooks(bfn).Activate
        bfshn.Select
    End If
    
    bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1 '30s83 =sum(a:a)+1からこちらへ
    Exit Sub
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー外部結合
Sub gaiketsu()   '一番上↑にパブリック変数あり。見逃し注意　　’旧「外部結合」
    Dim f As String, xsheet As Worksheet, xbook As Workbook, wsfag As Boolean
    Dim bun As Long, bni As Long
    Dim mr() As String, er() As Currency
    Dim nkg As Long, kahi As Long, cted(1) As Long, rrr As Long 'ppp→rrr 86_014r
    Dim pap() As Long   'pap配列変数化(86_019d)
    Dim er2() As Currency, er3() As Currency, er5() As Long, er78() As Currency, er9() As Currency, er34 As String
    Dim mr2() As String, mr3() As String, mr5() As String, mr8() As String, mr9() As String
    Dim a As Long, pkt As Long, nn As Long, n1 As Long, qq As Long, pqp As Long
    Dim am1 As String, am2 As String, h As Long, m As Long, k0 As Long, h0 As Long, n2 As Long, kg1 As String
    Dim qap As Long, ii As Long, jj As Long, trt As Long, tst As Long, dif As Long, axa As Long
    Dim saemp '←今も型が設定されていない
    Dim hirt As Variant, hiru As Variant, tameshi As Range, ct8 As String, tempo As String, baba As String
    Dim kasan As Variant, c5 As Long, c99 As String, c98 As String, ct3 As String, hk1 As String, rog As String
    Dim zzz As String, zyyz() As String, xxxx() As String, zxxz() As String, nifuku As Long
  
    '再計算を一旦自動に
    Application.Calculation = xlCalculationAutomatic
    'おまじない(「コードの実行が中断されました」対処)
    Application.EnableCancelKey = xlDisabled
    UserForm1.StartUpPosition = 1 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
    UserForm1.Show vbModeless
    UserForm1.Repaint

    kyosydou  '共通の初動
 
    bfshn.Cells(sr(3), 5).Value = chekku     'Excelソート仕様表記　30s86_020t　Excelバージョンチェック
    bfshn.Cells(sr(5), 5).Value = "hrkt_" & hrkt     'Excelソート仕様表記　30s86_020t　Excelバージョンチェック
    
    If bfshn.Cells(sr(3), 5).Value = "旧ｿｰﾄ" And hrkt = 0 Then MsgBox "旧ｿｰﾄ(2013)でhrkt=0(phonetic不使用)です。キーに同じひらがなカタカナある時(「あ」「ア」など)誤作動の危険あり注意。"
    
    twt.Cells.Clear

    Workbooks(bfn).Activate
    bfshn.Select
        
    If dd1 = 0 Then Call oshimai("", bfn, shn, 1, 0, "dd1がゼロです")

    'その日の初回チェック
    ii = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(ii, 1).Value = ""
        ii = ii + 1
        If ii = 50000 Then
            MsgBox "空白行が見つからないようです"
            Exit Sub
        End If
    Loop
    
    If Workbooks(twn).Sheets(shog).Cells(ii - 1, 4).Value <> Val(Format(Now(), "yyyymmdd")) Then
        Call oshimai("", bfn, shn, 1, 0, "その日の初回は、最初に[FIRST]ボタンを押して下さい。")
    End If
    
    If bfshn.Cells(sr(0) - 1, 5) = "" Then
        Call oshimai("", bfn, shn, sr(0) - 1, 5, "★集計名を入力して下さい。")
    End If

    Call iechc(hk1) '旧igchc(hk1)
    hk1 = ""
    flag = False

    ii = 1
    '当シート側のall1探し
    Do Until bfshn.Cells(ii, 1).Value = "all1"
        If IsNumeric(bfshn.Cells(ii, 1).Value) And bfshn.Cells(ii, 1).Value <> "" Then '17s
            Call oshimai("", bfn, shn, ii, 1, "一列目は数値を入れないで下さい")
        End If
        ii = ii + 1
        If ii = 100 Then
            Call oshimai("", bfn, shn, 1, 0, "当シート「all1」が見つからないようです")
        End If
    Loop

    k = ii + 1     'k確定'kはデータ開始行(サンプル行ではなくなった。)
    bfshn.Cells(1, 2).Value = k     'データ開始行

    'この時点でのiiは、当シートの「all1」記載行
    Do Until bfshn.Cells(ii, 1).Value = ""
        ii = ii + 1
    Loop
    'ここでのiiは当シートall1列の空白になった行、データ無しの場合はデータ開始行
    
    'オートフィルタが設定されてれば、解除
    If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    nn = 0 'セル空白チェックフラグ
    If dd1 <= 5 Then Call oshimai("", bfn, shn, 1, 0, "タテヨコ終了、6列目以降が対象です")    'よ
    If dd2 >= mghz Then Call oshimai("", bfn, shn, 1, 0, "枠外が選択されてます")    '86_108i

    
For a = dd1 To dd2 '選択範囲列分の繰り返し　ら
    
    Application.Goto bfshn.Cells(k, a - 3), True
    bfshn.Cells(k, a).Select
    
    kg1 = "" 'リセット　kg2はmr(2,3,bni)のredimでリセットされる。
    bni = 1  'リセット
    bun = 1  'リセット
    nkg = 0 'リセット
    nifuku = 0
    
    '複文節用区切り文字の確定　fnywti→rvsrz3を流用
    kg1 = Mid(rvsrz3(bfshn.Cells(sr(7), a).Value, 3, "ｦ", 0), 1, 1)  'kg2定義は先 30s73:nkg2→0

    If bfshn.Cells(sr(6), a).Value <= -90 Then
        bun = 1 '-90台は強制1
    ElseIf kg1 = "" Then
        nkg = 1 'ゐ区切り無指定(kg1="")は文節ゼロ（区切りしない）
    Else '複文節数(bun)確定のためのルーチン(kg1<>"")
        For ii = 1 To 8
            Do Until rvsrz3(bfshn.Cells(sr(ii), a).Value, bni, kg1, 0) = ""
                bni = bni + 1
                If bni > 120 Then Call oshimai("", bfn, shn, 4, 2, "bni120超え")
            Loop
            bni = bni - 1
            If ii = 2 And bni > 1 Then
                nifuku = 1
                MsgBox "2行目での複文節あり。注意を。"  '解禁へ　86_016v
            End If
            If bun < bni Then bun = bni '（この時点でbun確定）
            bni = 1
        Next
    End If

'◆Ａ◆単文節
    bni = 1 'リセット
    ReDim pap(9, bun) 'As long  30s86_019c
    ReDim er(11, bun) 'As Currency 10→11 30s83
    ReDim mr(4, 11, bun) 'As String 30s81第３因子化 10→11 30s83 ,mr(3→mr(4：ｦ6ｦ→ｦFｦ 用
    
    '再計算を手動に
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    For ii = 1 To 7 'mr(0,記述　　ii=2からでも大丈夫と思われる。
        mr(0, ii, 1) = rvsrz3(bfshn.Cells(sr(ii), a).Value, 1, kg1, nkg)  'n行目の全因子　bni:1
    Next
    
    '1文節目特注
    mr(1, 1, 1) = rvsrz3(mr(0, 1, 1), 2, "ｦ", 2)  'シート名
    mr(2, 1, 1) = rvsrz3(mr(0, 1, 1), 3, "ｦ", 0)  'ファイル名
    
    mr(2, 7, 1) = kg1 'ゐ　７行目第二因子
If StrConv(Left(mr(1, 1, 1), 1), 8) = "*" Then 'り　30s57左一文字が＊だけの時（＊?の時は?がmr(2,0,1)に入る）
        bfshn.Cells(sr(0), a).Value = bfshn.Cells(sr(0) + 3, a).Value
        bfshn.Cells(sr(0) + 1, a).Value = bfshn.Cells(sr(0) + 4, a).Value
        bfshn.Cells(sr(0) + 2, a).Value = bfshn.Cells(sr(0) + 5, a).Value
Else  'り(*だけでない通常時、外結最後部まで続く)
    
    For ii = 1 To 7  '空白確認（bni=1）
        If bfshn.Cells(sr(ii), a).Value = "" Then nn = sr(ii)
    Next
    
    If bfshn.Cells(sr(6), a).Value > -90 Then     'ここでのiiは8 -99は8行目確認しない
        If bfshn.Cells(sr(8), a).Value = "" Then nn = sr(8) 'ii→8　（同値）
    End If
    If nn > 0 Then Call oshimai("", bfn, shn, nn, a, "外結設定情報が空欄の所があります")
'◇Ａ◇単文節ここまで↑

'◆Ｂ◆複文節for（：準備編）ふあ↓　文節毎のforがここから始まる
    For bni = 1 To bun
        nn = 0 'セル空白チェックフラグ→？
        mr(2, 0, bni) = bfn '30s82
        mr(1, 0, bni) = shn  '30s82
        For ii = 1 To 7
            mr(0, ii, bni) = rvsrz3(bfshn.Cells(sr(ii), a).Value, bni, kg1, nkg)  'n行目の全因子
            If mr(0, ii, bni) = "" Then mr(0, ii, bni) = mr(0, ii, bni - 1)  '2文節目以降、空欄なら前節コピペ
        Next
        
        If bfshn.Cells(sr(6), a).Value <= -90 Then
            mr(1, 1, bni) = shn  'シート名　'-90台はセル見ず、bfshn強制　セルは日付書式など自由に書けられる。
            mr(2, 1, bni) = bfn  'ファイル名
        Else
            mr(1, 1, bni) = rvsrz3(mr(0, 1, bni), 2, "ｦ", 2)  'シート名 ""にはならない
            mr(2, 1, bni) = rvsrz3(mr(0, 1, bni), 3, "ｦ", 0)  'ファイル名30s73:nkg2→0
        End If
        If Left(StrConv(mr(1, 1, bni), 8), 1) = "\" Then mr(1, 1, bni) = shn '￥→シート名変換
        If mr(1, 1, bni) = "" Then Call oshimai("", bfn, shn, sr(1), a, "対象シート名(" & bni & "文節目)が空欄です")
        If mr(2, 1, bni) = "" Then mr(2, 1, bni) = bfn
        mr(4, 1, bni) = mr(1, 1, bni)    '86_014s これ、抜けてた。
        wsfag = False   'ファイル・シート有無chk
        Do Until wsfag = True
            For Each xbook In Workbooks
                If xbook.Name = mr(2, 1, bni) Then wsfag = True
            Next xbook
            If wsfag = False Then
'                If MsgBox("ファイルが確認できませんが、" & vbCrLf & "続行しますか？", 289, "ファイル不明") = vbCancel Then 'キャンセル時
                    '↑この手法（Yes,No選択方式）はここでは意味なかったので、このやり方は中止。↓こちらへ
                    Call oshimai("", bfn, shn, sr(1), a, "実施中止しました。" & vbCrLf & "ファイルが確認できません")
'                end if
            End If
        Loop
    
        wsfag = False
        For Each xsheet In Workbooks(mr(2, 1, bni)).Sheets
            If xsheet.Name = mr(1, 1, bni) Then wsfag = True
        Next xsheet
        If wsfag = False Then Call oshimai("", bfn, shn, sr(1), a, mr(1, 1, bni) & " のシートが不明です(" & bni & "文節目)")
    
        nn = 0 '一旦リセット
            
            For ii = 2 To 6  '(ヱ)はここで定まる。
                mr(2, ii, bni) = rvsrz3(mr(0, ii, bni), 3, "ｦ", 0)
            Next
            mr(3, 4, bni) = rvsrz3(mr(0, 4, bni), 4, "ｦ", 0) '第3因子（工事中）
            
            If bni >= 2 Then mr(2, 7, bni) = mr(2, 7, bni - 1) '戻った(ゐ)強制前節コピー

            'mr(2, 11, bni)・・・項準とかが入る。↓er(11,bni)はリファラ変数
            mr(2, 11, bni) = koudicd(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), mr(0, 4, bni))
            
            mr(1, 2, bni) = yhwat1(bfn, shn, 2, sr(2), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 2, bni))
            mr(4, 2, bni) = zhwat1(bfn, shn, 2, sr(2), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 2, bni))

            For ii = 3 To 6
                mr(1, ii, bni) = yhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(ii), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, ii, bni))
                mr(4, ii, bni) = zhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(ii), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, ii, bni))
            Next
            
            mr(1, 7, bni) = yhwat1(bfn, shn, 2, sr(7), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 7, bni))
            mr(4, 7, bni) = zhwat1(bfn, shn, 2, sr(7), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 7, bni))
       
            For ii = 2 To 7
                er(ii, bni) = Val(mr(1, ii, bni)) 'ヱ対応時：val関数は数字と認識出来る所までを数値変換する。要は第一列目を。
            Next
    
        If nn > 0 Then Call oshimai("", bfn, shn, nn, a, "（処理中止）" & vbCrLf & "数値以外の情報があります。")
        If er(5, 1) >= 0 And er(7, 1) < 0 And er(6, 1) > -90 Then Call oshimai("", bfn, shn, 1, 0, "（処理中止：文法エラー）" & vbCrLf & "er(5,0)>=0　で　er(7,0)<0　です。")  'bni→1
    
        If er(6, bni) > -90 Then '8番目（転載列）の処理
            c98 = bfshn.Cells(sr(8), a).Value  '26ｓ
            mr(0, 8, bni) = rvsrz3(bfshn.Cells(sr(8), a).Value, bni, kg1, nkg)
            mr(2, 8, bni) = rvsrz3(mr(0, 8, bni), 3, "ｦ", 0) ' '30s73:nkg2→0
            mr(1, 8, bni) = yhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2有_30s56
            mr(4, 8, bni) = zhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2有_30s56
        
            If bni >= 2 Then '2文節目以降からというのがミソ、1文節目は下記
                If mr(0, 8, bni) = "" Then mr(0, 8, bni) = mr(0, 8, bni - 1)
                If mr(2, 8, bni) = "" Then mr(2, 8, bni) = mr(2, 8, bni - 1) 'mr(2,8,0)誤動作阻止 2文節目以降、空欄なら前節コピペ
                mr(1, 8, bni) = yhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2有_30s56
                mr(4, 8, bni) = zhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2有_30s56
            End If
            er(8, bni) = Val(mr(1, 8, bni))
            pap(8, 0) = kgcnt(mr(1, 8, bni), mr(2, 4, bni)) '30s86_019  papx→pap(x,0) 以下同じ ,0)は既存改築分(いじらない,papzouseiしない)
            
'            pap(8, bni) = kgcnt(mr(1, 8, bni), mr(2, 4, bni)) '86_019e  bni:1?　,bni)は新規造成分
            '↓等価
            Call papzousei(pap(), mr(), 8, bni) '86_019e  bni:1?　,bni)はpapzouseiは新規造成分が対象(既存改築分はやらない)
            
            If Round(er(6, bni)) <> -2 And er(8, bni) < 0 Then   '複数列解禁86_013f
                If pap(8, 0) = 0 Then
                    If mr(2, 8, bni) = "" Then mr(2, 8, bni) = "、" '区切り文字デフォは「、」（-15でも適用）
                Else
                    If rvsrz3(mr(2, 8, bni), 0 + 1, mr(2, 4, bni), 0) = "" Then mr(2, 8, bni) = "、" & mr(2, 8, bni) '30s86_018c
                End If
            End If
        
            If er(6, bni) <= -3 And er(8, bni) < 0 Then Call oshimai("", bfn, shn, 1, 0, "「c=-3以下は重複連なり型(e<0)は実行できないです。")
            
            '30s86_012s追加↓
            If er(5, bni) < 0 And er(7, bni) = 0 Then Call oshimai("", bfn, shn, sr(7), a, "差分時7行0は実施されなくなりました。")
            
            '30s82d追加↓                      er(5,1)→er(5,bni) 86_010
            If (er(6, 1) = -1 Or er(6, 1) = -2) And er(5, bni) >= 0 And kg1 <> "" And rvsrz3(bfshn.Cells(sr(7), a).Value, 2, kg1, 0) <> "" Then Call oshimai("", bfn, shn, nn, a, "6行-1-2の時でnot差分時(通常時)は7行複文節不可です。")
 
            '↓順変更　30s86_012w
            If Round(er(6, bni)) = -2 Then
                mr(0, 9, bni) = mr(0, 8, bni)
                mr(2, 9, bni) = mr(2, 8, bni)
                mr(1, 9, bni) = mr(1, 8, bni)
                er(9, bni) = Val(mr(1, 9, bni))
            ElseIf er(6, bni) > 0 Then
                mr(0, 9, bni) = mr(0, 6, bni)
                mr(2, 9, bni) = mr(2, 6, bni)
                mr(1, 9, bni) = mr(1, 6, bni)
                er(9, bni) = Val(mr(1, 9, bni))
            End If

            If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0) Then
                'MsgBox "社数抽出該当" '←特命条件追加30s86_012w
                bfshn.Cells(sr(6), 5).Value = "特命" '←20191118　こちらへ　86_015c
            Else
                bfshn.Cells(sr(6), 5).Value = ""
            End If
          
            If Not (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) > 0) Then '特命条件変更30s86_012w、追加30s86_012s
            '   6行ー２でない、　　　　　　かつ　８行０でない、　　　かつ　（５行＋ あるいは　６行０以下）　ならば
                If (Not Round(er(6, bni)) = -2) And (Not er(8, bni) = 0) And (er(5, bni) >= 0 Or er(6, bni) <= 0) Then '新仕様(30s83)
                '10行目処理(er(10,0)は実質の当シート転載列) (特命条件の連載型は実施する30s86_012w)
                    mr(0, 10, bni) = mr(0, 8, bni)
                    mr(2, 10, bni) = mr(2, 8, bni)
                    mr(1, 10, bni) = mr(1, 8, bni)
                    er(10, bni) = Val(mr(1, 10, bni))
                    If pap(8, 0) = 0 Then
                        If mr(2, 8, bni) <> "" And er(10, bni) = 0.1 Then
                            er(10, bni) = -0.1
                            er(8, bni) = -0.1
                        End If
                    Else
                        If rvsrz3(mr(2, 8, bni), 0 + 1, mr(2, 4, bni), 0) <> "" And er(10, bni) = 0.1 Then
                            er(10, bni) = -0.1
                            er(8, bni) = -0.1
                        End If
                    End If
                End If
            End If '特命条件追加30s86_012s
        End If

        nn = 0 '一旦リセット

        'log処理部
        ii = 1

        If StrConv(Left(mr(0, 1, bni), 1), 8) <> "*" And bfn <> twn Then '*接頭辞や、ツールそのものの外結はログ生成対象外
            Do Until ii = 0
                If rvsrz3(rog, ii, "ヱ", 0) = "" Then
                    rog = mr(2, 1, bni) & "\" & mr(1, 1, bni) & "ヱ" & rog  'rogはその外結実施範囲での対象シートファイルの集合体(ヱで連結)
                    ii = 0
                ElseIf rvsrz3(rog, ii, "ヱ", 0) = mr(2, 1, bni) & "\" & mr(1, 1, bni) Then
                    ii = 0
                Else
                    ii = ii + 1
                End If
                If ii = 200 Then Call oshimai("", bfn, shn, k, a, "うまくいってない。")
            Loop

        
        End If
        If er(2, bni) > 0 And er(9, bni) > 0 And er(10, bni) <> 0 And er(7, bni) = 0 And er(5, bni) >= 0 Then '↓特命条件考慮
            Call oshimai("", bfn, shn, 1, 0, "処理中止。c加算ありで同列転載しようとしています。確認を。")
        End If
        Call papzousei(pap(), mr(), 2, bni) '86_019e　papzouseiは、新規区画分(bni:1?)を造成
    Next
'◇Ｂ◇複文節ふあ（準備編）ここまで↑

'◆Ｃ◆単文節（本番プレ）ここから↓
    bni = 1 '1文節目で判断、実施部分'うあ

    '前回データコピー
    If StrConv(Left(bfshn.Cells(sr(1), a).Value, 2), 8) <> "**" Then '**はやらないへ024_検証2
        bfshn.Cells(sr(0), a).Value = bfshn.Cells(sr(0) + 3, a).Value
        bfshn.Cells(sr(0) + 1, a).Value = bfshn.Cells(sr(0) + 4, a).Value
        bfshn.Cells(sr(0) + 2, a).Value = bfshn.Cells(sr(0) + 5, a).Value
    End If
    h = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1 + k - 2 '現状の最下行(以下伸びていく)※データ無しの時は項目行となり、h=k-1となるので注意
    
    '体裁(teisai)ここから
    trt = 0 'リセット trt・・・タイプの定義　trt:セルのタイプ
    'trt：0初期ﾘｾｯﾄ,-9・-99型,-2当列加算型,-1当列転載型,1強制文字列型
    'trt：-9・-99型 ,-1当列転載型、-2当列加算型(含特命加算型)、1強制文字列型(含特命連載型)
    'trt定義部
    If er(6, bni) <= -90 Then '-99はこちら
        trt = -9
        '↓特命条件（加算or連載型）
    ElseIf er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0 Then
        If er(8, bni) > 0 Then
            trt = -2 '特命加算型
        Else
            trt = 1 '特命連載型
        End If
    ElseIf er(6, bni) > 0 Or Round(er(6, bni)) = -2 Then
        trt = -2  '当列加算型 差分時も
    ElseIf er(5, bni) >= 0 Then
        If er(10, bni) < 0 Then
            trt = 1   '強制文字列型
            'MsgBox "強制文字列です。"
        Else
            trt = -1    '当列転載型
            'MsgBox "連載じゃないです。"
        End If
    ElseIf er(5, bni) < 0 Then  '差分文字比較時
        If mr(2, 6, bni) = "1" Then
            Call oshimai("", bfn, shn, sr(6), a, "差分時の6行op1の運用は終わっている")
        ElseIf mr(2, 6, bni) = "-1" Then
            trt = -1 '差分を文字(ﾊﾞﾘｭｰ)で表現 →転載型
        Else
            trt = -1 '差分を数(0,1)で表現、-2→-1へ(最初文字、あとで数値のため)
        End If
    End If
    If trt = 0 Then Call oshimai("", bfn, shn, 1, 0, " trt = 0 ？")  'trtゼロ状態で無いことの確認念のため
        
    'tst定義部ここから
    tst = -1 'リセット　tst:セルの型
    'tst : -1初期ﾘｾｯﾄ , ｰ2通貨,   0標準,   1文字列,　 7一括踏襲型 , 8ちまちま
        
    If trt = -9 Then
        If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then 'ーー
            tst = 1     '文字列型
        ElseIf (er(4, bni) < 0 Or er(4, bni) = 0.1) Then
            tst = -2     '通貨型
        ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then '対象シート名を模倣　a11_1列→対象シートに変更
            tst = 7 '一括踏襲型(模倣型)
        Else
            tst = 0 '標準型
        End If
    ElseIf trt = -2 Then
        If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  'ーー
            tst = 0     '標準型
        ElseIf (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  '＋ー
            tst = 7     '一括型
        ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then  'ー＋
                        '何もしない(tst = -1)
                        '↓こちらへ転向 30s86_014f
            tst = 8     '一括型
            MsgBox "ちまちま(tst = 8)"
        Else                                              '＋＋
            tst = -2    '通貨型
        End If
    ElseIf trt = 1 Then '強制文字列
        tst = 1 '文字列型
    ElseIf trt = -1 Then
        If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  'ーー
            tst = 1     '文字列型
        ElseIf (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  '＋ー
            tst = 7     '一括型
        ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then  'ー＋
            tst = 8     'ちまちまセル踏襲型
            MsgBox "ちまちま(tst = 8)"
        Else                                              '＋＋
            tst = 0    '標準型
        End If
    End If
    
    If tst = -1 Then Call oshimai("", bfn, shn, 1, 0, " tst = -1 ？")
    
    If h >= k And StrConv(Left(bfshn.Cells(sr(1), a).Value, 2), 8) <> "**" Then '既存行無い時はこちら実行しないへ,「＊＊あ」も実行しない
        '選択列の既存データクリア(事前既存行)
        With Range(bfshn.Cells(k, a), bfshn.Cells(h, a))
            .ClearContents
            .ClearComments '30s79コメントもクリア
            .Interior.Pattern = xlNone '30s79
            .NumberFormatLocal = "G/標準"  '30s83デフォルトでまず
        End With


        '当列事前の体裁一律化 86_013d trt,tst化
        If trt = -9 Then '-99はこちら   86_020r記→　◇E◇で同じこと実施するようになったので、こちらは不要かも
            If tst = 1 Then '←当列転載文字型
                Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "@"
            ElseIf tst = -2 Then  '通貨型
                Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[赤]-#,##0"
            ElseIf tst = 7 Then '対象シート名を模倣　a11_1列→対象シートに変更
                Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = bfshn.Cells(sr(1), a).NumberFormatLocal 'sr(4)→sr(1)へ30s68
            End If
        '新仕様-99以外 文字列指名or連載　er(7, bni) > 0は-1の複数列対応
        ElseIf tst = 1 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "@"
        ElseIf tst = -2 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        ElseIf tst = 0 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "G/標準"
        End If
        '体裁(teisai)ここまで 86_012g
        
        If er(2, 1) < 0 Or kgcnt(mr(1, 2, 1), mr(2, 4, bni)) > 0 Then  '30s45右端クリアへ
            With Range(bfshn.Cells(k, mghz), bfshn.Cells(h, mghz + 2)) '3目列目まで削除へ30s80
                .ClearContents
                .NumberFormatLocal = "@" '文字列
            End With
            With Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(h, mghz + 1))  '2列目は数値用
                .ClearContents
                .NumberFormatLocal = "G/標準"
            End With
        End If
    End If
    
    '4列目に要素転記(-90台も実施、８行目別途) ※文字列変換済み
    For ii = 1 To 7
        hk1 = "" '第3因子対応
        If StrConv(Left(mr(0, ii, bni), 1), 8) = "*" Then  '30s75（*有無も追加）
            bfshn.Cells(sr(ii), 4).Value = "*ｦ" & mr(4, ii, bni) & "ｦ" & mr(2, ii, bni) & hk1
        Else
            bfshn.Cells(sr(ii), 4).Value = "ｦ" & mr(4, ii, bni) & "ｦ" & mr(2, ii, bni) & hk1
        End If
    Next

    '8行内容を4列目に要素転記　※文字列変換済み
    If mr(2, 8, bni) = Chr(Val("&H" & "0A")) Then '行の膨らみ阻止16s
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then
            bfshn.Cells(sr(8), 4).Value = "*ｦ" & mr(4, 8, bni) & "ｦ(LF)"
        Else
            bfshn.Cells(sr(8), 4).Value = "ｦ" & mr(4, 8, bni) & "ｦ(LF)"
        End If
    Else
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then
            bfshn.Cells(sr(8), 4).Value = "*ｦ" & mr(4, 8, bni) & "ｦ" & mr(2, 8, bni)
        Else
            bfshn.Cells(sr(8), 4).Value = "ｦ" & mr(4, 8, bni) & "ｦ" & mr(2, 8, bni)
        End If
    End If
    
    If er(6, bni) < -90 Then
        '85_024検証6　-90台で５行目ヱ対応用
        pap(5, 0) = kgcnt(mr(1, 5, bni), mr(2, 4, bni)) '5行目ヱの数 30s48
        ReDim er5(pap(5, 0))
        ReDim mr5(pap(5, 0))
        er5(0) = Val(mr(1, 5, bni)) 'ヱでない時向け　erx()は通貨型なのでvalを被せざるを得ない。
        mr5(0) = mr(2, 5, bni)
        If pap(5, 0) > 0 Then 'mr(2, 4, bni) <> ""
            For ii = 0 To pap(5, 0)
                er5(ii) = Val(rvsrz3(mr(1, 5, bni), ii + 1, mr(2, 4, bni), 0))
                mr5(ii) = rvsrz3(mr(2, 5, bni), ii + 1, mr(2, 4, bni), 0)
            Next
        End If
    End If
'?????
   
  bfshn.Cells(sr(0), 4).Value = bni & "/" & bun & "文節"   '30s86_012新設
  bfshn.Cells(sr(0) + 1, 4).Value = rvsrz3(bfshn.Cells(1, a).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0) & "列"
   
'◇Ｃ◇単文節（本番プレ）ここまで↑
 
 'るA 以降-99or*,**はここ通過
  If er(6, bni) > -90 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then

'◆Ｄ◆複文節for（：本番実行編）ふい↓　文節毎のforがここから始まる ループ部
    For bni = 1 To bun    '上に移行へ　86_016w
    
    '5列目より左転載防止 86_012_m
    If er(7, bni) >= 1 And er(7, bni) <= 5 And er(5, bni) >= 0 Then
        Call oshimai("", bfn, shn, 1, 0, "5列目より左に転載しないで下さい。")
    End If
    
    'ループ前の初期値
    n1 = k  'n1：前回突合処理対象行_当シートでの、kはデータ開始行（固定）m2→n1
    pap(2, 0) = kgcnt(mr(1, 2, bni), mr(2, 4, bni)) '86_016w mr(1, 2, 1)→mr(1, 2, bni)
    ReDim er2(pap(2, 0))
    ReDim mr2(pap(2, 0))
    
    er2(0) = Val(mr(1, 2, bni)) 'mr2活性化(30s86_017a)　mr(1, 2, 1)→mr(1, 2, bni)
    mr2(0) = mr(2, 2, bni)
    If pap(2, 0) > 0 Then
        For ii = 0 To pap(2, 0)
            er2(ii) = Val(rvsrz3(mr(1, 2, bni), ii + 1, mr(2, 4, bni), 0))
            mr2(ii) = rvsrz3(mr(2, 2, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
        
    '下側からこちらへ　86_020c
    pap(3, 0) = kgcnt(mr(1, 3, bni), mr(2, 4, bni)) '3行目ヱの数
    ReDim er3(pap(3, 0))
    ReDim mr3(pap(3, 0))
    er3(0) = Val(mr(1, 3, bni)) 'ヱでない時向け　erx()は通貨型なのでvalを被せざるを得ない(以下同)。
    mr3(0) = mr(2, 3, bni)
    If pap(3, 0) > 0 Then
        For ii = 0 To pap(3, 0)
            er3(ii) = Val(rvsrz3(mr(1, 3, bni), ii + 1, mr(2, 4, bni), 0))
            mr3(ii) = rvsrz3(mr(2, 3, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
        
    '近似高速種別確認　619 当面単文節で(∵近似可否er2(pap(2,0))のため。)
    mr(1, 11, bni) = "" '一旦リセット 純高速/近似高速/ノーマル　が入る。

    If er(2, 1) < 0 Then '　bni→1に書換え(ここの条件分岐全体を)　86_016w
        If er(6, 1) < 0 Then mr(1, 11, 1) = "近似高速" Else mr(1, 11, 1) = "純高速"
    Else
        mr(1, 11, 1) = "ノーマル"
    End If 'bni:2以降はmr(1,11,bni)nullなので注意
    
    bfshn.Cells(sr(2), 5).Value = mr(1, 11, 1)     'ノーマル・純・近似表記５列　bni→1
    
    'mghz　４行目に区切り文字を入れる mghzの要素生成で使用  mghz→mghz+1へ
'    With bfshn.Cells(sr(4), mghz)
    '↓30s86_022mh
    With bfshn.Cells(sr(4), mghz + 1)
        .NumberFormatLocal = "@"  '文字列として入れる
        .Value = mr(2, 4, bni)
    End With
    
    '◆◆◆
        If (er(2, 1) < 0 Or pap(2, 0) > 0) Then   '2行目にヱがある時か高速時実施(1文節目で判断)、1文節目ヱある時→全文節ヱ適用される。

        If k <= h Then '当シート既存情報あり
            If StrConv(bfshn.Cells(sr(1), a), 8) = "\" And er(6, 1) >= 0 And mr(1, 2, bni) <> mr(1, 3, bni) And pap(2, 0) > 0 Then
                    'Call oshimai("", bfn, shn, sr(2), a, "\の時で仮想キー使用＆6行>0のときは2行3行一致が必要です。") '無限ループ防止s
                    'MsgBox "2行3行不一致(無限ループの可能性有)" '86_016r　86_021b解除、無限ループするなら戻す
            End If
            
            'mghz2列の情報埋め込み（値と数値の書式で貼付)
            If er(2, 1) < 0 Then '高速（単数列も複数列も）

                bfshn.Cells(k, mghz).Select '動作可視化
                DoEvents    '86_019r
                Application.Calculation = xlCalculationAutomatic    '数式計算方法自動に　'新形式　85_007

'                tempo = "R[0]C[" & LTrim(Str(Abs(er2(0)) - mghz)) & "]"
                '↓30s86_022k
                tempo = "R[0]C[" & LTrim(Str(Abs(er2(0)) - (mghz + 1))) & "]"

                baba = tempo
                
                For ii = 1 To pap(2, 0) '数式の素生成
'                    tempo = "&" & "R" & LTrim(Str(Abs(sr(4)))) & "C[0]" & "&" & "R[0]C[" & LTrim(Str(Abs(er2(ii)) - mghz)) & "]"
                '↓30s86_022m
                    tempo = "&" & "R" & LTrim(Str(Abs(sr(4)))) & "C[0]" & "&" & "R[0]C[" & LTrim(Str(Abs(er2(ii)) - (mghz + 1))) & "]"
                    baba = baba & tempo
'MsgBox sr(4)
                Next

'                Range(bfshn.Cells(k, mghz), bfshn.Cells(h, mghz)).NumberFormatLocal = "G/標準"    'mghz列一旦標準へ（数式入れるため）
                '↓30s86_022k
                Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(h, mghz + 1)).NumberFormatLocal = "G/標準" 'mghz列一旦標準へ（数式入れるため）

'                bfshn.Cells(sr(8), mghz).FormulaR1C1 = "=" & baba & "&" & """" & """"     '←数式右端に「&""」を付加(空白セルが「0」になる対策)　30s86_020s
                '↓30s86_022k
                bfshn.Cells(sr(8), mghz + 1).FormulaR1C1 = "=" & baba & "&" & """" & """"   '←数式右端に「&""」を付加(空白セルが「0」になる対策)　30s86_020s
                

                '↓関数貼り付け化　86_020g　8行目→当シートに
'                Call copipe(bfn, shn, sr(8), mghz, sr(8), mghz, bfn, shn, k, mghz, h, mghz, 3)  '3→FormulaR1C1
                '↓30s86_022k
                Call copipe(bfn, shn, sr(8), mghz + 1, sr(8), mghz + 1, bfn, shn, k, mghz + 1, h, mghz + 1, 3) '3→FormulaR1C1


                
                
                '値に変換 '◆30s86_022jここが悪さ
'                Call copipe(bfn, shn, k, mghz, h, mghz, bfn, shn, k, mghz, h, mghz, 1)  '1→Valueコピペ　旧cpp2、-4163
                '↓30s86_022k(16ケタの数値が化ける対策)文字列として貼り付けで解決、但し、空白が昇順でてっぺんに来る副作用発生satesate変数で解決
                Range(bfshn.Cells(k, mghz), bfshn.Cells(h, mghz)).NumberFormatLocal = "@"  'mghz文字列へ
                Call copipe(bfn, shn, k, mghz + 1, h, mghz + 1, bfn, shn, k, mghz, h, mghz, 1) '1→Valueコピペ　旧cpp2、-4163
                


                 '30s86_020s
                If hrkt = 16 Then baba = "ASC(PHONETIC(R[0]C[-1]))" Else baba = "ASC(R[0]C[-1])"  '前列をコピる数式
'MsgBox baba
                bfshn.Cells(k, mghz + 1).FormulaR1C1 = "=" & baba

                '↓関数貼り付け化　86_020g
                Call copipe(bfn, shn, k, mghz + 1, k, mghz + 1, bfn, shn, k + 1, mghz + 1, h, mghz + 1, 3) '3→FormulaR1C1

                'mghz+1→mghz、値化
                Call copipe(bfn, shn, k, mghz + 1, h, mghz + 1, bfn, shn, k, mghz, h, mghz, 1)    '1→Value

                'mghz+1列(行番号)の処理(フィル活用)　※ここにあるのはmghz+1の数式をさっさと消すため
                bfshn.Cells(k, mghz + 1).Value = k

                If h > k Then
                    Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(k, mghz + 1)).AutoFill Destination:=Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(h, mghz + 1)), Type:=xlFillSeries '←連続する数値
                End If

                '30s86_022mh mghzの「""」のセルを、空白に（ソートで上に出ないようにするため）
                Call saato(bfn, shn, mghz, k, mghz, h, mghz + 1, 1)
                satesate = k
                Do Until satesate = h + 1
                    If bfshn.Cells(satesate, mghz).Value <> "" Then
                        If satesate > k Then
MsgBox "空白発動"
                            Range(bfshn.Cells(k, mghz), bfshn.Cells(satesate - 1, mghz)).NumberFormatLocal = "G/標準"
                            Call copipe(bfn, shn, k, mghz, satesate - 1, mghz, bfn, shn, k, mghz, satesate - 1, mghz, 1) '1→Valueコピペ　旧cpp2、-4163
                        End If
                        Exit Do
                    End If
                    satesate = satesate + 1
                Loop
                Call saato(bfn, shn, mghz + 1, k, mghz, h, mghz + 1, 1)
                Range(bfshn.Cells(k, mghz), bfshn.Cells(h, mghz)).NumberFormatLocal = "@"  'mghz文字列へ
                'ここまで30s86_022mh mghzの「""」のセルを、空白に（ソートで上に出ないようにするため）ここまで
                    

                'ヴ→ｳﾞ(・・ASC で処理されないので。strconv24では、ヴ→ｳﾞされるので、ここまでやって初めてASCとstrconv8が等価になる。
                Range(Workbooks(bfn).Sheets(shn).Cells(k, mghz), Workbooks(bfn).Sheets(shn).Cells(h, mghz)).Replace What:="ヴ", _
                    Replacement:="ｳﾞ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False

                '↓30s86_022k 上側に移動
'               Range(bfshn.Cells(k, mghz), bfshn.Cells(h, mghz)).NumberFormatLocal = "@"  'mghz列一旦文字列へ戻す（数式入れ、値化したので）
                DoEvents

'◆◆◆
                
                Application.Calculation = xlCalculationManual  '再計算再び手動に（重くなるため）30s66
                     
                If pap(2, 0) > 0 Then '30s86_020c　020q高速専用として復帰　２行複数列で下側がall空白chk 30s86_020a
                    ii = h
                    Do Until bfshn.Cells(ii, mghz) <> StrConv(wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3), 8 + hrkt) '←使ってるのmr(2,4,bni)だけ
                        ii = ii - 1
                    Loop
                    If ii < h Then

                        Range(bfshn.Cells(ii + 1, mghz), bfshn.Cells(h, mghz)).ClearContents
                    Else
                        'MsgBox "mghzクリアはありませんでした(2行複数列)"
                    End If
                End If
            Else '低速（単数列も複数列も）
                bfshn.Cells(sr(8), mghz).Value = "(不使用(ちまちまpap(2,0)≠0)"
                For ii = k To h
                    bfshn.Cells(ii, mghz).Value = wetaiou(bfn, shn, ii, er2(), mr(2, 4, bni), mr(1, 11, 1), mr2(), 2) 'mghz
                    bfshn.Cells(ii, mghz + 1).Value = ii 'mghz+1処理
                    Call hdrst(ii, a)         '左下ステータス表示部
                Next
                
                '30s86_020m 低速にも適用させるため、こちらにも　２行複数列で下側がall空白chk 30s86_020a
                If pap(2, 0) > 0 Then '30s86_020c　　020q低速専用として開設
                    ii = h
                    Do Until bfshn.Cells(ii, mghz) <> wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) '←使ってるのmr(2,4,bni)だけ
                        ii = ii - 1
                    Loop
                    If ii < h Then
                        Range(bfshn.Cells(ii + 1, mghz), bfshn.Cells(h, mghz)).ClearContents
                    Else
                        'MsgBox "mghzクリアはありませんでした(2行複数列)"
                    End If
                End If
            End If
            cnt = 0
        End If '当シート既存情報あり
        
        If er(2, 1) > 0 Then er(2, bni) = mghz Else er(2, bni) = -mghz '←86_016w
    End If

    '(ｈ変更)非独立集計　（「ど」対応あべこべに） '＊ど」にも対応 2行複数列にも対応となっている。ゐ複文節にも対応
    If Not (InStr(1, mr(0, 2, bni), "ｦ") > 0 And InStr(1, rvsrz3(mr(0, 2, bni), 1, "ｦ", 0), "ど") > 0) Then  'h変える
        ii = h
        Do Until ii = k - 1
            If bfshn.Cells(ii, Abs(er(2, 1))) <> "" Then Exit Do
            ii = ii - 1
        Loop
        h = ii
    Else  'h変えない
        MsgBox "ど対応"
    End If
    
    DoEvents
    
    ct3 = ""
    ct8 = "" '86_019o
    With Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)) 'フィルタ解除する（コピペでの、はしょられ防止）
        If .FilterMode Then .ShowAllData
    End With

    cted(0) = ctdg(mr(2, 1, bni), mr(1, 1, bni), er(4, bni), a)
   ' If cted(0) > 500000 Then Call oshimai("", bfn, shn, 1, 0, "対象シート行数50万行超え")  '30s86_021e
    
    cted(1) = ctdr(mr(2, 1, bni), mr(1, 1, bni), er(4, bni), a)

'◆◆◆◆◆◆
    If h >= k And (mr(1, 11, 1) = "純高速" Or mr(1, 11, 1) = "近似高速") Then   'mghzソート
        '昇順キー(mghz列)を昇順に=xlAscending 昇順処理は近似高速(一文節のみ)・純高速(全文節)共に実施へ
        Call saato(bfn, shn, mghz, k, mghz, h, mghz + 1, 1)
    
    
    
    End If


    k0 = k  '複文節一旦リセット仕様へ
    h0 = h '複文節一旦リセット仕様へ
   
    If h >= k And (mr(1, 11, 1) = "近似高速") Then
        am1 = "" '近似時は初回値強制nullへ。(k行採用できないため)
        n1 = k
    ElseIf er(2, 1) < 0 Then  '高速時はこちら
        am1 = bfshn.Cells(k, Abs(er(2, bni))).Value  '30s83こちらで復活
        n1 = bfshn.Cells(k, Abs(er(2, bni)) + 1).Value
    Else  '通常
        am1 = bfshn.Cells(k, Abs(er(2, bni))).Value  '30s83こちらで復活
        n1 = k
    End If
    'ｓか」

    pap(5, 0) = kgcnt(mr(1, 5, bni), mr(2, 4, bni)) '5行目ヱの数
    pap(9, 0) = kgcnt(mr(1, 9, bni), mr(2, 4, bni)) 'pap6から変更
    
    ReDim er5(pap(5, 0))
    ReDim mr5(pap(5, 0))
    ReDim er9(pap(9, 0))
    ReDim mr9(pap(9, 0))
    
    er9(0) = Val(mr(1, 9, bni))
    mr9(0) = mr(2, 9, bni)

    If pap(9, 0) > 0 Then
        For ii = 0 To pap(9, 0)
            er9(ii) = Val(rvsrz3(mr(1, 9, bni), ii + 1, mr(2, 4, bni), 0))
            mr9(ii) = rvsrz3(mr(2, 9, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
    
    er5(0) = Val(mr(1, 5, bni))
    mr5(0) = mr(2, 5, bni)
    If pap(5, 0) > 0 Then
        For ii = 0 To pap(5, 0)
            er5(ii) = Val(rvsrz3(mr(1, 5, bni), ii + 1, mr(2, 4, bni), 0))
            mr5(ii) = rvsrz3(mr(2, 5, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
    
    pap(8, 0) = kgcnt(mr(1, 8, bni), mr(2, 4, bni)) '8行目ヱの数  ＊の時も仮で
   
    'pap(8,0)再定義（＊のとき）30ｓ70
    If StrConv(Left(rvsrz3(mr(0, 8, bni), 1, "ｦ", 0), 1), 8) = "*" Then '30s74改良
        paq8 = pap(8, 0) / 2 'paq8は半数(：＊のグルーピング数　.5もあり得る)
        qap = 0
        For ii = 0 To Int(paq8)
            If ii = Int(paq8) And paq8 - Int(paq8) = 0 Then '最終周かつpap(8,0)偶数(孤立)
                qap = qap + 1
            Else
                fma = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 1, mr(2, 4, bni), 0)) 'from列
                If fma = 0.1 Then fma = cted(1)  '85_020
                tob = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 2, mr(2, 4, bni), 0)) 'to列
                If tob = 0.1 Then tob = cted(1)  '85_020
                qap = qap + Abs(fma - tob) + 1
            End If
        Next
        pap(8, 0) = qap - 1 'pap(8,0)再定義完了
    End If
            '↓-1:複数列比較で使用30s79、0:7側、1：8側→終了
    ReDim er78(-1 To 1, -1 To pap(8, 0)) '←pap(7,0)<=pap(8,0)　という前提,-1はc=-1-2の時だけ使用(tensai側で)
    ReDim mr8(-2 To pap(8, 0)) '30s75
    
    er78(1, 0) = Val(mr(1, 8, bni))
    
    If er(10, bni) = -0.1 Then er78(1, 0) = er(10, bni) '30s86_018d 追加（行番号連載補正対応）
    
    If pap(8, 0) > 0 Then
        mr8(0) = rvsrz3(mr(2, 8, bni), 0 + 1, mr(2, 4, bni), 0)
        If StrConv(Left(bfshn.Cells(sr(8), a).Value, 1), 8) = "*" Then
            qaap = 0
            For ii = 0 To Int(paq8) '周
                fma = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 1, mr(2, 4, bni), 0)) 'from列
                If fma = 0.1 Then fma = cted(1)
                er78(1, qaap) = fma
                If ii < Int(paq8) Or paq8 - Int(paq8) = 0.5 Then   '最終周でない、あるいは最終周でtoあり
                    tob = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 2, mr(2, 4, bni), 0)) 'to列
                    If tob = 0.1 Then tob = cted(1)  '85_020
                    If fma > tob Then sa8 = -1 Else sa8 = 1
                    For qap = qaap + 1 To qaap + 1 + Abs(fma - tob) - 1 'from＝toはfor特性上、実施されない
                        er78(1, qap) = er78(1, qaap) + (sa8) * (qap - qaap)
                    Next
                End If
                qaap = qap
            '＊の時は、mr8(qap)は当面不使用とする　'30s75
            Next
        Else
            For ii = 1 To pap(8, 0) 'これまで通り
                er78(1, ii) = Val(rvsrz3(mr(1, 8, bni), ii + 1, mr(2, 4, bni), 0))
                mr8(ii) = rvsrz3(mr(2, 8, bni), ii + 1, mr(2, 4, bni), 0)
            Next
        End If
           
        '30s86_017v 特例措置（これ連載でないときも実施されているようである。問題のときはプログラム改善）86_020j
        If mr(2, 8, bni) = mr(2, 4, bni) Then
            MsgBox "連載時の特例措置mr8(0)→" & mr(2, 8, bni) & "へ"
            mr8(0) = mr(2, 8, bni)
        End If
    Else
        mr8(0) = mr(2, 8, bni)
    End If
    
    'pap(8,0)ここまで。ここからpap(7,0)
    'fma,tobはpap(7,0)として再リセットされて使用される(pap(8,0)のが踏襲されて使用されない)
    
    soroeru = 0
    er78(0, -1) = 0 '30s62 -1はc=-1-2の時だけ使用(tensai側で)
    er78(0, 0) = Val(mr(1, 7, bni)) '(1,0)→(0,0)修正30s61_4
    
    pap(7, 0) = kgcnt(mr(1, 7, bni), mr(2, 4, bni)) '7行目ヱの数  ＊の時も仮で
    If pap(8, 0) > 0 Then 'er78を定める。
        '７行目
        If StrConv(Left(bfshn.Cells(sr(7), a).Value, 1), 8) = "*" Then  '＊の周回ルーチン
            If Val(rvsrz3(mr(1, 7, bni), pap(7, 0) + 1, mr(2, 4, bni), 0)) = 0.1 Then '7行右が「?」
                soroeru = 1
                paq7 = (pap(7, 0) - 1) / 2
                '一番右に「?」アリ→soroeruビットを立てて、ーは無い前提(：pap(7,0)-1)でpaq7を設定
            Else
                paq7 = pap(7, 0) / 2 'paq7は半数(：＊のグルーピング数　.5もあり得る)
            End If
            
            qaap = 0
            For ii = 0 To Int(paq7) '＊ペアグループ毎で周回
                fma = Val(rvsrz3(mr(1, 7, bni), ii * 2 + 1, mr(2, 4, bni), 0)) 'from列
                er78(0, qaap) = fma  'ペアグループ毎の一個目
                qap = qaap  'ここでのqapは現fmaの配列位置(0,2,,,)
                pap(7, 0) = qaap
                If ii < Int(paq7) Or paq7 - Int(paq7) = 0.5 Then
                '(右のーは無い仮定での)最終周でない、あるいは最終周でtoあり
                '※右がーの処理は下のsoreoeru=1　の所で実施される。
                    tob = Val(rvsrz3(mr(1, 7, bni), ii * 2 + 2, mr(2, 4, bni), 0)) 'to列
                    
                    If qaap + 1 + Abs(fma - tob) - 1 > pap(8, 0) Then
                        Call oshimai("", bfn, shn, 1, 0, "転載列数：7行目>8行目です。確認を。")
                    End If
                    
                    If fma > tob Then sa7 = -1 Else sa7 = 1
                    For qap = qaap + 1 To qaap + 1 + Abs(fma - tob) - 1 '30ｓ70fromto導入
                        er78(0, qap) = er78(0, qaap) + (sa7) * (qap - qaap)
                    Next
                    pap(7, 0) = qap - 1
                    qaap = qap 'ここでのqap,qaapは、ある＊グループ精算後のnextのfma入れ込む位置
                End If
            Next
            
            If pap(7, 0) > pap(8, 0) Then Call oshimai("", bfn, shn, 1, 0, "pap(7,0)>pap(8,0)です。確認を。")
 
            If soroeru = 1 Then '7行右が「?」・・pap(8,0)と揃える
                For qaap = pap(7, 0) + 1 To pap(8, 0)
                    er78(0, qaap) = er78(0, pap(7, 0)) + (qaap - pap(7, 0))
                Next
                If er78(0, qaap - 1) >= mghz Then
                    Call oshimai("", bfn, shn, 1, Int(er78(0, qaap - 1)), "mghzはみ出てます。" & vbCrLf _
                    & rvsrz3(bfshn.Cells(1, Int(er78(0, qaap - 1))).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0) & "列までデータあり")
                End If
                pap(7, 0) = pap(8, 0)
            End If
        Else '普通のルーチン
            If pap(7, 0) > pap(8, 0) Then Call oshimai("", bfn, shn, 1, 0, "pap(7,0)>pap(8,0)です。確認を。")
            For ii = 1 To pap(8, 0)
                er78(0, ii) = Val(rvsrz3(mr(1, 7, bni), ii + 1, mr(2, 4, bni), 0))
            Next
        End If
    End If
'pap系の値策定はここまで。以降でpapの値はないと思われる。
    'MsgBox IsNumeric("&H" & mr8(0))
    '↓mr8(pap(7,0))→mr8(0) 86_017v
    If Len(mr8(0)) > 1 And (er(8, bni) < 0 Or Round(er(6, bni), 0) = -15 Or Round(er(6, bni), 0) = -14) Then   '-14追加202005
        'If Val("&H" & mr8(0)) <> 0 Then
        '↓30s86_022r
        If IsNumeric("&H" & mr8(0)) Then
            'MsgBox "16進変換アリ " & mr8(0) & "→" & Chr(Val("&H" & mr8(0)))
            mr8(0) = Chr(Val("&H" & mr8(0)))
        Else
            MsgBox "16進変換サレズ：" & mr8(0)
        End If
        '16進変換対象は、マニュアル連載のmr8だけである。
    End If

    If (er(5, bni) < 0 And pap(7, 0) <> pap(8, 0)) Then Call oshimai("", bfn, shn, sr(7), a, "列数が一致しません（差分の複数列比較）")
    
    '4列目に要素転記(８行目別途) ※２文節以降も実施,4列目反映へ
    For ii = 1 To 7
        hk1 = "" '第3因子対応
        If mr(3, ii, bni) <> "" Then
        'MsgBox "第3因子:" & mr(3, ii, bni)
        hk1 = "ｦ" & mr(3, ii, bni)
        End If
        If StrConv(Left(mr(0, ii, bni), 1), 8) = "*" Then  '30s75（*有無も追加）　'mr(1,→mr(4 へ
            bfshn.Cells(sr(ii), 4).Value = "*ｦ" & mr(4, ii, bni) & "ｦ" & mr(2, ii, bni) & hk1
        Else
            bfshn.Cells(sr(ii), 4).Value = "ｦ" & mr(4, ii, bni) & "ｦ" & mr(2, ii, bni) & hk1
        End If
    Next

    '8行内容を4列目に要素転記　※文字列変換済み
    If mr(2, 8, bni) = Chr(Val("&H" & "0A")) Then '行の膨らみ阻止16s
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then  'mr(1,→mr(4 へ
            bfshn.Cells(sr(8), 4).Value = "*ｦ" & mr(4, 8, bni) & "ｦ(LF)"
        Else
            bfshn.Cells(sr(8), 4).Value = "ｦ" & mr(4, 8, bni) & "ｦ(LF)"
        End If
    Else
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then
            bfshn.Cells(sr(8), 4).Value = "*ｦ" & mr(4, 8, bni) & "ｦ" & mr(2, 8, bni)
        Else
            bfshn.Cells(sr(8), 4).Value = "ｦ" & mr(4, 8, bni) & "ｦ" & mr(2, 8, bni)
        End If
    End If
    
    bfshn.Cells(sr(0), 4).Value = bni & "/" & bun & "文節"
    bfshn.Cells(sr(4), 5).Value = mr(2, 11, bni)    'koum（項準とか）
    cnt = 0  '左下カウンタリセット
    qq = er(11, bni)

    '対象シートの「1」探し 30s81 上側から引っ越し
    If Left(mr(2, 11, bni), 2) = "項無" And qq = 0 Then qq = 1 '項無の場合（・）
    If Left(mr(2, 11, bni), 2) <> "項準" Then '・項目行ある場合はその次行、ない場合は１行目から探す
        Do Until Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(4, bni))).Value = "all1" Or Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(4, bni))).Value = 1
            qq = qq + 1
            If qq = 2000 Then Call oshimai("", bfn, shn, k, a, "対象シートのall1列の「1」が見つからないようです。")
        Loop
        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(4, bni))).Value = "all1" Then qq = qq + 1
    Else '項準：次行から・　項順2、項順aも
        qq = qq + 1
    End If
    
    If pap(8, 0) > 0 Then
        '7行分 一旦仮完成、保留　30s79 複数列コメント挿入
        If pap(7, 0) > 0 Then '←参照シートに項目行が無い場合を除く
            Call tnsai(ct8, tst, ct3, er78(), a, sr(7), bni, 1, k - 1, -7, mr(), er(), pap(7, 0), mr8())
        End If
        '8行完成、一旦保留
        If er(11, bni) > 0 And pap(8, 0) > 0 Then '←参照シートに項目行が無い場合を除く
            If mr(2, 11, bni) = "項固" Then
                Call tnsai(ct8, tst, ct3, er78(), a, sr(8), bni, 1, qq - 1, -8, mr(), er(), 0, mr8())
            Else
                Call tnsai(ct8, tst, ct3, er78(), a, sr(8), bni, 1, Int(er(11, bni)), -8, mr(), er(), 0, mr8())
            End If
        End If
    End If

    '◆高速シート作成30s82f
        If mr(1, 11, 1) = "純高速" Then
            Call kskst(pap(7, 0), h, er78(), er9(), mr9(), er3, mr3(), er5(), mr5(), c5, pap(3, 0), pap(5, 0), bni, qq, rrr, mr(), er(), a, cted())
            hirt = Range(twt.Cells(1, 1), twt.Cells(rrr + 1, 9)).Value '※rrr→rrr+1(項無或いはデータすっからかん対策)　86_018i
        End If '高速シート作成ここまで
    
    ii = qq 'rrrは高速シートのデータ終了行(含ロック因子)を温存へ、iiが増えていく
    cnt = 0 'カウンタリセット
    pqp = 0 'ロックオン可否リセット

    If mr(1, 11, 1) = "純高速" Or mr(1, 11, 1) = "近似高速" Then
        hiru = Range(bfshn.Cells(1, mghz), bfshn.Cells(h, mghz + 1)).Value
    End If

    Application.Goto bfshn.Cells(k, a - 3), True  '86_020n
    bfshn.Cells(k, a).Select
    
    '◆ここから行毎
    Do While ii <= cted(0) '628より
        ct3 = ""
        ct8 = "" '86_019o
        If Abs(er(4, bni)) >= 1 Then
            If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(ii, Abs(er(4, bni))).Value = "" Then
                Exit Do  '対象シートのall_1列
            End If
        End If

        If mr(1, 11, 1) = "純高速" Then  '高速ロックオン判定
            If hirt(ii, 2) = "000_0000000" Then  '純高速時は高速シート参照へ　30s86_012r   0→0000000  86_019c
                UserForm4.StartUpPosition = 2 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
                UserForm4.Show vbModeless
                UserForm4.Repaint
                bfshn.Cells(sr(2), 5).Value = "純高ﾛｯｸ" '20191118
                k0 = h0  'これがポイント
                pqp = 1
                Unload UserForm4
                UserForm1.Repaint
            End If
            
            ct3 = hirt(ii + pqp, 6)  'ここでct3注入(特命条件で使用、転載エレメント)
            ct8 = hirt(ii + pqp, 8)  '   30s86_019o(特命条件で使用、転載エレメント(値の方))
            
            '↓これするために、５行０でも、高速シート３列目に１を埋め込む。へ。
            If hirt(ii + pqp, 3) = "" Then
                c5 = 0
            Else  '↓c5最終選考(純高速)追加(86_020f)
                                           '↓使用されてるのmr(2,4,bni)だけ
                If pap(2, 0) > 0 And wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) = hirt(ii + pqp, 1) Then '､､､､だけ→不採用
                    c5 = 0
                    MsgBox "2行目複数列の環境で対象列がall空白、出ました、無視されます(純高速)。避けたければ2行目単数列で。"
                Else '採用
                    c5 = 1
                    qq = Val(rvsrz3(hirt(ii + pqp, 2) & "", 2, "_", 0)) '高速シート→対象シートの行に換算,配列変数導入、c5=1のみに適用へ  86_019c
                End If
            End If
        
        Else  '純高速以外　ct3は必ず""
            c5 = kaunta(mr(), ii, pap(5, 0), bni, er5(), mr5()) 'ここでのiiは高速シートの0カウント行
             '（ここでのc5 = 1は、まだ仮採用状態）
            If c5 = 1 Then  'で、←c5最終選考(not純高速)　'↓使用されてるのmr(2,4,bni)だけ
                If pap(2, 0) > 0 And wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) = wetaiou(mr(2, 1, bni), mr(1, 1, bni), ii, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) Then
                    c5 = 0 '採用予定→不採用へ　86_020f追加
                    MsgBox "2行目複数列の環境で対象列がall空白、出ました、無視されます(not純高速)。避けたければ2行目単数列で。"
                Else '採用予定→本採用へ
                    qq = ii 'c5=1確定時のみに適用へ(一応バグ)　86_015b　従来はこちらでけであった
                End If
            End If
        End If
        
        'カウント対象であれば実施（で無ければ飛ばす）
        If c5 = 1 Then
            'ヱ複文節の返り値を返す(○ヱ▽ヱ◇形式)。複文節で無くても常時使われる(am2)
            If mr(1, 11, 1) = "純高速" Then  '高速ロックオン判定
'                am2 = hirt(ii + pqp, 1)   '30s85_027検証6
                am2 = WorksheetFunction.Clean(hirt(ii + pqp, 1))   '30s86_022i
            Else
                am2 = wetaiou(mr(2, 1, bni), mr(1, 1, bni), qq, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) '30s51
            End If
            
            If am2 = "" Then am2 = "ー(情報空白行)ー"
            hk1 = ""

            '突合キーunicode→?になる対策開始　86_016q
            If InStr(am2, "?") = 0 And InStr(StrConv(am2, 8 + hrkt), "?") > 0 Then MsgBox "uniあり注意（" & am2  '86_020s　24→8 + hrkt

            'p判定--■--■--■--　p→pkt　86_014n
            If mr(1, 11, 1) = "純高速" Or mr(1, 11, 1) = "近似高速" Then  '高速ロックオン判定
               pkt = kskup(am1, am2, n1, n2, h, er(2, bni), er(6, bni), k0, h0, pap(2, 0), er2(), mr(1, 11, 1), pqp, er5(0), er3(), hiru)
            Else '低速
               pkt = tskup(am1, am2, n1, n2, h, er(2, bni), er(6, bni), k0, h0, pap(2, 0), er2(), mr(1, 11, 1), pqp, er5(0), er3())
            End If

            kahi = 0 '(加算可否判別フラグ) zyou→kahiへ　86_014j
            kasan = 0 '16s
            
            If pkt = -1 Then Exit Do
            If pkt = -2 Then '★629(85_001) pap(2,0)はゼロである。　高速ベタ
                MsgBox "高速ベタ"
                '特命条件での高速ベタはどうなるのでしょうか？↓今は強制終了
                If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0) Then Call oshimai("", bfn, shn, 1, 0, "特命条件、大丈夫？")
                
                If rrr - qq <= 1 Then '対象シートデータ2行以下なら高速ベタでなく通常方法に戻す(p=2)
                     pkt = 2
                Else
                    MsgBox "表空白ベタ貼り転載開始します。" & vbCrLf & _
                    "行数：(予定)" & rrr - qq + 1 & vbCrLf & _
                    "列数：(予定)" & 1 _
                    , 64, "ベタ貼り「転載」開始"
                    
                    '一列二列、５列のオートフィル跡地
                    Call betat4(bfn, shn, k, 0.4, k + rrr - qq, 0.4, bfn, shn, k, 1, "pp", "1")
                    
                    If Not StrConv(Left(mr(0, 2, bni), 1), 8) = "*" Then 'ー２にも2行＊概念反映へ
                        Call betat4(bfn, shn, k, 0.4, k + rrr - qq, 0.4, bfn, shn, k, 2, "pp", "c")
                    End If
                    
                    Call cpp2("", Now(), 0, 0, 0, -1, bfn, shn, k, 5, k + rrr - qq, 5, -4163) '5列目（now）これに凝縮 速
                    
                    '突合列　'↓文字列での値化（速達のため）
                    Call betat4(twn, "高速シート_" & syutoku(), qq, 1, rrr, 1, bfn, shn, k, Abs(er2(0)), "mm", "")
                    
                    '当列
                    If mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then 'ｦｦ定数加算対応
                        Call betat4(bfn, shn, k, 0.4, k + rrr - qq, 0.4, bfn, shn, k, a, "pm", mr(2, 9, bni))
                    ElseIf (er(8, bni) = 0 Or er(7, bni) <> 0) And er(6, bni) > 0.5 Then  '加算処理部(同列転載なら実行しない)
                        Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Int(er(6, bni)), rrr, Int(er(6, bni)), bfn, shn, k, a, "pm", "")
                    End If
                   
                    If pap(7, 0) <> pap(8, 0) Then MsgBox "p=-2でpap(7,0) <> pap(8,0)です.留意を(処理は続行)"
                        'MsgBox "横方向ベタ貼りです"
                    If er(8, bni) <> 0 Then
                        For jj = 0 To pap(7, 0) 'pap(7,0),pap(8,0)＝0の時は一回ポッキリ実施
                            qap = jj
                            If er78(1, jj) <> 0 Then
                                If er78(0, jj) = 0 Then
                                    axa = a
                                    If er(6, bni) > 0 And er(5, bni) >= 0 Then
                                        Call oshimai("", bfn, shn, 1, 0, "処理中止。a加算ありで同列転載しようとしています。確認を。")
                                    End If
                                Else
                                    axa = er78(0, jj)
                                End If
                            End If
                    
                            If pap(7, 0) > 0 And jj <> pap(7, 0) Then
                                Do While er78(0, jj + 1) - er78(0, jj) = 1 And er78(1, jj + 1) - er78(1, jj) = 1
                                    jj = jj + 1   ' for内のjjを増やす。
                                    If jj = pap(7, 0) Then Exit Do
                                Loop
                                'If qap <> jj Then MsgBox "横方向ベタ貼りです"
                            End If
                            DoEvents
                            Application.StatusBar = "横ベタ中、" & Str(jj) & " / " & Str(pap(7, 0)) & " 、 " & Str(Abs(qap - jj) + 1) & "列"
                
                            If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then '文字列時
                                er34 = "mm" '文字列指定(ーー)
                            ElseIf er(3, bni) > 0.2 And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then '通貨時
                                er34 = "pm"  '通貨指定（＋ー）
                            ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then
                                er34 = "mp"  'セル踏襲 "ap"→"mp（ー＋）"　遅・コピペパターン
                            Else
                                er34 = "pp"  '（＋＋）
                            End If
                            If Abs(er78(1, qap)) = 0.1 Then MsgBox "betat4挙動注意a" '86_019a
                            Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, er78(1, qap), rrr, er78(1, jj), bfn, shn, k, axa, er34, mr8(jj))
                        Next
                    End If
                    Exit Do
                End If
                Application.StatusBar = False
            End If  'p=-2ここまで

            '突合先にAMある時(p=1)の処理　n2：当シート処理行　qq：対象シート取込行
            If pkt = 1 Then
                If Round(er(6, bni)) = -2 Then
                    'c=-2の加算処理(ね、ひ)
                    'If er(7, bni) <> 0 And er(5, bni) >= 0 Then
                    '↓30s86_020n bni→1
                    If er(7, bni) <> 0 And er(5, 1) >= 0 Then
                        Call oshimai("", bfn, shn, 1, 0, "「-2」で他列操作をしようとしています。修正を。")
                    End If
                    '項準適用廃止30s64
                    If mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then
                        kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                        kahi = 1
                    ElseIf Round(er(9, bni)) <> 0 Then
                        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(9, bni))).Value <> "" Then
                            kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                            kahi = 1
                        End If
                    End If
                    If kahi = 1 Then bfshn.Cells(n2, a).Value = bfshn.Cells(n2, a).Value + kasan
                ElseIf er(8, bni) <> 0 And Not (er(5, bni) < 0 And er(6, bni) > 0) Then
                    '転載処理(連載型) 強制同列・(上書型) え、う　他列(上書型) さ、せ、ち、と (強制同列)
                    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) > 0) Then  '30s86_012w '特命加算型の方
                    '特命条件　特命時(加算型)はここはスルー
                    'MsgBox "iei"
                    Else  'これまではこちら↓　特命時(連載型)もこちら
                        If er(6, bni) = -5 Or er(6, bni) = -7 Then
                            bfshn.Cells(n2, a).Value = hunpan(mr(2, 1, bni), mr(1, 1, bni), qq, er(3, bni), er(4, bni), er(5, bni), er(6, bni), er(8, bni), mr(2, 6, bni), hk1)
                            If hk1 <> "" And er(7, bni) <> 0 Then bfshn.Cells(n2, a).Value = hk1
                        Else '特命条件(連載型)では↓こちらを通る。
                            Call tnsai(ct8, tst, ct3, er78(), a, n2, bni, pkt, qq, 0, mr(), er(), pap(7, 0), mr8()) '30s62一元化
                        End If
                    End If
                End If
                
                '加算処理(同列転載なら実行しない) か、え、き、あ
                If (er(10, bni) = 0 Or Abs(er(7, bni)) > 0.2) And (er(6, bni) > 0 Or Round(er(6, bni)) = -1) Then
                    If (er(5, bni) >= 0 And er(6, bni) < 0 And er(7, bni) <> 0) Or _
                    (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And Int(er(7, bni)) = 0 And er(8, bni) <> 0) Then
                        kasan = 1  '6行-1で他列転載時 or 特命条件時86_14s
                        kahi = 1
                    ElseIf mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then 'ｦｦ定数加算対応
                        '特命状況の加算型はここを通るケースある（既存p=1なら）→通らないへ86_14s
                        kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                        kahi = 1
                    ElseIf Round(er(9, bni)) > 0 Then
                        '特命状況の加算型はここを通るケースある（既存p=1なら）少→通らないへ86_14s
                        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = "" _
                            Or (mr(2, 9, bni) = "n0" And Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = 0) Then
                            '加算対象セルが空白or0なら加算処理実施しない
                        Else
                            '特命状況の加算型はここを通るケースある（既存p=1なら）多→通らないへ86_14s
                            kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                            kahi = 1
                        End If
                    End If
                    
                    If kahi = 1 Then
                        If er(2, 1) < 0 And n2 > h0 Then  '高速かつ新規の既存
                            hirt(n2 - h0 + 1, 5) = hirt(n2 - h0 + 1, 5) + kasan    '5列目処理
                            '(高速シート５列目使用、高速かつ新規の既存、（特命条件もこのケースあり）
                        Else
                            bfshn.Cells(n2, a).Value = bfshn.Cells(n2, a).Value + kasan
                            '特命条件でこのケースもある　最初から既存の時
                        End If
                    End If
                End If
                n1 = n2
            End If '(p=1ここまで)
            
            '突合先にAM無い時（追加）(p=2)　n2：当シート処理行　qq：対象シート取込行
            If pkt = 2 And er(6, bni) >= 0 Then
                '低速新規はちまちま1,2,5列追加へ（「ど」集計時の対応仕様）最初に　86_011q
                If mr(1, 11, 1) = "ノーマル" And bfshn.Cells(n2, 1).Value <> 1 Then
                    bfshn.Cells(n2, 1).Value = 1
                    If Not StrConv(Left(mr(0, 2, bni), 1), 8) = "*" Then '30s73
                        bfshn.Cells(n2, 2).Value = "c"
                    End If
                    bfshn.Cells(n2, 5).Value = Now() 'タイムスタンプ追記
                End If
            
                If bfshn.Cells(n2, Abs(er(2, bni))).Value <> "" Then
                    Call oshimai("", bfn, shn, 1, 0, "新規予定行に既に情報があります。確認を。")
                End If
                    
                If bfshn.Cells(n2 + 1, a).Value <> "" Then Call oshimai("", bfn, shn, 1, 0, "新規次行に既に情報があります。確認を。")

                If er(2, 1) < 0 Then  '85_007 高速時場合分け
                        hirt(n2 - h0 + 1, 4) = am2      '4列目処理
                Else  'ノーマル（複数列・単列）
                    If pap(2, 0) > 0 Then
                        '86_012j追加バグ対応
                        With bfshn.Cells(n2, Abs(er(2, bni)))  'キー追加（mghz)ノーマル複数列時
                            .NumberFormatLocal = "@"
                            .Value = am2
                        End With
                    Else  '
                        With bfshn.Cells(n2, Abs(er2(0)))  'キー追加（mghzのケースは発生しない。er(2,x)ではなくer(x)なので
                            .NumberFormatLocal = "@"
                            .Value = am2
                        End With
                    End If
                End If
                
                If pap(2, 0) > 0 Then   '２行目ヱ対応(複数列のとき) ※全部の複数列が対象　mghzではない
                    For jj = 0 To pap(2, 0) 'qap→jj
                        If Round(Abs(er2(jj))) > 0 Then  '30s70 0.4(null,ー)は無視とする(エラー防止のため)
                            With bfshn.Cells(n2, Abs(er2(jj)))
                                .NumberFormatLocal = "@"
                                .Value = rvsrz3(am2, jj + 1, mr(2, 4, bni), 0)
                            End With
                        End If
                    Next
                End If
             
                '転載処理部 連結型 あ、い、し、そ、つ、な(読点あり)強制同列・上書型 え、う （他列転載）・上書型 さ、せ 強制同列
                If er(8, bni) <> 0 And Not (er(5, bni) < 0 And (Round(er(6, bni)) = -2 Or er(6, bni) > 0)) Then
                    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) > 0) Then  '特命加算型の方
                        '特命条件追加30s86_012s　'特命時転載はしてはいけませんです。　'特命条件　特命時(加算型)はここはスルー
                    Else  'これまではこちら↓　特命時(連載型)はこちら　特命条件(連載型)では↓こちらを通る。
                        Call tnsai(ct8, tst, ct3, er78(), a, n2, bni, pkt, qq, 0, mr(), er(), pap(7, 0), mr8()) '30s62一元化
                    End If
                End If

                '加算処理部(同列転載なら実行しない) か、え、き,あ
                If (er(10, bni) = 0 Or er(7, bni) <> 0) And er(6, bni) > 0 Then  '特命条件対応版へ
                    '特命条件加算型はここでも処理される。　'項準適用廃止30s64
                    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And Int(er(7, bni)) = 0 And er(8, bni) <> 0) Then
                        '特命条件一律これで　30s86s
                        kahi = 1
                        kasan = 1
                    ElseIf mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then
                        kahi = 1
                        kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                    ElseIf Round(er(9, bni)) > 0 Then
                        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = "" Or _
                                (mr(2, 9, bni) = "n0" And _
                            Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = 0) Then
                        Else
                            kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                            kahi = 1
                        End If
                    End If
                    If kahi = 1 Then
                        If er(2, 1) < 0 Then
                            hirt(n2 - h0 + 1, 5) = kasan
                            '特命条件こっちっぽい　高速なので
                        Else
                            bfshn.Cells(n2, a).Value = kasan
                        End If
                    End If
                End If
                    
                If er(2, 1) > 0 Then h0 = n2 'h0も更新(高速時以外) ※一文節目で判断30s22
                '１列２列５列　ちまちま→文節最後にまとめてで
                h = h + 1
            
            End If '(p=2ここまで)
            
            If (pkt = 2 Or pkt = 1) Then  '-1-2新規以外の時 新設
                am1 = am2
                n1 = n2
            End If
        End If 'カウント対象時行うここまで

        ii = ii + 1
        If er(2, 1) < 0 Then
            Call hdrst2(ii, a, 10000, k0, h0)
        Else
            Call hdrst2(ii, a, 1000, k0, h0) '201904　100→1000
        End If
    Loop '◇ここから行毎ここまで
  
    'ここでのqq・・対象シートで最終でカウント対象とした行
    bfshn.Cells(2, 4).Value = k0 'Loop終了後行う30s86_002より
    bfshn.Cells(3, 4).Value = h0 'Loop終了後行う30s86_002より
   
    If mr(1, 11, 1) = "純高速" Or mr(1, 11, 1) = "近似高速" Then Erase hiru
    
    If mr(1, 11, 1) = "純高速" Then '高速シートAーE列入れ込み戻し、当列、突合列へベタ
        Range(twt.Cells(1, 1), twt.Cells(rrr + 1, 5)).Value = hirt  '※項無対策(rrr→rrr+1)
        Erase hirt

        'ノーマル同様、先に1,2,5列処理へ
        bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
    
        '一列２列５列新規増分分のまとめて埋め込み　高速時のみに(低速は↑で実施済み)
        If bfshn.Cells(1, 4).Value - 1 + k - 1 < h Then  '「ど」考慮型
            '２列目（ｃ）こちらが先
            If Not StrConv(Left(mr(0, 2, bni), 1), 8) = "*" Then 'ｐ＝ー２にも2行＊概念反映へ
                Call betat4(bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 0.4, h, 0.4, bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 2, "pp", "c")
            End If
            '5列目（now）
            Call cpp2("", Now(), 0, 0, 0, -1, bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 5, h, 5, -4163) '5列目（now）これに凝縮　-1→-2
            '一列目（１）は最後で
            Call betat4(bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 0.4, h, 0.4, bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 1, "pp", "1")
            DoEvents
        End If
        
        If h0 < h Then  '新規あるときのみ ←h1→h0へ。準高速onlyなので　And pap(2,0)撤廃
            '突合列新規ベタ(単数列時のみ)
            If pap(2, 0) = 0 Then
                Call betat4(twn, "高速シート_" & syutoku(), 2, 4, h - h0 + 1, 4, bfn, shn, h0 + 1, Abs(er2(0)), "mm", "")
            End If
            
            '当列新規ベタ（加算時のみ）※同列転載なら実行しない)
            '↓特命対応
            If (er(10, bni) = 0 Or er(7, bni) <> 0) And er(6, bni) > 0 Then
                '当列控え
                Call betat4(twn, "高速シート_" & syutoku(), 2, 5, h - h0 + 1, 5, bfn, shn, h0 + 1, a, "pm", "")
                '※この後、高速シート５列目は、↓のstrconv24用途で使用
            End If
            
            'mghz列新規ベタ 30s86_011
            Application.Calculation = xlCalculationAutomatic    '数式計算方法自動に　'新形式　85_007
            DoEvents '86_019r

            Application.Calculation = xlCalculationManual  '再計算再び手動に（重くなるため）30s66
            
        End If
    End If
    
    cnt = 0
    Call hdrst(ii, a)   'exitdoを考慮し、ここにも
    bfshn.Cells(1, a).Value = cted(0)
    '↓不要かもだが、念のため
    bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1

    'サマリ値処理（転載分）p=-2は、転載も実施
    If er(5, bni) >= 0 And er(6, bni) >= 0 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then
        '再計算を一旦自動に戻す
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        If pap(7, 0) > pap(8, 0) Then ii = pap(8, 0) Else ii = pap(7, 0)
        'もともとp=-2の転載処理
        jj = ii  'jjは最終節
        
        bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
        
        For ii = 0 To jj 'pap(7,0),pap(8,0)＝0の時は一回ポッキリ実施
            Application.StatusBar = "サマル中、" & Str(ii) & " / " & Str(jj)
            If er78(0, ii) > 0.2 And er78(0, ii) <> a Then  '86_013f 7行0.4対処→0.1対処
                Call samaru(Int(er78(0, ii)), mr(1, 1, 1)) '複数列さまる
            End If
        Next
        '再計算を再び手動に
        Application.Calculation = xlCalculationManual
        Application.DisplayStatusBar = True
    End If
    
    Next
'◇Ｄ◇複文節ふい（実行編）ここまで↑

'★★＃papchk予定地＃
Call papchk(pap(), 0, bun)

'◆Ｅ◆単文節ふい（事後編）ここから↓
    bni = 1 '1文節目の値で判断、実行 86_020p　bun→1戻す。一括踏襲のところは、bni→bunへ置換
    mg2 = 0 'mg2の変数はとりあえずpublicで入れている。整理後ほど

    'ここに -1-2 の空白埋め込みルーチン　複数列対応30s61_4
    If er(6, bni) = -1 Or er(6, bni) = -2 Or er(6, bni) <= -3 Then
        If nifuku = 1 Then
            'MsgBox "2文節目での複文節あり。注意を。"  '解禁へ　86_016v
            Call oshimai("", bfn, shn, 1, 0, "2文節目複文節での指定は許容されていません")
        End If
      
      'c<=-1から変更(C<=-3 の空白埋め込みは近似高速時のみやるに復活へ　85_014
        DoEvents
        cnt = 0
        k0 = k  'k0一旦リセット(高速時、またここから増えていく)
        
        If mr(1, 11, bni) = "近似高速" Then  '近似型
            If er(6, bni) = Round(er(6, bni), 0) Then '-15.1→複写対応しない
                mg2 = 1
                For ii = h To k Step -1  'i→ii
                    Do Until bfshn.Cells(ii, mghz).Value <> ""  '使われていない？
                        ii = ii - 1
                    Loop
                    '次行(前行)一致＆aに情報あり　'↓こちらへ(strconv「2」被せる)
                    If StrConv(bfshn.Cells(ii, mghz).Value, 2) = StrConv(bfshn.Cells(ii + 1, mghz).Value, 2) Then
                        If bfshn.Cells(bfshn.Cells(ii + 1, mghz + 1).Value, a).Value <> "" Then
                            Call tnsai(ct8, tst, ct3, er78(), a, bfshn.Cells(ii, mghz + 1).Value, bni, 1, 0, bfshn.Cells(ii + 1, mghz + 1).Value, mr(), er(), pap(7, 0), mr8())
                        End If
                    End If
                    Call hdrst(h - ii, a)  '左下ステータス表示部
                Next
                cnt = 0
            End If
        ElseIf er(2, bni) < 0 Then
            Call oshimai("", bfn, shn, sr(2), a, "旧高速使用終了です。")
        Else 'ノーマル
            If er(6, bni) = -1 Or er(6, bni) = -2 Then
                For ii = k + 1 To h  'i→ii
                    Do Until bfshn.Cells(ii, Abs(a)).Value = NullString
                        ii = ii + 1
                    Loop
                    If bfshn.Cells(ii, Abs(er(2, bni))).Value = bfshn.Cells(ii - 1, Abs(er(2, bni))).Value Then '前行と同じ場合
                        If bfshn.Cells(ii - 1, a).Value <> "" Then '-1他列転載時、これまでのようにmatchてっぺんに必ず情報があるとは限らなくなったので、その対処）
                            Call tnsai(ct8, tst, ct3, er78(), a, ii, bni, 1, 0, ii - 1, mr(), er(), pap(7, 0), mr8()) '30s81qq無効化
                        End If
                    Else
                        '高速時は↓以下通らない(キーが昇順の前提であるため。必ずエラーになる)。
                        If Not IsError(Application.Match(bfshn.Cells(ii, Abs(er(2, bni))).Value, Range(bfshn.Cells(k0, Abs(er(2, bni))), bfshn.Cells(ii - 1, Abs(er(2, bni)))), 0)) Then  'match使用
                            m = Application.WorksheetFunction.Match(bfshn.Cells(ii, Abs(er(2, bni))).Value, Range(bfshn.Cells(k0, Abs(er(2, bni))), bfshn.Cells(ii - 1, Abs(er(2, bni)))), 0) 'h→ii-1に修正30s24
                            If bfshn.Cells(k0 + m - 1, a).Value <> "" Then  '-1他列転載時、これまでのようにmatchてっぺんに必ず情報があるとは限らなくなったので、その対処
                                Call tnsai(ct8, tst, ct3, er78(), a, ii, bni, 1, 0, k0 + m - 1, mr(), er(), pap(7, 0), mr8()) '30s63集約化
                            End If
                        End If
                    End If
                    Call hdrst(ii, a)    '左下ステータス表示部
                    bfshn.Cells(2, 4).Value = k0 'Cells(13, 3)→Cells(2, 4)
                Next
                cnt = 0
            End If
        End If
    End If
    cnt = 0

    '面取り
    '↓連載型特命条件も追加へ　86_013d
    If ((er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) < 0) Or _
        (er(9, bni) = 0 And er(10, bni) < 0)) And _
        Not ((er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1)) Then '(当列)最後の「、」を取る。
        
        Application.Cursor = xlWait '85_026
      
        If h >= k Then 'データ無の時は実施しないへ　86_018f
            hirt = Range(bfshn.Cells(k, a), bfshn.Cells(h + 1, a)).Value 'h→h+1 variant で配列、ただしセルが1x1だけの時、配列として認識してくれない対策　30s86_017v
            For ii = k To h
                'hirtでは、ｋ→１、ｈ→ｈ?ｋ＋１　、forのiiに、-k+1をかぶせる　ii→ii-k+1
                '↓86_017m(連載複数文字対応) ↓86_017v(修正・区切り文字は複数列一番右→一番左に仕様変更へ（この方が自然）
                If Right(hirt(ii - k + 1, 1), Len(mr8(0))) = mr8(0) Then hirt(ii - k + 1, 1) = Left(hirt(ii - k + 1, 1), Len(hirt(ii - k + 1, 1)) - Len(mr8(0)))
                If hirt(ii - k + 1, 1) = "" Then hirt(ii - k + 1, 1) = NullString '　""を空白に（カウントさせないため）←ここはケガの功名にならなかった。
                Call hdrst(ii, a)   '左下ステータス表示部
            Next
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).Value = hirt '新バ
            Erase hirt '新バ
        End If
        Application.Cursor = xlDefault
    End If
    cnt = 0
    
    '結合
    If (er(5, bni) < 0) And er(6, bni) <= 0 And Round(er(6, bni)) <> -2 And er(7, bni) < 0 And Not IsNumeric(bfshn.Cells(sr(7), a).Value) Then  '2s
        For ii = k To h  'i→ii
            bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value & bfshn.Cells(ii, Abs(er(7, bni))).Value '文字結合
        Next
        Call hdrst(ii, a)    '左下ステータス表示部
    End If
        
    '差分処理(er(5,0)<0)
    If (er(5, bni) < 0) Then  '2s
        cnt = 0
        axa = 0
        dif = 1
        If IsNumeric(bfshn.Cells(sr(7), a).Value) Then
            axa = a + er(7, bni) '数値：位置が相対値
        Else
            axa = Abs(er(7, bni)) '文字：位置が絶対値
            If er(7, bni) < 0 Then dif = -1 '和・結合実施フラグ
        End If
        bfshn.Cells(sr(0) + 3, 3).Value = ""

        Application.Cursor = xlWait
        For ii = k To h
            saemp = 0
            If pap(7, 0) <> 0 Or (bfshn.Cells(ii, a).Value <> "" Or bfshn.Cells(ii, axa).Value <> "") Then '両方空欄なら実施しない(：単列なな以外実施)
                If IsNumeric(bfshn.Cells(ii, axa).Value) Or IsDate(bfshn.Cells(ii, axa).Value) Then
                    saemp = bfshn.Cells(ii, axa).Value
                End If
                '文字比較(新対応)※従来型は下部elseifへ
                If er(6, bni) <= 0 And Round(er(6, bni)) <> -2 And dif = 1 Then
                    zzz = bfshn.Cells(ii, a).Value '既存のデータ箱
                    ReDim zxxz(pap(7, 0)) 'pap(7,0)=8の前提（複数差分比較時）
                    ReDim xxxx(pap(7, 0)) 'pap(7,0)=8の前提（複数差分比較時）
                    ReDim zyyz(pap(7, 0)) 'pap(7,0)=8の前提（複数差分比較時）20190614、013q
                    
                    Range(Cells(ii, a), Cells(ii, a)).ClearContents  '一旦空白に　86_020bからこの記述
                        
                    zyyz() = Split(zzz, mr(2, 4, bni))  'Split関数で対象シート側個々要素を一気に入れる
                    
                    For jj = 0 To pap(7, 0) '比較列を個々見ていく ii→jjへ
                        xxxx(jj) = bfshn.Cells(ii, er78(0, jj)).Value '当列側比較対象列
                        'kahi(各々のセル比較結果の策定)
                        If zzz = "" Then
                            kahi = -1
                        '↓ここでエラー
                        ElseIf zyyz(jj) = "" And xxxx(jj) = "" Then 'zyz→zyyz(jj)
                            kahi = 2  'なしなし
                        ElseIf zyyz(jj) = xxxx(jj) Then
                            kahi = 0
                        Else
                            kahi = 1
                        End If
                            
                        If mr(2, 6, bni) = "-1" Then 'op:1　差分は具体的差分表記を羅列パターン )1→-1に変更
                            If kahi <> -1 Then  '-1は羅列すらしない
                                If kahi = 1 Then '差分情報付記
                                    zxxz(jj) = zyyz(jj)  'セル転記は最後↓、join関数にて
                                Else  '情報掲載せず（kahi=0,2)
                                End If
                            End If
                        Else 'op:2、無、c、cc（差分は数値フラグ表記パターン）
                            If mr(2, 6, bni) = "-2" Then  '数値フラグを羅列で表記　※セル転記は最後↓、join関数にて　2→-2に変更
                                '数値フラグ設定
                                If kahi <> 2 Then
                                    zxxz(jj) = CStr(kahi) '=LTrim(Str(kahi)) -1,0,1が刻まれる
                                    '挿入箇所(ヒヅケセルあたりのコメント）
                                End If
                            Else 'op:無、c、cc　 '数値フラグを合算で表記
                                If kahi <> 2 Then  '比較結果２→合算操作自体を行わない
                                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value + kahi  'セルに結果記載（合計値）
                                        
                                    'differentセルの色付け
                                    If kahi = 1 And Left(mr(2, 6, bni), 1) = "c" Then '色付け(c,cc)
                                        bfshn.Cells(ii, a).Interior.Color = RGB(254, 254, 238) '← = 15662846　→　30s86_017p
                                        bfshn.Cells(ii, er78(0, jj)).Interior.Color = RGB(254, 254, 238) '← = 15662846　30s86_017p
                                        bfshn.Cells(ii, er78(0, jj)).ClearComments
                                        If mr(2, 6, bni) = "cc" Then 'さらにコメント付加（比較元側）(cc)処理重い
                                            bfshn.Cells(ii, er78(0, jj)).AddComment
                                            bfshn.Cells(ii, er78(0, jj)).Comment.Text Text:=zyyz(jj) 'zyz→zyyz(jj)
                                            bfshn.Cells(ii, er78(0, jj)).Comment.Shape.TextFrame.AutoSize = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    '羅列型(１か２)はここで記載↓join関数使用
                    If ((mr(2, 6, bni) = "-1" And zzz <> "") Or mr(2, 6, bni) = "-2") Then
                        bfshn.Cells(ii, a).Value = Join(zxxz, mr(2, 4, bni))
                    End If
                    '文字比較(新対応)201708追加ここまで
                ElseIf er(6, bni) <= 0 And Round(er(6, bni)) <> -2 And dif = -1 Then
                    '文字結合→移行によりここでは何もしないことに
                ElseIf dif = -1 And er(8, bni) >= 0 Then  '1m =0→>=0 30s69バク修正
                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value + saemp '和
                ElseIf dif = -1 And er(8, bni) < 0 Then  '<-0.5→<0　30s69バク修正
                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value * saemp '積
                ElseIf er(8, bni) >= 0 Then 'e=0→e>=0修正16ｓ
                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value - saemp '差
                ElseIf er(8, bni) < 0 Then  '1m　<-0.5→<0　30s69バク修正
                    bfshn.Cells(ii, a).Value = saemp - bfshn.Cells(ii, a).Value '差(反転)
                End If
            End If
            '比較の移行先はここ←？
            Call hdrst(ii, a)              '左下ステータス表示部
        Next  '差分のある一行処理ルーチンはここまで
        
        Application.Cursor = xlDefault
        cnt = 0
    End If
  
    DoEvents
    '再計算を自動に戻す
    Application.Calculation = xlCalculationAutomatic
  
'MsgBox "hi"
    '突合列さまる(-99、＊では実施しないので、こちらに）
    For ii = 0 To pap(2, 0)  '30s86_022qhs
        'If StrConv(Left(mr(1, 1, 1), 1), 8) <> "*" Then
        If Int(Abs(er2(ii))) <> 0 Then
            Call samaru(Int(Abs(er2(ii))), mr(1, 1, 1))   '当シート突合列（２行目）さまる
        End If
        'End If
        
    Next
  
    '再計算を手動に戻す
    Application.Calculation = xlCalculationManual
  
  End If 'るA　-99or*,**は通過、ここまで
  
' 86_020p　戻す。一括踏襲のところは、bni→bunへ置換
  '終了時の文字調整(強制orちまちま→不要、通常の加算や転載or突合元情報の最終行の書式に合わせる時のみ要、)
  If er(6, bun) >= 0 Or (er(6, bun) < 0 And tst = 7 And trt <> -9) Then 'c<0のvlookup型も実施しないに(tst7,-99以外)　86_013d
        If tst = -2 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        ElseIf tst = 0 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "G/標準"
        ElseIf tst = 7 Then
            If trt = -2 Then  '加算型の一括踏襲
                If Abs(er(9, bun)) = 0.4 Or Abs(er(9, bun)) = 0.1 Then
                    MsgBox "加算固定値(0.4、0.1ー)なので一括踏襲は行われません。"
                Else
                    If qq > 0 Then
                    Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = Workbooks(mr(2, 1, bun)).Sheets(mr(1, 1, bun)).Cells(qq, Abs(er(9, bun))).NumberFormatLocal
                    Else
                    MsgBox "一括踏襲はスルーします(情報無いので)"  '86_016z
                    End If
                End If
            ElseIf trt = -1 Then   '転載型の一括踏襲
                If er(10, bun) < 0.5 Then
                    MsgBox "連載、ｦｦ○(：転載情報書式が存在しない) は対象外"
                Else
                    If qq > 0 Then
                    Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = Workbooks(mr(2, 1, bun)).Sheets(mr(1, 1, bun)).Cells(qq, Abs(er(10, bun))).NumberFormatLocal
                    Else
                    MsgBox "一括踏襲はスルーとなります(情報無いので)""  '86_016z"
                    End If
                End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "こういうケースがあるのだろうか？")
            End If
        End If

        '折り返して全体を表示しない (LF対処)16s　一律適用へ18s
        If trt <> -2 Then Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).WrapText = False
  End If
'◇Ｅ◇ここまで(not98)
        
        
'-99､-98用ここから
  'If er(6, bni) <= -90 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then 'るC　20190214　*ありは実施しないに
    '↓30s86_022qhs
  If er(6, bni) <= -90 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" And h >= k Then 'るC　20190214　*ありは実施しないに
    
        '独立集計(ｈ変更)
        'こちらの「ど」の対応はどうするか要検討
        If InStr(1, mr(0, 2, bni), "ｦ") > 0 And InStr(1, rvsrz3(mr(0, 2, bni), 1, "ｦ", 0), "ど") > 0 Then
            ii = h
            Do Until ii = k - 1  'ケツから上方向へサーチ
                If bfshn.Cells(ii, Abs(er(2, 1))) <> "" Then Exit Do  'bfshn.を被せた
                ii = ii - 1
            Loop
            h = ii
        End If
                    
        DoEvents
        Application.Calculation = xlCalculationAutomatic    '数式計算方法自動に　'新形式　85_007

        Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "G/標準"    'mghz列一旦標準へ（数式入れるため）
    

    
    
        '数式でコピペ
        If mr(2, 6, bni) = "c" Then '85_024_まだ使用されていない。
            MsgBox "使われている？"
            bfshn.Cells(sr(8), a).Copy  'ほぼ避けられない
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).PasteSpecial Paste:=xlPasteAll 'すべて -4104
        Else '従来型 -4123→xlPasteFormulas→.copy.paste
            Call copipe(bfn, shn, sr(8), a, sr(8), a, bfn, shn, k, a, h, a, 3)  '3→FormulaR1C1(脱.copy.paste)　86_020d
        End If
    
        '30s86_021cこちらへ
        If tst = 1 Then '←当列転載文字型
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "@"
        ElseIf tst = -2 Then  '通貨型
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        ElseIf tst = 7 Then '対象シート名を模倣　a11_1列→対象シートに変更
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = bfshn.Cells(sr(1), a).NumberFormatLocal 'sr(4)→sr(1)へ30s68
        '↓ここがREMな理由：日付とかが、実数になってしまう
       'ElseIf tst = 0 Then '標準型
       'Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "G/標準"
        End If
    
        If er(6, bni) = -99 Then 'セルを値に変換
            Call cpp2(bfn, shn, k, a, h, a, bfn, shn, k, a, 0, 0, -4163) '何と""→Nullstringされる(ケガの功名、浄化作用)
        End If


        bfshn.Cells(sr(8), 4).Value = "(-99系)"  '86_019q
        bfshn.Cells(sr(8), a).ClearComments   '86_019q　既存コメントは消していく仕様へ
        bfshn.Cells(1, a).Value = h  'h0→h 30s64
    
        DoEvents
        Application.Calculation = xlCalculationManual  '再計算再び手動に（重くなるため）30s66
    
        cnt = 0
        If er(5, bni) <> 1 Then  'カウント基準列が空白行はこちらも空白に 30s68
            Application.Cursor = xlWait
            For ii = k To h
                If kaunta(mr(), ii, pap(5, 0), bni, er5(), mr5()) = 0 Then '30s85_024検証6
                    'bfshn.Cells(ii, a).Value = bfshn.Cells(sr(0) + 2, 1).Value
                    '↓30s86_020b
                    Range(Cells(ii, a), Cells(ii, a)).ClearContents
                End If
                Call hdrst(ii, a)            '左下ステータス表示部
            Next
            Application.Cursor = xlDefault
        End If
    
        cnt = 0
        If er(7, bni) = 0 Then '30s84_617
            Application.Cursor = xlWait
            yayuyo = 0
            For ii = k To h  'h
                If IsError(bfshn.Cells(ii, a)) Then '30s68_2バグ改良
                    yayuyo = yayuyo + 1 '85_005
                    If yayuyo = 10 Then Call oshimai("", bfn, shn, k, a, "やゆよ10回")
                ElseIf bfshn.Cells(ii, a) = "" Then
                    Range(Cells(ii, a), Cells(ii, a)).ClearContents    '""→空白　30s86_020b　NullStringからこちらへ
                End If
                Call hdrst(ii, a)            '左下ステータス表示部
            Next
        End If
        Application.Cursor = xlDefault
  '-99用ここまで

'  Else
'
'    '突合列さまる(-99では実施しないので、こちらに）
'    For ii = 0 To pap(2, 0)  '30s86_022qhs
'        If StrConv(Left(mr(1, 1, 1), 1), 8) <> "*" Then
'        If Int(Abs(er2(ii))) <> 0 Then
'            Call samaru(Int(Abs(er2(ii))), mr(1, 1, 1))   '当シート突合列（２行目）さまる
'        End If
'        End If
'        MsgBox "hi"
'    Next
  
  End If  'るC
    
    
    
    Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).WrapText = False  '86_016t（折り返ししないAPIのXML膨張防止）こっちに
    
  '左下ステータス表示部　ここだけ特殊バージョン
  DoEvents
  If flag = True Then Call oshimai("", bfn, shn, k, a, "中止しました")    '中止ボタン処理
  Application.StatusBar = Str(cnt) & "、" & Str(a - dd1 + 1) & " / " & Str(dd2 - dd1 + 1)
  cnt = 0
 
End If 'り(not"*")
     
    '再計算を自動に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
     
    'サマリ値処理（転載分）ここでも(-1,-2のみ)
    If er(5, bni) >= 0 And er(6, bni) > -90 And er(6, bni) < 0 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then
        If pap(7, 0) > pap(8, 0) Then jj = pap(8, 0) Else jj = pap(7, 0) 'ii→jj
        
        bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
        
        For ii = jj To 0 Step -1
            If er78(0, ii) > 0.2 And er78(0, ii) <> a Then Call samaru(Int(er78(0, ii)), mr(1, 1, 1))  '複数列さまる
        Next
    End If
    DoEvents
    bfshn.Cells(sr(0) + 4, a).Select
    
    Call samaru(a, mr(1, 1, 1))  '当列さまる



    bfshn.Cells(sr(0) + 3, a).Value = Now() '活用例1タイムスタンプ入れる
    bfshn.Cells(sr(0) + 2, 4).Value = Now() '活用例1タイムスタンプ入れる sr(0) + 3,→sr(0) + 2　３０ｓ７４
    '↓一行毎にインクリメント（ID変わったときのみリセットされる） 。※bfn、shn側は更新されない（初回複写時の値が載ってるだけ）。
    twbsh.Cells(14, 3).Value = twbsh.Cells(14, 3).Value + 1 '7s (3,3)→(14,3)25s
Next '選択範囲列分の繰り返し　ら
'ここでの「a」は、d+1である。
Call giktzg(a, rog)  'ログ部別プロシージャー化_201905
   
End Sub '外部結合ここまで
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub kskst(pap7 As Long, h As Long, er78() As Currency, er9() As Currency, mr9() As String, er3() As Currency, mr3() As String, er5() As Long, mr5() As String, c5 As Long, pap3 As Long, pap5 As Long, bni As Long, qq As Long, rrr As Long, mr() As String, er() As Currency, a As Long, cted() As Long) '高速シートは徐々にこちらへ
    Dim hirt As Variant, ii As Long, tempo As String, baba As String
    ThisWorkbook.Activate
    Sheets("高速シート_" & syutoku()).Select
    DoEvents
        
    twt.Cells.Clear
    twt.Cells.Delete Shift:=xlUp
    DoEvents
    twt.Columns("A:A").NumberFormatLocal = "@"  '一列目文字列に
    twt.Columns("G:G").NumberFormatLocal = "@"  '7列目(旧4列目)文字列に（転載前キー列用途）　30s85_027
    twt.Columns("F:F").NumberFormatLocal = "@"  '6列目も文字列に（転載前キー列用途）　30s86_014g
    twt.Columns("D:D").NumberFormatLocal = "@"  '４列目文字列に（新規分情報a※用途のみ）へ　30s85_027
'    twt.Columns("K:V").NumberFormatLocal = "@"  '86_020h
        
    rrr = qq

    If Abs(er(4, bni)) >= 1 Then  '最終行と抜け有無チェック(単列のみ)
        'ここでのhirtは対象シートの「ALL一列」の1列・・seekの最後検知用
        hirt = Range(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(1, Abs(er(4, bni))), Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(cted(0) + 1, Abs(er(4, bni)))).Value
        Do Until hirt(rrr, 1) = ""
            rrr = rrr + 1
            Call hdrst2(rrr, a, 10000, 0, 0)
        Loop
        Erase hirt
    Else
        rrr = cted(0) + 1  '項準ｂ対応　30d85_018
    End If
    DoEvents
    rrr = rrr - 1 'rrrは対象シートの最終行　qqは対象シートのベタ貼り開始行。 対象シートのデータ部がすっからかんのとき、rrr<qqとなってしまう
    cnt = 0
    
    If rrr < qq Then  '対象シートデータ部がすっからかんな時
        MsgBox "対象シートのデータ部がすっからかんです。動作は続きます"
    Else '対象シートデータ部がすっからかんでない時の処理ここから(通常はこちら)
        '2列目(行番号)の処理(フィル活用)　ベタ・ちま共用
        Call betat4(twn, "高速シート_" & syutoku(), qq, 0.1, rrr, 0.1, twn, "高速シート_" & syutoku(), qq, 2, "pp", "000_")
        
        '3列目(カウント列)に入れ込む処理◆常時入ることに
        If Abs(er(5, bni)) <> 0 Then  'カウント列の情報(複数時は一番左)→3列目へ
            '86_012r：20190515★大改造
            If pap5 = 0 And mr(2, 5, bni) = "" Then  '以前は一律この仕様↓
                Call cpp2(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er(5, bni)), rrr, Abs(er(5, bni)), twn, "高速シート_" & syutoku(), qq, 3, 0, 0, -4163) '85_027検証3
            Else '追加仕様がこちら　86_012r
                cnt = 0
                For ii = qq To rrr
                    c5 = kaunta(mr(), ii, pap5, bni, er5(), mr5())
                    If c5 = 1 Then  'カウント対象なら実施
                        If pap5 = 0 Then 'カウンタ単数列・・カウンタセルをコピペ
                            Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(ii, 3).Value = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(ii, Abs(er(5, bni)))
                        Else  'カウンタ複数列・・・1を入れる仕様
                            Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(ii, 3).Value = 1
                        End If
                    End If
                    Call hdrst2(ii, a, 1000, 0, 0)
                Next
            End If
        Else   'カウント列ゼロ時・・・高速3列に入れ込む仕様に(これまで入れて無かった)
            '1を入れ込んでる。
            Call betat4(twn, "高速シート_" & syutoku(), qq, 0.4, rrr, 0.4, twn, "高速シート_" & syutoku(), qq, 3, "pp", "1")
        End If
        
        '一列目作成(pap3単数時・複数時場合分け)
        If pap3 = 0 Then '単数列時（ pap3=0 ）：ベタ貼り仕様

            'こっちは高速シート演算使ってない。 hiro→hirt
            hirt = Range(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er3(0))), Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(rrr + 1, Abs(er3(0))))
            For ii = 1 To 1 + rrr - qq
                If hirt(ii, 1) = "" Then
                    hirt(ii, 1) = "ー(情報空白行)ー" '013tこちらへ
                Else
                    hirt(ii, 1) = hirt(ii, 1) & ""      '数値→文字列とさせる技　86_012
                End If
            Next
            Range(Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(qq, 1), Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(rrr, 1)).Value = hirt
            '↑hirtの一番下(ダミー行)は無視されるだけ
            Erase hirt
            Application.Cursor = xlWait
            Application.Cursor = xlDefault
            
        Else  '複数列時（ pap3>0 ） babaの暫定変数は後々恒久措置検討(既存使い回せないか？)。
                'ｍｍ：文字列型、pm：通貨型、
            
            
            If er3(0) <> 0.1 Then '（６→11列目に転載）第1列でのー(：行番号転載)は実施しない。
                Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er3(0)), rrr, Abs(er3(0)), twn, "高速シート_" & syutoku(), qq, 11, "mm", mr3(0))
            
            ElseIf er3(0) = 0.1 And mr3(0) <> "" Then '86_022g 第1列でー(：行番号転載)で、固有転載あり(定文字対応） のとき
                Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(0.4), rrr, Abs(0.4), twn, "高速シート_" & syutoku(), qq, 11, "mm", mr3(0))
            
            
            End If
            
            For ii = 1 To pap3 '（７→１２列目以降に転載）第2列からはこちら　行番号転載もあればする。
                If er3(ii) = 0.1 And mr3(ii) <> "" Then MsgBox "betat4挙動注意b" '30s86_019m
                Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, er3(ii), rrr, er3(ii), twn, "高速シート_" & syutoku(), qq, 11 + ii, "mm", mr3(ii))
            Next

'            tempo = "R[0]C[1]"
            
            tempo = "CLEAN(R[0]C[1])"  '30s86_022i
            
            baba = tempo
            For ii = 1 To pap3 '数式の素生成
'                tempo = "&" & """" & mr(2, 4, bni) & """" & "&" & "R[0]C[" & LTrim(Str(ii + 1)) & "]"
                
                tempo = "&" & """" & mr(2, 4, bni) & """" & "&" & "CLEAN(R[0]C[" & LTrim(Str(ii + 1)) & "])" '30s86_022i
                
                baba = baba & tempo
            Next
                
            Application.Calculation = xlCalculationAutomatic   '数式計算方法自動に(次行演算の為)
            twt.Cells(qq, 10).FormulaR1C1 = "=" & baba  '10列目1行目に数式の素を注入

            Call copipe(twn, "高速シート_" & syutoku(), qq, 10, qq, 10, twn, "高速シート_" & syutoku(), qq + 1, 10, rrr, 10, 3) '3→FormulaR1C1 '86_019s
        
            '↓ようやく10列目を１列目にコピー（数式→値化）
            Call betat4(twn, "高速シート_" & syutoku(), qq, 10, rrr, 10, twn, "高速シート_" & syutoku(), qq, 1, "mm", "")
            
            twt.Columns("J:K").ClearContents  '10列11列クリア(12列以降は特にクリアはしてない)
        End If  '一列目作成ここまで
                   
 '複文節工事予定ここまで　86_018m
        
        '1列目→11列目(カタカナ、半角化)
        
        twt.Columns("K:K").NumberFormatLocal = "G/標準"
        If hrkt = 16 Then '30s86_020s
'            baba = "ASC(PHONETIC(R[0]C[-10]))"
            baba = "CLEAN(ASC(PHONETIC(R[0]C[-10])))"   '86_022i
        Else
'            baba = "ASC(R[0]C[-10])"
            baba = "CLEAN(ASC(R[0]C[-10]))"    '86_022i
        End If
        
        twt.Cells(qq, 11).FormulaR1C1 = "=" & baba  '10列目1行目に数式の素を注入
        Call copipe(twn, "高速シート_" & syutoku(), qq, 11, qq, 11, twn, "高速シート_" & syutoku(), qq + 1, 11, rrr, 11, 3) '3→FormulaR1C1
        twt.Columns("J:J").NumberFormatLocal = "@"
            
        '↓11列目を１0列目にコピー（数式→値化）
        Call betat4(twn, "高速シート_" & syutoku(), qq, 11, rrr, 11, twn, "高速シート_" & syutoku(), qq, 10, "mm", "")
        twt.Columns("K:K").ClearContents  '11列クリア
        
        'ヴ→ｳﾞ(・・ASC(PHONETIC で処理されないので。strconv24では、ヴ→ｳﾞされるので、ここまでやって初めてASC(PHONETICとstrconv24が等価になる。
        Range(Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(qq, 10), Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(rrr, 10)).Replace What:="ヴ", _
            Replacement:="ｳﾞ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False

        Application.Calculation = xlCalculationManual  '再計算再び手動に
        
        '特命条件↓　7列(旧4列)ベタ・事前ソート入れる。
        If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0) Then
            '8行目(転載列・最初数列)を7列目にベタっと転載
            Call cpp2(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er(8, bni)), rrr, Abs(er(8, bni)), twn, "高速シート_" & syutoku(), qq, 7, 0, 0, -4163)
            '8行目(転載列・最後数列)を6列目にベタっと転載
            Call cpp2(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er78(1, UBound(er78(), 2))), rrr, Abs(er78(1, UBound(er78(), 2))), twn, "高速シート_" & syutoku(), qq, 6, 0, 0, -4163)
            
            '加算列の転載(9列)　86_013d↓
            If er9(0) <> 0.1 Then
                Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er9(0)), rrr, Abs(er9(0)), twn, "高速シート_" & syutoku(), qq, 9, "pp", mr9(0))
            Else '６行ーなら、こっち(30s86_019o　運用終了)
                Call oshimai("", bfn, shn, sr(6), a, "６行ーの運用は終了しました。")
                'Call betat4(twn, "高速シート_" & syutoku(), qq, 0.4, rrr, 0.4, twn, "高速シート_" & syutoku(), qq, 9, "pp", "1")
                'MsgBox "6行ーです。"
            End If
            Call saato(twn, "高速シート_" & syutoku(), 3, qq, 1, rrr, 10, 99) '3列目降ソート　99は降順の意 1列?10範囲で3列目を、
            Call saato(twn, "高速シート_" & syutoku(), 7, qq, 1, rrr, 10, 1) '7列目(旧4列目)昇ソート
            Call saato(twn, "高速シート_" & syutoku(), 10, qq, 1, rrr, 10, 1) '昇ソート 1→10列目(Excel2019対策・平・片同一視されない)
            
            ii = rrr
            cnt = 0
            
            '3列目消込プログラム　　　'rrr→rrr+1(空白コピペ用) 86_013v
            hirt = Range(Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(qq, 1), Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(rrr + 1, 10)).Value

            Do Until ii = qq  '  30s86_012s　qq行(データ開始行)は以下の操作やらない。qq+1行までが対象
                '1列目は常に情報がある。
                If hirt(-qq + 1 + ii, 3) <> "" Then '3列目情報ありなら、以下 hirt へ
                    If hirt(-qq + 1 + ii, 7) = "" Then '3列情報有,7列空なら、3列空白(・)、8列何もしない
                        hirt(-qq + 1 + ii, 3) = hirt(rrr - qq + 2, 3) '←セル空白化
                    Else  '3列＆4列共に情報あり　8、9
                        If hirt(-qq + 1 + ii, 8) = "" And hirt(-qq + 1 + ii, 9) <> "" Then hirt(-qq + 1 + ii, 8) = hirt(-qq + 1 + ii, 9)
                            
                            '↓処理行と上の行の(1→10列＆7列)が一致ならば以下（肝）7
                        If hirt(-qq + 1 + ii, 10) = hirt(-qq + 1 + ii - 1, 10) And hirt(-qq + 1 + ii, 7) = hirt(-qq + 1 + ii - 1, 7) Then  ',1)→,10)
                            hirt(-qq + 1 + ii, 3) = hirt(rrr - qq + 2, 3) '←空白化
                            hirt(-qq + 1 + ii, 7) = hirt(rrr - qq + 2, 3) '←空白化
                            hirt(-qq + 1 + ii, 6) = hirt(rrr - qq + 2, 3) '←空白化
                            
                            If hirt(-qq + 1 + ii, 8) <> "" Or hirt(-qq + 1 + ii - 1, 9) <> "" Then '86_014g if条件追加
                                hirt(-qq + 1 + ii - 1, 8) = hirt(-qq + 1 + ii, 8) + hirt(-qq + 1 + ii - 1, 9) '8列上行＝8列当行+9列上行
                                hirt(-qq + 1 + ii, 8) = hirt(rrr - qq + 2, 3) '処理後、8列当行は空白化
                            End If
                        End If
                    End If
                ElseIf hirt(-qq + 1 + ii, 7) <> "" Then '3列情報ナシで7列ありならこちら
                    hirt(-qq + 1 + ii, 7) = hirt(rrr - qq + 2, 3) '7列空白化
                    hirt(-qq + 1 + ii, 6) = hirt(rrr - qq + 2, 3) '6列空白化
                End If
                ii = ii - 1
                Call hdrst2(rrr - ii, a, 10000, 0, 0)
            Loop  '3列目消込プログラムここまで
            '※ここでのiiはqq(開始行)
            
            Range(Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(qq, 1), Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(rrr + 1, 10)).Value = hirt
            Erase hirt
            
            Application.Cursor = xlWait
            
            '↓開始行特別処置A　8列先頭行　'↓９列nullで８列ゼロ入り込む阻止用
            If twt.Cells(qq, 3).Value <> "" And twt.Cells(qq, 7).Value <> "" And twt.Cells(qq, 8).Value = "" And twt.Cells(qq, 9).Value <> "" Then
                twt.Cells(qq, 8).Value = twt.Cells(qq, 9).Value  'qq=iiですね
            End If
            
            '↓開始行特別処置B　6列、7列先頭行　（無くてもバグらないかもしれないが）
            If twt.Cells(qq, 3).Value = "" And twt.Cells(qq, 7).Value <> "" Then
                twt.Cells(qq, 7).Value = twt.Cells(rrr + 10, 3).Value  '←空白化
                twt.Cells(qq, 6).Value = twt.Cells(rrr + 10, 3).Value  '←空白化
            End If
            
            '↓８列→3列目にコピー（数式→値化）パターンA?Cどの場合も途中段階として
            Call betat4(twn, "高速シート_" & syutoku(), qq, 8, rrr, 8, twn, "高速シート_" & syutoku(), qq, 3, "pp", "")  '8(旧5)
            Application.Cursor = xlDefault
        End If '特命条件処理（４列ベタ）ここまで
        
        cnt = 0
        rrr = rrr + 1  'rrrは最終行の次行(all1的には空白になった行)

        'ロック識別子挿入
        If h >= k Then
            twt.Cells(rrr, 1).Value = bfshn.Cells(h, Abs(er(2, bni)))  'mghz(strconv24済)の最下行
            twt.Cells(rrr, 10).Value = StrConv(bfshn.Cells(h, Abs(er(2, bni))), 8 + hrkt) 'mghzの最下行　1→10列目(Excel2019対策
            twt.Cells(rrr, 2).Value = "000_0000000"
            rrr = rrr + 1
        End If

        cnt = 0
        rrr = rrr - 1
        'この時点のrrrは高速シートのデータ終了行(含ロック因子)、qqは相変わらずデータ開始行(対象シート及び高速シート)
        
        '元来の高速シート昇順降順はこっち
        Call saato(twn, "高速シート_" & syutoku(), 3, qq, 1, rrr, 10, 99)  '3列(ｶｳﾝﾄ列)降順
            '↑これがないと、特命条件の連載が数の多い順にならない。特命条件のルーチンい移してもよいが。　20210309記
        Call saato(twn, "高速シート_" & syutoku(), 10, qq, 1, rrr, 10, 1)   '昇ソート 1→10列目(Excel2019対策・平・片同一視されない)
    End If '対象シートデータ部がすっからかんでない時の処理ここまで

    Workbooks(bfn).Activate  '86_017h　上からこちらへ（正常性確認中）
    Sheets(shn).Select
    DoEvents

End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub giktzg(a As Long, rog As String)     '外結事後
    Dim ii As Long, jj As Long
    '↓一ボタン毎にインクリメント（ID変わったときのみリセットされる） 。※bfn、shn側は更新されない（初回複写時の値が載ってるだけ）。
    twbsh.Cells(13, 3).Value = twbsh.Cells(13, 3).Value + 1 '7s (2,3)→(13,3)25s
    
    'ログ部ここから　30s76
    If rog <> "" Then     'if,rogがnullなら(＊だけの時）記載しない
    'MsgBox rog
        jj = 1
        Do Until Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = ""
            jj = jj + 1
            If jj = 10000 Then
                MsgBox "空白行が見つからないようです"
                Exit Sub
            End If
        Loop

        Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = 1 '項目名
        Workbooks(twn).Sheets(shog).Cells(jj, 2).Value = jj '項番

        'log
        Workbooks(twn).Sheets(shog).Cells(jj, 3).Value = "結.　外結、" & Mid(twn, 1, Len(twn) - 5) _
        & "、ｦ" & shn & "ｦ" & bfn & "、ｦ" & shn & "ｦ" & bfn & "、from" & dd1 & "to" & dd2 _
        & "、項目名、b" & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "、" & twbsh.Cells(2, 2).Value & "、" _
        & Format(Now(), "yyyymmdd_hhmmss") & "、" & bfshn.Cells(sr(8), 5).Value & "、" & Application.WorksheetFunction.Sum(bfshn.Range("A:A")) & "、" & dd2 - dd1 + 1
        '末備から、外結処理列数(min:1)、all1(行数)←(＋１では

        Workbooks(twn).Sheets(shog).Cells(jj, 4).Value = Format(Now(), "yyyymmdd")  'date
        Workbooks(twn).Sheets(shog).Cells(jj, 5).Value = Format(Now(), "yyyymmdd_hhmmss")  'timestamp
        Workbooks(twn).Sheets(shog).Cells(jj, 7).Value = bfn & "\" & shn  'to

        frmx = 9 'fromの開始列
        ii = 1
        Do Until rvsrz3(rog, ii, "ヱ", 0) = ""
            Workbooks(twn).Sheets(shog).Cells(jj, ii + frmx - 1).Value = rvsrz3(rog, ii, "ヱ", 0) 'from
            ii = ii + 1
            If ii = 200 Then
                Call oshimai("", bfn, shn, k, a, "うまくいってない2。")
            End If
        Loop
        Workbooks(twn).Sheets(shog).Cells(jj, 8).Value = ii + frmx - 2 '最右列の列に入れる値
    End If  'ログ部ここまで
    
    Application.CutCopyMode = False
    bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
    DoEvents
    Call oshimai("", bfn, shn, k, dd2, "")
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub papchk(pap() As Long, nmb As Long, bun As Long)
    Dim ii As Long
    For ii = 1 To bun
    '    If pap(nmb, 0) <> pap(nmb, ii) Then MsgBox "pap(nmb," & ii & ")不一致"  '当面、nmbは不活性で
        If pap(2, 0) <> pap(2, ii) Then MsgBox "pap(2," & ii & ")不一致"
    Next
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub papzousei(pap() As Long, mr() As String, nmb As Long, bni As Long)
    If bni = 0 Then Call oshimai("", bfn, shn, 1, 0, "(papzousei)bni:0は対象外です。")
'    pap(2, bni) = kgcnt(mr(1, 2, bni), mr(2, 4, bni))
    pap(nmb, bni) = kgcnt(mr(1, nmb, bni), mr(2, 4, bni))
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub saato(fbk As String, fsh As String, sas As Long, fmg1 As Long, fmr1 As Long, fmg2 As Long, fmr2 As Long, sot As Long)
    'ソート列、fmg1、fmg2　まずは昇順、１ー４列限定で
    'https://excelwork.info/excel/cellsortcollection/
    If sot = 1 Then '昇順
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Clear
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Add Key:=Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas), Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With Workbooks(fbk).Worksheets(fsh).Sort 'Sortオブジェクトに対して '並べ替える範囲を指定し↓
            .SetRange Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Worksheets(fsh).Cells(fmg2, fmr2))
            .Header = xlNo '1行目がタイトル行かどうかを指定し（規定値：xlNo）
            .MatchCase = False '大文字と小文字を区別するかどうかを指定し
            .Orientation = xlTopToBottom '並べ替えの方向(行/列)を指定し  (規定値：xlTopToBottom)
            .SortMethod = xlPinYin 'ふりがなを使うかどうかを指定し  (規定値：xlPinYin)
            .Apply '並べ替えを実行します 　省略はしない方が無難、前回のを引き継ぐらしいので
        End With
    ElseIf sot = 99 Then '降順
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Clear
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Add Key:=Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas), Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas)), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With Workbooks(fbk).Worksheets(fsh).Sort
            .SetRange Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Worksheets(fsh).Cells(fmg2, fmr2))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Else
        Call oshimai("", bfn, shn, 1, 0, "sotの引数が変です")
    End If
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function kaunta(mr() As String, qq As Long, pap5 As Long, bni As Long, er5() As Long, mr5() As String) As Long
    kaunta = 1 '減点方式
    For qap = 0 To pap5  '複数列の減点方式
        If Abs(er5(qap)) < 1 Then
            'A,０や０．４：カウント対象
        ElseIf mr5(qap) = "ヰ#N/A" Then
            If WorksheetFunction.IsNA(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) Then
                '↓エラーになるｗｈｙ？
                kaunta = 0 'カウント非対象
                Exit For '←こうしないと複数列時エラーになる。
            ElseIf Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value = "" Then
                kaunta = 0 'カウント非対象
            End If
        '以下はセルにN/Aあるとエラーになる。20200207
        ElseIf Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value = "" Then
            'B,セル空白はカウント非対象
                kaunta = 0
        '以下セルに情報あり
        ElseIf mr5(qap) = "" Then
            'C,カウント対象
        ElseIf Left(mr5(qap), 1) = "≧" Then
            'D1
            If IsDate(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) Then '日付比較新設 86_018j
                If CDate(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) < CDate(Mid(mr5(qap), 2)) Then
                    kaunta = 0 'カウント非対象
                End If
            Else '↓従来(isnumeric)
                If Val(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) < Val(Mid(mr5(qap), 2)) Then
                    kaunta = 0 'カウント非対象
                End If
            End If
        ElseIf Left(mr5(qap), 1) = "≦" Then
            'D1
            If IsDate(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) Then '日付比較新設 86_018j
                If CDate(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) > Val(Mid(mr5(qap), 2)) Then
                    kaunta = 0 'カウント非対象
                End If
            Else '↓従来(isnumeric)
                If Val(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) > Val(Mid(mr5(qap), 2)) Then
                    kaunta = 0 'カウント非対象
                End If
            End If
        ElseIf Left(mr5(qap), 1) = "ー" Or Left(mr5(qap), 1) = "ヰ" Or mr5(qap) = "n0" Then  '「ー」追加86_012y
            'E　※Dのn0→Eで処理（文字列の0にも対応可能
            '30s75 strcomp 初導入(ヰ△の△とセル側比較がうまく行かないため)
            If StrComp(Mid(mr5(qap), 2), Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value, vbBinaryCompare) = 0 Then
                kaunta = 0 'カウント非対象
            End If
        ElseIf mr5(qap) = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value Then
            'F,n0：カウント対象
        Else
            'Z,上記以外：カウント非対象　　ｦ〇ｦ□　で不一致のケースが想定される。
            kaunta = 0
        End If
    Next
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function tszn(er9() As Currency, bni As Long, mr() As String, qq As Long, pap9 As Long, mr9() As String) As Variant  '加算(足し算)
    'ヱ対応6or8行、当列a,書き込み行hx,文節bni,P値(AM有無)、同列0多列1(使用終了)、-1-2複写,mr,er,6行目８行目ヱの数(pap9),mr9
    Dim qap(2) As Long
    'Dim nuez As Double
    qap(1) = 0
    qap(2) = pap9
    For qap(0) = qap(1) To qap(2)  '複数列取扱時、ループする。
        If Abs(er9(qap(0))) = 0.4 Or Abs(er9(qap(0))) = 0.1 Then
            nuex = Val(mr9(qap(0)))
        Else '通常時
            nuex = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er9(qap(0)))).Value 'er(11, bni)→ee→qq
        End If
        If qap(0) = 0 Then
            'ヱが０個のとき→何もしない
        Else
            nuex = nuey * nuex '複数列の２回目以降→乗算実施
        End If
        nuey = nuex
    Next
    tszn = nuex
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub samaru(aa As Long, mr101 As String)  'サマリ値処理30ｓ59より
    'さまる列、mr(1, 1, 1)←シート名
    Dim gx As Long, gy As Long, ii As Long

    For ii = sr(0) + 4 To sr(0) + 5
        uu = 0
        gx = 0
        gy = 0
        If IsError(bfshn.Cells(ii, aa)) Then '30s66_2バグ改良
            uu = 1  'セルがエラー→値が入っている。
        ElseIf bfshn.Cells(ii, aa) <> "" Then
            uu = 1
        End If
        
        If uu = 1 Then 'ん
            '30s71_3場所若干変更、３０ｓ７４さらに変更
            If bfshn.Cells(ii, aa).Font.Color = RGB(38, 37, 38) Or bfshn.Cells(ii, aa).Font.Color = RGB(255, 0, 0) Then '黒[赤]sum
                gy = sr(8) + 5
                gx = 4
            ElseIf bfshn.Cells(ii, aa).Font.Color = RGB(1, 102, 255) Then '青数値＞０
            
                gy = sr(8) + 5 'sr(8)は対:転載列の行
                gx = 2  '4
            ElseIf bfshn.Cells(ii, aa).Font.Color = RGB(198, 90, 17) Then '茶文字列
                gy = sr(8) + 6 '17
                gx = 4 '2
            ElseIf bfshn.Cells(ii, aa).Font.Color = RGB(1, 176, 80) Then '緑nonzero
                gy = sr(8) + 6 '16
                gx = 2
            End If
        End If  'ん
             
        If gx > 0 Then  'コピペ部
             Call copipe(bfn, shn, gy, gx, gy, gx, bfn, shn, ii, aa, ii, aa, 3)  '3→FormulaR1C1(-4123向け) (脱.copy.paste)　86_020d
            
            'コピペしたセルを値に変換
            If StrConv(Left(mr101, 1), 8) <> "*" Then
                Call cpp2(bfn, shn, ii, aa, ii, aa, bfn, shn, ii, aa, 0, 0, -4163) '-4163は値をコピー　速
            Else
                '１行目が「＊」だけの時、ここ通過してる。
                'MsgBox "ここを通るケースは１行目が「＊」だけの時"
            End If
            bfshn.Cells(sr(0) + 3, aa).Value = Now() '活用例1タイムスタンプ入れる
        End If
    Next
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function tnele(ct8 As String, ax As Long, qap0 As Long, ct3 As String, er78() As Currency, a As Long, hx As Long, bni As Long, qq As Long, fuk12 As Long, mr() As String, er() As Currency, mr8() As String) As String '転載エレメント決定　86_013j
    '転載エレメント決定　30s86_013j function化
    If fuk12 > 0 Then '30s62 ｃ：-1-2複写時
        tnele = bfshn.Cells(fuk12, ax).Value
    ElseIf fuk12 = -7 Then 'ここから以下、fuk12が0以下
        If er78(0, qap0) = 0 Then   'bx→0へ
            tnele = Trim$(CStr(Workbooks(mr(2, 0, bni)).Sheets(mr(1, 0, bni)).Cells(qq, Abs(a)).Value))
        Else
            tnele = Trim$(CStr(Workbooks(mr(2, 0, bni)).Sheets(mr(1, 0, bni)).Cells(qq, Abs(er78(0, qap0))).Value))
        End If
    '以降、bx:1
    ElseIf er78(1, qap0) = 0.4 Then 'マニュアル値(mr8)は、er8ないときだけ使われる。
        tnele = Trim$(mr8(qap0))     '30s75複数列対応化
    ElseIf Round(er(6, bni), 0) = -15 Then    '-15 ある文節エレメント抽出
        tnele = rvsrz3(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), Val(mr(2, 6, bni)), mr8(0), 0)
    ElseIf Round(er(6, bni), 0) = -14 Then '区切り数　30s86_016m
        tnele = kgcnt(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), mr8(0)) + Val(mr(2, 6, bni))
    ElseIf Round(er(6, bni), 0) = -10 Then 'naka2
        tnele = Mid(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), Val(mr(2, 6, bni)))
    ElseIf Round(er(6, bni), 0) = -9 Then 'hiduke -9
        tnele = Format(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), mr(2, 6, bni))
    ElseIf Round(er(6, bni), 0) = -8 Then 'mojihen    -13→-8
        'MsgBox "使われてますね。6行が-8"
        tnele = StrConv(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), mr(2, 6, bni))
    ElseIf Abs(er78(1, qap0)) = 0.1 Then
        tnele = Format(qq, "0000000")  '30s85_014 行番号転載対応の修正
    ElseIf er78(1, qap0) > 0.2 And mr8(qap0) <> "" And bfshn.Cells(hx, a).Value <> "" And UBound(er78(), 2) = 0 Then
        '20180813新井対応→202102複数列では実施せずの仕様に(動作おかしくなるので) 86_020j
        '何もしない（tnele=""のまま）第二因子がｖ用　 "6448"
    Else '通常時　fuk12=-8はここ通る。
        If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) < 0) Then
            If qap0 = 0 Then
                tnele = ct8  '特命条件(連載型使用時、８列加算値、86_019oこちらへ)
            Else
                tnele = ct3  '特命条件(連載型使用時、６列値（従来）)
            End If
        Else '従来の通常パターン
            tnele = Trim$(CStr(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))).Value))
        End If
    End If
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub tnsai(ct8 As String, tst As Long, ct3 As String, er78() As Currency, a As Long, hx As Long, bni As Long, p As Long, qq As Long, fuk12 As Long, mr() As String, er() As Currency, pap7 As Long, mr8() As String)

    'ヱ対応7行8行、当列a,書込行hx(0はコメント用),文節bni,P値(AM有無)、参照行(コメント時のケース(項目行が参照行)もあり)、-1-2複写,mr,er,7行目ヱの数(pap7)
    Dim ax As Long '書込列（一方、aは当列）
    Dim tenx As String, teny As String, tenz As String
    Dim qap(3) As Long, apa7 As Long    'qap3は将来用
    Dim bx As Long, mm As Long
    'bx策定
    
    If fuk12 = -7 Then bx = 0 Else bx = 1  '-7は７行コメント用,fuk12:-1-2複写フラグのこと 通常はbx:1
    
    If er(5, bni) < 0 Or fuk12 = -7 Then apa7 = 0 Else apa7 = pap7
    '差分時あるいは7行のコメント時はpap7成分無効化(以降当ルーチンapa7使用、pap7不使用(-7以外))
    
    If fuk12 > 0 Then '-1-2複写時 fuk12 は転載元の行
        qap(1) = -1
        qap(2) = apa7
    Else
        qap(1) = 0
        If fuk12 = -7 Then  '7行のコメントの特殊時
            If pap7 <= UBound(er78(), 2) Then '普通な時
                qap(2) = 0    'pap7→0へ
            Else '普通でないとき　30s85_021新設
                'MsgBox "ここはもう来ないのでは"    '７行ヱの数＞8行ヱの数　はどこかでoshimai処理だったような。。
                Call oshimai("", bfn, shn, sr(7), a, "ここはもう来ないのでは")
                qap(2) = UBound(er78(), 2)
            End If
        ElseIf fuk12 = -8 Then
            qap(2) = 0    'UBound(er78(), 2)→0へ
        Else '通常時
            qap(2) = apa7  '（通常）UBound(er78(), 2)[８行目のヱの数]→apa7[7行目のヱの数]へ
        End If
    End If

    For qap(0) = qap(1) To qap(2)  '複数列転載時ループ(通常：7行目の複数列個分)する。単数列はループせず1回のみ。-1-2複写時、-1からループする。qap(2)までループする。
                             'qap(2)は通常は8行目ヱの数。7行コメントだけ7行ヱの数。8行ヱ数≠7行ヱ数の時意識せよ。
        '★７行ーの時はスルーへ　86_013
        If er78(0, qap(0)) = 0.1 Then
            'MsgBox "7行0.1(ー)、省略数目あり"
        ElseIf er78(0, qap(0)) = 0.4 Then
            Call oshimai("", bfn, shn, sr(7), a, "7行目0.4は今の所あり得ないかと。")
        Else
            '同列か他列か？（ax策定）転載元・転載先で使用
            If qap(0) <= apa7 Then '単数列＆複数列　apa7最小値ゼロ（：7行ヱ無し）
                'ax更新(ヱの、はみ出し分は更新しない)
                If er78(0, qap(0)) > 0 And er(5, bni) >= 0 Then '差分時でなく、7行正値の時のみ他列許容、(普通の転載)
                    ax = er78(0, qap(0))  '他列
                Else
                    '↓30s86_012w 特命条件対応版
                    If er(9, bni) > 0 And er(10, bni) > 0 And er(5, bni) >= 0 And Not (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And Int(er(7, bni)) = 0 And er(8, bni) <> 0) Then
                    If er(20, bni) = 1 Then MsgBox "yaa"
                        Call oshimai("", bfn, shn, 1, 0, "b加算ありで同列転載しようとしています。確認を。")
                    Else
                        ax = a  '差分時でも一発目だけは有効　-1の時
                    End If
                End If
            End If
    
            tenx = ""   '不要(tneleでリセットされる)だが、一応念のため
            tenx = tnele(ct8, ax, qap(0), ct3, er78(), a, hx, bni, qq, fuk12, mr(), er(), mr8())
            '↑30s86_013j function化
            
            '連載処理〔Ｂ〕・・〔Ａ〕よりこちらが先に
            If fuk12 <= 0 Then  'ｃが-1、-2以外 30s82=→<= に(コメント-7-8対応)
                If qap(0) = qap(2) Then '最終数のみ以下
                    If UBound(er78(), 2) > qap(2) Then 'はみ出ている場合のみ、以下はみ出し分連結処理〔Ｂ〕
                        For mm = qap(2) + 1 To UBound(er78(), 2)
                            If er78(0, mm) <> 0.1 Then
                                tenx = tenx & mr(2, 4, bni) & tnele(ct8, ax, mm, ct3, er78(), a, hx, bni, qq, fuk12, mr(), er(), mr8())
                            End If
                        Next
                    End If
                End If
            End If
       
            '連載処理〔Ａ〕(tenx更新)〔Ａ〕いつもの当初からの連載処理(語尾付加)
            If fuk12 <= 0 Then  'ｃが-1、-2以外
                    'ヱなし、同列、６行ゼロor-1台、８行マイナスが条件、複数列許容へ
                            '↓ノーマル連載は-2以下は許容してない。あるいは、
                    If (er(6, bni) = 0 Or Round(er(6, bni)) = -1) Or _
                        (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) < 0) Then '<0→<0.3 85s023
                        '↑特命条件連載型をorで追加
                        '↓86_018d　複数列初列で判断へ
                        If mr8(0) <> "" And (er78(bx, 0) < 0) Then  'Or er78(bx, qap(0)) = 0.1は不要だった(行番号は連載型許容してない)
                            tenx = bfshn.Cells(hx, ax).Value & tenx & mr8(0) 'あ 特命条件もこっち mr(2,8,bni)→mr8(qap(0))→mr8(0)
                        ElseIf er78(bx, 0) < -0.5 Then    '(bx, qap(0))→(bx, 0)
                            Call oshimai("", bfn, shn, sr(8), a, "もう通ることはないと思われる(※８行目は新仕様に修正して下さい)")
                            tenx = bfshn.Cells(hx, ax).Value & tenx & "、"  'い
                        Else
                            '行番号上書き型(0.1、ここでは何もしない)
                        End If
                        '↑特命条件以外での他列複数列連載型は↑この変動動作が整理されていない状況である。202005
                    End If
            End If
            
            '複数列の最終数目以外をここで書き込み、-7-8コメント時はここで実施しない(最後で)。ｰ1ｰ2複写処理はここでは行われない。
            '複数列転載行為の最終数目(単数列はその一回)は、ここでは実施せずfornextの後ろで実施(最終まとめ分のtenxがまとまってないため。)
            If tenx <> "" And hx <> 0 And fuk12 >= 0 Then  'And qap(0) < apa7 の条件撤廃　86_013q
                'tst、trt↓使えそう。-7-8コメントはここ使われていないし。
                If p = 2 And tst = 1 Then
                    With bfshn.Cells(hx, ax)  '文字列新規ちまちま(-3<0&-4<0　か連載新規),※fuk12はここ弾かれる。
                        .NumberFormatLocal = "@"
                        .Value = tenx
                    End With
                ElseIf tst = 8 And fuk12 = 0 Then  '元シートセルちまちま
                    If Abs(er78(1, qap(0))) = 0.4 Or Abs(er78(1, qap(0))) = 0.1 Then Call oshimai("", bfn, shn, sr(8), a, "この形式では踏襲型ちまちまで指定できませんa")
                    With bfshn.Cells(hx, ax)
                        .NumberFormatLocal = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap(0)))).NumberFormatLocal
                        .Value = tenx
                    End With
                ElseIf tst = 8 And er(6, bni) <> -2 And fuk12 > 0 Then   'セル複写ちまちま(fuk12)86_012k
                    With bfshn.Cells(hx, ax)
                        .NumberFormatLocal = bfshn.Cells(fuk12, ax).NumberFormatLocal
                        .Value = tenx
                    End With
                Else 'normal(感知しない、そのまま値で貼り付け、標準でもない（後で標準なり通貨なり処理）)
                    bfshn.Cells(hx, ax).Value = tenx
                End If
            End If
        End If '７行ーでない時の処理はここまで
    Next '複数列転載時、ループする。-1-2複写時、-1からループする。は、ここまで
    
    qap(0) = qap(0) - 1 '1戻す(Next後のインクリ戻す)
    'ここでのtenxは、最終まとめ分のtenxがまとまっている状態
    '単数列転載は必ずここで行われる。for側では行われない。ｰ1ｰ2複写はここで行われる。
    If hx = 0 Then MsgBox "hx=0ということがあり得るだろうか？"

    If (fuk12 = -8 Or fuk12 = -7) And tenx <> "" Then '30s79 8行コメント用
        bfshn.Cells(hx, a).ClearComments
        bfshn.Cells(hx, a).AddComment
        bfshn.Cells(hx, a).Comment.Text Text:=tenx
        'bfshn.Cells(hx, a).Comment.Shape.TextFrame.AutoSize = True    '86_022b not use for mac
    End If
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function wetaiou(fn As String, f As String, ii As Long, erx() As Currency, kg2 As String, spd As String, mrx() As String, gyu As Long) As String
    'ファイル名、シート名、行数、列数(複数)、区切り文字、高速可否、列第二因子(複数)、２行目か３行目か
    Dim am2 As String, am3 As String, qap As Long, jj As Long
    If fn = "" Then '30s86_020a
        If UBound(erx()) > 0 Then 'ヱがあるときのみ（二発目以降）。単文節は通らない。
            jj = UBound(erx())   '2.11(従来通り）　３もある？
            For qap = 1 To jj
                am2 = am2 & kg2 & am3
            Next
        End If
    Else '既存ver
        'ヱ複数列の返り値を返す(○ヱ▽ヱ◇形式)。複数列で無くても常時使われる(am2)
        If Abs(erx(0)) < 1 Then
'            If erx(0) > 0.3 Then
            If erx(0) >= 0.1 Then   '86_022f
                If mrx(0) = "" Then '第二因子あればそちらの記載優先で　85_027検証8
'                    am2 = "" '単数列はここはあり得ない。複文節でかつ列指定無い時(0.4)が該当。
                                                                                '↓コメント修正 '86_022f
                     am2 = "" '単数列はここはあり得ない。複文節でかつ列指定無い時(0.4　あるいは0.1)が該当。
                Else
                    am2 = mrx(0)  '85_027検証8　３行目第二因子
                End If
            End If
        ElseIf spd = "純高速" Or spd = "近似高速" Then  '単文節時or複数列一発目(0番目)高速
'            am2 = CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(0))).Value)  '高速時はtrim実施しないで揃えるへ　85＿027検証6
            
            am2 = WorksheetFunction.Clean(Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(0))).Value)))  '高速時はtrim再開＆CLEAN関数　86_022i
        Else '単文節時or複数列一発目(0番目)低速 ※返り値null(セルが空白返り値)の場合もある。
'            am2 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(0))).Value))
            am2 = WorksheetFunction.Clean(Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(0))).Value))) 'CLEAN関数適用へ　86_022i
        End If
    
        If UBound(erx()) > 0 Then 'ヱがあるときのみ（二発目以降）。単文節は通らない。
            jj = UBound(erx())   '2.11(従来通り）　３もある？
            For qap = 1 To jj
                If Abs(erx(qap)) < 1 Then
                    If erx(qap) < 0.3 Then
                        am3 = Format(ii, "0000000") '「ー」の時(0.1)、行番号をキーとする。
                    ElseIf mrx(qap) = "" Then '(0.4)　New投入　85_027検証8
                        am3 = "" '従来(0.4)※３行目第二因子なし　※従来型
                    Else
                        am3 = mrx(qap)  '※３行目第二因子あり　※New　85_027検証8
                    End If
                ElseIf spd = "純高速" Or spd = "近似高速" Then    '高速時はtrim実施しないで揃えるへ　85＿027検証6
'                    am3 = CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(qap))).Value)
                    am3 = WorksheetFunction.Clean(Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(qap))).Value)))  '高速時はtrim再開＆CLEAN関数　86_022i
                
                Else  '低速は従来通り
'                    am3 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(qap))).Value))
                    am3 = WorksheetFunction.Clean(Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(qap))).Value))) 'CLEAN関数適用へ　86_022i
                End If
                am2 = am2 & kg2 & am3
            Next
        End If
    End If
    
    'ここでam3は一旦用済みなので、ここから、別の用途で 30s86_021g
    am3 = am2
    am2 = WorksheetFunction.Clean(am3)
    
'    If am2 <> am3 Then MsgBox "Clean関数発動(" & am2 & ")"  '←必要に応じて表示する
    
    wetaiou = am2
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub hdrst(ii As Long, a As Long)  '21ｓ簡素化
    '左下ステータス表示部  100→1000へ201904
    If cnt < Int(ii / 1000) * 1000 Then  'cnt はパブリック変数
        cnt = Int(ii / 1000) * 1000
        DoEvents
        If flag = True Then Call oshimai("", bfn, shn, 1, 0, "中止しましたです") '中止ボタン処理
        Application.StatusBar = Str(cnt) & "、" & Str(a - dd1 + 1) & " / " & Str(dd2 - dd1 + 1)
    End If
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub hdrst2(ii As Long, a As Long, ak As Long, kkk As Long, hhh As Long)
    '左下ステータス表示部
    If ak <= 0 Then ak = 100 '異常値のときは100のデフォ値(従来変わらず)で
    If cnt < Int(ii / ak) * ak Then  'cnt はパブリック変数
        If kkk <> 0 Then
            bfshn.Cells(2, 4).Value = kkk
            bfshn.Cells(3, 4).Value = hhh
        End If
        cnt = Int(ii / ak) * ak
        DoEvents
        If flag = True Then Call oshimai("", bfn, shn, 1, 0, "中止しましたよ") '中止ボタン処理
        Application.StatusBar = Str(cnt) & "、" & Str(a - dd1 + 1) & " / " & Str(dd2 - dd1 + 1)
    End If
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function hunpan(fn As String, f As String, ii As Long, g As Currency, e7 As Currency, e5 As Currency, c As Currency, e As Currency, mc As String, hk1 As String) As String
    Dim et5 As String, et As String
    '暗号鍵の文字列
    If mc = "1" Then 'mc・・６行op、鍵の種類
        et5 = "1"
    ElseIf mc = "e5" Then  '使われている　5行目カウント基準列にある文字列が鍵(a234567　とか)
        et5 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(e5))))
    ElseIf mc = "e" Then  '8行目転載列にある文字列が鍵(a234567　とか)　使われていない
        et5 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(e))))
    Else     '3行目転載列にある文字列が鍵(a234567　とか)　使われていない
        et5 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(g))))  '
    End If
    et = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(e))))  '暗号鍵確定
    hunpan = hunk2(c, et, et5, hk1)    '暗号鍵
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function hunk2(cc As Currency, et As String, et5 As String, hk1 As String) As String
    '暗号復号、対象文字、鍵、エラービット 変数二文字化　86_014k
    Dim mjs As Long '転載文字の文字数
    Dim gjs As Long '突合文字の文字数
    Dim pp As Long, vv As Long
    Dim hh As String '変換用一文字
    Dim uu As String '変換後文字
    Dim tt As Long '変換unicodeシフト定数
    Dim jj As Long '範囲チェック
    Dim qq As Long '余り
    Dim mb As Long 'まぶし係数
    Dim ww As Long, tn As Long, ss As Long

    uu = ""
    tn = 0
    gjs = Len(et5) '暗号鍵の文字数
    mjs = Len(et) '暗号対象文字の文字数
    vv = 0
    For pp = 1 To gjs
        ss = AscW(Mid(et5, pp, 1))
        vv = vv + ss
    Next
    ww = vv Mod 10  '暗号鍵のまぶしバイアス値
    vv = 0
    
    For pp = 1 To mjs 'ppは転載文字の、ある文字目
        qq = (pp - tn) Mod gjs + 1  '暗号鍵の文字数循環
        ss = AscW(Mid(et5, qq, 1))
        mb = (ss + pp - tn + ww) Mod 50 'まぶし係数(０?＋?４９)
        tt = 25000 - 30 + 99 * mb '仮のtt
        If cc = -5 Then
            jj = 0
            tt = tt
        ElseIf cc = -7 Then
            jj = tt
            tt = (-1) * tt
        End If
        If Mid(et, pp, 1) = "、" Then
            hh = "、"
            tn = pp
        ElseIf (AscW(Mid(et, pp, 1)) >= jj + 32 And AscW(Mid(et, pp, 1)) <= jj + 126) Then
            hh = ChrW(AscW(Mid(et, pp, 1)) + (tt))
        Else
            vv = 1
        End If
        If vv = 0 Then uu = uu & hh
    Next

    If vv = 0 Then
        hunk2 = uu
    Else
        hunk2 = "規定外文字あり：" & et
        hk1 = "1"
    End If
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function yhwat1(fn As String, f As String, ee As Currency, yhs As Long, a As Long, bni As Long, kg1 As String, nkg As Long, kg2 As String, bugyo As String) As String
    '対_ファイル名、対_シート名、対_項目行、当_行番号、当_列番号、文節数、区切り文字(ゐ)、無し区切りビット、区切り文字(ヱ)、その文節の原本情報（＊も含む）
    '仮想キー　◇ヱ◇ヱヱ◇対応30s43(doloop→fornextに）
    Dim ymj1 As String, ymj2 As String, mh As Long, yap As Long '←ヱの個数
    Dim pm As Long, ii As Long, wenk As Long  'ー無有　、ヱの順番目(kgwe→ii)　、ヱの区切り可否
    Dim chwat1 As String, chwat2 As String '返り値の区切り毎生成物　、返り値の素（蓄積型）
    
    If kg2 = "" Then wenk = 1 'wenk=1：ヱ区切り無し、wenk=0 ：ヱ区切り有り
    pm = 1
    If Right(fn, 4) = ".xls" Then mh = 256 Else mh = 2000 '85_009　.xlsにも対応
    ymj1 = rvsrz3(bugyo, 2, "ｦ", 2) 'nkg：２、色んな行で＊あヱい　形式（＊ありヲ不使用）を許容（30s73記）
    If bni >= 2 And bugyo = "" Then  '30s75 ymj1→bugyo に(第二文節以降がｦｦの時、前節情報踏襲になるバグ対処)
        yhwat1 = ""   '←yhwat1が文字関数にならざるを得ない根源
    Else 'bni=1 のとき、若しくは、ymj1情報ありのとき
        yap = kgcnt(ymj1, kg2)  '←ヱの個数
        For ii = 1 To yap + 1
            pm = 1  '30s81_7追加（バグ、pmはfor毎にリセットしなければならない）
            ymj2 = rvsrz3(ymj1, ii, kg2, wenk)
            If Mid(ymj2, 1, 1) = "ー" Then
                pm = -1
                If ymj2 = "ー" Then ymj2 = "" Else ymj2 = Mid(ymj2, 2)
            End If
            chwat1 = ""
            If IsNumeric(ymj2) Or IsDate(ymj2) Then
                chwat1 = ymj2
            ElseIf ymj2 = "" Then
                chwat1 = pm * 0.4  'ｦｦ→戻り値0.4、ｦーｦ→戻り値-0.4が入る
                If chwat1 = -0.4 Then chwat1 = 0.1 '85_024 -0.4→0.1
            ElseIf ee > 0 Then
                If IsError(Application.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)) Then
                    Call oshimai("", bfn, shn, yhs, a, "（処理中止）" & vbCrLf & "シート名：" & f & " 上の" & vbCrLf & "項目名「" & ymj2 & "」が見つかりません")
                Else
                    chwat1 = pm * Application.WorksheetFunction.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)
                End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "「項目名」が「なし」で文字検索しています")
            End If
            If ii = 1 Then chwat2 = chwat1 Else chwat2 = chwat2 & kg2 & chwat1
        Next
        yhwat1 = chwat2
    End If
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function zhwat1(fn As String, f As String, ee As Currency, yhs As Long, a As Long, bni As Long, kg1 As String, nkg As Long, kg2 As String, bugyo As String) As String
    '対_ファイル名、対_シート名、対_項目行、当_行番号、当_列番号、文節数、区切り文字(ゐ)、無し区切りビット、区切り文字(ヱ)、その文節の原本情報（＊も含む）
    '仮想キー　◇ヱ◇ヱヱ◇対応30s43(doloop→fornextに）
    Dim ymj1 As String, ymj2 As String, mh As Long, yap As Long '←ヱの個数
    Dim pm As Long, ii As Long, wenk As Long  'ー無有　、ヱの順番目(kgwe→ii)　、ヱの区切り可否
    Dim chwat1 As String, chwat2 As String '返り値の区切り毎生成物　、返り値の素（蓄積型）
    
    If kg2 = "" Then wenk = 1 'wenk=1：ヱ区切り無し、wenk=0 ：ヱ区切り有り
    pm = 1
    If Right(fn, 4) = ".xls" Then mh = 256 Else mh = 2000 '85_009　.xlsにも対応
    ymj1 = rvsrz3(bugyo, 2, "ｦ", 2) 'nkg：２、色んな行で＊あヱい　形式（＊ありヲ不使用）を許容（30s73記）
    If bni >= 2 And bugyo = "" Then  '30s75 ymj1→bugyo に(第二文節以降がｦｦの時、前節情報踏襲になるバグ対処)
        zhwat1 = ""   '←zhwat1が文字関数にならざるを得ない根源
    Else 'bni=1 のとき、若しくは、ymj1情報ありのとき
        yap = kgcnt(ymj1, kg2)  '←ヱの個数
        For ii = 1 To yap + 1
            pm = 1  '30s81_7追加（バグ、pmはfor毎にリセットしなければならない）
            ymj2 = rvsrz3(ymj1, ii, kg2, wenk)
            If Mid(ymj2, 1, 1) = "ー" Then
                pm = -1
                If ymj2 = "ー" Then ymj2 = "" Else ymj2 = Mid(ymj2, 2)
            End If
            chwat1 = ""
            If IsNumeric(ymj2) Or IsDate(ymj2) Then
                chwat1 = ymj2
            ElseIf ymj2 = "" Then
                chwat1 = pm * 0.4  'ｦｦ→戻り値0.4、ｦーｦ→戻り値-0.4が入る
                If chwat1 = -0.4 Then chwat1 = 0.1 '85_024 -0.4→0.1
            ElseIf ee > 0 Then
                If IsError(Application.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)) Then
                    Call oshimai("", bfn, shn, yhs, a, "（処理中止）" & vbCrLf & "シート名：" & f & " 上の" & vbCrLf & "項目名「" & ymj2 & "」が見つかりません")
                Else
                    chwat1 = pm * Application.WorksheetFunction.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)
                End If
                    If Val(chwat1) > 0 Then
                        chwat1 = rvsrz3(Workbooks(fn).Sheets(f).Cells(1, Abs(Val(chwat1))).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0)
                    Else
                        chwat1 = "ｰ" & rvsrz3(Workbooks(fn).Sheets(f).Cells(1, Abs(Val(chwat1))).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0)
                    End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "「項目名」が「なし」で文字検索しています")
            End If
            If ii = 1 Then chwat2 = chwat1 Else chwat2 = chwat2 & kg2 & chwat1
        Next
        zhwat1 = chwat2
    End If
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function rvsrz3(cef As String, bni As Long, kgr As String, nkg As Long) As String '30s03
    '対象文字列(A:*ｦあｦい)、文節数、区切り文字、区切りしないフラグ
    'bni=0 (文節ゼロ)→そのままフルで返す
    Dim ipc As Long, bb As Long, cc As Long, prd As String
    If kgr = "" Then
        'Call oshimai("", bfn, shn, 1, 0, "このケース、あるのかな？")　'→ありました
        prd = "、"
    Else
        prd = kgr
    End If
    If nkg = 1 Then
        rvsrz3 = cef  '区切りなしフラグ有効→そのまま返す
    'nkg=2→ｦなし時の(bni)2文節(ｦ◇ｦ)対処ルーチン。※ｦの解析の時(sr)しか使われない。
    ElseIf bni = 2 And nkg = 2 And StrConv(Left(cef, 1), 8) <> "*" And InStr(1, cef, "ｦ") = 0 Then 'nkg:2→ｦ対応　あ（ヲ無し）　bni2限定に,色んな行で使用 30s81接頭ど対応
        rvsrz3 = cef 'bni=1→2　30ｓ61
    ElseIf bni = 2 And nkg = 2 And StrConv(Left(cef, 1), 8) = "*" And InStr(1, cef, "ｦ") = 0 Then  '＊あ（ヲ無し）、＊　bni2限定に,色んな行で使用 30s81 接頭ど対応
        If Len(cef) > 1 Then
            If StrConv(Mid(cef, 2, 1), 8) = "*" Then '**
                If Len(cef) = 2 Then
                    Call oshimai("", bfn, shn, 1, 0, "「**」は使用されないです。" & vbCrLf & "「**あ」形式でよろしく")
                Else 'New 建設中　＊＊あ
                    rvsrz3 = Mid(cef, 3) '＊＊あ　→あ」を返す
                End If
            Else '従来型
                rvsrz3 = Mid(cef, 2) '＊あ　→あ」を返す
            End If
        Else
            rvsrz3 = cef '＊
        End If
    Else 'nkg=0, nkg=2：ヲ有り（先頭＊含む）
        If bni > 0 Then
            Do
                ipc = ipc + 1
                bb = cc
                cc = InStr(bb + 1, cef, prd)
            Loop Until cc = 0 Or ipc = bni 'それ以降該当なしor規定文節到達で抜ける。
        End If
        If cc = 0 And ipc = bni Then  'ジャスト規定文節で区切無しに　０文節もこちらに入る。
          If ipc > 1 Then '86_022s
            'rvsrz3 = Mid(cef, bb + 1)
            '↓30s86_022r
            rvsrz3 = Mid(cef, bb + Len(prd))
          
          Else '1文節目で未到達
          
'         MsgBox "cef:" & cef & "、bb:" & bb & "、" & "、kgr:" & kgr & "、" & Mid(cef, bb + 1)
'            rvsrz3 = ""
'            rvsrz3 = Mid(cef, bb + 1)
            '↓86_022s  (↑bbが0以外のケースはないと思われるので↓へ)
            rvsrz3 = cef
            If bb <> 0 Then MsgBox "bbが０以外っす"    'もしbb＜＞0のケースあれば、別途検討へ　86_022s

          End If
        ElseIf cc = 0 And ipc < bni Then  '規定文節未到達(よって該当文節は"")
            rvsrz3 = ""
        Else '規定文節で区切り文字もまだある。
'            rvsrz3 = Mid(cef, bb + 1, cc - bb - 1)
            '↓30s86_022r
            
            'rvsrz3 = Mid(cef, bb + Len(prd), cc - bb - Len(prd))
        '　　↓86_022s
          If ipc = 1 Then  '86_022s
            rvsrz3 = Mid(cef, 1, cc - bb - 1)
          Else
            rvsrz3 = Mid(cef, bb + Len(prd), cc - bb - 1)
          End If
        
        
        End If
    End If
'    If cef = "SELECT" Then
'    MsgBox "Hi"
'    End If
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function ctdg(rtyu As String, tyui As String, qwer As Currency, wert As Long) As Long '　最終行を返す。'こちらは行の方
    'ブック名、シート名、er4系、当列
    ctdg = ctreg(rtyu, tyui)    '←こちらに凝縮、pubikouへ
    '項準b(項零)での制限事項↓
    If ctdg > 200000 And Abs(qwer) < 1 Then  '1000→10000　86_014r　10000→200000　86_022j
        Call oshimai("", bfn, shn, sr(1), wert, "対象シートが20万行超え(" & ctdg & ")です")
    End If
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function ctdr(rtyu As String, tyui As String, qwer As Currency, wert As Long) As Long '　最右列を返す。'こちらは列の方
    'ブック名、シート名、er4系、当列
    '初期化 ↓対象シートの最右列　628
    ctdr = Workbooks(rtyu).Sheets(tyui).Range("A1").SpecialCells(xlLastCell).Column()
    If ctdr > 300 And Abs(qwer) < 1 Then
        Call oshimai("", bfn, shn, sr(1), wert, "対象シートが300列超え(" & ctdr & ")です")
    End If
    ctdr = ctdr + 1
    Do Until Workbooks(rtyu).Sheets(tyui).Cells(1, ctdr).EntireColumn.Hidden = False
        ctdr = ctdr + 1
    Loop  'ctrl+endの次列がhiddenだった場合の対処（85_020)
    ctdr = ctdr - 1
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function pasuwardo(knt As Integer) As String  '旧passwordGet
    Dim ii As Integer, aa As Integer, bb As String
    'PW文字数は２文字以上ないとエラーになる。 変数二文字化　86_014k
    For ii = 4 To knt
        aa = 0
        Do Until aa = 1 'Randomize
            bb = rndchr("nsuu")
            If InStr(1, pasuwardo, bb) = 0 Then aa = 1
        Loop
        pasuwardo = pasuwardo & bb
    Next ii
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function rndchr(suei As String) As String
    Dim aa As Long, bb As Integer, san As Integer
    '乱数(数字英字)取得
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
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub oshimai(msx As String, fn As String, ff As String, ii As Long, aa As Long, msbv As String)
         '↑msxは使われてないですね→再使用へ　201912
    Unload UserForm1
    Unload UserForm3
    Workbooks(fn).Activate '30s76追加（外結ログ対策）
    DoEvents
    Worksheets(ff).Select  '30s76追加（外結ログ対策）
    If aa > 0 And ii > 0 Then
        Workbooks(fn).Sheets(ff).Cells(ii, aa).Select
        If msx = "" Then
            If msbv <> "" Then MsgBox msbv & vbCrLf & "(選択セル)"
        Else
            If msbv <> "" Then MsgBox msbv & vbCrLf & "(選択セル)", 289, msx
        End If
    Else
        If msx = "" Then
            If msbv <> "" Then MsgBox msbv
        Else
            If msbv <> "" Then MsgBox msbv, 289, msx
        End If
    End If
    '名前の定義の削除
    Dim nmn As Name
    For Each nmn In ActiveWorkbook.Names
        On Error Resume Next  ' エラーを無視。
        nmn.Delete
    Next
    'フィルタ再
    If k > 1 Then bfshn.Rows(k - 1).AutoFilter         '一度つけて、
    
    Application.Calculation = xlCalculationAutomatic  '再計算自動に戻す
    Application.StatusBar = False
    If aa > 0 And ii > 0 Then Workbooks(fn).Sheets(ff).Cells(ii, aa).Select
    Application.Cursor = xlDefault
    End
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub fukusya()  ' 選択範囲を別シートにコピー　　旧「複写」
    Dim a, j, x As Integer   'g1→gg1,g→gg2(ローカル撤廃,グローバル化)　30s74
    Dim i As Long, jj(2) As Long, kk
    Dim shemei, wd, dg1, dg2 As String, hk1 As String, pasu As String
    Dim se_name As String, fimei As String, c99 As String
    
    'おまじない(「コードの実行が中断されました」対処)
    Application.EnableCancelKey = xlDisabled
    pasu = ActiveWorkbook.Path
    
    kyosydou
    If dd1 = 0 Then Call oshimai("", bfn, shn, 1, 0, "dd1がゼロです")
    
    'その日の初回チェック
    j = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(j, 1).Value = ""
        j = j + 1
        If j = 50000 Then
            MsgBox "空白行が見つからないようです"
            Exit Sub
        End If
    Loop
    If Workbooks(twn).Sheets(shog).Cells(j - 1, 4).Value <> Val(Format(Now(), "yyyymmdd")) Then
        Call oshimai("", bfn, shn, 1, 0, "その日の初回は、最初に[初回]ボタンを押して下さい。")
    End If
    
    If bfshn.Cells(sr(0) - 1, 5) = "" Then
        Call oshimai("", bfn, shn, sr(0) - 1, 5, "集計名を入力して下さい(緑色セル)")
    End If

    Call iechc(hk1)  '旧igchc(hk1)
    
    a = Len(bfn)
    zikan = Format(Now(), "yymmdd_hhmmss")
    shemei = bfshn.Cells(1, 3).Value & "_" & zikan
    se_name = bfshn.Cells(1, 3).Value
    If bfshn.Cells(1, 5).Value = "←" Then '30s81
        fimei = shemei
    ElseIf bfshn.Cells(1, 5).Value <> "" Then
        fimei = bfshn.Cells(1, 5).Value & "_" & zikan
    End If
    
    '再計算を自動に
    Application.Calculation = xlCalculationAutomatic
    j = 6
    i = 19967
    Do Until j >= mghz
        If bfshn.Cells(2, j).Value <> "" Then
            If Not IsError(Application.Match(bfshn.Cells(2, j).Value, Range(bfshn.Cells(2, j + 1), bfshn.Cells(2, mghz)), 0)) Then
                Range(bfshn.Cells(2, Application.Match(bfshn.Cells(2, j).Value, Range(bfshn.Cells(2, j + 1), bfshn.Cells(2, mghz)), 0) + j), bfshn.Cells(2, Application.Match(bfshn.Cells(2, j).Value, Range(bfshn.Cells(2, j + 1), bfshn.Cells(2, mghz)), 0) + j)).Select
                Call oshimai("", bfn, shn, 2, Int(j), "文字重複セルあり（" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & "セル）")
            End If
        End If
        If bfshn.Cells(2, j).Value <> "" And IsNumeric(bfshn.Cells(3, j).Value) Then '30s76 若干改良(３行目文字の時は無視)
            If bfshn.Cells(3, j).Value > i Then i = bfshn.Cells(3, j).Value
        End If
        j = j + 1
        Application.StatusBar = "重複確認、" & Str(j) & " / " & Str(mghz) '9s
    Loop
    i = i + 1
    Application.StatusBar = False
    If gg2 <= 3 Then
      '---Unicode部ここから---   '再計算を手動に
        Application.Calculation = xlCalculationManual
        For j = dd1 To dd2
            Application.StatusBar = "埋め込み中、" & Str(j - dd1 + 1) & " / " & Str(dd2 - dd1 + 1) '9s
            If gg2 = 3 And bfshn.Cells(2, j).Value = "" And bfshn.Cells(3, j).Value = "" Then '30s76
                bfshn.Cells(3, j).Value = i
                i = i + 1
            End If
            If gg1 <= 2 And bfshn.Cells(2, j).Value = "" And bfshn.Cells(3, j).Value <> "" Then
                If IsError(Application.Match(ChrW(bfshn.Cells(3, j).Value), Range(bfshn.Cells(2, 6), bfshn.Cells(2, mghz)), 0)) Then
                    bfshn.Cells(2, j).Value = ChrW(bfshn.Cells(3, j).Value)
                Else
                    Range(bfshn.Cells(2, Application.Match(ChrW(bfshn.Cells(3, j).Value), Range(bfshn.Cells(2, 6), bfshn.Cells(2, mghz)), 0) + 5), bfshn.Cells(2, Application.Match(ChrW(bfshn.Cells(3, j).Value), Range(bfshn.Cells(2, 6), bfshn.Cells(2, mghz)), 0) + 5)).Select
                    Call oshimai("", bfn, shn, 2, Int(j), "挿入予定文字「" & ChrW(bfshn.Cells(3, j).Value) & "」：" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & "セルと重複")
                End If
            ElseIf gg1 = 3 And bfshn.Cells(2, j).Value <> "" And bfshn.Cells(3, j).Value = "" Then
                bfshn.Cells(3, j).Value = AscW(bfshn.Cells(2, j).Value)
                If bfshn.Cells(3, j).Value < 0 Then bfshn.Cells(3, j).Value = bfshn.Cells(3, j).Value + 65536
                bfshn.Cells(3, j).Value = "手" & bfshn.Cells(3, j).Value
            End If
        Next
        '再計算を自動に
        Application.Calculation = xlCalculationAutomatic
        bfshn.Cells(3, 3).Value = i
        Application.StatusBar = False
      '---Unicode部ここまで---
    ElseIf gg2 = sr(8) And gg1 = sr(8) Then
       '8行コメントコピー 30s86_012_a
        If dd1 = dd2 And bfshn.Cells(sr(6), dd1).Value = -99 Then
            If TypeName(ActiveCell.Comment) = "Comment" Then 'コメント有りの場合
                c99 = ActiveCell.Comment.Text  'コメント内容
                If Left(c99, 1) = "=" Then
                    If MsgBox("セルをコメントに掲載の式:" & vbCrLf & c99 & vbCrLf & "に、置き換えていいですか？", 289, "事前かくにん") = vbOK Then 'ok時
                        c99 = Replace(c99, "ｦｦ", "=")
                        ActiveCell.Value = c99
                        ActiveCell.Replace What:="ｦｦ", Replacement:="=", LookAt:=xlPart, _
                            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                            ReplaceFormat:=False
                    End If
                Else
                    MsgBox "実施対象外"
                End If
            End If
        End If
    Else  '本来
        '上側からこちらに引っ越し　86_014e
        If bfshn.Cells(1, 3) = "" Then
            Call oshimai("", bfn, shn, 1, 3, "複写名を入力して下さい(紫色セル)")
        ElseIf Len(bfshn.Cells(1, 3)) > 14 Then
            Call oshimai("", bfn, shn, 1, 3, "複写名は14文字以内に収めて下さい(" & Len(bfshn.Cells(1, 3)) & ")")
        End If
                
        'Aa：当シート、ここから（メイン部）
        Selection.Copy
        'Aa：データ開始行採取（後のフィルタリング、ウィンドウの固定のため）
        j = 1
        Do Until bfshn.Cells(j, 1).Value = 1 Or bfshn.Cells(j, 1).Value = "all1"
            j = j + 1
            If j = 200 Then
                MsgBox "「1」が見つからないようです"
                Exit Sub
            End If
        Loop
        If bfshn.Cells(j, 1).Value = 1 Then k = j Else k = j + 1 'kはデータ開始行(サンプル行ではなくなった)
        'Aa：当シート、ここまで
        'Ab：新シート、ここから
        Worksheets.Add  '新シート生成
        ActiveSheet.Name = shemei
    
        Range(Cells(2, 3), Cells(2, 3)).Select
        'Ab：貼り付け（値と数値の書式で貼付）
        Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats '85_24検証3から12→数式と数値の書式で貼付 へ
        
        'Ab：貼り付け（コメントを貼付）30s79
        Selection.PasteSpecial Paste:=xlPasteComments  '-4144
        'Ab：書式を貼り付け
        Selection.PasteSpecial Paste:=xlPasteFormats '-4122
        
        'Ab：罫線編集
        Selection.Borders.Color = RGB(191, 191, 191) '←= -4210753 　　30s86_017p
            
        '85_024検４当シート
        If gg1 < k Then
            If gg1 > sr(8) Then x = gg1 Else x = sr(8) + 1 'まずここでx使用(値のコピー開始行)
            bfshn.Select
            Range(Cells(x, dd1), Cells(k - 1, dd2)).Copy
            
            'Bbop：新シート
            Workbooks(bfn).Sheets(shemei).Select
            Range(Cells(x - gg1 + 2, 3), Cells(x - gg1 + 2, 3)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats  '12 'Bb1op：貼り付け（値と数値の書式で貼付）
        End If
            
        Range(Cells(1, 1), Cells(1, dd2 - dd1 + 3)).Select  '一行目セルうす緑に
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(179, 255, 231) 'うす緑　元15204275　30s86_017p
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
        Range(Cells(1, 2), Cells(1, dd2 - dd1 + 3)).Select  '一行目unicode文字色ほぼ透明に
        With Selection.Font
            .Color = RGB(226, 239, 218)  '元-2428958 30s86_017p
            .TintAndShade = 0
        End With
    
        Cells.FormatConditions.Delete      '条件付き書式解除
    
        x = k - gg1 + 1 'Ab：x：転載シートのフィルタリング基準行

        If gg2 < k - 1 Then x = gg2 - gg1 + 1 + 1 '特例パターン
        
        If x > 0 Then  'x>0でないパターンはないかと。ｘは２以上
            Range(Cells(2, 1), Cells(x, 2)).Select  'Ab：1列2列タテ上部をうす緑に
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(179, 255, 231) 'うす緑　'元15204275　 30s86_017p
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        
        Selection.ColumnWidth = 2  'Ab：1列2列幅調整
        Range(Cells(1, 2), Cells(1, 2)).Select
        Cells(1, 2).Value = "."   'if文撤廃、30s74
    
        x = k - gg1 + 1 'Ab：x：転載シートのフィルタリング基準行(再)
    
        'Ab：新シート、ここまで
        If bfshn.Cells(1, 5).Value <> "" Then '不要かと625
            'dw = ""  '625以降新バージョン
            '↓30s86_019j、30s86_019m
            If bfshn.Cells(2, 3).Value <> Left(twn, 7) And bfshn.Cells(2, 3).Value <> syutoku() & "r" And bfshn.Cells(2, 3).Value <> "" Then
                dw = bfshn.Cells(2, 3).Value  '625以降新バージョン
            Else 'これまでのパターン
                dw = pasuwardo(10)
            End If
            
            UserForm2.Show vbModal 'PWありボタン押した場合、dwにPWが格納される。
            wd = dw
            If Application.Version < 16 And fmt = "csv" Then
                Call oshimai("", bfn, shn, 1, 0, "excel2013以前はcsv(utf-8)で保存はできません)")
            End If
        End If
        'Ba：当シート、左端タテ系部
        bfshn.Select
    
        If x > 1 And k - gg2 < 1 Then  'k - gg2 < 2→1 (項目行下端は対象外に)
    
            Range(Cells(gg1, 1), Cells(gg2, 2)).Copy     '1列2列タテ大雑把コピー(all1とカウンタcの所)
    
            'Bbop：新シート
            Workbooks(bfn).Sheets(shemei).Select '１列２列大雑把貼り付け(装飾系あとで)
            Range(Cells(2, 1), Cells(2, 1)).Select
            'Bb1op：貼り付け（値と数値の書式で貼付）
            Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats  '12
            With Selection.Font
                .Color = RGB(128, 128, 128) '30s86_017p ねずみ色に(=8421504)
                .TintAndShade = 0
            End With
    
            'Bb2op：オートフィルタ設定
            Rows(x).AutoFilter
            
            'Bb2op：ウィンドウ枠の固定
            Range(Cells(x + 1, 3), Cells(x + 1, 3)).Select
            ActiveWindow.FreezePanes = True
        
            Range(Cells(2, 1), Cells(x, 2)).ClearContents   '30s74　Bb2op：1列2列タテ上部をうす緑に
            
            Cells(x, 1).Value = "項目"
            Cells(x, 2).Value = "ｃ：集計対象　　"
        
            Range(Cells(x, 1), Cells(x, 2)).Select
            With Selection 'Bb2op：1列2列の項目行は上寄せ
                .VerticalAlignment = xlTop
                .Orientation = -90
            End With
            Selection.Font.Size = 9 'Bb2op：1列2列の項目行フォント調整
       
            If x > 3 Then
                Range(Cells(2, 1), Cells(x - 1, 1)).Select 'Bb2op：1列2列の上部フォントほぼ透明化
                With Selection.Font
                    .Color = RGB(179, 255, 231) '←　= -1572941　30s86_017p
                    .TintAndShade = 0
                End With
            End If
    
            If gg1 < sr(0) + 6 Then  'subtotalコピペ　N/A対応もコピペへ
                
                'subtotal事前
                'i(mghzコピペ開始用)定義(gg1：選択範囲開始行に左右される)
                
                If gg1 = sr(0) + 5 Then 'Bb2op 30s77改良
                    i = sr(0) + 4 '3行特殊
                    jj(1) = 1
                ElseIf gg1 < sr(0) + 5 Then
                    i = sr(0) + 3  '４行標準
                    jj(1) = i - gg1 + 2 + 1
                End If
        
                'subtotal本番、Bb2op_opC
                
                For j = mghz + 1 To mghz Step -1  '85_024検証9
                    'Bb2op_opCa：ここから、当シート(subtotalコピペ)
                    bfshn.Select     'B：右端subtotalコピー
                    Range(bfshn.Cells(i, j), bfshn.Cells(sr(0) + 6, j)).Copy
                
                    'Bb2op_opCb：ここから、新シート（subtotalコピペ）
                    Workbooks(bfn).Sheets(shemei).Select

                    If gg1 > sr(0) + 3 Then
                        Range(Cells(1, 2), Cells(1, 2)).Select 'はみ出しペースト
                    Else
                        Range(Cells(i - gg1 + 2, 2), Cells(i - gg1 + 2, 2)).Select '標準ペースト
                    End If
                    ActiveSheet.Paste
                
                    Range(Cells(jj(1), 2), Cells(jj(1) + 1, 2)).Columns.AutoFit
                
                    For jj(0) = jj(1) To jj(1) + 1
                        jj(2) = Range(Cells(jj(0), 2), Cells(jj(0), 2)).Font.Color
                        
                        Range(Cells(jj(0), 2), Cells(jj(0), 2)).Copy
                                                                
                        For ii = jj(1) To jj(1) + 1
                            For kk = 3 To dd2 - dd1 + 3
                                If WorksheetFunction.IsErr(Cells(ii, kk).Value) Then 'N/A条件追加,DIV/0条件追加（IsNA→IsErr）30s86_020x
                                    If Cells(ii, kk).Font.Color = jj(2) Then
                                        Range(Cells(ii, kk), Cells(ii, kk)).Select
                                        ActiveSheet.Paste
                                    End If
                                ElseIf Cells(ii, kk).Font.Color = jj(2) And Cells(ii, kk) <> "" Then '30s77null条件追加(従来)
                                        Range(Cells(ii, kk), Cells(ii, kk)).Select
                                        ActiveSheet.Paste
                                End If
                            Next
                        Next
                    Next
                Next
            'Bb2op_opCここまで
            End If
        'Bb2opここまで
        End If
        
        Range(Cells(2, 1), Cells(2, 1)).Select
        'Bb：新シート、ここまで
    
        'D：ここから、当シート(体裁調整)
        'Da
        bfshn.Select
    
        Range(Cells(1, 1), Cells(sr(0) + 5, mghz - 1)).Select  'Da1：1-18行(標準)を調整
        With Selection.Font  'Da：フォント調整
            .Name = "ＭＳ Ｐゴシック"  'Da：←これでWin10影響(メイリオなど)影響受けなくなる。
            .Size = 9
        End With
    
        Range(Cells(sr(0) + 3, 5), Cells(sr(0) + 3, mghz - 1)).Select 'Da1：標15行調整
        With Selection.Font  'Da：フォント、色調整
            .Name = "Haettenschweiler"
            .Size = 10
        End With
        Selection.NumberFormatLocal = "m/d h:mm"
    
        Range(Cells(sr(0), 5), Cells(sr(0), mghz - 1)).Select  'Da1：標12行調整
        With Selection.Font  'フォント、色調整
            .Name = "Haettenschweiler"
            .Size = 10
        End With
        Selection.NumberFormatLocal = "m/d h:mm"
        
        'Da1：前歴集計値調                                              '赤が出る事あり注意↓
        Range(Cells(sr(0) + 1, 5), Cells(sr(0) + 2, mghz - 1)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        'Da1：前歴集計値調                                              '赤が出る事あり注意↓
        Range(Cells(sr(0) + 4, 5), Cells(sr(0) + 5, mghz - 1)).NumberFormatLocal = "#,##0;[赤]-#,##0"
    
        '↓86_014m
        Range(Cells(sr(0), 3), Cells(sr(0) + 1, 5)).Select 'Da1：赤セル部フォント再調整1＆2
        With Selection.Font
            .Name = "ＭＳ Ｐゴシック"
            .Size = 9
        End With
        Selection.NumberFormatLocal = "G/標準" '←withに入れるとエラーになる。
    
        'Da2：当シート列幅をコピー
        Range(Cells(2, dd1), Cells(2, dd2)).Copy '列幅コピる（2行目にて）
        
        Range(Cells(gg1, dd1), Cells(gg1, dd1)).Select  '開始セル（：左上）
        dg1 = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) '開始セル転記
        Range(Cells(gg2, dd2), Cells(gg2, dd2)).Select  '終了セル（：右下）
        dg2 = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) '終了セル転記
        Range(Cells(gg1, dd1), Cells(gg2, dd2)).Select '選択範囲は戻す
        'Da：当シート、ここまで
        
        'Db：ここから、新シート（列幅のみコピペ）
        Workbooks(bfn).Sheets(shemei).Select
        
        Range(Cells(1, 3), Cells(1, 3)).PasteSpecial Paste:=xlPasteColumnWidths '8'列幅を貼り付け
    
        If x > 1 And k - gg2 < 2 Then  'Dbop_Unicodeの貼り付け(項目行(k-1)が含まれていることが貼り付けの条件)
        'Dbop：x：フィルタリング基準行
            Range(Cells(1, 3), Cells(1, 3)).PasteSpecial Paste:=xlPasteValues '-4163
        End If
        
        'Db
        Cells(1, 1).Value = "複.　複写、" & twn & "、" & shemei & "、ｦ" & shn _
        & "ｦ" & bfn & "、" & dg1 & dg2 & "、列固有名"
        Range(Cells(2, 2), Cells(2, 2)).Select
    
        'ログ部ここから→kyosydouへ移設(Activate処理撤廃)
        j = 1
        Do Until Workbooks(twn).Sheets(shog).Cells(j, 1).Value = ""
            j = j + 1
            If j = 10000 Then
                MsgBox "空白行が見つからないようです"
                Exit Sub
            End If
        Loop

        Workbooks(twn).Sheets(shog).Cells(j, 1).Value = 1 '項目名
        Workbooks(twn).Sheets(shog).Cells(j, 2).Value = j  '項番
        
        Workbooks(twn).Sheets(shog).Cells(j, 3).Value = Workbooks(bfn).Sheets(shemei).Cells(1, 1).Value & "、z" & wd & "z" & "b" _
        & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "、" & twbsh.Cells(2, 2).Value & "、" & _
        Format(Now(), "yyyymmdd_hhmmss") & "、" & bfshn.Cells(sr(8), 5).Value & "、" & gg2 - gg1 + 1 & "、" & dd2 - dd1 + 1
        
        Workbooks(twn).Sheets(shog).Cells(j, 4).Value = Format(Now(), "yyyymmdd") 'date
        Workbooks(twn).Sheets(shog).Cells(j, 5).Value = Format(Now(), "yyyymmdd_hhmmss") 'timestamp
        If fimei = "" Then
            Workbooks(twn).Sheets(shog).Cells(j, 7).Value = shemei 'to
        Else
            Workbooks(twn).Sheets(shog).Cells(j, 7).Value = fimei & "." & fmt & "\" & shemei 'to
        End If
        Workbooks(twn).Sheets(shog).Cells(j, 8).Value = 9  '最右列(複写は固定値)
        Workbooks(twn).Sheets(shog).Cells(j, 9).Value = bfn & "\" & shn  'from
        
        'ログ部ここまで
 
        'シートを別ファイルとしても複製(E1セル情報有の時)30s81
        If bfshn.Cells(1, 5).Value <> "" Then  '抜本改良　30s86_017x~z
            '627メモ帳のPW→Excel新シートへ
            If Left(fmt, 3) <> "csv" Then 'csv系はPWシート作らない xlsxPWなしも一旦PWシートは作る(office2007にファイルのテーマ変えるため)
                
                ThisWorkbook.Activate
                Sheets("高速シート_" & syutoku()).Select
                Sheets("高速シート_" & syutoku()).Copy

'                Sheets("高速シート_" & syutoku()).Name = shemei & "のPW"
                Sheets("高速シート_" & syutoku()).Name = "仮ですaaa"    '30s86_021b
                Worksheets.Add
                ActiveSheet.Name = shemei & "のPW"
                
                Application.DisplayAlerts = False
                Worksheets("仮ですaaa").Delete
                Application.DisplayAlerts = True

                '86_020v　改良
                Cells(5, 1).Value = "-------------------------------------------------------------------------------"
                
                Cells(7, 1).Value = "ファイル名： " & fimei & "." & fmt
                Cells(8, 1).Value = "ファイルのＰＷ：" & wd
                Cells(9, 1).Value = "シート名：" & shemei
                
                Cells(11, 1).Value = "-------------------------------------------------------------------------------"
                
                Range(Cells(2, 1), Cells(19, 1)).Select

'                If wd <> "" Then Shell "c:\windows\system32\notepad.exe", vbNormalFocus 'PW用メモ帳立ち上げ 86_020z　廃止へ
                '30s79追加、新ファイルのフォントを游ゴシックではなく、MSPゴシック仕様に
                If Application.Version < 16 Then
                    'MsgBox "excel2013以前です"
                Else                  'MsgBox "excel2016以降です"
                    If IsNumeric(syutoku()) Then
                        '↓レノボのPC
                        ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                        "C:\Program Files (x86)\Microsoft Office\Root\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml")
                    Else
                        '↓fmvのPC
                        'ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                            "C:\Program Files\WindowsApps\Microsoft.Office.Desktop_16051.12228.20364.0_x86__ああああ\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml" _
                            )
                        '20200115にRev上がったかと。まだこの頃は、store版
                        'ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                            "C:\Program Files\WindowsApps\Microsoft.Office.Desktop_16051.12325.20288.0_x86__ああああ\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml" _
                            )
                        '202004のRev　(store版からDL版へ)
                        ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                            "C:\Program Files (x86)\Microsoft Office\Root\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml" _
                            )
                    End If
                End If
            Else 'csv系
                Workbooks(bfn).Activate
                Sheets(shemei).Select
                Sheets(shemei).Copy
            End If
            
            'ファイルを所定の様式・所定のファイル名で保存
            If fmt = "xlsx" Then
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlOpenXMLWorkbook, Password:=wd, CreateBackup:=False    'xlsxフォーマット
            ElseIf fmt = "csv" Then
                Rows("1:1").Select
                Selection.Delete Shift:=xlUp
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlCSVUTF8, CreateBackup:=False                          'csv(utf-8)フォーマット
            ElseIf fmt = "csvsjis" Then
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlCSV, CreateBackup:=False                              'csvフォーマット
            Else
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlExcel12, Password:=wd, CreateBackup:=False            'xlsbフォーマット
            End If
            
            If Left(fmt, 3) <> "csv" Then
            
                'Eb　（従来：対象のシートを新規ファイルとしてコピーしていた）
                Workbooks(bfn).Activate
                Sheets(shemei).Select
                Sheets(shemei).Copy Before:=Workbooks(fimei & "." & fmt).Sheets(shemei & "のPW")
            
                Workbooks(fimei & "." & fmt).Activate
                Sheets(shemei & "のPW").Select
                Sheets(shemei & "のPW").Move 'PWシートは新ファイルとして移す
            
                If wd = "" Then
                    Sheets(shemei & "のPW").Name = "不要"  '※PWなしxlsx→捨てファイル
                Else '30s86_020a PWシートをUTF-8のテキストとして保存を追加(メモ帳立ち上げは廃止)
                    ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei & "のPW.txt", _
                    FileFormat:=xlUnicodeText, CreateBackup:=False
                End If
            
                Workbooks(fimei & "." & fmt).Activate
                ActiveWorkbook.Save      '上書保存
            
            End If
        End If
    End If
    Application.CutCopyMode = False
End Sub  '複写ここまで
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function chekku() As String
    'Excelバージョンチェック,　86_020s挿入部ここから　020u改良
    twt.Cells(1, 1).Value = "あ"  '詳細記述は初回のみ（）の方で実施
    twt.Cells(2, 1).Value = "ア"
    
    twt.Sort.SortFields.Clear
    twt.Sort.SortFields.Add Key:=Range(twt.Cells(1, 1), twt.Cells(1, 1)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With twt.Sort
        .SetRange Range(twt.Cells(1, 1), twt.Cells(2, 1))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    If twt.Cells(1, 1).Value = "ア" Then
        chekku = "新ｿｰﾄ" '(excel2019のあるver以降)"
    Else
        chekku = "旧ｿｰﾄ" '順序そのまま(excel2016以前、excel2007)"
    End If
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub syokainomi()  '旧「初回のみ」
    Dim hk1 As String, mghx As Long, xlvrn As String
    
    kyosydou  '共通初動
 
    xlvrn = chekku 'Excelバージョンチェック
    twt.Cells(3, 1).Value = "ｿｰﾄ後"
    twt.Cells(1, 2).Value = "あ"
    twt.Cells(2, 2).Value = "ア"
    twt.Cells(3, 2).Value = "ｿｰﾄ前"
    twt.Cells(1, 3).Value = "ア"
    twt.Cells(2, 3).Value = "あ"
    twt.Cells(3, 3).Value = "新仕様(片仮名が上へ)"
    twt.Cells(1, 4).Value = "あ"
    twt.Cells(2, 4).Value = "ア"
    twt.Cells(3, 4).Value = "旧仕様(片・平同一視)"
    twt.Cells(1, 6).Value = "↓textjoin結果"  '30s86_021e
    twt.Cells(2, 6).Value = "ｦｦTEXTJOIN(""、"",TRUE,A2,B2)"
    twt.Cells(3, 6).Value = "↑「あ、ア」なら正常"
    twt.Cells(4, 6).Value = "　「#NAME?」ならこのOfficeではサポートされていない(Excel再起動でサポートされることもある)"
    'ｦｦ→=
    Range(twt.Cells(2, 6), twt.Cells(2, 6)).Replace What:="ｦｦ", _
        Replacement:="=", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    If hrkt = 16 Then MsgBox "hrkt=16(phonetic使用[従来型])です。キーに「ヶ」があると重くなり、誤作動の可能性あり。2013でも2019でも"
    If xlvrn = "旧ｿｰﾄ" And hrkt = 0 Then MsgBox "旧ｿｰﾄ(2013)でhrkt=0(phonetic不使用)です。キーに同じひらがなカタカナある時(「あ」「ア」など)誤作動の危険あり注意。"
    
    If Not (shn = "▲集計_雛形" And bfn = twn) Then
        bfshn.Cells(sr(0) + 1, 3).Value = ""    '13→sr(0)+1
        bfshn.Cells(sr(0) + 2, 3).Value = ""     '14→sr(0)+2
    End If

    Call iechc(hk1)

    mghx = Application.Match("。", Range(twbsh.Cells(1, 1), twbsh.Cells(1, 5000)), 0) 'mghxはマクロファイルの「。」の列,mghzは当シートの、
    
    jj = 6 'j→jj 30s82
    ii = 19967 'i→ii 30s82
    Do Until jj >= mghz
        If bfshn.Cells(2, jj).Value <> "" Then
            If Not IsError(Application.Match(bfshn.Cells(2, jj).Value, Range(bfshn.Cells(2, jj + 1), bfshn.Cells(2, mghz)), 0)) Then
                Range(bfshn.Cells(2, Application.Match(bfshn.Cells(2, jj).Value, Range(bfshn.Cells(2, jj + 1), bfshn.Cells(2, mghz)), 0) + jj), bfshn.Cells(2, Application.Match(bfshn.Cells(2, jj).Value, Range(bfshn.Cells(2, jj + 1), bfshn.Cells(2, mghz)), 0) + jj)).Select
                Call oshimai("", bfn, shn, 2, Int(jj), "文字重複セルあり（" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & "セル）")
            End If
        End If
        If bfshn.Cells(2, jj).Value <> "" Then '17s
            If bfshn.Cells(3, jj).Value > ii Then ii = bfshn.Cells(3, jj).Value
        End If
        jj = jj + 1
        Application.StatusBar = "重複確認、" & Str(jj) & " / " & Str(mghz) '9s
    Loop

    '再計算を一旦自動に
    Application.Calculation = xlCalculationAutomatic
    Application.ExtendList = False 'データ範囲拡張:オフ（：右隣セルが勝手に書式変わられるのを阻止）"
    
    'オートフィルタが設定されているかどうか判断＆解除
    If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    Application.EnableAutoComplete = False  'オートコンプリート
       
    ThisWorkbook.Activate
    
    jj = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = ""
        jj = jj + 1
        If jj = 10000 Then
            MsgBox "空白行が見つからないようです"
            Exit Sub
        End If
    Loop
    
    Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = 1 '項目名
    Workbooks(twn).Sheets(shog).Cells(jj, 2).Value = jj  '項番
    
    Workbooks(twn).Sheets(shog).Cells(jj, 3).Value = _
    "初.　初回、" & _
    twn _
    & "、ｦ" & shn & "ｦ" & bfn _
    & "、ｦ▲集計_雛形" & "ｦ" & twn _
    & "、from" & dd1 & "to" & dd2 _
    & "、項目名、b" & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "、" & twbsh.Cells(2, 2).Value & "、" _
    & Format(Now(), "yyyymmdd_hhmmss") & "、" & bfshn.Cells(sr(8), 5).Value & "、0、0"
    Workbooks(twn).Sheets(shog).Cells(jj, 4).Value = Format(Now(), "yyyymmdd") 'date
    Workbooks(twn).Sheets(shog).Cells(jj, 5).Value = Format(Now(), "yyyymmdd_hhmmss") 'timestamp
    Workbooks(twn).Sheets(shog).Cells(jj, 7).Value = bfn & "\" & shn 'to
    Workbooks(twn).Sheets(shog).Cells(jj, 8).Value = 9  '最右列(複写は固定値)
    Workbooks(twn).Sheets(shog).Cells(jj, 9).Value = twn & "\▲集計_雛形"   'from
    'ログ部ここまで
    
    Call cpp2(twn, "▲集計_雛形", 14, 1, 18, 5, bfn, shn, sr(0) - 1 + 5 - 2, 1, 0, 0, -4104) 'サマリ関数周辺丸ごとコピペなので４１０４
    Call cpp2(twn, "▲集計_雛形", 15, mghx, 18, mghx + 2, bfn, shn, sr(0) - 1 + 5 - 1, mghz, 0, 0, -4104) '同、サマリ関数周辺丸ごとコピペ(mghz側)
    Call cpp2(twn, "▲集計_雛形", 11, 5, 15, 6, bfn, shn, sr(0) - 1, 5, 0, 0, -4122)  '同、サマリ関数周辺丸ごとコピペ(mghz側)
    
    Workbooks(bfn).Activate '↑の.copy後、これをここに入れると、セルが複数個所選択されている妙な映りは解消されるっぽい
    Sheets(shn).Select
    
    Range(bfshn.Cells(2, mghz - 2), bfshn.Cells(21, mghz)).Borders.Color = RGB(191, 191, 191)  '←　=-4210753　 30s86_017p　'30s86_012i
    
    '再計算を手動に
    Application.Calculation = xlCalculationManual

    With Application.AutoCorrect      'オートコレクトさせない　３０ｓ５２
        .TwoInitialCapitals = False
        .CorrectSentenceCap = False
        .CapitalizeNamesOfDays = False
        .CorrectCapsLock = False
        .ReplaceText = False
        .DisplayAutoCorrectOptions = True
    End With

    bfshn.Cells(sr(0) - 1, 5).Select   '緑色セル
    Range(bfshn.Cells(gg1, dd1), bfshn.Cells(gg2, dd2)).Select '選択範囲は戻す bfshn被せた
    Call oshimai(syutoku(), bfn, shn, 1, 0, "初回処理完了")
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub betat4(fbk As String, fsh As String, fmg1 As Long, fmr1 As Currency, fmg2 As Long, fmr2 As Currency, tbk As String, tsh As String, tog As Long, tor As Long, er34 As String, mr_8 As String)
    'e ver
    Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).ClearContents  'betatnはcppと異なり、仕様として、クリアすることとする。
    If er34 = "pp" Then '標準型調整
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).NumberFormatLocal = "G/標準"
    ElseIf er34 = "mm" Then '文字列型調整
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).NumberFormatLocal = "@"
    ElseIf er34 = "pm" Then '通貨型調整
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).NumberFormatLocal = "#,##0;[赤]-#,##0"
    Else
        'mp
    End If
    
    If Abs(fmr1) = 0.4 Or Abs(fmr1) = 0.1 Then 'ｦｦ定文字対応、ｦーｦ行番号対応　（フィル型）
        'こちらにも(ｦｦ、0.4向け)
        If Abs(fmr1) = 0.1 Then 'ｦーｦ行番号対応　※暫定運用
            Workbooks(tbk).Sheets(tsh).Cells(tog, tor).Value = Trim$(mr_8) & Format(fmg1, "0000000") '86_019a
            If fmg2 > fmg1 Then '範囲が1行か2行しか無い場合の対処(以下同文)
                If mr_8 <> "" And mr_8 <> "000_" Then MsgBox "betat4挙動注意bb " & mr_8 '30s86_019a
                Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor).Value = Trim$(mr_8) & Format(fmg1 + 1, "0000000")
            'ここ、高速シートの2列目をコピーすることも検討し得る。
            End If
            If fmg2 > fmg1 + 1 Then  'フィルは３行以上ある場合のみ実施
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor)).AutoFill Destination:=Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor))
            End If
        Else 'ｦｦ定文字対応(通常) (0.4)
            Workbooks(tbk).Sheets(tsh).Cells(tog, tor).Value = Trim$(mr_8)
            If fmg2 > fmg1 Then
                Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor).Value = Trim$(mr_8)
            End If
            If fmg2 > fmg1 + 1 Then  'フィルは３行以上ある場合のみ実施
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor)).AutoFill Destination:=Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor))
            End If
        End If
    Else '一般対応時、含複数列ベタ型
        If er34 = "mp" Then  'セル踏襲
            Call cpp2(fbk, fsh, fmg1, Int(fmr1), fmg2, Int(fmr2), tbk, tsh, tog, tor, 0, 0, 12) '12:値と数値の書式 '遅・コピペパターン
        Else 'mm,pp,pm
            Call cpp2(fbk, fsh, fmg1, Int(fmr1), fmg2, Int(fmr2), tbk, tsh, tog, tor, 0, 0, -4163) '-4163:値
        End If
    End If
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub copipe(fbk As String, fsh As String, fmg1 As Long, fmr1 As Long, fmg2 As Long, fmr2 As Long, tbk As String, tsh As String, tog1 As Long, tor1 As Long, tog2 As Long, tor2 As Long, mdo As Long)
    'mdo: 1:value 、2:formula、3:formulaR1C1
    'クリップボードを使用しないコピペの実現
    If mdo = 1 Then  'Value
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)) = _
          Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Value
    ElseIf mdo = 2 Then
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)) = _
          Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Formula
    ElseIf mdo = 3 Then '-4123向け
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)).FormulaR1C1 = _
          Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).FormulaR1C1 '-4123向け
    Else
        Call oshimai("", bfn, shn, 1, 0, "まだ造成中d")
    End If
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub cpp2(fbk As String, fsh As String, fmg1 As Long, fmr1 As Long, fmg2 As Long, fmr2 As Long, tbk As String, tsh As String, tog1 As Long, tor1 As Long, tog2 As Long, tor2 As Long, mdo As Long)
        
    'mdo:12　・・・値と数値の書式　(2002以降)　　xlPasteValuesAndNumberFormats　.copy.paste
    '　　　　-4123 (-99,samaru)　数式　xlPasteFormulas　　　.copy.paste
    '　　　　-4104 (ログ部)　すべて　xlPasteAll　　.copy.paste　・・・クリップボードに情報保持される(繰返し可能)
    '　　　　-4163 (値)　　xlPasteValues　　　　　.copy.pasteしない→だから早い
    'ー4163以外は速度遅い.Copy.Paste なので
    
    If tog2 = 0 And tor2 = 0 Then '従来パターン
        'コピペルーチン　30s85_004
        If mdo = -4163 Then  '新型速度早い。
            Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1 + fmg2 - fmg1, tor1 + fmr2 - fmr1)) = _
              Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Value
    
        Else '従来型(も改良へ) コピペなので遅い(数式コピペはこれ、避けられない)。
            UserForm3.StartUpPosition = 3 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
            UserForm3.Show vbModeless
            UserForm3.Repaint
            Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Copy
            Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1).PasteSpecial Paste:=mdo
            '(参考)↑.copyのコピーメソッドは、selectチックに範囲は指定される挙動である。
            Unload UserForm3
            UserForm1.Repaint
        End If
    ElseIf fmr2 <= 0 Then  '新パターン（一行を複数行にコピペ）-99や近似高速のｦｦASC(PHONETIC())とかで使われる。
        If fmr2 = 0 Then
            If mdo = -4163 Then  '新型速度早い。
                Call oshimai("", bfn, shn, 1, 0, "まだ造成中a")
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)) = _
                Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Value
            Else '従来型(も改良へ) コピペなので遅い(数式コピペはこれ、避けられない)。
                UserForm3.StartUpPosition = 3 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
                UserForm3.Show vbModeless
                UserForm3.Repaint
                Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1)).Copy
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)).PasteSpecial Paste:=mdo
                Unload UserForm3
                UserForm1.Repaint
            End If
        ElseIf fmr2 = -1 Then
            If tog1 > tog2 Then Call oshimai("", bfn, shn, 1, 0, "tog1がtog2よりでかいa")
            UserForm3.StartUpPosition = 3 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
            UserForm3.Show vbModeless
            UserForm3.Repaint
            Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1)) = fsh
            If tog2 > tog1 Then
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1)).Copy
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor1)).PasteSpecial Paste:=mdo 'tor2は使ってない
            End If
            Unload UserForm3
            UserForm1.Repaint
        '4163分岐なし
        ElseIf fmr2 = -2 Then  'フィル(連番のみ、固定フィルはやらないに）
          '4163分岐なし
            If tog1 > tog2 Then Call oshimai("", bfn, shn, 1, 0, "tog1がtog2よりでかいb")
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1)) = fmg2
                If tog2 > tog1 Then
                    Range(Workbooks(tbk).Sheets(tsh).Cells(tog1 + 1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1 + 1, tor1)) = fmg2 + 1
                    If tog2 > tog1 + 1 Then  'フィル
                        Range(bfshn.Cells(tog1, tor1), bfshn.Cells(tog1 + 1, tor1)).AutoFill Destination:=Range(bfshn.Cells(tog1, tor1), bfshn.Cells(tog2, tor1))   'tor2は使ってない
                    End If
                End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "どうなるか未定")
            End If
        Else
        Call oshimai("", bfn, shn, 1, 0, "まだ造成中b")
    End If
    '↓86_014c追加（excel2019 対策向けtest）
    Application.CutCopyMode = False
End Sub
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function kskup(am1 As String, am2 As String, n1 As Long, n2 As Long, h As Long, b As Currency, c As Currency, k0 As Long, h0 As Long, pap2 As Long, er2() As Currency, spd As String, pqp As Long, e5 As Long, er3() As Currency, hiru As Variant) As Long  'じあ
    'p(kskup)判定新仕様、n3は近似用　,(引数)m2→n1へ リファラ：n3、n2,h0,k0, pqp追加624,e5&er3()追加629
    Dim m As Long, n3 As Long
    Dim tskosk As Long 'strconv用(定速：2(従来通り)、高速：26(New85_006))
    Dim n3a As Long
    '◆高速側(純・近似)専用。低速側はtskupへ
    tskosk = 10 + hrkt '30s86_020s
    kskup = 0 'リセット(不要だが)
    n3 = 0 'ゼロスタート
    krpm2 = 0 '高速ロック＆p=-2判定フラグ

    If h < k Then '当シートに何も無い場合(初回のみ通過ゾーン) kあ
        If c < 0 Then
            kskup = -1    '-1-2この時点でexitdo(p:-1とする。)
            MsgBox "表空白の-1-2です(p=-1,exitdo、このまま終了されます)。"
        Else  '※c>=0が前提となる。
            kskup = 2
            n2 = h + 1
            
            If er2(0) < 0 Then  '高速ロックオン(表空白)です。"
                UserForm4.StartUpPosition = 2 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
                UserForm4.Show vbModeless
                UserForm4.Repaint
                bfshn.Cells(sr(2), 5).Value = "純高ﾛｯｸ(表空白)" 'そのうち
                If pap2 = 0 And Abs(e5) < 1 And UBound(er3()) > 0 Then 'p=-2の判定　６２９
                    For jj = 1 To UBound(er3())
                        If er3(jj) = 0.1 Then krpm2 = 1  'er3(jj)<0→=0.1 へ
                    Next
                    If krpm2 = 1 Then kskup = -2 'p=-2の確定　６２９
                End If
                Unload UserForm4
                UserForm1.Repaint
            End If
        End If
    ElseIf StrConv(am2, tskosk) = StrConv(am1, tskosk) Then '(アドバラピド)前回一致
        kskup = 1
        n2 = n1
    ElseIf spd = "純高速" And (k0 > h0 Or pqp = 1) Then  '(アドバラピド)高速ロックオン状態　pqp追加624
        If c < 0 Then
            kskup = -1    '-1-2この時点でexitdo(p:-1とする。)
            MsgBox "ここはもう通らないはず(-1-2)以後当シート側情報ないです(p=-1)。"
        Else
            kskup = 2
            n2 = h + 1
        End If
    '(アドバラピド)次行一致 c = Round(c, 0)追加85＿026（：-15.１実施しない、-15実施する）
    ElseIf Round(c) <> -1 And Round(c) <> -2 And c = Round(c, 0) And (spd = "純高速" Or spd = "ノーマル") And n1 < h0 And _
            StrConv(am2, tskosk) = StrConv(bfshn.Cells(n1 + 1, Abs(b)).Value, tskosk) Then
        kskup = 1
        If b < 0 Then '純高速時
            n3 = n1 + 1 'n2は今回突合処理対象行(仮)
            n2 = hiru(n3, 2) '近似の読み替え ８５＿027検証ｃ
            k0 = n3
        Else 'ノーマル
            n2 = n1 + 1
        End If
    Else
        'p=0
    End If
    
    If kskup = 0 Then 'まだ決まらず(p=0)→マッチング実施
        If k0 > h0 Then Call oshimai("", bfn, shn, 1, 0, "k0>h0でmatchにいくようなことはあってはならない。")
        If IsError(Application.Match(StrConv(am2, tskosk), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 1)) Then '近似でもエラーは発生する。
            m = 0
        Else
            m = Application.WorksheetFunction.Match(StrConv(am2, tskosk), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 1) '仮策定
            If StrConv(am2, tskosk) <> StrConv(bfshn.Cells(k0 + m - 1, Abs(b)), tskosk) Then '一致してなければ二段階へ
                m = 0
            Else 'ok近似実施
                kskup = 1
                n3 = k0 + m - 1 'n2は今回突合処理対象行(仮)
                n2 = hiru(n3, 2) '近似の読み替え ８５＿027検証ｃ
                If spd = "純高速" Then k0 = n3
            End If
        End If
        If spd = "純高速" And m = 0 Then '新規で純高速は当処理行う
            kskup = 2
            n2 = h + 1
        End If
    End If
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function tskup(am1 As String, am2 As String, n1 As Long, n2 As Long, h As Long, b As Currency, c As Currency, k0 As Long, h0 As Long, pap2 As Long, er2() As Currency, spd As String, pqp As Long, e5 As Long, er3() As Currency) As Long  'じあ
    'p(tskup)判定新仕様、n3は近似用　,(引数)m2→n1へ リファラ：n3、n2,h0,k0, pqp追加624,e5&er3()追加629
    Dim m As Long, n3 As Long
    Dim tskosk As Long 'strconv用(定速：2(従来通り)、高速：26(New85_006))
    '◇低速側専用
    tskosk = 2  '大文字→小文字化
    tskup = 0 'リセット(不要だが)
    n3 = 0 'ゼロスタート
    krpm2 = 0 '高速ロック＆p=-2判定フラグ
    If h < k Then  '当シートに何も無い場合(初回のみ通過ゾーン) kあ
        If c < 0 Then
            tskup = -1    '-1-2この時点でexitdo(p:-1とする。)
            MsgBox "表空白の-1-2です(p=-1,exitdo、このまま終了されます)。"
        Else  '※c>=0が前提となる。
            tskup = 2
            n2 = h + 1
            h0 = n2 '(=k0,高速時ロックオン) h0処理は基本tskupより後、ここだけ特例(hk逆転イレギュラー解消措置
            If er2(0) < 0 Then
                MsgBox "こちらは低速専用になりました。高速でこちら通るのはおかしい。"
            End If
        End If
    ElseIf LCase(am2) = LCase(am1) Then '(アドバラピド)前回一致　'←86_016q(uni対策)　StrConv(am2→LCase(am2)
        tskup = 1
        n2 = n1
    ElseIf spd = "純高速" And k0 = h0 And b < 0 And pqp = 1 Then '(アドバラピド)高速ロックオン状態　pqp追加624
        MsgBox "こちらは低速専用になりました。純高速でこちら通るのはおかしい。"
    
    '(アドバラピド)次行一致 c = Round(c, 0)追加85＿026（：-15.１実施しない、-15実施する）
    ElseIf Round(c) <> -1 And Round(c) <> -2 And c = Round(c, 0) And (spd = "純高速" Or spd = "ノーマル") And n1 < h0 And _
            LCase(am2) = LCase(bfshn.Cells(n1 + 1, Abs(b)).Value) Then    '←86_016q(uni対策)　StrConv(am2→LCase(am2)
        tskup = 1
        If b < 0 Then '純高速時
            MsgBox "こちらは低速専用になりました。高速でこちら通るのはおかしい。"
        Else 'ノーマル
            n2 = n1 + 1
        End If
    Else
        'p=0
    End If
    If tskup = 0 Then 'まだ決まらず(p=0)→マッチング実施
        If c < 0 And (spd = "近似高速") Then  'spd = "旧近似高速" Or spd = "旧近似ノーマル" Or　は除外 85_006
            MsgBox "こちらは低速専用になりました。近似高速でこちら通るのはおかしい。"
        Else 'ノーマルor純　 am2→strcnv化　85_008
            'If IsError(Application.Match(LCase(am2), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 0)) Then  '新規2
            '↓86_022th
            If IsError(Application.Match(Replace(LCase(am2), "~", "~~"), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 0)) Then  '新規2
                If c >= 0 Then '-1-2はやらないに（p=0のまま終了へ　'←86_016q(uni対策)　StrConv(am2→LCase(am2)
                    tskup = 2
                    n2 = h + 1
                End If
            Else
                tskup = 1
                m = Application.WorksheetFunction.Match(Replace(LCase(am2), "~", "~~"), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 0)  '完全一致
                If spd = "純高速" Then '高速時
                    MsgBox "こちらは低速専用になりました。純高速でこちら通るのはおかしい。"
                Else 'not高速時
                    n2 = k0 + m - 1
                End If
            End If
        End If
    End If
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function iptfg(jmj As String, czc As Long, ww As String) As String
    Do Until jj >= czc '←>を入れてるのは無限ループ防止
        fzf = tzt
        tzt = InStr(fzf + 1, jmj, ww)
        jj = jj + 1
    Loop
    If tzt = 0 Then tzt = Len(jmj) + 1 '4オク対応
    iptfg = Mid(jmj, fzf + 1, tzt - fzf - 1)
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Function koudicd(fn As String, ff As String, er0 As Currency, mr04 As String) As String
    'er0返り値は項目行の行数（返り値0も有：項無の場合）
    '・対象シートの「項目名」探し
    Dim ywe As String
    For er0 = 1 To 2000  '・
        If Workbooks(fn).Sheets(ff).Cells(er0, 1).Value = "項目名" Then
            koudicd = "項有"
            Exit For
        End If
        If Right(Workbooks(fn).Sheets(ff).Cells(er0, 1).Value, 4) = "列固有名" Then  '30s82より
            koudicd = "項固"
            Exit For
        End If
    Next
    
    If koudicd = "" Then koudicd = "項無"  '←仮
        
        If InStr(1, mr04, "ｦ") > 0 Then 'ｦあり[あ]
            'yweは文字
            ywe = iptfg(mr04, 3, "ｦ") '先に「ヱ」把握 iptfg・・NewVersion[い]
            If ywe <> "" And InStr(1, iptfg(mr04, 2, "ｦ"), ywe) > 0 Then 'ｦ△ヱ▲ｦヱ
                ywmoji = iptfg(iptfg(mr04, 2, "ｦ"), 1, ywe) '[う]
                yw10 = iptfg(iptfg(mr04, 2, "ｦ"), 2, ywe)  '[え]
            Else 'ヱなし
                ywmoji = iptfg(mr04, 2, "ｦ") 'yw10はnull[お]
            End If
        Else  'ｦなし
            ywmoji = mr04  'yw10はnull
        End If
        If Mid(ywmoji, 1, 1) = "ー" Then ywmoji = Mid(ywmoji, 2)
        
        '・・対象シートの「項目名」項指項有▲(ｦ△ヱ▲ｦヱ)or項零
        If ywmoji = "" Or ywmoji = "0" Then  '項零628
            If yw10 = "" Then
                er0 = 0
            ElseIf IsNumeric(yw10) Then
                er0 = Val(yw10) - 1
            Else
                Call oshimai("", bfn, shn, 1, 0, "項準bが実施できないようです。")
            End If
            koudicd = "項準b"  '後々項零に変えたい。
            'MsgBox "項準b（項零）"
        ElseIf yw10 <> "" Then   '▲(ｦ△ヱ▲ｦヱ)
            er0 = 1
            If Not IsNumeric(ywmoji) Then '場合分け622
                '従来パターン
                Do Until Workbooks(fn).Sheets(ff).Cells(er0, 1).Value = yw10
                    If er0 = 20000 Then '2000→20000
                        Call oshimai("", bfn, shn, 1, 0, "項指の項目行が見つからないようです")
                    End If
                    er0 = er0 + 1  '・
                Loop
                koudicd = "項準2"  '・　30s64　項指→項準2(旧項指)に変更
            Else        'Newパターン　項準a ここのyw10はall1の代替
                Do Until Workbooks(fn).Sheets(ff).Cells(er0, Abs(Val(ywmoji))).Value = yw10
                    If er0 = 20000 Then '2000→20000
                        Call oshimai("", bfn, shn, 1, 0, "項指の項目行が見つからないようです")
                    End If
                    er0 = er0 + 1  '・
                Loop
                koudicd = "項準a"  '・　30s64　項指→項準2に変更
            End If
        End If
        
        '・・対象シートの「項目名」及び項指もないが、all1が数値記載でない場合(準項目名) 終了措置628
        If koudicd = "項無" And Not IsNumeric(ywmoji) Then
            er0 = 1
            Do Until Workbooks(fn).Sheets(ff).Cells(er0, 1).Value = ywmoji
                If er0 = 20000 Then '2000→20000
                    Call oshimai("", bfn, shn, 1, 0, "項準の項目行が見つからないようです")
                End If
                er0 = er0 + 1  '・
            Loop
            koudicd = "項準" 'yw10は""　・
            Call oshimai("", bfn, shn, 1, 0, "項準は終了しました。項準aに移行して下さい。") '終了措置628
        End If
    If koudicd = "項無" Then er0 = 0
End Function
'ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー◆ー
Sub iechc(hk1 As String)  '以前はigchc
    If twbsh.Cells(3, 2).Value = "" Then  '86_017g 初回誰でも使えるように。
        twbsh.Cells(3, 2).Value = hunk2(-5, syutoku(), "1", hk1)
        If hk1 = "1" Then
            Call oshimai("", twn, "▲集計_雛形", 3, 2, "(処理終了)IDキーが作れません" & vbCrLf & "※ID:" & syutoku())
        End If
    ElseIf syutoku() = hunk2(-7, twbsh.Cells(3, 2).Value, "1", hk1) Then
        'MsgBox "正解です"
    Else  '"不正解です"
        ThisWorkbook.Activate
        Sheets("▲集計_雛形").Select
        twbsh.Cells(1, 1).Select  '緑色セル
        Call oshimai("", twn, "▲集計_雛形", 3, 2, "(処理終了)IDキー不一致" & vbCrLf & "※ID:" & syutoku())
    End If
    
    hk1 = Left(twn, 7) & "r"   'ラピド固定
    twbsh.Cells(2, 2).Value = syutoku() & "r" '新設86_016e
End Sub



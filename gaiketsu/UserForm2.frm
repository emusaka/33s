VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "複写処理"
   ClientHeight    =   2310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3375
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton6_Click()
    'PWなしです　csv(sjis)です
    fmt = "csvsjis"
    Unload Me
End Sub
Private Sub UserForm_Initialize()
'    MsgBox "やぁ"
    TextBox1.Value = dw
End Sub
Private Sub CommandButton1_Click()
    'PWありです xlsxです
    'MsgBox TextBox1.Value
    dw = TextBox1.Value
    fmt = "xlsx"
    Unload Me
End Sub
Private Sub CommandButton2_Click()
    'PWなしです　xlsxです
    'MsgBox TextBox1.Value
    fmt = "xlsx"
    Unload Me
End Sub
Private Sub CommandButton3_Click()
    'PWありです xlsbです
    'MsgBox TextBox1.Value
    dw = TextBox1.Value
    fmt = "xlsb"
    Unload Me
End Sub
Private Sub CommandButton4_Click()
    'PWなしです　xlsbです
    'MsgBox TextBox1.Value
    fmt = "xlsb"
    Unload Me
End Sub
Private Sub CommandButton5_Click()
    'PWなしです　csv(utf-8)です
    fmt = "csv"
    Unload Me
End Sub

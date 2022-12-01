VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} name 
   Caption         =   "新規取引先情報入力"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "name.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'新規取引先登録
Private Sub cmdEnter_Click()

    '//入力内容確認
    Dim errMsg As String: errMsg = Validate
    If errMsg <> "OK" Then
        MsgBox errMsg, vbInformation, "入力エラー"
        Exit Sub
    End If
    
    '最終行(空欄)をコピーして挿入し新規の取引先情報を入力
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    With Rows(lastRow)
        .Copy
        .Insert xlDown
    End With
    Application.CutCopyMode = False
    With Cells(lastRow, 1)
        .Value = txtCode.Text
        .Offset(, 1).Value = txtKana.Text & ":" & txtName.Text
    End With
    '取引先を五十音順で並び替え
    With Sheets(1).Sort.SortFields
        .Clear
        .Add Key:=Range("B3"), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    End With
    With Sheets(1).Sort
        .SetRange Range(Cells(5, 1), Cells(lastRow + 1, Cells(2, Columns.Count).End(xlToLeft).Column))
        .Header = xlYes
        .Apply
    End With
    txtCode.Text = ""
    txtName.Text = ""
    txtKana.Text = ""
    txtCode.SetFocus
End Sub
Private Function Validate() As String
    
        '入力内容確認
    If txtCode.Text = "" Then
        Validate = "取引先コードを入力して下さい!"
        Exit Function
    ElseIf IsNumeric(txtCode.Text) = False Then
        Validate = "取引先コードは数字以外は入力できません!"
        Exit Function
    ElseIf txtKana.Text = "" Then
        Validate = "取引先名ｶﾅを入力してください!"
        Exit Function
    ElseIf txtName.Text = "" Then
        Validate = "取引先名を入力してください!"
        Exit Function
    End If
    
    '//コードの重複がないか確認
    Dim usedCustomer As String '//既にコードを使用している会社名
    
    If Application.WorksheetFunction.CountIf(Columns(1), txtCode.Value) > 0 Then
        usedCustomer = Columns(1).Find(what:=txtCode.Value, lookat:=xlWhole).Offset(, 1).Value
        Validate = "「" & txtCode.Value & "」" & "は既に使用されています。" & vbLf & vbLf & "※取引先コードは一意である必要があります。" & _
            vbLf & "(「" & txtCode.Value & " " & usedCustomer & "」 で使用中)"
        Exit Function
    End If
    
    Validate = "OK"
    
End Function

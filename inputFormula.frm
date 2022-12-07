VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} inputFormula 
   Caption         =   "式入力"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4395
   OleObjectBlob   =   "inputFormula.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "inputFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const INCREASE_COLUMN As Long = 5  '月の列番号から見て増加分が何列右にあるかを表す
Const DECREASE_COLUMN As Long = 1 '月の列番号から見て支払/入金が何列右にあるかを表す
Const OFFSET_COLUMN As Long = 3     '月の列番号から見て相殺列が何列右にあるかを表す
Const FIRST_CELL_ROW As Long = 6    '入力を開始するセルの行番号
Const AMOUNT_ROW As Long = 4        '合計の値が入力されているセルの行番号

Private Sub cmb_cancel_Click()
    Unload Me
End Sub

'メインプログラム
Private Sub cmb_enter_Click()
    
    If cmb_month.Value = "" Then
        MsgBox "対象月を選択してください!"
        Exit Sub
    ElseIf chb_increase.Value = False And chb_decrease.Value = False Then
        MsgBox "増加分か支払/入金かを選択してください!"
        Exit Sub
    ElseIf chb_increase.Value = True And chb_decrease.Value = True Then
        MsgBox "増加分と支払/入金のどちらか一つのみを選択してください!"
    ElseIf chk_all.Value = False And chk_partial.Value = False Then
        MsgBox "全てのセルに式を入力するか、部分的に入力するかを選択して下さい!", vbExclamation
        Exit Sub
    ElseIf chk_all.Value = True And chk_partial.Value = True Then
        MsgBox "全てのセルに式を入力するか、部分的に入力するかどちらか一つのみを選択して下さい!", vbExclamation
        Exit Sub
    End If
    
    '1 monthクラス宣言&インスタンス生成
    Dim mymonth As Month: Set mymonth = New Month
    
    '1-1 指定月の列番号取得
    Dim monthColumn As Long: monthColumn = mymonth.monthColumn(Replace(cmb_month.Value, "月", ""))
    Set mymonth = Nothing
    
    '2 チェックボックス 「増加分」にチェックが入っていたら増加分の式を入力、「支払/入金」にチェックが入っていたら支払/入金の式入力
    If chb_increase.Value = True Then
        If InputIncrease(monthColumn + INCREASE_COLUMN, Replace(cmb_month.Value, "月", "")) = True Then
            MsgBox "処理が完了しました", vbInformation, ThisWorkbook.name
        Else
            Exit Sub
        End If
    ElseIf chb_decrease.Value = True Then
        If InputDecrease(monthColumn, Replace(cmb_month.Value, "月", "")) = True Then
            MsgBox "処理が完了しました。", vbInformation, ThisWorkbook.name
        Else
            Exit Sub
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    'コンボボックスに月を追加
    Dim i As Long
    With cmb_month
        For i = 4 To 12
            .AddItem i & "月"
        Next
        For i = 1 To 3
            .AddItem i & "月"
        Next
    End With
End Sub

'指定の月の増加分セルに式を入力
Private Function InputIncrease(ByVal monthColumn As Long, trgMonth As Long) As Boolean
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'メッセージ表示
    Dim trgColumn As String: trgColumn = Replace(Cells(1, monthColumn).Address(True, False), "$1", "")   '指定月の列番号をアルファベットに変換(メッセージボックスで使用)
    Dim ans As VbMsgBoxResult: ans = MsgBox(trgMonth & "月の増加分(" & trgColumn & "列)に式を入力します。" & vbLf & "実行してよろしいですか?", vbYesNo + vbQuestion)
    If ans = vbNo Then
        InputIncrease = False
        Exit Function
    Else
        If Cells(AMOUNT_ROW, monthColumn).Value <> 0 And chk_all.Value = True Then
            ans = MsgBox("既に値が入力されていますが上書きしてよろしいですか?", vbYesNo + vbExclamation)
            If ans = vbNo Then
                InputIncrease = False
                Exit Function
            End If
        End If
    End If
    
    '式入力
    Dim i As Long
    For i = FIRST_CELL_ROW To Cells(Rows.Count, 1).End(xlUp).Row
        If chk_partial.Value = False Or Not Cells(i, monthColumn).Value > 0 Then
            Cells(i, monthColumn).Formula = "=iferror(if(ワーク!$d$1=" & trgMonth & ",vlookup(indirect(address(row(),1,1,1,),1),ワーク!a:c,3,0),0),0)"
        End If
    Next
    InputIncrease = True
    
    Application.Calculation = xlCalculationAutomatic
    
End Function

'指定の月の支払/入金列のセルに式入力
Private Function InputDecrease(ByVal monthColumn As Long, ByVal trgMonth As Long) As Boolean
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '1 式を入力する列番号をコレクションに格納
    Dim trgColumn As Collection: Set trgColumn = New Collection
    With trgColumn
        .Add monthColumn
        .Add monthColumn + DECREASE_COLUMN
        .Add monthColumn + OFFSET_COLUMN
    End With
    
    Dim i As Long, msgColumn As String '式を入力する列番号をアルファベットに変換したもの(メッセージボックスで使用)
    
    For i = 1 To trgColumn.Count
        msgColumn = msgColumn & Replace(Cells(1, trgColumn(i)).Address(True, False), "$1", "") & "列 "
    Next
    Set trgColumn = Nothing
    
    '2 メッセージ表示
    Dim ans As VbMsgBoxResult: ans = MsgBox(trgMonth & "月支払/入金(" & msgColumn & ")に式を入力します。" & vbLf & "実行してよろしいですか?", vbYesNo + vbQuestion)
    
    If ans = vbNo Then
        InputDecrease = False
        Exit Function
    Else
        If chk_all.Value = True Then
            If Cells(AMOUNT_ROW, monthColumn + DECREASE_COLUMN).Value <> 0 Or _
                Cells(AMOUNT_ROW, monthColumn + OFFSET_COLUMN).Value <> 0 Then
                ans = MsgBox("既に値が入力されていますが上書きしてよろしいですか?", vbYesNo + vbExclamation)
                If ans = vbNo Then
                    InputDecrease = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    '3 式入力
    For i = FIRST_CELL_ROW To Cells(Rows.Count, 1).End(xlUp).Row
        With Cells(i, monthColumn)
            If chk_partial.Value = False Or Not .Offset(, DECREASE_COLUMN).Value > 0 And Not .Offset(, OFFSET_COLUMN).Value > 0 Then
                .Formula = "=iferror(if(ワーク!$f$1=" & trgMonth & ",vlookup(indirect(address(row(),1,1,1),1),ワーク!a:e,5,0),""""),"""")"
                .Offset(, DECREASE_COLUMN).Formula = "=iferror(if(ワーク!$f$1=" & trgMonth & ",vlookup(indirect(address(row(),1,1,1),1),ワーク!a:e,3,0),0),0)"
                .Offset(, OFFSET_COLUMN).Formula = "=iferror(if(ワーク!$f$1=" & trgMonth & ",vlookup(indirect(address(row(),1,1,1),1),ワーク!a:e,4,0),0),0)"
            End If
        End With
    Next
    InputDecrease = True
    
    Application.Calculation = xlCalculationAutomatic
    
End Function

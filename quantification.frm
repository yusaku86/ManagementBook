VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} quantification 
   Caption         =   "数値化"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "quantification.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "quantification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const INCREASE_COLUMN As Long = 5  '月の列番号から見て増加分が何列右にあるかを表す
Const FIRST_CELL_ROW As Long = 6 '数値化を開始するセルの行番号
Const DECREASE_COLUMN As Long = 1 '月の列番号から見て支払/入金が何列右にあるかを表す
Const OFFSET_COLUMN As Long = 3 '月の列番号から見て相殺列が何列右にあるかを表す
Private Sub cmd_close_Click()
    Unload Me
End Sub
'メインプログラム
Private Sub cmd_enter_Click()
    '0 前処理&入力内容確認
    Application.Calculation = xlCalculationManual
    If cmb_month.Value = "" Then
        MsgBox "対象月を選択してください!", vbExclamation
        Exit Sub
    ElseIf chb_increase.Value = False And chb_decrease.Value = False Then
        MsgBox "増加分か支払/入金かを選択してください!", vbExclamation
        Exit Sub
    ElseIf chb_increase.Value = True And chb_decrease.Value = True Then
        MsgBox "増加分か支払/入金かの一つのみ選択してください!", vbExclamation
        Exit Sub
    ElseIf chb_all.Value = False And chb_partial.Value = False Then
        MsgBox "「全て数値化する」か「0以外を数値化する」かを選択して下さい!", vbExclamation
        Exit Sub
    ElseIf chb_all.Value = True And chb_partial.Value = True Then
        MsgBox "「全て数値化する」か「0以外を数値化する」かの一つのみを選択して下さい!", vbExclamation
        Exit Sub
    End If
    
    '1 数値化
    '1-1 対象月の列番号取得
    Dim mymonth As Month: Set mymonth = New Month
    Dim monthColumn As Long: monthColumn = mymonth.monthColumn(Replace(cmb_month.Value, "月", ""))
    Set mymonth = Nothing
    
    '1-2 数値化する列番号をコレクションに格納
    Dim myCollection As Collection: Set myCollection = New Collection
    With myCollection
        If chb_increase.Value = True Then   '増加分の場合
            .Add monthColumn + INCREASE_COLUMN
        ElseIf chb_decrease.Value = True Then   '支払/入金の場合
            .Add monthColumn
            .Add monthColumn + DECREASE_COLUMN
            .Add monthColumn + OFFSET_COLUMN
        End If
    End With
        
    '1-3 数値化
    Dim dtype As String 'Qualifyのメッセージに表示するdTYpe設定
    If chb_increase.Value = True Then
        dtype = "増加分"
    Else
        dtype = "支払/入金"
    End If
    If Quantify(myCollection, Replace(cmb_month, "月", ""), dtype, chb_all.Value) = True Then
        MsgBox "処理が完了しました。"
    End If
    
    'プログラム終了
    Set myCollection = Nothing
    Application.Calculation = xlCalculationAutomatic
    Unload Me
End Sub

'コンボボックスに選択肢を追加
Private Sub UserForm_Initialize()
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
'指定の列を数値化(主に増加分や支払/入金を入力した後に使用)
Private Function Quantify(ByVal monthColumn As Collection, ByVal trgMonth As Long, ByVal dtype As String, ByVal includeZero As Boolean) As Boolean
    
    '1 メッセージを表示
    Dim trgMonthColumn As String, i As Long
    For i = 1 To monthColumn.Count
        trgMonthColumn = trgMonthColumn + Replace(Cells(1, monthColumn(i)).Address(True, False), "$1", "") & "列 " '数値化する列番号をアルファベットに変換(msgboxで使用)
    Next
    Dim ans As VbMsgBoxResult: ans = MsgBox(trgMonth & "月" & dtype & "(" & trgMonthColumn & ")の数値化を実行します。" & vbLf & "よろしいですか?", vbYesNo + vbQuestion)
    If ans = vbNo Then
        Quantify = False
        Exit Function
    End If
    
    '2 指定の列の数値化(全て数値化する場合は0のセルは空欄にする)
    Dim j As Long
    For i = FIRST_CELL_ROW To Cells(Rows.Count, 1).End(xlUp).Row
        For j = 1 To monthColumn.Count
            If includeZero = False Then
                If Cells(i, monthColumn(j)).Value <> 0 And Cells(i, monthColumn(j)).Value <> "" Then
                    Cells(i, monthColumn(j)).Value = Val(Cells(i, monthColumn(j)).Value)
                End If
            ElseIf includeZero = True Then
                If Cells(i, monthColumn(j)).Value = 0 Or Cells(i, monthColumn(j)).Value = "" Then
                    Cells(i, monthColumn(j)).ClearContents
                Else
                    Cells(i, monthColumn(j)).Value = Val(Cells(i, monthColumn(j)).Value)
                End If
            End If
        Next j
    Next i
    Quantify = True
End Function

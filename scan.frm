VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} scan 
   Caption         =   "請求書漏れ確認"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   OleObjectBlob   =   "scan.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "scan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const INTERVAL As Long = 9 '月と月の間の列数
Const FIRST_CELL_ROW As Long = 6 '最初の取引先行
Const INCREASE_COLUMN As Long = 5 '月列から見て増加分の列が何列右にあるかを表す
Const STANDARD_COLUMN As Long = 5 '期首月の列番号
'cmb_condition_2に選択肢を追加
Private Sub cmb_condition_Change()
    If cmb_condition.Value = "" Then
        Exit Sub
    End If
    Dim i As Long
    With cmb_condition_2
            .Clear
        For i = 1 To Replace(cmb_condition, "ヵ月", "")
            .AddItem i & "回"
        Next
    End With
End Sub
'cmb_conditionに選択肢を追加
Private Sub cmb_month_Change()
    
    '指定月の列数を求め、期首から指定月まで何か月経過しているか求める
    Dim mymonth As Month: Set mymonth = New Month
    Dim num As Long: num = (mymonth.monthColumn(Replace(cmb_month.Value, "月", "")) - STANDARD_COLUMN) / INTERVAL
    Set mymonth = Nothing
    
    '経過した月の分だけ選択肢を追加
    Dim i As Long
    With cmb_condition
        .Clear
        For i = 1 To num
            .AddItem i & "ヵ月"
        Next
    End With
End Sub

Private Sub cmd_cancel_Click()
    Unload Me
End Sub
'メインプログラム
Private Sub cmd_enter_Click()
    '0 前処理&入力内容確認
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    If cmb_month.Value = "" Or cmb_condition.Value = "" Or cmb_condition_2.Value = "" Then
        MsgBox "対象月を選択して下さい!", vbExclamation
        Exit Sub
    End If
    
    '1 請求書漏れの対象の取引先抽出
    '1-1 対象月の列番号取得
    Dim mymonth As Month: Set mymonth = New Month
    Dim monthColumn As Long: monthColumn = mymonth.monthColumn(Replace(cmb_month.Value, "月", ""))
    Set mymonth = Nothing
    
    '1-2 対象の取引先行をコレクションに格納
    Dim targetRow As Collection: Set targetRow = New Collection
    Dim i As Long, j As Long, counter As Long
    For i = FIRST_CELL_ROW To Cells(Rows.Count, 1).End(xlUp)
        For j = 1 To Replace(cmb_condition.Value, "ヵ月", "")
            If Cells(i, monthColumn + INCREASE_COLUMN - INTERVAL * j).Value > 0 Then '取引が合った回数をconterに格納
                counter = counter + 1
            End If
        Next j
        If counter >= Replace(cmb_condition_2.Value, "回", "") And Cells(i, monthColumn + INCREASE_COLUMN).Value = 0 Then
            targetRow.Add i
        End If
        counter = 0
    Next i
    
    '2 CSV出力
    If targetRow.Count > 0 Then
        Dim filePath As String: filePath = ExportCSV(targetRow, monthColumn, Replace(cmb_condition.Value, "ヵ月", ""))
    End If
    
    '3メッセージ表示&ファイル起動
    If targetRow.Count > 0 Then
        Dim ans As VbMsgBoxResult: ans = MsgBox("処理が完了しました。該当取引先は" & targetRow.Count & "件です。" & vbLf & "CSVファイルを開きますか?", vbYesNo + vbQuestion)
        If ans = vbYes Then
            Workbooks.Open filePath
        End If
    Else
        MsgBox "処理が完了しました。該当取引先はありません。"
    End If
    Set targetRow = Nothing
    
    'プログラム終了
    Application.Calculation = xlCalculationAutomatic
    Unload Me
End Sub

'cmb_monthに選択肢を追加
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
'CSVファイルを出力
Private Function ExportCSV(ByRef targetRow As Collection, ByVal monthColumn As Long, ByVal span As Long) As String
    
    '1 ファイルパス設定
    Dim myWsh As WshShell: Set myWsh = New WshShell
    Dim filePath As String: filePath = myWsh.SpecialFolders(4) & "\該当取引先.csv"
    Set myWsh = Nothing
    
    '2 CSV出力
    Dim myFSO As FileSystemObject: Set myFSO = New FileSystemObject
    Dim i As Long, j As Long, txt As String
    With myFSO.CreateTextFile(filePath)
        '2-1 ヘッダー作成
        txt = "取引先コード,取引先名,"
        For i = Replace(cmb_condition.Value, "ヵ月", "") To 1 Step -1
            txt = txt & Replace(cmb_month.Value, "月", "") - i & "月分,"
        Next
        .WriteLine txt
        '2-2 「取引先コード/取引先名/過去の取引金額(指定の期間分)」となるように記入
        For i = 1 To targetRow.Count
            txt = Cells(targetRow(i), 1).Value & "," & Cells(targetRow(i), 2).Value & ","
            For j = span To 1 Step -1
                txt = txt & Cells(targetRow(i), monthColumn + INCREASE_COLUMN - j * INTERVAL).Value & ","
            Next j
            .WriteLine txt
        Next i
    End With
    
    'プログラム終了
    Set myFSO = Nothing
    ExportCSV = filePath
End Function

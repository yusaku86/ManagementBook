VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Month"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const MONTH_INTERVAL As Long = 9 '月と月の間の列数
Const STANDARD_MONTH_COLUMN As Long = 5 '期首月の列番号
Public Event change(ByVal counter As Boolean) 'イベントの定義

'月列番号設定
Public Property Get monthColumn(ByVal newMonth As Long) As Long
    If newMonth >= 1 And newMonth <= 12 Then
        monthColumn = CalculateMonthColumn(newMonth)
    End If
End Property

'月の列番号を取得
Private Function CalculateMonthColumn(ByVal trgMonth As Long) As Long

    'インスタンス生成と設定
    Dim myReg As RegExp: Set myReg = New RegExp
    With myReg
        .Pattern = "[^0-9]"
        .Global = True
    End With
    
    '期首(standardMonth)が何月かを求め、取得したい月(trgMonth）との差に月と月の間の列数をかけ、期首の列番号を足す
    Dim standardMonth As Long: standardMonth = myReg.Replace(Sheets(1).Cells(1, 5).Value, "")
    
    Set myReg = Nothing
    
    If trgMonth - standardMonth >= 0 Then
        CalculateMonthColumn = (trgMonth - standardMonth) * MONTH_INTERVAL + STANDARD_MONTH_COLUMN
    Else
        CalculateMonthColumn = (trgMonth + 12 - standardMonth) * MONTH_INTERVAL + STANDARD_MONTH_COLUMN
    End If
    
End Function



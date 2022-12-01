VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_increase 
   Caption         =   "増加分入力"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   OleObjectBlob   =   "f_increase.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "f_increase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'■■====================
'[f_increase]
'作成者:Yusaku Suzuki(2022/03/15)
'====================■■
Option Explicit
Private Sub cmd_close_Click()
    Unload Me
End Sub
Private Sub cmd_enter_Click()
 
    ' 0 入力内容確認&前処理
    If cmb_code.Value = "" Or cmb_amount.Value = "" Or cmb_month.Value = "" Or cmb_name.Value = "" Then
        MsgBox "入力内容に不備があります!", vbExclamation, "エラー"
        Exit Sub
    ElseIf cmb_code.Value = cmb_amount.Value Or cmb_code.Value = cmb_name.Value Or cmb_amount.Value = cmb_name.Value Then
        MsgBox "選択した列が重複しています!", vbExclamation, "エラー"
        Exit Sub
    End If
    
    With Application                                    '/**
        .ScreenUpdating = False                      '* 前処理
        .EnableEvents = False                         '*/
        .Calculation = xlCalculationManual
    End With
    
    With Sheets("増加分列設定")                 '/**
        .Cells(2, 1).Value = cmb_code.Value        '* デフォルト列設定
        .Cells(2, 2).Value = cmb_name.Value       '*/
        .Cells(2, 3).Value = cmb_amount.Value
    End With
    
    Sheets("ワーク").Cells.Clear                    '/**
    With Sheets("ワーク2")                           '* セルに式が入っている可能性があるため値として保存する
        .Cells.Copy                                       '*/
        .Cells(1, 1).PasteSpecial xlPasteValues
    End With
    
    '1 コードと取引先名、金額のみをワークに表示
    '1-1 コード列と取引先名列、金額列をワーク2のA列～C列に移動
    Dim codeColumn As Long: codeColumn = Cells(1, Replace(cmb_code.Value, "列", "")).Column
    Dim nameColumn As Long: nameColumn = Cells(1, Replace(cmb_name.Value, "列", "")).Column
    Dim amountColumn As Long: amountColumn = Cells(1, Replace(cmb_amount.Value, "列", "")).Column
    
    Range(Columns(1), Columns(3)).Insert xlToRight
    Columns(codeColumn + 3).Copy Destination:=Cells(1, 1)
    Columns(nameColumn + 3).Copy Destination:=Cells(1, 2)
    Columns(amountColumn + 3).Copy Destination:=Cells(1, 3)
    
    '1-2 コードと取引先名、コードと金額を格納した連想配列を2つ作成
    Dim myDic1 As Dictionary: Set myDic1 = New Dictionary
    Dim i As Long
    
    For i = 1 To Sheets("ワーク2").Cells(Rows.Count, 1).End(xlUp).Row
        If myDic1.Exists(Cells(i, 1).Value) = False Then
            myDic1.Add Cells(i, 1).Value, Cells(i, 2).Value
        End If
    Next
    
    Dim myDic2 As Dictionary: Set myDic2 = New Dictionary
    
    For i = 1 To Sheets("ワーク2").Cells(Rows.Count, 1).End(xlUp).Row
        If myDic2.Exists(Cells(i, 1).Value) = False Then
            myDic2.Add Cells(i, 1).Value, Cells(i, 3).Value
        ElseIf myDic2.Exists(Cells(i, 1).Value) = True Then
            myDic2(Cells(i, 1).Value) = myDic2(Cells(i, 1).Value) + Cells(i, 3).Value
        End If
    Next
    
    '1-3 格納したコード、取引先名、金額をワークのA列～C列に入力
    i = 1
    Dim myKey As Variant
    With Sheets("ワーク")
        For Each myKey In myDic1.Keys
            .Cells(i, 1).Value = myKey
            .Cells(i, 2).Value = myDic1.Item(myKey)
            .Cells(i, 3).Value = myDic2.Item(myKey)
            i = i + 1
        Next
    End With
    Set myDic1 = Nothing
    Set myDic2 = Nothing
    
    '2 計上月入力
    Sheets("ワーク").Cells(1, 4).Value = Replace(cmb_month.Value, "月", "")
    
    '3 新規取引先抽出&「数値化」& A1の色変更
    Sheets("ワーク2").Cells.Clear
    Call FindNewCustomer '新規取引先抽出
    With Sheets(1)
        .Activate
        .Cells(1, 1).Interior.ColorIndex = 3
    End With
    
    'プログラム終了
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    Unload Me
End Sub
'/**
 '* フォーム起動時処理
 '*/
 Private Sub UserForm_Initialize()
    
    Dim i As Long
    
    '取引先コード列・取引先名列・金額列のコンボボックスに選択肢を追加
    With cmb_code
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "列")
        Next
        .Value = Sheets("増加分列設定").Cells(2, 1).Value
        
     End With
     
     With cmb_name
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "列")
        Next
        .Value = Sheets("増加分列設定").Cells(2, 2).Value
     
     End With
     
     With cmb_amount
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "列")
        Next
        .Value = Sheets("増加分列設定").Cells(2, 3).Value
        
    End With
    
    '計上月のコンボボックスに選択肢を追加
    With cmb_month
        
        For i = 4 To 12
            .AddItem i & "月"
        Next
        For i = 1 To 3
            .AddItem i & "月"
        Next
        .Value = Month(DateSerial(Year(Now), Month(Now) - 1, Day(Now))) & "月"
        
    End With
    
End Sub
'管理帳に登録されてない取引先を抽出&メッセージ表示
Private Sub FindNewCustomer()

    Dim i As Long
    Dim myDic As Dictionary: Set myDic = New Dictionary
    
    With Sheets("ワーク")
        For i = 1 To Sheets("ワーク").Cells(Rows.Count, 1).End(xlUp).Row
            If IsNumeric(.Cells(i, 1).Value) = True And .Cells(i, 1).Value <> "" Then
                If Application.WorksheetFunction.CountIf(Sheets(1).Columns(1), .Cells(i, 1).Value) = 0 Then '管理帳に登録されていない取引先コードを連想配列に追加
                    myDic.Add .Cells(i, 1).Value, .Cells(i, 2).Value
                End If
            End If
        Next
    End With
    
    '管理帳に登録されてない取引先の文字色を変更(ワーク)
    With Sheets("ワーク")
        For i = 0 To myDic.Count - 1
             .Cells(.Cells.Find(what:=myDic(myDic.Keys(i)), lookat:=xlWhole).Row, 1).Interior.ColorIndex = 50
        Next
    End With
    
    'メッセージ表示
    Dim customer As String
    
    For i = 0 To myDic.Count - 1
        customer = customer & vbLf & myDic.Keys(i) & ":" & myDic(myDic.Keys(i))
    Next
    
    Dim msg As String
    
    If myDic.Count >= 1 Then
        msg = "処理が完了しました。" & vbLf & "新規取引先(管理帳に登録されていない取引先)は以下の通りです。" & vbLf & customer
    Else
        msg = "処理が完了しました。"
    End If
    MsgBox msg
    Set myDic = Nothing
    
End Sub

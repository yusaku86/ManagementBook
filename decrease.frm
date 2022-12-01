VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} decrease 
   Caption         =   "支払/入金分入力"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   OleObjectBlob   =   "decrease.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "decrease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmd_cancel_Click()
    Unload Me
End Sub
'メインプログラム
Private Sub cmd_enter_Click()
    
    '0 前処理&入力内容確認
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    If cmb_code.Value = "" Or cmb_name.Value = "" Or cmb_amount.Value = "" Then
        MsgBox "入力内容に不備があります!", vbExclamation
        Exit Sub
    ElseIf chk_charge.Value = False And cmb_charge.Value = "" Then
        MsgBox "入力内容に不備があります!", vbExclamation
        Exit Sub
    ElseIf chk_date.Value = False And cmb_date.Value = "" Then
        MsgBox "入力内容に不備があります!", vbExclamation
        Exit Sub
    ElseIf cmb_code.Value = cmb_name.Value Or cmb_code.Value = cmb_amount.Value Or cmb_code.Value = cmb_charge.Value Or cmb_code.Value = cmb_date.Value Or _
            cmb_name.Value = cmb_amount.Value Or cmb_name.Value = cmb_charge.Value Or cmb_name.Value = cmb_date.Value Or cmb_amount.Value = cmb_charge.Value Or _
            cmb_amount.Value = cmb_date.Value Then
        MsgBox "選択した列が重複しています!", vbExclamation
        Exit Sub
    ElseIf cmb_charge.Value <> "" And cmb_charge.Value = cmb_date.Value Then
        MsgBox "選択した列が重複しています!", vbExclamation
        Exit Sub
    ElseIf chk_date.Value = True And txt_date.Value = "" Then
        MsgBox "日付を入力してください!", vbExclamation
        Exit Sub
    ElseIf chk_date.Value = True And IsNumeric(txt_date.Value) = False Then
        MsgBox "日付には数字を入力して下さい!", vbExclamation
        Exit Sub
    End If
    
    With Sheets("入金(支払)分列設定")                   '/**
        .Cells(2, 1).Value = cmb_code.Value                  '* デフォルト列更新
        .Cells(2, 2).Value = cmb_name.Value                 '*/
        .Cells(2, 3).Value = cmb_amount.Value
        .Cells(2, 4).Value = cmb_charge.Value
        .Cells(2, 5).Value = cmb_date.Value
    End With
    
    With Sheets("ワーク2")
        .Cells.Copy
        .Cells(1, 1).PasteSpecial xlPasteValues
    End With
        
    '1 ワーク2のA～E列にコード、取引先名、金額列を移動
     Sheets("ワーク2").Range(Columns(1), Columns(5)).Insert xlToRight
    TransferColumn Cells(1, Replace(cmb_code.Value, "列", "")).Column + 5, 1
    TransferColumn Cells(1, Replace(cmb_name.Value, "列", "")).Column + 5, 2
    TransferColumn Cells(1, Replace(cmb_amount.Value, "列", "")).Column + 5, 3
        
    If chk_charge.Value = False Then
        TransferColumn Cells(1, Replace(cmb_charge.Value, "列", "")).Column + 5, 4
    End If
    If chk_date.Value = False Then
        TransferColumn Cells(1, Replace(cmb_date.Value, "列", "")).Column + 5, 5
    Else
        Dim i As Long
        For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
            Cells(i, 5).Value = txt_date.Value
        Next
    End If
    
    '2 ワーク2のA～E列の内容をコードが重複するものを統合してワークに表示
    Sheets("ワーク").Cells.Clear
    
    CreateDictionary 1, 2, "name"
    CreateDictionary 1, 3, "sum"
    If chk_charge.Value = False Then
        CreateDictionary 1, 4, "sum"
    End If
    CreateDictionary 1, 5, "date"
    
    '3 ワークのF1に計上月入力&メッセージ表示
    Sheets("ワーク2").Cells.Clear
    Sheets("ワーク").Cells(1, 6).Value = Replace(cmb_month.Value, "月", "")
    Sheets(1).Cells(1, 1).Interior.ColorIndex = 3
    MsgBox "処理が完了しました。"
    
    
    'プログラム終了
    Application.Calculation = xlCalculationAutomatic
    Unload Me

End Sub
'指定の行を移動
Private Sub TransferColumn(ByVal trgColumn As Long, ByVal destinationColumn As Long)
    Columns(trgColumn).Copy Destination:=Cells(1, destinationColumn)
End Sub

'連想配列を作成し、ワークに内容を入力
'keyColumn→連想配列のキーとする列番号(今回は全て取引先コード列)
'valueColumn→連想配列の値とする列番号
'aggfunc→同じコードの取引先の値の計算方法(sumなら合計、dateなら 10・30のように表記)
Private Sub CreateDictionary(ByVal keyColumn As Long, ByVal valueColumn As Long, ByVal aggfunc As String)
    
    '連想配列の作成
    Dim myDic As Dictionary: Set myDic = New Dictionary
    Dim i As Long
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, keyColumn).Value <> "" And IsNumeric(Cells(i, keyColumn).Value) = True Then
            If myDic.Exists(Cells(i, keyColumn).Value) = False Then
                myDic.Add Cells(i, keyColumn).Value, Cells(i, valueColumn).Value
            Else
                If aggfunc = "sum" Then
                    myDic(Cells(i, keyColumn).Value) = myDic(Cells(i, keyColumn).Value) + Cells(i, valueColumn).Value
                ElseIf aggfunc = "date" Then
                    If myDic(Cells(i, keyColumn).Value) <> Cells(i, valueColumn).Value Then
                        myDic(Cells(i, keyColumn).Value) = myDic(Cells(i, keyColumn).Value) & "・" & Cells(i, valueColumn).Value
                    End If
                End If
            End If
        End If
    Next
    
    '連想配列の値をワークに入力
    With Sheets("ワーク")
        For i = 1 To myDic.Count
            .Cells(i, keyColumn).Value = myDic.Keys(i - 1)
            .Cells(i, valueColumn).Value = myDic(myDic.Keys(i - 1))
        Next
    End With
    Set myDic = Nothing
    
End Sub
Private Sub UserForm_Initialize()

    'コンボボックスに選択肢を追加&デフォルト値設定
    Dim i As Long
    
    With cmb_code
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "列")
        Next
        .Value = Sheets("入金(支払)分列設定").Cells(2, 1).Value
        
     End With
     
     With cmb_name
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "列")
        Next
        .Value = Sheets("入金(支払)分列設定").Cells(2, 2).Value
        
     End With
     
     With cmb_amount
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "列")
        Next
        .Value = Sheets("入金(支払)分列設定").Cells(2, 3).Value
    
    End With
    
    With cmb_charge
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "列")
        Next
        .Value = Sheets("入金(支払)分列設定").Cells(2, 4).Value
    
    End With
    
    With cmb_date
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "列")
        Next
        .Value = Sheets("入金(支払)分列設定").Cells(2, 5).Value
    
    End With
    
    With cmb_month
        For i = 4 To 12
            .AddItem i & "月"
        Next
        For i = 1 To 3
            .AddItem i & "月"
        Next
    End With

End Sub

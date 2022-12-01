VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} carriedForward 
   Caption         =   "管理帳繰越"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "carriedForward.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "carriedForward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'// 管理帳の次期繰越を行うフォーム
Option Explicit

'// フォーム起動時の処理
Private Sub UserForm_Initialize()

    txtFileName.Value = ThisWorkbook.name

End Sub

'// 参照を押したときの処理
Private Sub cmdDialog_Click()

    '// ダイアログを表示し保存先選択 selectFolder [ダイアログタイトル], [ダイアログの初期フォルダ]
    Dim destinationFolder As String: destinationFolder = selectFolder("管理帳繰越:保存先フォルダ選択", ThisWorkbook.Path & "\")
    
    If destinationFolder <> "" Then
        txtFolder.Value = destinationFolder
    End If

End Sub

'// 実行を押したときの処理
Private Sub cmdEnter_Click()

    '// バリデーション
    '// 保存先フォルダが入力されているか
    If txtFolder.Value = "" Then
        MsgBox "保存先フォルダを入力してください。", vbQuestion, "管理帳繰越"
        Exit Sub
    End If
    
    Dim fso As New FileSystemObject
    
    '// 保存先フォルダが存在するか
    If fso.FolderExists(txtFolder.Value) = False Then
        MsgBox "保存先フォルダが存在しません。", vbQuestion, "管理帳繰越"
        Set fso = Nothing
        Exit Sub
    End If
    
    Set fso = Nothing
    
    '// ファイル名が入力されているか
    If txtFileName.Value = "" Then
        MsgBox "ファイル名を入力してください。", vbQuestion, "管理帳繰越"
        Exit Sub
    End If
    
    '// ファイル名が変更されているか
    If txtFileName.Value = ThisWorkbook.name Then
        MsgBox "新しい管理帳のファイル名は現在のものと違う名前にしてください。", vbQuestion, "管理帳繰越"
        Exit Sub
    End If
    
    Me.Hide
    
    '// 繰越実行
    Call createNextYearChart(txtFolder.Value & "\" & txtFileName.Value)
    
    Unload Me

End Sub

'// 閉じるを押したときの処理
Private Sub cmdClose_Click()

    Unload Me

End Sub

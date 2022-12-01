Attribute VB_Name = "functions"
 '// 関数を定義するモジュール
Option Explicit

'// ダイアログを表示してフォルダを選択
Public Function selectFolder(ByVal dialogTitle As String, ByVal initialFolder As String)

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = initialFolder
        .AllowMultiSelect = False
        .Title = dialogTitle
        
        If .Show Then
            selectFolder = .SelectedItems(1)
        Else
            selectFolder = ""
        End If
    End With
    
End Function



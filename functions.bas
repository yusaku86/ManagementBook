Attribute VB_Name = "functions"
 '// �֐����`���郂�W���[��
Option Explicit

'// �_�C�A���O��\�����ăt�H���_��I��
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



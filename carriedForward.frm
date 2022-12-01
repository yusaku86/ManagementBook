VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} carriedForward 
   Caption         =   "�Ǘ����J�z"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "carriedForward.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "carriedForward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'// �Ǘ����̎����J�z���s���t�H�[��
Option Explicit

'// �t�H�[���N�����̏���
Private Sub UserForm_Initialize()

    txtFileName.Value = ThisWorkbook.name

End Sub

'// �Q�Ƃ��������Ƃ��̏���
Private Sub cmdDialog_Click()

    '// �_�C�A���O��\�����ۑ���I�� selectFolder [�_�C�A���O�^�C�g��], [�_�C�A���O�̏����t�H���_]
    Dim destinationFolder As String: destinationFolder = selectFolder("�Ǘ����J�z:�ۑ���t�H���_�I��", ThisWorkbook.Path & "\")
    
    If destinationFolder <> "" Then
        txtFolder.Value = destinationFolder
    End If

End Sub

'// ���s���������Ƃ��̏���
Private Sub cmdEnter_Click()

    '// �o���f�[�V����
    '// �ۑ���t�H���_�����͂���Ă��邩
    If txtFolder.Value = "" Then
        MsgBox "�ۑ���t�H���_����͂��Ă��������B", vbQuestion, "�Ǘ����J�z"
        Exit Sub
    End If
    
    Dim fso As New FileSystemObject
    
    '// �ۑ���t�H���_�����݂��邩
    If fso.FolderExists(txtFolder.Value) = False Then
        MsgBox "�ۑ���t�H���_�����݂��܂���B", vbQuestion, "�Ǘ����J�z"
        Set fso = Nothing
        Exit Sub
    End If
    
    Set fso = Nothing
    
    '// �t�@�C���������͂���Ă��邩
    If txtFileName.Value = "" Then
        MsgBox "�t�@�C��������͂��Ă��������B", vbQuestion, "�Ǘ����J�z"
        Exit Sub
    End If
    
    '// �t�@�C�������ύX����Ă��邩
    If txtFileName.Value = ThisWorkbook.name Then
        MsgBox "�V�����Ǘ����̃t�@�C�����͌��݂̂��̂ƈႤ���O�ɂ��Ă��������B", vbQuestion, "�Ǘ����J�z"
        Exit Sub
    End If
    
    Me.Hide
    
    '// �J�z���s
    Call createNextYearChart(txtFolder.Value & "\" & txtFileName.Value)
    
    Unload Me

End Sub

'// ������������Ƃ��̏���
Private Sub cmdClose_Click()

    Unload Me

End Sub

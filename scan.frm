VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} scan 
   Caption         =   "�������R��m�F"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   OleObjectBlob   =   "scan.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "scan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const INTERVAL As Long = 9 '���ƌ��̊Ԃ̗�
Const FIRST_CELL_ROW As Long = 6 '�ŏ��̎����s
Const INCREASE_COLUMN As Long = 5 '���񂩂猩�đ������̗񂪉���E�ɂ��邩��\��
Const STANDARD_COLUMN As Long = 5 '���񌎂̗�ԍ�
'cmb_condition_2�ɑI������ǉ�
Private Sub cmb_condition_Change()
    If cmb_condition.Value = "" Then
        Exit Sub
    End If
    Dim i As Long
    With cmb_condition_2
            .Clear
        For i = 1 To Replace(cmb_condition, "����", "")
            .AddItem i & "��"
        Next
    End With
End Sub
'cmb_condition�ɑI������ǉ�
Private Sub cmb_month_Change()
    
    '�w�茎�̗񐔂����߁A���񂩂�w�茎�܂ŉ������o�߂��Ă��邩���߂�
    Dim mymonth As Month: Set mymonth = New Month
    Dim num As Long: num = (mymonth.monthColumn(Replace(cmb_month.Value, "��", "")) - STANDARD_COLUMN) / INTERVAL
    Set mymonth = Nothing
    
    '�o�߂������̕������I������ǉ�
    Dim i As Long
    With cmb_condition
        .Clear
        For i = 1 To num
            .AddItem i & "����"
        Next
    End With
End Sub

Private Sub cmd_cancel_Click()
    Unload Me
End Sub
'���C���v���O����
Private Sub cmd_enter_Click()
    '0 �O����&���͓��e�m�F
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    If cmb_month.Value = "" Or cmb_condition.Value = "" Or cmb_condition_2.Value = "" Then
        MsgBox "�Ώی���I�����ĉ�����!", vbExclamation
        Exit Sub
    End If
    
    '1 �������R��̑Ώۂ̎���撊�o
    '1-1 �Ώی��̗�ԍ��擾
    Dim mymonth As Month: Set mymonth = New Month
    Dim monthColumn As Long: monthColumn = mymonth.monthColumn(Replace(cmb_month.Value, "��", ""))
    Set mymonth = Nothing
    
    '1-2 �Ώۂ̎����s���R���N�V�����Ɋi�[
    Dim targetRow As Collection: Set targetRow = New Collection
    Dim i As Long, j As Long, counter As Long
    For i = FIRST_CELL_ROW To Cells(Rows.Count, 1).End(xlUp)
        For j = 1 To Replace(cmb_condition.Value, "����", "")
            If Cells(i, monthColumn + INCREASE_COLUMN - INTERVAL * j).Value > 0 Then '������������񐔂�conter�Ɋi�[
                counter = counter + 1
            End If
        Next j
        If counter >= Replace(cmb_condition_2.Value, "��", "") And Cells(i, monthColumn + INCREASE_COLUMN).Value = 0 Then
            targetRow.Add i
        End If
        counter = 0
    Next i
    
    '2 CSV�o��
    If targetRow.Count > 0 Then
        Dim filePath As String: filePath = ExportCSV(targetRow, monthColumn, Replace(cmb_condition.Value, "����", ""))
    End If
    
    '3���b�Z�[�W�\��&�t�@�C���N��
    If targetRow.Count > 0 Then
        Dim ans As VbMsgBoxResult: ans = MsgBox("�������������܂����B�Y��������" & targetRow.Count & "���ł��B" & vbLf & "CSV�t�@�C�����J���܂���?", vbYesNo + vbQuestion)
        If ans = vbYes Then
            Workbooks.Open filePath
        End If
    Else
        MsgBox "�������������܂����B�Y�������͂���܂���B"
    End If
    Set targetRow = Nothing
    
    '�v���O�����I��
    Application.Calculation = xlCalculationAutomatic
    Unload Me
End Sub

'cmb_month�ɑI������ǉ�
Private Sub UserForm_Initialize()
    Dim i As Long
    With cmb_month
        For i = 4 To 12
            .AddItem i & "��"
        Next
        For i = 1 To 3
            .AddItem i & "��"
        Next
    End With
End Sub
'CSV�t�@�C�����o��
Private Function ExportCSV(ByRef targetRow As Collection, ByVal monthColumn As Long, ByVal span As Long) As String
    
    '1 �t�@�C���p�X�ݒ�
    Dim myWsh As WshShell: Set myWsh = New WshShell
    Dim filePath As String: filePath = myWsh.SpecialFolders(4) & "\�Y�������.csv"
    Set myWsh = Nothing
    
    '2 CSV�o��
    Dim myFSO As FileSystemObject: Set myFSO = New FileSystemObject
    Dim i As Long, j As Long, txt As String
    With myFSO.CreateTextFile(filePath)
        '2-1 �w�b�_�[�쐬
        txt = "�����R�[�h,����於,"
        For i = Replace(cmb_condition.Value, "����", "") To 1 Step -1
            txt = txt & Replace(cmb_month.Value, "��", "") - i & "����,"
        Next
        .WriteLine txt
        '2-2 �u�����R�[�h/����於/�ߋ��̎�����z(�w��̊��ԕ�)�v�ƂȂ�悤�ɋL��
        For i = 1 To targetRow.Count
            txt = Cells(targetRow(i), 1).Value & "," & Cells(targetRow(i), 2).Value & ","
            For j = span To 1 Step -1
                txt = txt & Cells(targetRow(i), monthColumn + INCREASE_COLUMN - j * INTERVAL).Value & ","
            Next j
            .WriteLine txt
        Next i
    End With
    
    '�v���O�����I��
    Set myFSO = Nothing
    ExportCSV = filePath
End Function

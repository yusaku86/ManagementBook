VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} inputFormula 
   Caption         =   "������"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4395
   OleObjectBlob   =   "inputFormula.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "inputFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const INCREASE_COLUMN As Long = 5  '���̗�ԍ����猩�đ�����������E�ɂ��邩��\��
Const DECREASE_COLUMN As Long = 1 '���̗�ԍ����猩�Ďx��/����������E�ɂ��邩��\��
Const OFFSET_COLUMN As Long = 3     '���̗�ԍ����猩�đ��E�񂪉���E�ɂ��邩��\��
Const FIRST_CELL_ROW As Long = 6    '���͂��J�n����Z���̍s�ԍ�
Const AMOUNT_ROW As Long = 4        '���v�̒l�����͂���Ă���Z���̍s�ԍ�



Private Sub cmb_cancel_Click()
    Unload Me
End Sub
'���C���v���O����
Private Sub cmb_enter_Click()
    '0 �O����&���͓��e�m�F
    Application.Calculation = xlCalculationManual
    
    If cmb_month.Value = "" Then
        MsgBox "�Ώی���I�����Ă�������!"
        Exit Sub
    ElseIf chb_increase.Value = False And chb_decrease.Value = False Then
        MsgBox "���������x��/��������I�����Ă�������!"
        Exit Sub
    ElseIf chb_increase.Value = True And chb_decrease.Value = True Then
        MsgBox "�������Ǝx��/�����̂ǂ��炩��݂̂�I�����Ă�������!"
    ElseIf chk_all.Value = False And chk_partial.Value = False Then
        MsgBox "�S�ẴZ���Ɏ�����͂��邩�A�����I�ɓ��͂��邩��I�����ĉ�����!", vbExclamation
        Exit Sub
    ElseIf chk_all.Value = True And chk_partial.Value = True Then
        MsgBox "�S�ẴZ���Ɏ�����͂��邩�A�����I�ɓ��͂��邩�ǂ��炩��݂̂�I�����ĉ�����!", vbExclamation
        Exit Sub
    End If
    
    '1 month�N���X�錾&�C���X�^���X����
    Dim mymonth As Month: Set mymonth = New Month
    '1-1 �w�茎�̗�ԍ��擾
    Dim monthColumn As Long: monthColumn = mymonth.monthColumn(Replace(cmb_month.Value, "��", ""))
    Set mymonth = Nothing
    
    '2 �`�F�b�N�{�b�N�X �u�������v�Ƀ`�F�b�N�������Ă����瑝�����̎�����́A�u�x��/�����v�Ƀ`�F�b�N�������Ă�����x��/�����̎�����
    If chb_increase.Value = True Then
        If InputIncrease(monthColumn + INCREASE_COLUMN, Replace(cmb_month.Value, "��", "")) = True Then
            MsgBox "�������������܂���"
        Else
            Exit Sub
        End If
    ElseIf chb_decrease.Value = True Then
        If InputDecrease(monthColumn, Replace(cmb_month.Value, "��", "")) = True Then
            MsgBox "�������������܂����B"
        Else
            Exit Sub
        End If
    End If
    
    '�v���O�����I��
    Application.Calculation = xlCalculationAutomatic
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    '�R���{�{�b�N�X�Ɍ���ǉ�
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
'�w��̌��̑������Z���Ɏ������
Private Function InputIncrease(ByVal monthColumn As Long, trgMonth As Long) As Boolean
    
    '���b�Z�[�W�\��
    Dim trgColumn As String: trgColumn = Replace(Cells(1, monthColumn).Address(True, False), "$1", "")   '�w�茎�̗�ԍ����A���t�@�x�b�g�ɕϊ�(���b�Z�[�W�{�b�N�X�Ŏg�p)
    Dim ans As VbMsgBoxResult: ans = MsgBox(trgMonth & "���̑�����(" & trgColumn & "��)�Ɏ�����͂��܂��B" & vbLf & "���s���Ă�낵���ł���?", vbYesNo + vbQuestion)
    If ans = vbNo Then
        InputIncrease = False
        Exit Function
    Else
        If Cells(AMOUNT_ROW, monthColumn).Value <> 0 And chk_all.Value = True Then
            ans = MsgBox("���ɒl�����͂���Ă��܂����㏑�����Ă�낵���ł���?", vbYesNo + vbExclamation)
            If ans = vbNo Then
                InputIncrease = False
                Exit Function
            End If
        End If
    End If
    
    '������
    Dim i As Long
    For i = FIRST_CELL_ROW To Cells(Rows.Count, 1).End(xlUp).Row
        If chk_partial.Value = False Or Not Cells(i, monthColumn).Value > 0 Then
            Cells(i, monthColumn).Formula = "=iferror(if(���[�N!$d$1=" & trgMonth & ",vlookup(indirect(address(row(),1,1,1,),1),���[�N!a:c,3,0),0),0)"
        End If
    Next
    InputIncrease = True
    
End Function
'�w��̌��̎x��/������̃Z���Ɏ�����
Private Function InputDecrease(ByVal monthColumn As Long, ByVal trgMonth As Long) As Boolean
    
    '1 ������͂����ԍ����R���N�V�����Ɋi�[
    Dim trgColumn As Collection: Set trgColumn = New Collection
    With trgColumn
        .Add monthColumn
        .Add monthColumn + DECREASE_COLUMN
        .Add monthColumn + OFFSET_COLUMN
    End With
    Dim i As Long, msgColumn As String '������͂����ԍ����A���t�@�x�b�g�ɕϊ���������(���b�Z�[�W�{�b�N�X�Ŏg�p)
    For i = 1 To trgColumn.Count
        msgColumn = msgColumn & Replace(Cells(1, trgColumn(i)).Address(True, False), "$1", "") & "�� "
    Next
    Set trgColumn = Nothing
    
    '2 ���b�Z�[�W�\��
    Dim ans As VbMsgBoxResult: ans = MsgBox(trgMonth & "���x��/����(" & msgColumn & ")�Ɏ�����͂��܂��B" & vbLf & "���s���Ă�낵���ł���?", vbYesNo + vbQuestion)
    If ans = vbNo Then
        InputDecrease = False
        Exit Function
    Else
        If Cells(AMOUNT_ROW, monthColumn + DECREASE_COLUMN).Value <> 0 Or Cells(AMOUNT_ROW, monthColumn + OFFSET_COLUMN).Value <> 0 Then
            ans = MsgBox("���ɒl�����͂���Ă��܂����㏑�����Ă�낵���ł���?", vbYesNo + vbExclamation)
            If ans = vbNo Then
                InputDecrease = False
                Exit Function
            End If
        End If
    End If
    
    '3 ������
    For i = FIRST_CELL_ROW To Cells(Rows.Count, 1).End(xlUp).Row
        With Cells(i, monthColumn)
            .Formula = "=iferror(if(���[�N!$f$1=" & trgMonth & ",vlookup(indirect(address(row(),1,1,1),1),���[�N!a:e,5,0),""""),"""")"
            .Offset(, DECREASE_COLUMN).Formula = "=iferror(if(���[�N!$f$1=" & trgMonth & ",vlookup(indirect(address(row(),1,1,1),1),���[�N!a:e,3,0),0),0)"
            .Offset(, OFFSET_COLUMN).Formula = "=iferror(if(���[�N!$f$1=" & trgMonth & ",vlookup(indirect(address(row(),1,1,1),1),���[�N!a:e,4,0),0),0)"
        End With
    Next
    InputDecrease = True
    
End Function

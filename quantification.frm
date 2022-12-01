VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} quantification 
   Caption         =   "���l��"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "quantification.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "quantification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const INCREASE_COLUMN As Long = 5  '���̗�ԍ����猩�đ�����������E�ɂ��邩��\��
Const FIRST_CELL_ROW As Long = 6 '���l�����J�n����Z���̍s�ԍ�
Const DECREASE_COLUMN As Long = 1 '���̗�ԍ����猩�Ďx��/����������E�ɂ��邩��\��
Const OFFSET_COLUMN As Long = 3 '���̗�ԍ����猩�đ��E�񂪉���E�ɂ��邩��\��
Private Sub cmd_close_Click()
    Unload Me
End Sub
'���C���v���O����
Private Sub cmd_enter_Click()
    '0 �O����&���͓��e�m�F
    Application.Calculation = xlCalculationManual
    If cmb_month.Value = "" Then
        MsgBox "�Ώی���I�����Ă�������!", vbExclamation
        Exit Sub
    ElseIf chb_increase.Value = False And chb_decrease.Value = False Then
        MsgBox "���������x��/��������I�����Ă�������!", vbExclamation
        Exit Sub
    ElseIf chb_increase.Value = True And chb_decrease.Value = True Then
        MsgBox "���������x��/�������̈�̂ݑI�����Ă�������!", vbExclamation
        Exit Sub
    ElseIf chb_all.Value = False And chb_partial.Value = False Then
        MsgBox "�u�S�Đ��l������v���u0�ȊO�𐔒l������v����I�����ĉ�����!", vbExclamation
        Exit Sub
    ElseIf chb_all.Value = True And chb_partial.Value = True Then
        MsgBox "�u�S�Đ��l������v���u0�ȊO�𐔒l������v���̈�݂̂�I�����ĉ�����!", vbExclamation
        Exit Sub
    End If
    
    '1 ���l��
    '1-1 �Ώی��̗�ԍ��擾
    Dim mymonth As Month: Set mymonth = New Month
    Dim monthColumn As Long: monthColumn = mymonth.monthColumn(Replace(cmb_month.Value, "��", ""))
    Set mymonth = Nothing
    
    '1-2 ���l�������ԍ����R���N�V�����Ɋi�[
    Dim myCollection As Collection: Set myCollection = New Collection
    With myCollection
        If chb_increase.Value = True Then   '�������̏ꍇ
            .Add monthColumn + INCREASE_COLUMN
        ElseIf chb_decrease.Value = True Then   '�x��/�����̏ꍇ
            .Add monthColumn
            .Add monthColumn + DECREASE_COLUMN
            .Add monthColumn + OFFSET_COLUMN
        End If
    End With
        
    '1-3 ���l��
    Dim dtype As String 'Qualify�̃��b�Z�[�W�ɕ\������dTYpe�ݒ�
    If chb_increase.Value = True Then
        dtype = "������"
    Else
        dtype = "�x��/����"
    End If
    If Quantify(myCollection, Replace(cmb_month, "��", ""), dtype, chb_all.Value) = True Then
        MsgBox "�������������܂����B"
    End If
    
    '�v���O�����I��
    Set myCollection = Nothing
    Application.Calculation = xlCalculationAutomatic
    Unload Me
End Sub

'�R���{�{�b�N�X�ɑI������ǉ�
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
'�w��̗�𐔒l��(��ɑ�������x��/��������͂�����Ɏg�p)
Private Function Quantify(ByVal monthColumn As Collection, ByVal trgMonth As Long, ByVal dtype As String, ByVal includeZero As Boolean) As Boolean
    
    '1 ���b�Z�[�W��\��
    Dim trgMonthColumn As String, i As Long
    For i = 1 To monthColumn.Count
        trgMonthColumn = trgMonthColumn + Replace(Cells(1, monthColumn(i)).Address(True, False), "$1", "") & "�� " '���l�������ԍ����A���t�@�x�b�g�ɕϊ�(msgbox�Ŏg�p)
    Next
    Dim ans As VbMsgBoxResult: ans = MsgBox(trgMonth & "��" & dtype & "(" & trgMonthColumn & ")�̐��l�������s���܂��B" & vbLf & "��낵���ł���?", vbYesNo + vbQuestion)
    If ans = vbNo Then
        Quantify = False
        Exit Function
    End If
    
    '2 �w��̗�̐��l��(�S�Đ��l������ꍇ��0�̃Z���͋󗓂ɂ���)
    Dim j As Long
    For i = FIRST_CELL_ROW To Cells(Rows.Count, 1).End(xlUp).Row
        For j = 1 To monthColumn.Count
            If includeZero = False Then
                If Cells(i, monthColumn(j)).Value <> 0 And Cells(i, monthColumn(j)).Value <> "" Then
                    Cells(i, monthColumn(j)).Value = Val(Cells(i, monthColumn(j)).Value)
                End If
            ElseIf includeZero = True Then
                If Cells(i, monthColumn(j)).Value = 0 Or Cells(i, monthColumn(j)).Value = "" Then
                    Cells(i, monthColumn(j)).ClearContents
                Else
                    Cells(i, monthColumn(j)).Value = Val(Cells(i, monthColumn(j)).Value)
                End If
            End If
        Next j
    Next i
    Quantify = True
End Function

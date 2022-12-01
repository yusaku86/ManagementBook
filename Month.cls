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
Const MONTH_INTERVAL As Long = 9 '���ƌ��̊Ԃ̗�
Const STANDARD_MONTH_COLUMN As Long = 5 '���񌎂̗�ԍ�
Public Event change(ByVal counter As Boolean) '�C�x���g�̒�`

'����ԍ��ݒ�
Public Property Get monthColumn(ByVal newMonth As Long) As Long
    If newMonth >= 1 And newMonth <= 12 Then
        monthColumn = CalculateMonthColumn(newMonth)
    End If
End Property

'���̗�ԍ����擾
Private Function CalculateMonthColumn(ByVal trgMonth As Long) As Long

    '�C���X�^���X�����Ɛݒ�
    Dim myReg As RegExp: Set myReg = New RegExp
    With myReg
        .Pattern = "[^0-9]"
        .Global = True
    End With
    
    '����(standardMonth)�������������߁A�擾��������(trgMonth�j�Ƃ̍��Ɍ��ƌ��̊Ԃ̗񐔂������A����̗�ԍ��𑫂�
    Dim standardMonth As Long: standardMonth = myReg.Replace(Sheets(1).Cells(1, 5).Value, "")
    
    Set myReg = Nothing
    
    If trgMonth - standardMonth >= 0 Then
        CalculateMonthColumn = (trgMonth - standardMonth) * MONTH_INTERVAL + STANDARD_MONTH_COLUMN
    Else
        CalculateMonthColumn = (trgMonth + 12 - standardMonth) * MONTH_INTERVAL + STANDARD_MONTH_COLUMN
    End If
    
End Function


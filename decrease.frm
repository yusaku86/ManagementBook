VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} decrease 
   Caption         =   "�x��/����������"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   OleObjectBlob   =   "decrease.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
'���C���v���O����
Private Sub cmd_enter_Click()
    
    '0 �O����&���͓��e�m�F
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    If cmb_code.Value = "" Or cmb_name.Value = "" Or cmb_amount.Value = "" Then
        MsgBox "���͓��e�ɕs��������܂�!", vbExclamation
        Exit Sub
    ElseIf chk_charge.Value = False And cmb_charge.Value = "" Then
        MsgBox "���͓��e�ɕs��������܂�!", vbExclamation
        Exit Sub
    ElseIf chk_date.Value = False And cmb_date.Value = "" Then
        MsgBox "���͓��e�ɕs��������܂�!", vbExclamation
        Exit Sub
    ElseIf cmb_code.Value = cmb_name.Value Or cmb_code.Value = cmb_amount.Value Or cmb_code.Value = cmb_charge.Value Or cmb_code.Value = cmb_date.Value Or _
            cmb_name.Value = cmb_amount.Value Or cmb_name.Value = cmb_charge.Value Or cmb_name.Value = cmb_date.Value Or cmb_amount.Value = cmb_charge.Value Or _
            cmb_amount.Value = cmb_date.Value Then
        MsgBox "�I�������񂪏d�����Ă��܂�!", vbExclamation
        Exit Sub
    ElseIf cmb_charge.Value <> "" And cmb_charge.Value = cmb_date.Value Then
        MsgBox "�I�������񂪏d�����Ă��܂�!", vbExclamation
        Exit Sub
    ElseIf chk_date.Value = True And txt_date.Value = "" Then
        MsgBox "���t����͂��Ă�������!", vbExclamation
        Exit Sub
    ElseIf chk_date.Value = True And IsNumeric(txt_date.Value) = False Then
        MsgBox "���t�ɂ͐�������͂��ĉ�����!", vbExclamation
        Exit Sub
    End If
    
    With Sheets("����(�x��)����ݒ�")                   '/**
        .Cells(2, 1).Value = cmb_code.Value                  '* �f�t�H���g��X�V
        .Cells(2, 2).Value = cmb_name.Value                 '*/
        .Cells(2, 3).Value = cmb_amount.Value
        .Cells(2, 4).Value = cmb_charge.Value
        .Cells(2, 5).Value = cmb_date.Value
    End With
    
    With Sheets("���[�N2")
        .Cells.Copy
        .Cells(1, 1).PasteSpecial xlPasteValues
    End With
        
    '1 ���[�N2��A�`E��ɃR�[�h�A����於�A���z����ړ�
     Sheets("���[�N2").Range(Columns(1), Columns(5)).Insert xlToRight
    TransferColumn Cells(1, Replace(cmb_code.Value, "��", "")).Column + 5, 1
    TransferColumn Cells(1, Replace(cmb_name.Value, "��", "")).Column + 5, 2
    TransferColumn Cells(1, Replace(cmb_amount.Value, "��", "")).Column + 5, 3
        
    If chk_charge.Value = False Then
        TransferColumn Cells(1, Replace(cmb_charge.Value, "��", "")).Column + 5, 4
    End If
    If chk_date.Value = False Then
        TransferColumn Cells(1, Replace(cmb_date.Value, "��", "")).Column + 5, 5
    Else
        Dim i As Long
        For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
            Cells(i, 5).Value = txt_date.Value
        Next
    End If
    
    '2 ���[�N2��A�`E��̓��e���R�[�h���d��������̂𓝍����ă��[�N�ɕ\��
    Sheets("���[�N").Cells.Clear
    
    CreateDictionary 1, 2, "name"
    CreateDictionary 1, 3, "sum"
    If chk_charge.Value = False Then
        CreateDictionary 1, 4, "sum"
    End If
    CreateDictionary 1, 5, "date"
    
    '3 ���[�N��F1�Ɍv�㌎����&���b�Z�[�W�\��
    Sheets("���[�N2").Cells.Clear
    Sheets("���[�N").Cells(1, 6).Value = Replace(cmb_month.Value, "��", "")
    Sheets(1).Cells(1, 1).Interior.ColorIndex = 3
    MsgBox "�������������܂����B"
    
    
    '�v���O�����I��
    Application.Calculation = xlCalculationAutomatic
    Unload Me

End Sub
'�w��̍s���ړ�
Private Sub TransferColumn(ByVal trgColumn As Long, ByVal destinationColumn As Long)
    Columns(trgColumn).Copy Destination:=Cells(1, destinationColumn)
End Sub

'�A�z�z����쐬���A���[�N�ɓ��e�����
'keyColumn���A�z�z��̃L�[�Ƃ����ԍ�(����͑S�Ď����R�[�h��)
'valueColumn���A�z�z��̒l�Ƃ����ԍ�
'aggfunc�������R�[�h�̎����̒l�̌v�Z���@(sum�Ȃ獇�v�Adate�Ȃ� 10�E30�̂悤�ɕ\�L)
Private Sub CreateDictionary(ByVal keyColumn As Long, ByVal valueColumn As Long, ByVal aggfunc As String)
    
    '�A�z�z��̍쐬
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
                        myDic(Cells(i, keyColumn).Value) = myDic(Cells(i, keyColumn).Value) & "�E" & Cells(i, valueColumn).Value
                    End If
                End If
            End If
        End If
    Next
    
    '�A�z�z��̒l�����[�N�ɓ���
    With Sheets("���[�N")
        For i = 1 To myDic.Count
            .Cells(i, keyColumn).Value = myDic.Keys(i - 1)
            .Cells(i, valueColumn).Value = myDic(myDic.Keys(i - 1))
        Next
    End With
    Set myDic = Nothing
    
End Sub
Private Sub UserForm_Initialize()

    '�R���{�{�b�N�X�ɑI������ǉ�&�f�t�H���g�l�ݒ�
    Dim i As Long
    
    With cmb_code
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "��")
        Next
        .Value = Sheets("����(�x��)����ݒ�").Cells(2, 1).Value
        
     End With
     
     With cmb_name
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "��")
        Next
        .Value = Sheets("����(�x��)����ݒ�").Cells(2, 2).Value
        
     End With
     
     With cmb_amount
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "��")
        Next
        .Value = Sheets("����(�x��)����ݒ�").Cells(2, 3).Value
    
    End With
    
    With cmb_charge
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "��")
        Next
        .Value = Sheets("����(�x��)����ݒ�").Cells(2, 4).Value
    
    End With
    
    With cmb_date
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "��")
        Next
        .Value = Sheets("����(�x��)����ݒ�").Cells(2, 5).Value
    
    End With
    
    With cmb_month
        For i = 4 To 12
            .AddItem i & "��"
        Next
        For i = 1 To 3
            .AddItem i & "��"
        Next
    End With

End Sub

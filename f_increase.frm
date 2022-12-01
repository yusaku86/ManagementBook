VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_increase 
   Caption         =   "����������"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   OleObjectBlob   =   "f_increase.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "f_increase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'����====================
'[f_increase]
'�쐬��:Yusaku Suzuki(2022/03/15)
'====================����
Option Explicit
Private Sub cmd_close_Click()
    Unload Me
End Sub
Private Sub cmd_enter_Click()
 
    ' 0 ���͓��e�m�F&�O����
    If cmb_code.Value = "" Or cmb_amount.Value = "" Or cmb_month.Value = "" Or cmb_name.Value = "" Then
        MsgBox "���͓��e�ɕs��������܂�!", vbExclamation, "�G���["
        Exit Sub
    ElseIf cmb_code.Value = cmb_amount.Value Or cmb_code.Value = cmb_name.Value Or cmb_amount.Value = cmb_name.Value Then
        MsgBox "�I�������񂪏d�����Ă��܂�!", vbExclamation, "�G���["
        Exit Sub
    End If
    
    With Application                                    '/**
        .ScreenUpdating = False                      '* �O����
        .EnableEvents = False                         '*/
        .Calculation = xlCalculationManual
    End With
    
    With Sheets("��������ݒ�")                 '/**
        .Cells(2, 1).Value = cmb_code.Value        '* �f�t�H���g��ݒ�
        .Cells(2, 2).Value = cmb_name.Value       '*/
        .Cells(2, 3).Value = cmb_amount.Value
    End With
    
    Sheets("���[�N").Cells.Clear                    '/**
    With Sheets("���[�N2")                           '* �Z���Ɏ��������Ă���\�������邽�ߒl�Ƃ��ĕۑ�����
        .Cells.Copy                                       '*/
        .Cells(1, 1).PasteSpecial xlPasteValues
    End With
    
    '1 �R�[�h�Ǝ���於�A���z�݂̂����[�N�ɕ\��
    '1-1 �R�[�h��Ǝ���於��A���z������[�N2��A��`C��Ɉړ�
    Dim codeColumn As Long: codeColumn = Cells(1, Replace(cmb_code.Value, "��", "")).Column
    Dim nameColumn As Long: nameColumn = Cells(1, Replace(cmb_name.Value, "��", "")).Column
    Dim amountColumn As Long: amountColumn = Cells(1, Replace(cmb_amount.Value, "��", "")).Column
    
    Range(Columns(1), Columns(3)).Insert xlToRight
    Columns(codeColumn + 3).Copy Destination:=Cells(1, 1)
    Columns(nameColumn + 3).Copy Destination:=Cells(1, 2)
    Columns(amountColumn + 3).Copy Destination:=Cells(1, 3)
    
    '1-2 �R�[�h�Ǝ���於�A�R�[�h�Ƌ��z���i�[�����A�z�z���2�쐬
    Dim myDic1 As Dictionary: Set myDic1 = New Dictionary
    Dim i As Long
    
    For i = 1 To Sheets("���[�N2").Cells(Rows.Count, 1).End(xlUp).Row
        If myDic1.Exists(Cells(i, 1).Value) = False Then
            myDic1.Add Cells(i, 1).Value, Cells(i, 2).Value
        End If
    Next
    
    Dim myDic2 As Dictionary: Set myDic2 = New Dictionary
    
    For i = 1 To Sheets("���[�N2").Cells(Rows.Count, 1).End(xlUp).Row
        If myDic2.Exists(Cells(i, 1).Value) = False Then
            myDic2.Add Cells(i, 1).Value, Cells(i, 3).Value
        ElseIf myDic2.Exists(Cells(i, 1).Value) = True Then
            myDic2(Cells(i, 1).Value) = myDic2(Cells(i, 1).Value) + Cells(i, 3).Value
        End If
    Next
    
    '1-3 �i�[�����R�[�h�A����於�A���z�����[�N��A��`C��ɓ���
    i = 1
    Dim myKey As Variant
    With Sheets("���[�N")
        For Each myKey In myDic1.Keys
            .Cells(i, 1).Value = myKey
            .Cells(i, 2).Value = myDic1.Item(myKey)
            .Cells(i, 3).Value = myDic2.Item(myKey)
            i = i + 1
        Next
    End With
    Set myDic1 = Nothing
    Set myDic2 = Nothing
    
    '2 �v�㌎����
    Sheets("���[�N").Cells(1, 4).Value = Replace(cmb_month.Value, "��", "")
    
    '3 �V�K����撊�o&�u���l���v& A1�̐F�ύX
    Sheets("���[�N2").Cells.Clear
    Call FindNewCustomer '�V�K����撊�o
    With Sheets(1)
        .Activate
        .Cells(1, 1).Interior.ColorIndex = 3
    End With
    
    '�v���O�����I��
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    Unload Me
End Sub
'/**
 '* �t�H�[���N��������
 '*/
 Private Sub UserForm_Initialize()
    
    Dim i As Long
    
    '�����R�[�h��E����於��E���z��̃R���{�{�b�N�X�ɑI������ǉ�
    With cmb_code
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "��")
        Next
        .Value = Sheets("��������ݒ�").Cells(2, 1).Value
        
     End With
     
     With cmb_name
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "��")
        Next
        .Value = Sheets("��������ݒ�").Cells(2, 2).Value
     
     End With
     
     With cmb_amount
        
        For i = 1 To 26
            .AddItem (Replace(Cells(1, i).Address(True, False), "$1", "") & "��")
        Next
        .Value = Sheets("��������ݒ�").Cells(2, 3).Value
        
    End With
    
    '�v�㌎�̃R���{�{�b�N�X�ɑI������ǉ�
    With cmb_month
        
        For i = 4 To 12
            .AddItem i & "��"
        Next
        For i = 1 To 3
            .AddItem i & "��"
        Next
        .Value = Month(DateSerial(Year(Now), Month(Now) - 1, Day(Now))) & "��"
        
    End With
    
End Sub
'�Ǘ����ɓo�^����ĂȂ������𒊏o&���b�Z�[�W�\��
Private Sub FindNewCustomer()

    Dim i As Long
    Dim myDic As Dictionary: Set myDic = New Dictionary
    
    With Sheets("���[�N")
        For i = 1 To Sheets("���[�N").Cells(Rows.Count, 1).End(xlUp).Row
            If IsNumeric(.Cells(i, 1).Value) = True And .Cells(i, 1).Value <> "" Then
                If Application.WorksheetFunction.CountIf(Sheets(1).Columns(1), .Cells(i, 1).Value) = 0 Then '�Ǘ����ɓo�^����Ă��Ȃ������R�[�h��A�z�z��ɒǉ�
                    myDic.Add .Cells(i, 1).Value, .Cells(i, 2).Value
                End If
            End If
        Next
    End With
    
    '�Ǘ����ɓo�^����ĂȂ������̕����F��ύX(���[�N)
    With Sheets("���[�N")
        For i = 0 To myDic.Count - 1
             .Cells(.Cells.Find(what:=myDic(myDic.Keys(i)), lookat:=xlWhole).Row, 1).Interior.ColorIndex = 50
        Next
    End With
    
    '���b�Z�[�W�\��
    Dim customer As String
    
    For i = 0 To myDic.Count - 1
        customer = customer & vbLf & myDic.Keys(i) & ":" & myDic(myDic.Keys(i))
    Next
    
    Dim msg As String
    
    If myDic.Count >= 1 Then
        msg = "�������������܂����B" & vbLf & "�V�K�����(�Ǘ����ɓo�^����Ă��Ȃ������)�͈ȉ��̒ʂ�ł��B" & vbLf & customer
    Else
        msg = "�������������܂����B"
    End If
    MsgBox msg
    Set myDic = Nothing
    
End Sub

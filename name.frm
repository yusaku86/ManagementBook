VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} name 
   Caption         =   "�V�K����������"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "name.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'�V�K�����o�^
Private Sub cmdEnter_Click()

    '//���͓��e�m�F
    Dim errMsg As String: errMsg = Validate
    If errMsg <> "OK" Then
        MsgBox errMsg, vbInformation, "���̓G���["
        Exit Sub
    End If
    
    '�ŏI�s(��)���R�s�[���đ}�����V�K�̎����������
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    With Rows(lastRow)
        .Copy
        .Insert xlDown
    End With
    Application.CutCopyMode = False
    With Cells(lastRow, 1)
        .Value = txtCode.Text
        .Offset(, 1).Value = txtKana.Text & ":" & txtName.Text
    End With
    '�������܏\�����ŕ��ёւ�
    With Sheets(1).Sort.SortFields
        .Clear
        .Add Key:=Range("B3"), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    End With
    With Sheets(1).Sort
        .SetRange Range(Cells(5, 1), Cells(lastRow + 1, Cells(2, Columns.Count).End(xlToLeft).Column))
        .Header = xlYes
        .Apply
    End With
    txtCode.Text = ""
    txtName.Text = ""
    txtKana.Text = ""
    txtCode.SetFocus
End Sub
Private Function Validate() As String
    
        '���͓��e�m�F
    If txtCode.Text = "" Then
        Validate = "�����R�[�h����͂��ĉ�����!"
        Exit Function
    ElseIf IsNumeric(txtCode.Text) = False Then
        Validate = "�����R�[�h�͐����ȊO�͓��͂ł��܂���!"
        Exit Function
    ElseIf txtKana.Text = "" Then
        Validate = "����於�ł���͂��Ă�������!"
        Exit Function
    ElseIf txtName.Text = "" Then
        Validate = "����於����͂��Ă�������!"
        Exit Function
    End If
    
    '//�R�[�h�̏d�����Ȃ����m�F
    Dim usedCustomer As String '//���ɃR�[�h���g�p���Ă����Ж�
    
    If Application.WorksheetFunction.CountIf(Columns(1), txtCode.Value) > 0 Then
        usedCustomer = Columns(1).Find(what:=txtCode.Value, lookat:=xlWhole).Offset(, 1).Value
        Validate = "�u" & txtCode.Value & "�v" & "�͊��Ɏg�p����Ă��܂��B" & vbLf & vbLf & "�������R�[�h�͈�ӂł���K�v������܂��B" & _
            vbLf & "(�u" & txtCode.Value & " " & usedCustomer & "�v �Ŏg�p��)"
        Exit Function
    End If
    
    Validate = "OK"
    
End Function

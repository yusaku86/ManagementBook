Attribute VB_Name = "Module1"
Option Explicit

Sub InputIncrease()
    f_increase.Show vbModeless
End Sub

Sub AddCustomer()
    name.Show
End Sub

Sub InputCellsFormula()
    inputFormula.Show vbModeless
End Sub

Sub QuantifyCells()
    quantification.Show vbModeless
End Sub

Sub ScanData()
    scan.Show
End Sub

Sub InputDecrease()
    decrease.Show vbModeless
End Sub

'// �Ǘ����J�z�̃t�H�[���\��
Sub openFormToCarriedForward()

    carriedForward.Show

End Sub


'// �������̊Ǘ������쐬����
Public Sub createNextYearChart(ByVal newFileName As String)
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        
    '/**
     '* �u�b�N����ǉ����Ď������̊Ǘ������쐬
    '**/
    
    Dim newFile As Workbook: Set newFile = Workbooks.Add
        
    ThisWorkbook.Sheets("Sheet1").Copy after:=newFile.Sheets(1)
    newFile.Sheets(1).Delete
    newFile.Sheets(1).name = "Sheet1"
    
    
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    newFile.Sheets(1).Activate

    '// �I�[�g�t�B���^�[���������Ă�ꍇ�͉���
    If newFile.Sheets(1).AutoFilterMode = True Then
        Cells(1, 1).AutoFilter
    End If

    '// �O�����c�����R�s�[���Ċ���c���ɓ\��t��
    Range(Cells(6, 111), Cells(lastRow, 111)).Copy
    Cells(6, 3).PasteSpecial xlPasteValues
    
    '// �O���̃f�[�^���N���A(��������c���Ȃǂ̐����������Ă���Ƃ���͏���)
    Dim i As Long
    
    For i = 4 To Cells(2, Columns.Count).End(xlToLeft).Column
        
        '// �ƍ���̏ꍇ
        If i Mod 9 = 4 Then
            Range(Cells(4, i), Cells(lastRow, i)).ClearContents
        
        '// ���ϓ���̏ꍇ
        ElseIf i Mod 9 = 5 Then
            Range(Cells(6, i), Cells(lastRow, i + 3)).ClearContents
            i = i + 4
        
        '// ��������̏ꍇ
        ElseIf i Mod 9 = 1 Then
            Range(Cells(6, i), Cells(lastRow, i + 1)).ClearContents
            i = i + 2
        End If
    Next
    
    '// �ucheck0��0���v�Ɠ��͂��Ă���Z�����N���A
    Union(Cells(1, 12), _
          Cells(1, 21), _
          Cells(1, 30), _
          Cells(1, 39), _
          Cells(1, 48), _
          Cells(1, 57), _
          Cells(1, 66), _
          Cells(1, 75), _
          Cells(1, 84), _
          Cells(1, 93), _
          Cells(1, 102), _
          Cells(1, 111)) _
    .ClearContents
          
    Cells(1, 2).ClearContents
    
    ThisWorkbook.Sheets("���[�N").Copy after:=newFile.Sheets(newFile.Sheets.Count)
    ThisWorkbook.Sheets("���[�N2").Copy after:=newFile.Sheets(newFile.Sheets.Count)
    
    newFile.Sheets(1).Activate
    Cells(1, 1).Select
    
    newFile.SaveAs newFileName, xlOpenXMLWorkbookMacroEnabled
    newFile.Close
    
    MsgBox "�Ǘ����̌J�z���������܂����B", vbInformation, "�Ǘ����J�z"
    
End Sub

Attribute VB_Name = "Module1"
Option Explicit

'// ���������͂̃t�H�[���N��
Sub InputIncrease()

    f_increase.Show vbModeless

End Sub

'// �V�K�ǉ��̃t�H�[���N��
Sub AddCustomer()

    name.Show

End Sub

'// �����͂̃t�H�[���N��
Sub InputCellsFormula()

    inputFormula.Show vbModeless

End Sub

'// �����͂̃t�H�[���N��
Sub QuantifyCells()

    quantification.Show vbModeless

End Sub

'// �R����͂̃t�H�[���N��
Sub ScanData()
    scan.Show
End Sub

'// �x����/�������̃t�H�[���N��
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
     '* ���̃t�@�C���̃R�s�[���쐬���ė������̊Ǘ����ɂ���
    '**/
    ThisWorkbook.SaveCopyAs newFileName
    Dim newFile As Workbook: Set newFile = Workbooks.Open(newFileName)
            
    '// �I�[�g�t�B���^�[���������Ă�ꍇ�͉���
    If ActiveSheet.AutoFilterMode = True Then
        Cells(1, 1).AutoFilter
    End If

    Dim lastRow As Long: lastRow = Cells(Rows.Count, 1).End(xlUp).Row

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
    
    newFile.Sheets(1).Activate
    Cells(1, 1).Select
    
    newFile.Close True
    
    Set newFile = Nothing
    
    MsgBox "�Ǘ����̌J�z���������܂����B", vbInformation, "�Ǘ����J�z"
    
End Sub

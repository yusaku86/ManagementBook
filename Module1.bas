Attribute VB_Name = "Module1"
Option Explicit

'// 増加分入力のフォーム起動
Sub InputIncrease()

    f_increase.Show vbModeless

End Sub

'// 新規追加のフォーム起動
Sub AddCustomer()

    name.Show

End Sub

'// 式入力のフォーム起動
Sub InputCellsFormula()

    inputFormula.Show vbModeless

End Sub

'// 式入力のフォーム起動
Sub QuantifyCells()

    quantification.Show vbModeless

End Sub

'// 漏れ入力のフォーム起動
Sub ScanData()
    scan.Show
End Sub

'// 支払い/入金分のフォーム起動
Sub InputDecrease()
    decrease.Show vbModeless
End Sub

'// 管理帳繰越のフォーム表示
Sub openFormToCarriedForward()

    carriedForward.Show

End Sub

'// 来期分の管理帳を作成する
Public Sub createNextYearChart(ByVal newFileName As String)
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        
    '/**
     '* このファイルのコピーを作成して来期分の管理帳にする
    '**/
    ThisWorkbook.SaveCopyAs newFileName
    Dim newFile As Workbook: Set newFile = Workbooks.Open(newFileName)
            
    '// オートフィルターがかかってる場合は解除
    If ActiveSheet.AutoFilterMode = True Then
        Cells(1, 1).AutoFilter
    End If

    Dim lastRow As Long: lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    '// 前期末残高をコピーして期首残高に貼り付け
    Range(Cells(6, 111), Cells(lastRow, 111)).Copy
    Cells(6, 3).PasteSpecial xlPasteValues
    
    '// 前期のデータをクリア(未払高や残高などの数式が入っているところは除く)
    Dim i As Long
    
    For i = 4 To Cells(2, Columns.Count).End(xlToLeft).Column
        
        '// 照合列の場合
        If i Mod 9 = 4 Then
            Range(Cells(4, i), Cells(lastRow, i)).ClearContents
        
        '// 決済日列の場合
        ElseIf i Mod 9 = 5 Then
            Range(Cells(6, i), Cells(lastRow, i + 3)).ClearContents
            i = i + 4
        
        '// 増加高列の場合
        ElseIf i Mod 9 = 1 Then
            Range(Cells(6, i), Cells(lastRow, i + 1)).ClearContents
            i = i + 2
        End If
    Next
    
    '// 「check0月0日」と入力してあるセルをクリア
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
    
    MsgBox "管理帳の繰越が完了しました。", vbInformation, "管理帳繰越"
    
End Sub

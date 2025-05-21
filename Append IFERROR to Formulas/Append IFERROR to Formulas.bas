Attribute VB_Name = "Module1"
Sub Replace_Formulas()

    Dim cell As Range
    Dim FormulaRange As Range
    ActiveSheet.Range("DTS").Select
     
    Set FormulaRange = Range("DTS").Cells.SpecialCells(xlCellTypeFormulas)
    'Debug.Print FormulaRange.Address
    'Range("B27").Value = FormulaRange.Address
    
    For Each cell In FormulaRange
        cell.Formula = "=iferror(" & VBA.Mid(cell.Formula, 2) & ", """")"
    Next cell

End Sub


Sub Copy_Paste_Reset()
 ActiveSheet.Range("Reval").Select
 Selection.Copy
 Range("A8").Select
 ActiveSheet.Paste
 
End Sub

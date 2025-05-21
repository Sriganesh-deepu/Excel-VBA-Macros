Attribute VB_Name = "Module1"
Sub GenerateSimpleXLSReport()

Application.DisplayAlerts = False

Dim TempDataSource As String
Dim TempEmployeeName As String

'To open the file containing data

Workbooks.Open Filename:=Range("B4").Value
TempDataSource = ActiveWorkbook.Name
ThisWorkbook.Activate
Range("B7").Select

'Loop through employee name listed in the file 'ReportGeneratorV2'
'Extract data for that particular employee and create a new excel file'

Do While ActiveCell.Value <> ""
 TempEmployeeName = ActiveCell.Value
 Workbooks(TempDataSource).Activate
 Range("A1").Select
 Range(Selection, Selection.End(xlDown)).Select
 Range(Selection, Selection.End(xlToRight)).Select
 Selection.AutoFilter Field:=1, Criteria1:=TempEmployeeName
 Selection.SpecialCells(xlCellTypeVisible).Select
 Selection.Copy
 Workbooks.Add
 Range("A1").PasteSpecial
 Selection.Columns.AutoFit
 Application.CutCopyMode = False
 ActiveSheet.Name = TempEmployeeName
 ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & TempEmployeeName & ".xlsx"
 ActiveWorkbook.Close
 Selection.AutoFilter
 ThisWorkbook.Activate
 ActiveCell.Offset(1, 0).Select
Loop

Workbooks(TempDataSource).Close
Application.DisplayAlerts = True
MsgBox "Macro Complete"

End Sub
Sub generatePDF2()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Dim TempDataSource As String
Dim TempEmployeeName As String

'To open the file containing data

Workbooks.Open Filename:=Range("B4").Value
TempDataSource = ActiveWorkbook.Name
ThisWorkbook.Activate
Range("B7").Select

'Loop through employee name listed in the file 'ReportGeneratorV2'
'Extract data for that particular employee and create a PDF file'

Do While ActiveCell.Value <> ""
 TempEmployeeName = ActiveCell.Value
 Workbooks(TempDataSource).Activate
 Range("A1").Select
 Range(Selection, Selection.End(xlDown)).Select
 Range(Selection, Selection.End(xlToRight)).Select
 Selection.AutoFilter Field:=1, Criteria1:=TempEmployeeName
 Selection.SpecialCells(xlCellTypeVisible).Select
 Selection.Copy
 Workbooks.Add
 Range("A1").PasteSpecial
 Selection.Columns.AutoFit
 Application.CutCopyMode = False
 ActiveSheet.Name = TempEmployeeName
 Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
 ThisWorkbook.Path & "\" & TempEmployeeName & ".PDF", Quality:= _
 xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
 OpenAfterPublish:=False
 ActiveWorkbook.Close
 Selection.AutoFilter
 ThisWorkbook.Activate
 ActiveCell.Offset(1, 0).Select
Loop

Workbooks(TempDataSource).Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Macro Complete"

End Sub










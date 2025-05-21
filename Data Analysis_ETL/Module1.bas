Attribute VB_Name = "Module1"
Option Explicit


Sub overnight_Report()

                      ' # # Error handling - Input Data Set # #

' For No data

If (ActiveSheet.Range("A1").Value <> "") And (ActiveSheet.Range("I1").Value <> "") Then

MsgBox "Please Clear Old data by Clicking Clear Data Button"

Exit Sub

End If

' for entries other than buy or sell

Dim Erow As Long, E As Long, side As String

Erow = ActiveSheet.Cells(Rows.Count, 4).End(xlUp).Row
For E = 1 To Erow
side = UCase(Trim(Range("D" & E)))

If side <> "B" And side <> "S" Then

MsgBox "Data Contains Other entries Other than Buy Or Sell. Please enter only Buy or Sell Data"

Exit Sub

End If

Next E

' For No Prop Traders
 
Dim PH As String, GRSQ As String

Erow = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row

E = 0

For E = 1 To Erow

   GRSQ = Left(Trim(Range("B" & E)), 2)
   
    If GRSQ = "GC" Then
     PH = GRSQ
     
    End If
     
Next E

If PH = "" Then

MsgBox "No Prop Traders"

Exit Sub

End If


                        ' # # Actual Report Processing # #
                        

'For filling ID

Dim i As Long
Dim Lrowcount As Long
Lrowcount = Range("A1").End(xlDown).Row
For i = 1 To Lrowcount
Range("A" & i).Value = "HBCL"
Next i

' For creating headers

Dim namerange As Range

Range("I1").Value = "ID"
Range("J1").Value = "Tag"
Range("K1").Value = "Symbol"
Range("L1").Value = "Side"
Range("M1").Value = "Quantity"
Range("N1").Value = "Price"
Range("O1").Value = "Value"

Set namerange = Range("I1:O1")

With namerange.Font
.Bold = True
 
End With


'For Separating GC & GGBT accounts

Dim brsq As String

For i = 1 To Lrowcount
  brsq = Left(Trim(Range("B" & i)), 2)
    
     If brsq = "GC" Then
     
    Range("B" & i).Select
    ActiveCell.offset(0, -1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("I100000").End(xlUp).Select
    ActiveCell.offset(1, 0).Select
    ActiveSheet.Paste
     
    
    ElseIf Trim(Range("B" & i).Value) = "GGBT" Then
    
    Range("B" & i).Select
    ActiveCell.offset(0, -1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("I100000").End(xlUp).Select
    ActiveCell.offset(1, 0).Select
    ActiveSheet.Paste
    
    ElseIf Trim(Range("B" & i).Value) = "GGPV" Then
    
    Range("B" & i).Select
    ActiveCell.offset(0, -1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("I100000").End(xlUp).Select
    ActiveCell.offset(1, 0).Select
    ActiveSheet.Paste
    
     End If
    
    
Next i
    
    Range("I1").Select


' For Calculating dollar value

Dim r As Long
Dim Erowcount As Long
Erowcount = Range("J1").End(xlDown).Row

For r = 2 To Erowcount

Range("O" & r).Value = (Range("M" & r).Value * Range("N" & r).Value)

Next r

' for sorting rows

Range("I1").CurrentRegion.Sort key1:=Range("L1"), order1:=xlAscending, Header:=xlYes

'for creating a line between buy and sell side

Dim Frowcount As Long
 
 Frowcount = Range("L1").End(xlDown).Row
 
 Range("I1").Select
 
For i = 2 To Frowcount
 
   If Trim(Range("L" & i)) = "S" Then
     Range("L" & i).Select
     ActiveCell.offset(0, -3).Select
     Range(Selection, Selection.End(xlToRight)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
     Exit For
    
   End If
Next i

' for BUY subtotal

Dim Bcount As Long, Scount As Long

Range("I1").End(xlDown).Select
ActiveCell.offset(1, 6).Select
Bcount = ActiveCell.End(xlUp).Row
ActiveSheet.Range("O" & Bcount + 1).Value = WorksheetFunction.Sum(Range("O2:O" & Bcount))
ActiveSheet.Range("O" & Bcount + 1).Select
ActiveCell.Font.Bold = True

'For sell subtotal

ActiveCell.offset(1, 0).Select
 
 If ActiveCell.Value > 0 Then
    Scount = ActiveCell.End(xlDown).Row
 ActiveSheet.Range("O" & Scount + 1).Value = WorksheetFunction.Sum(Range("O" & Bcount + 2 & ":" & "O" & Scount))
 ActiveSheet.Range("O" & Scount + 1).Select
 ActiveCell.Font.Bold = True
 End If
 
 Range("I1").CurrentRegion.Borders.LineStyle = True
 Range("N:O").NumberFormat = "[$$-en-US]#,##0.00"
 
 
                    ' # # Error Handing - Output Data # #
                    
 
 'for 0 Price in the filtered result
 
Dim Zrowcount As Long, Z As Long, symb(300) As Variant, Sym As String
i = 0
Zrowcount = ActiveSheet.Cells(Rows.Count, 14).End(xlUp).Row

    For Z = 2 To Zrowcount
    
        If Range("N" & Z).Value = 0 Then
            Range("N" & Z).Select
            ActiveCell.offset(0, -3).Select
            symb(i) = ActiveCell.Value
            Sym = Sym & symb(i) & vbCrLf
            i = i + 1
         End If
        
    Next Z
    
If Sym <> "" Then

MsgBox "The Price for below Symbols are 0" & vbCr & vbCr & "Please fill the price manually from Yahoo Finance" & vbCr & vbCr & Sym
 
End If

End Sub


Sub for_GR_Accounts()

Worksheets("GBOVERNIGHT").Select
Range("A1").CurrentRegion.Select
Selection.Copy
Worksheets("GR").Select
Range("A1").PasteSpecial
Range("A1").Select

' To find 0 symbols with 0 price

Dim Zrowcount As Long, Z As Long, symb(300) As Variant, Sym As String, i As Integer
i = 0
Zrowcount = ActiveSheet.Cells(Rows.Count, 14).End(xlUp).Row

    For Z = 2 To Zrowcount
    
        If Range("E" & Z).Value = 0 Then
            Range("E" & Z).Select
            ActiveCell.offset(0, -3).Select
            symb(i) = ActiveCell.Value
            Sym = Sym & symb(i) & vbCrLf
            i = i + 1
         End If
        
    Next Z
        i = 0

'For filling MPID - HBCL


Dim Lrowcount As Long
Lrowcount = Range("A1").End(xlDown).Row
For i = 1 To Lrowcount
Range("A" & i).Value = "HBCL"
Next i

' For creating headers

Dim namerange As Range

Range("I1").Value = "MPID"
Range("J1").Value = "BRSQ"
Range("K1").Value = "Symbol"
Range("L1").Value = "Side"
Range("M1").Value = "Quantity"
Range("N1").Value = "Price"
Range("O1").Value = "Value"

Set namerange = Range("I1:O1")

With namerange.Font
.Bold = True
 
End With


'For Separating GC & GGBT accounts

Dim brsq As String

For i = 1 To Lrowcount
  brsq = Left(Trim(Range("B" & i)), 2)
    
     If brsq = "GR" Then
     
    Range("B" & i).Select
    ActiveCell.offset(0, -1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("I100000").End(xlUp).Select
    ActiveCell.offset(1, 0).Select
    ActiveSheet.Paste
    
     End If
    
Next i
    
    Range("I1").Select


' For Calculating dollar value

Dim r As Long
Dim Erowcount As Long
Erowcount = Range("J1").End(xlDown).Row

For r = 2 To Erowcount

Range("O" & r).Value = (Range("M" & r).Value * Range("N" & r).Value)

Next r

' for sorting rows

Range("I1").CurrentRegion.Sort key1:=Range("L1"), order1:=xlAscending, Header:=xlYes

'for creating a line between buy and sell side

Dim Frowcount As Long
 
 Frowcount = Range("L1").End(xlDown).Row
 
 Range("I1").Select
 
For i = 2 To Frowcount
 
   If Trim(Range("L" & i)) = "S" Then
     Range("L" & i).Select
     ActiveCell.offset(0, -3).Select
     Range(Selection, Selection.End(xlToRight)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
     Exit For
    
   End If
Next i

' for BUY subtotal

Dim Bcount As Long, Scount As Long

Range("I1").End(xlDown).Select
ActiveCell.offset(1, 6).Select
Bcount = ActiveCell.End(xlUp).Row
ActiveSheet.Range("O" & Bcount + 1).Value = WorksheetFunction.Sum(Range("O2:O" & Bcount))
ActiveSheet.Range("O" & Bcount + 1).Select
ActiveCell.Font.Bold = True

'For sell subtotal
Dim First_sell_address As Variant
Dim Last_sell_address As Variant

ActiveCell.offset(1, 0).Select

If ActiveCell.Value > 0 Then
First_sell_address = ActiveCell.Address
Debug.Print First_sell_address
Range("O" & Rows.Count).End(xlUp).Select
Last_sell_address = ActiveCell.Address
Debug.Print Last_sell_address

'Scount = ActiveCell.End(xlDown).Row

 ActiveSheet.Range(Last_sell_address).offset(1, 0).Value = WorksheetFunction.Sum(Range(First_sell_address & ":" & Last_sell_address))
 ActiveSheet.Range(Last_sell_address).offset(1, 0).Select
 ActiveCell.Font.Bold = True
 End If
 
 Range("I1").CurrentRegion.Borders.LineStyle = True
 Range("N:O").NumberFormat = "[$$-en-US]#,##0.00"
 
 
                    ' # # Error Handing - Output Data # #
                    
 
 'for 0 Price in the filtered result
 

    
If Sym <> "" Then

MsgBox "The Price for below Symbols are 0" & vbCr & vbCr & "Please fill the price manually from Yahoo Finance" & vbCr & vbCr & Sym
 
End If

End Sub



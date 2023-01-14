Attribute VB_Name = "Filter_By_Engine"
Option Explicit

Sub Filter_By_Engine()

Dim shRead As Worksheet, shWrite As Worksheet, Criteria As Range
Set shRead = ThisWorkbook.Worksheets("Eng_Change_Data")
Set shWrite = ThisWorkbook.Worksheets("DashBoard")
Set Criteria = ThisWorkbook.Worksheets("DashBoard").Range("F1")


With shWrite
    ' Clear the data in output worksheet
    .Range("E2", "K26").Value = ""
    
    ' Set the cell formats
    .Columns(8).NumberFormat = "dd-MMM-yy"
    .Columns(10).NumberFormat = "dd-MMM-yy"
End With

' Get the range

Dim rg As Range, rng1 As Range, rng2 As Range, rng3 As Range, last As Long
Set rg = shRead.Range("A1").CurrentRegion

' Remove any existing filters
rg.AutoFilter

' Apply the Autofilter
With rg
    If Criteria = "" Then
        MsgBox "Select Criteria", vbCritical, "Warning"  'For Warning if criteria is not selected
        Exit Sub
    Else
    .AutoFilter Field:=13, Criteria1:=Criteria
    End If
End With

' Copy the data using Range Copy
last = shRead.Cells(Rows.count, "L").End(xlUp).Row

Set rng1 = Union(shRead.Range("D1:D" & last), shRead.Range("F1:F" & last), shRead.Range("H1:H" & last))
Set rng2 = Union(shRead.Range("C1:C" & last), shRead.Range("K1:L" & last))
Set rng3 = shRead.Range("I1:I" & last)

rng1.SpecialCells(xlCellTypeVisible).Copy
shWrite.Range("E2").PasteSpecial xlPasteValues

rng2.SpecialCells(xlCellTypeVisible).Copy
shWrite.Range("H2").PasteSpecial xlPasteValues

rng3.SpecialCells(xlCellTypeVisible).Copy
shWrite.Range("K2").PasteSpecial xlPasteValues

' Remove any existing filters
rg.AutoFilter

' Active the output sheet so it is visible
shWrite.Activate

End Sub



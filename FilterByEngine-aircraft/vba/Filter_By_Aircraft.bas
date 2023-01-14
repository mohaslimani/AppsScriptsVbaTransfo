Attribute VB_Name = "Filter_By_Aircraft"
Option Explicit

Sub Filter_By_Aircraft()

Dim shRead As Worksheet, shWrite As Worksheet, Criteria As Range, Criteria2 As Range
Set shRead = ThisWorkbook.Worksheets("Eng_Change_Data")
Set shWrite = ThisWorkbook.Worksheets("DashBoard")
Set Criteria = ThisWorkbook.Worksheets("DashBoard").Range("B29")
Set Criteria2 = ThisWorkbook.Worksheets("DashBoard").Range("F1")

With shWrite
    ' Clear the data in output worksheet
    .Range("E28", "K52").Value = ""
    
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
    If Criteria2 = "" Then
        rg.AutoFilter Field:=4, Criteria1:=Criteria
    Else
        rg.AutoFilter Field:=4, Criteria1:=Criteria
        rg.AutoFilter Field:=1, Criteria1:=Criteria2
    End If
End With

' Copy the data using UNION
last = shRead.Cells(Rows.count, "C").End(xlUp).Row

Set rng1 = Union(shRead.Range("D1:D" & last), shRead.Range("F1:F" & last), shRead.Range("H1:H" & last))
Set rng2 = Union(shRead.Range("C1:C" & last), shRead.Range("K1:L" & last))
Set rng3 = shRead.Range("I1:I" & last)

rng1.SpecialCells(xlCellTypeVisible).Copy
shWrite.Range("E28").PasteSpecial xlPasteValues

rng2.SpecialCells(xlCellTypeVisible).Copy
shWrite.Range("H28").PasteSpecial xlPasteValues

rng3.SpecialCells(xlCellTypeVisible).Copy
shWrite.Range("K28").PasteSpecial xlPasteValues

' Remove any existing filters
rg.AutoFilter

' Active the output sheet so it is visible
shWrite.Activate

End Sub




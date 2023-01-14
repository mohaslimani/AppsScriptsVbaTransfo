Attribute VB_Name = "M001_PNqty"
Sub PNqty()

Set WSTEMPLATES = Sheets("TEMPLATES")

Dim PNqty As Integer, i As Integer, lastcell As Integer, j As Integer
Dim jtext As String

PNqty = WSTEMPLATES.Cells(12, 3).Value
lastcell = PNqty + 14

i = 0

WSTEMPLATES.Range("A15:E35").Select
    Selection.ClearContents
With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

For i = 15 To lastcell
    j = i - 14
    WSTEMPLATES.Cells(i, 1).Value = j
    WSTEMPLATES.Cells(i, 1).Select
    With Selection
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    'Selection.Font.Bold = True
 
Next
    
End Sub


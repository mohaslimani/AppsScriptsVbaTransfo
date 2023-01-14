Attribute VB_Name = "M002_PN_RESET"
Sub PNRESET()

Set WSTEMPLATES = Sheets("TEMPLATES")


WSTEMPLATES.Range("A15:E35").Select
    Selection.ClearContents
With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
WSTEMPLATES.Range("C12").Select
Selection.ClearContents
    
WSTEMPLATES.Range("C12").Select
    
End Sub



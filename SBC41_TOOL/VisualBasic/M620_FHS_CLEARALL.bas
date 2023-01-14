Attribute VB_Name = "M620_FHS_CLEARALL"
Sub FHS_CLEARALL()

Set WSTEMPLATES = Sheets("TEMPLATES")

Range("N22:O24").Select
    Selection.ClearContents
    Range("K29:P48").Select
    Selection.ClearContents
    Range("K27:P28").Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    WSTEMPLATES.Range("N22").Select
   
End Sub

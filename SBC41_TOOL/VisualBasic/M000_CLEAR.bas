Attribute VB_Name = "M000_CLEAR"
Sub CLEARDATA()
Attribute CLEARDATA.VB_ProcData.VB_Invoke_Func = " \n14"

Set WSTEMPLATES = Sheets("TEMPLATES")

    Range("C6").Select
    Selection.ClearContents
      Range("C7").Select
    Selection.ClearContents
      Range("E6").Select
    Selection.ClearContents
      Range("E7").Select
    Selection.ClearContents
      Range("E8").Select
    Selection.ClearContents
      Range("E9").Select
    Selection.ClearContents
      Range("C12:C13").Select
      Selection.ClearContents
    Range("C15:C21").Select
      Selection.ClearContents
    Range("E12:E13").Select
      Selection.ClearContents
        Range("E15:E21").Select
      Selection.ClearContents
    
    WSTEMPLATES.Range("B15:E35").Select
    Selection.ClearContents
With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    WSTEMPLATES.Range("C12").Select
Selection.ClearContents

WSTEMPLATES.Range("C6").Select
    
    'WSTEMPLATES.Cells(30, 3).Formula = "=SUM(B36:C36)"
    '"=IFERROR(VLOOKUP(C7&C6,CARS!D2:L19091,2,FALSE), "")"
    
End Sub

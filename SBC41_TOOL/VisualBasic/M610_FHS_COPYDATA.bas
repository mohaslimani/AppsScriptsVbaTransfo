Attribute VB_Name = "M610_FHS_COPYDATA"
Sub FHS_COPYDATA()

Set WSTEMPLATES = Sheets("TEMPLATES")

WSTEMPLATES.Cells(22, 14).Value = WSTEMPLATES.Cells(14, 3).Value
WSTEMPLATES.Cells(23, 14).Value = WSTEMPLATES.Cells(6, 3).Value
WSTEMPLATES.Cells(24, 14).Value = WSTEMPLATES.Cells(9, 3).Value
   
End Sub

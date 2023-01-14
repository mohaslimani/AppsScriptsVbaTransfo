VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Reset champs de recherche suite à un double click
Private Sub MyDocSearch_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Sheets("Doc_check").MyDocSearch.Value = ""
End Sub
'Touche Entrée sur la recherche MyDoc
Private Sub MyDocSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then SearchIcon_Click
End Sub

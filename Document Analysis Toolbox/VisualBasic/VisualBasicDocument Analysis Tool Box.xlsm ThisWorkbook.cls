VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
Dim Rnge As Range

    Sheets("doc_check").MyDocSearch = "Search in MyDoc"
    Sheets("Doc_Check").TextBox21.Value = ""
    Sheets("Doc_Check").TextBox23.Value = ""
    Sheets("Doc_Check").RefTest1.Value = ""
    Sheets("Doc_Check").RefTest2.Value = ""
    
    'Sheets("Doc_Check").ListObjects("Abb").HeaderRowRange.Value = "RQC2.22: Are all abbreviations mentionned in the glossary?" _
    & Chr(10) & "List of Abbreviations NOT found in the Glossary"

    Range("List_Abb").ClearContents
    Set Rnge = Range("Abb[#All]").Resize(5, 1)
    Sheets("Doc_Check").ListObjects("Abb").Resize Rnge
    ActiveSheet.ListObjects("Abb").ShowTotals = False
    
    Range("List_Ref").ClearContents
    Range("MyDocStatus").ClearContents
    Set Rnge = Range("DocRef[#All]").Resize(5, 2)
    Sheets("Doc_Check").ListObjects("DocRef").Resize Rnge
    ActiveSheet.ListObjects("DocRef").ShowTotals = False
    
    Range("List_Wrd").ClearContents
    Set Rnge = Range("ForWrd[#All]").Resize(5, 1)
    Sheets("Doc_Check").ListObjects("ForWrd").Resize Rnge
    ActiveSheet.ListObjects("ForWrd").ShowTotals = False
    
    Range("E13").Value = ""
    Range("D2").Select
'    MyDocSearch.Value = "Search in MyDoc"
    
End Sub

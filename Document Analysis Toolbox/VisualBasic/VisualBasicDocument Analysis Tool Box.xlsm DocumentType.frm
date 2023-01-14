VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DocumentType 
   Caption         =   "Document Type"
   ClientHeight    =   2085
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4824
   OleObjectBlob   =   "VisualBasicDocument Analysis Tool Box.xlsm DocumentType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DocumentType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
Unload DocumentType
End Sub

Private Sub OK_Click()
    Reset_Click 'ALE: demande faite par QMI en Oct 2017, vider les résultats précédents

    ' le choix du type du doc détermine la liste des mots interdits listé dans le userform "CheckType":
    ' cette liste est le contenu de la "table 2" dans la feuille "Forbidden words"
    ' si type = process alors contenu de la Colonne G sinon colonne H
    Dim i, y As Integer
    i = 7
    y = 3
    
    While Sheets("ForbiddenWords").Cells(2, i).Value <> ""
        If DocumentType.ListDocType.Value = Sheets("ForbiddenWords").Cells(2, i) Then
            While Sheets("ForbiddenWords").Cells(y, i).Value <> ""
                CheckType.ListForbWord.ListItems.Add , , Sheets("ForbiddenWords").Cells(y, i).Value
                y = y + 1
            Wend
        End If
        i = i + 1
    Wend
    'Coche par défaut chacun des mots interdits
    For i = 1 To CheckType.ListForbWord.ListItems.count
            CheckType.ListForbWord.ListItems(i).Checked = True
    Next
    Unload DocumentType
    CheckType.Show
    
End Sub
Private Sub UserForm_Initialize()
' intitialisation de la combobox "choix du type de document
' recupère le type de doc dans la feuille de paramétrage "forbidden_words" cellules G2 et H2
' ces deux cellules sont les "headers" de la "table 2"
' Pour ajouter un type il suffit d'ajouter une colonne entre G et H

    Dim DType As String
    i = 7
    While Sheets("ForbiddenWords").Cells(2, i).Value <> ""
        DType = Sheets("ForbiddenWords").Cells(2, i).Value
        ListDocType.AddItem DType
        i = i + 1
    Wend
    
End Sub

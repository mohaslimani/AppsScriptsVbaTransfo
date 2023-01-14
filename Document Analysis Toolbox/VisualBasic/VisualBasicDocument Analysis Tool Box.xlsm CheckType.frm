VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CheckType 
   Caption         =   "Type of Check"
   ClientHeight    =   5130
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5892
   OleObjectBlob   =   "VisualBasicDocument Analysis Tool Box.xlsm CheckType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CheckType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub All_checks_Click()
'Case à cocher "All Checks" pour selectionner/Deselectionner les check à réaliser
    If All_checks = False Then
        CheckForbWords = False
        CheckBoxAbb = False
        CheckBoxDocRef = False
    Else
        CheckForbWords = True
        CheckBoxAbb = True
        CheckBoxDocRef = True
    End If
End Sub
Private Sub CheckForbWords_Click()
'Paramètre les checkBox des mots interdits: Coche/Déchoche la liste en utilisant CheckForbWords
    If CheckForbWords = True Then
        For i = 1 To CheckType.ListForbWord.ListItems.count
                    CheckType.ListForbWord.ListItems(i).Checked = True
        Next
    Else
        For i = 1 To CheckType.ListForbWord.ListItems.count
                    CheckType.ListForbWord.ListItems(i).Checked = False
        Next
    End If
End Sub

Private Sub RunCheck_Click()

    Dim WordDoc As Word.Document
    Dim wordApp As Word.Application
    Dim oPageStart As Word.Range
    Dim oSec As Word.Section
    Dim strpath As String
    Dim ListWrd As String
    Dim TbFbWrd() As String
    Dim Wkb As Workbook
    Dim Rnge As Range
    
    Dim dateLancement As String
    Dim resultsBookName As String
    Dim resultsBookPath As String
    
    Set Wkb = ThisWorkbook
    Wkb.Activate
    
    'Enleve le calcul auto des tables automatiques:
    'Pour comprendre l'utilisation des tables auto:
    'https://www.thespreadsheetguru.com/blog/2014/6/20/the-vba-guide-to-listobject-excel-tables
    Application.AutoCorrect.AutoFillFormulasInLists = False
    
    'Remise à Zéro de la page d'acceuil de l'outil:
    'Nettoie et redimensionne les tables "List_Abb", "List_Ref" et "List_Wrd" utilisée
    'pour afficher les résultats
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
    
    With Application
        .UseSystemSeparators = False
    End With
    
    'ouverture d'une boite de dialogue pour selection d'un fichier word
    'autre methode pour ouverture des fichiers "Application.GetOpenFilename"
    
    'Ouverture de l'application word
    Set wordApp = CreateObject("Word.Application")
'    wordApp.Activate
'    Application.ChangeFileOpenDirectory ("C:\")
    Application.FileDialog(msoFileDialogOpen).Title = "Select the Document to check"
    Application.FileDialog(msoFileDialogOpen).Filters.Clear
    Application.FileDialog(msoFileDialogOpen).Filters.Add "Word Documents Only", "*.docx ; *.doc; *.docm"
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    If Application.FileDialog(msoFileDialogOpen).Show <> 0 Then
        'strpath = lien vers le document ce lien est utilisé après pour recuperer le document
        strpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        'wordDoc est l'object Word.Document
        Set WordDoc = wordApp.Documents.Open(strpath)
        'Récupere le nom du document et saisie dans la page d'acceuil:
        Range("E13").Value = "Current Document: " & WordDoc.Name 'Test ALE 05/09/2017
    Else
        'si rien n'est selectionner on quitte la macro
        Application.Visible = True
        Unload CheckType
        Exit Sub
    End If
    
    wordApp.Visible = True
    wordApp.Activate
    Application.Visible = True
    
    ' Lancement check des Forbidden words:
    
    ListWrd = ""
    'Variable contenant chaque mot interdit à checker dans le document séparé par ";"
    'Boucle sur chaque mot interdit coché dans le user_From "CheckType"
'    et incrémente la variable ListWrd avec les mots mots cochés par l'utilisateur:
    For i = 1 To ListForbWord.ListItems.count
        If ListForbWord.ListItems(i).Checked = True Then
            If ListWrd = "" Then ListWrd = ListForbWord.ListItems(i).Text Else ListWrd = ListWrd & ";" & ListForbWord.ListItems(i).Text
        End If
    Next
    

    
'   Transforme la chaine de caractère ListWord en variable tableau TbFbWrd et appel de la macro
'   CheckForbiddenWord
    If ListWrd <> "" Then
        TbFbWrd = Split(ListWrd, ";")
        Unload CheckType
        Call CheckForbiddenWord(wordApp, strpath, TbFbWrd())
    End If
    
    Unload CheckType
    
    'Créé les paramètres servant la création du fichier de sortie pour les résultats
    dateLancement = Format(Now, "dd.mm.yyyy_hh.mm.ss")
    resultsBookName = Left(WordDoc.Name, Len(WordDoc.Name) - 5) & "_Check_results_" & dateLancement & ".xlsx"
    resultsBookPath = ""
    
    'Lancement du check sur les Abbreviations:
    If CheckBoxAbb = True Then
        Call CheckAbbreviation(strpath, wordApp, dateLancement, resultsBookName, resultsBookPath)
    End If
    
    'Lancement du check des references:
    If CheckBoxDocRef = True Then
        Call CheckReferences(strpath, wordApp, dateLancement, resultsBookName, resultsBookPath)
    End If
    
    ' Lancement de l'export des résultats
    Call Create_Result(WordDoc.Name, dateLancement, resultsBookName, resultsBookPath)
    
    'MsgBox "Check is completed"
    Application.StatusBar = "Check completed ! " & Now
    Range("E13").Value = "Current Document: " & WordDoc.Name & " - CHECK COMPLETED" 'Test ALE 05/10/2017
    
    'Code pour récuperer la reference du Template depuis les propriétés du document word:
'    Templ = WordDoc.BuiltinDocumentProperties(wdPropertyTemplate)
'    If InStr(1, Templ, "FM") <> 0 Then
'        TempRef = Mid(Templ, InStr(1, Templ, "FM"), InStr(1, Templ, ".do") - InStr(1, Templ, "FM"))
'        TempIss = Mid(TempRef, InStr(1, TempRef, "v") + 1, Len(TempRef) - InStr(1, TempRef, "v") + 1)
'        TempRef = Mid(TempRef, 1, 9)
'        Call LastTemplate(TempRef, TempIss)
'    Else
'        Range("E13").Value = Range("E13").Value & Chr(10) & "! Warning ! the template reference (" & Templ & ") used is not valid"
'    End If
    
'    WordApp.Visible = True
End Sub

Private Sub CommandButton2_Click()
Unload CheckType
End Sub
Private Sub UserForm_Initialize()

    With ListForbWord
    
        With .ColumnHeaders
            .Clear
            .Add , , "Please select word(s)", 180
        End With
         
        .View = lvwReport
        .Gridlines = True
        .CheckBoxes = True
    End With
    
End Sub


Attribute VB_Name = "Macro_Check"
Const messageTableRefNotFound As String = "Table of Reference could not be found"
Const valErreurNA As String = "N/A"
Const valErreur100 As String = "100%"
Dim AllRef As String
Dim NotFoundRef As String
' *********************************** PASSWORD for project PROTECTION : "altranqemm" *************************************************

Sub CheckReferences(strpath As String, wordApp As Word.Application, dateLancement As String, resultsBookName As String, resultsBookPath As String)
    On Error GoTo errorCheckRef
    Dim NbRefTot, iPage, i, CtTest1, NbRef, iRef As Integer
    'NbRefTot= Nombre D'abbriation total dans la table de reference
    'iPage= Numéro de la page ou se trouve la table de reference
    'CtTest1 = Compte le nombre ref repondant négativement au Test1 "RQC2.09" "Est-ce que la reference listée dans la table est aussi citée dans le document"
    
    Dim StartRg, EndRg, sRgTb, eRgTb As Long
    Dim Reference, RefFound, PSDlist, Separator As String
    'AllRef=liste toute les références
    Dim TbAllRef() As String
    Dim wrdNameFull() As String
    Dim Com, TabName As Variant
    
    Dim TabRef As Word.Table
    Dim rng, Rng2 As Word.Range
    Dim CellRef As Word.Range
    Dim WordDoc As Word.Document
    
    Dim Rnge As Range
    Dim Tbl As ListObject
    
    Dim length As Integer
    
    Set Wkb = ThisWorkbook
    Wkb.Activate
    'recupère le document word grace au lien strpath
    Set WordDoc = GetObject(strpath)
    wordApp.Visible = True
    wordApp.Activate
    
    wordApp.Selection.Collapse Direction:=wdCollapseStart
    '******************************************************
    '*** - Trouve la table de reference des documents - ***
    '******************************************************
'    iPage = 1
'    'Boucle sur toutes les manière d'appeler la table des référence UK/FR/ESP/GER
'    TabName = Array("Referenced Documents", "Documents de Référence", "Referenzdokumente", "Documentos de Referencia")
'    For i = 0 To UBound(TabName)
'        With wordApp.Selection.Find
'                .Forward = False
'                .ClearFormatting
'                .MatchWholeWord = True
'                .MatchCase = False
'                .Wrap = wdFindContinue
'                .Text = TabName(i)
'                If .Execute = True Then
'                    iPage = wordApp.Selection.Information(wdActiveEndPageNumber)
'                    'Recupere le numéro de page dans laquel se trouve
'                    'la table des références
'                    Exit For
'                End If
'        End With
'    Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim iPages() As Integer
    Dim count As Integer
    Dim found As Boolean
    Dim consultedPage As String
    iPage = 1
    
    TabName = Array("Referenced Documents", "Documents de Référence", "Referenzdokumente", "Documentos de Referencia")
    consultedPage = ""
    
    For i = 0 To UBound(TabName)
        wordApp.Selection.GoTo what:=wdGoToHeading, which:=wdGoToFirst
        wordApp.Selection.Collapse Direction:=wdCollapseStart
        
        'With ActiveDocument.Content.Find
        With wordApp.Selection.Find
            Do While .Execute(findText:=TabName(i), Format:=False, MatchCase:=False, MatchWholeWord:=True, Forward:=True, Wrap:=wdFindContinue) = True
                
                wordApp.Selection.Collapse Direction:=wdCollapseEnd
                wordApp.Selection.GoTo what:=wdGoToTable, which:=wdGoToNext
                
                If InStr(consultedPage, wordApp.Selection.Information(wdActiveEndPageNumber)) = 0 Then
                    consultedPage = consultedPage + CStr(wordApp.Selection.Information(wdActiveEndPageNumber)) + ";"
                    
                    wordApp.Selection.Tables(1).Select
                    sRgTb = wordApp.Selection.Range.Start
                    eRgTb = wordApp.Selection.Range.End
                    Set TabRef = wordApp.Selection.Tables(1)
                    
                    stringCell = Trim(Left(Trim(TabRef.Rows(1).Cells(1).Range.Text), Len(TabRef.Rows(1).Cells(1).Range.Text) - 2))
                    
                    If stringCell = "Doc. Reference - Title" Or stringCell = "Doc. Reference" Then
                        iPage = wordApp.Selection.Information(wdActiveEndPageNumber)
                        Exit For
                    End If
                    
                    wordApp.Selection.Move
                    'wordApp.Selection.GoTo What:=wdGoToPage, which:=lNextPage.
                Else
                    GoTo EndLoop
                End If
            Loop
        End With
EndLoop:
    Next
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If iPage = 1 Then
        Sheets("Doc_Check").ListObjects("DocRef").DataBodyRange(1, 1).Value = messageTableRefNotFound
        Exit Sub
    End If
    'wordApp.Selection.GoTo What:=wdGoToTable , which:=wdGoToNext
    'wordApp.Selection.Tables(1).Select
    'sRgTb = wordApp.Selection.Range.Start
    'range du debut de la table des reference dans le document word
    'eRgTb = wordApp.Selection.Range.End
    'range de la Fin de la table
    'Set TabRef = wordApp.Selection.Tables(1)
    'NbRefTot = TabRef.Rows.count
        
    '***************************************************************
    '*** - Check l'ensemble des documents listés dans la table - ***
    '***************************************************************
    
    wordApp.Selection.HomeKey Unit:=wdStory
    
    Set rng = WordDoc.Content 'Rng = All Document Text
    rng.Start = WordDoc.Content.Start
    StartRg = rng.Start
    WordDoc.GoTo wdGoToPage, wdGoToAbsolute, iPage
    rng.End = wordApp.Selection.Start
    EndRg = rng.End
    
    'On copie et recolle la table au meme endroit pour enlever les lignes fantomes soulevant des erreurs
    TabRef.Select
    wordApp.Selection.MoveDown
    TabRef.Range.Cut
    wordApp.Selection.MoveUp
    wordApp.Selection.EndKey Unit:=wdLine
    wordApp.Selection.TypeParagraph
    wordApp.Selection.Paste
    'wordApp.Selection.TypeBackspace

    wordApp.Selection.Tables(1).Select
    sRgTb = wordApp.Selection.Range.Start
    eRgTb = wordApp.Selection.Range.End
    Set TabRef = wordApp.Selection.Tables(1)

    AllRef = ""
    CtTest1 = 0
    NbRefTot = TabRef.Rows.count
    
    'Parcours du tableau des references, merge des cellules de chaque ligne et ajout du tiret si il est manquant
    For NbRef = 1 To NbRefTot
        checkValue = TabRef.Rows(NbRef).Cells.count
        If checkValue > 1 Then
            If TabRef.Rows(NbRef).Cells(1).Range.ContentControls.count > 0 Then
                TabRef.Rows(NbRef).Cells(1).Range.ContentControls.Item(1).Delete
            End If
            If TabRef.Rows(NbRef).Cells(2).Range.ContentControls.count > 0 Then
                TabRef.Rows(NbRef).Cells(2).Range.ContentControls.Item(1).Delete
            End If
            Value1 = Trim(Left(TabRef.Rows(NbRef).Cells(1).Range.Text, Len(TabRef.Rows(NbRef).Cells(1).Range.Text) - 2))
            Value2 = Trim(Left(TabRef.Rows(NbRef).Cells(2).Range.Text, Len(TabRef.Rows(NbRef).Cells(2).Range.Text) - 2))
            TabRef.Rows(NbRef).Cells.Merge
            If (InStr(Value1, " - ") = 0 And Value1 <> "") Then
                TabRef.Rows(NbRef).Cells(1).Range.Text = Value1 + " - " + Value2
            ElseIf (Value1 = "") Then
                TabRef.Rows(NbRef).Delete
            End If
        End If
    Next
    
    NbRefTot = TabRef.Rows.count
    
'Ancien traitement qui faisait suite à une demande, à conserver au cas où...
'    'Parcours du tableau des references, merge des cellules de chaque ligne et ajout du tiret
'        For NbRef = 1 To NbRefTot
'            checkValue = TabRef.Rows(NbRef).Cells.Count
'            If checkValue > 1 Then
'                Value1 = Trim(Left(TabRef.Rows(NbRef).Cells(1).Range.Text, Len(TabRef.Rows(NbRef).Cells(1).Range.Text) - 2))
'                Value2 = Trim(Left(TabRef.Rows(NbRef).Cells(2).Range.Text, Len(TabRef.Rows(NbRef).Cells(2).Range.Text) - 2))
'                TabRef.Rows(NbRef).Cells.Merge
'                If (Value1 <> "" And Value2 <> "") Then
'                    TabRef.Rows(NbRef).Cells(1).Range.Text = Value1 + " - " + Value2
'                Else
'                    TabRef.Rows(NbRef).Delete
'                End If
'            End If
'        Next

'Nouveau traitement faisant suite à une demande, faisant l'inverse de celui au dessus
'Parcours du tableau des references, séparation au besoin des cellules de chaque ligne en deux parties
    For NbRef = 1 To NbRefTot
        checkValue = TabRef.Rows(NbRef).Cells.count
        If checkValue = 1 Then
            TabRef.Rows(NbRef).Select
            ValueBase = TabRef.Rows(NbRef).Cells(1).Range.Text
            ValueBase = Replace(ValueBase, Chr(45), "-")
            ValueBase = Replace(ValueBase, Chr(150), "-")
            ValueBase = Replace(ValueBase, Chr(151), "-")
            
            Value1 = Trim(Left(ValueBase, InStr(1, ValueBase, " - ") - 1))
            Value2 = Trim(Right(ValueBase, Len(ValueBase) - InStr(1, ValueBase, " - ") - 1))
            Value2 = Left(Value2, Len(Value2) - 2)
            
            TabRef.Rows(NbRef).Cells.Split NumRows:=1, NumColumns:=2
            If (Value1 <> "" And Value2 <> "") Then
                TabRef.Rows(NbRef).Cells(1).Range.Text = Value1
                TabRef.Rows(NbRef).Cells(1).Range.Font.Size = 9
                TabRef.Rows(NbRef).Cells(1).Range.Font.ColorIndex = 1
                TabRef.Rows(NbRef).Cells(2).Range.Text = Value2
                TabRef.Rows(NbRef).Cells(2).Range.Font.Size = 9
                TabRef.Rows(NbRef).Cells(2).Range.Font.ColorIndex = 1
                TabRef.Rows(NbRef).Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            Else
                TabRef.Rows(NbRef).Delete
            End If
        End If
    Next
    
    TabRef.Columns(1).PreferredWidthType = wdPreferredWidthPoints
    TabRef.Columns(1).Width = 90
    
    TabRef.Columns(2).PreferredWidthType = wdPreferredWidthPoints
    TabRef.Columns(2).Width = 400
    
    For NbRef = 1 To NbRefTot
        rng.Start = StartRg
        rng.End = EndRg
        
'2 formats de table possible:
' table avec 1 seule colonne -> la Ref. est la partie de Gauche de la cellule et est
' délimitée par un " - " donc on récupère toute la partie gauche jusqu'au tiret
        If Len(TabRef.Cell(NbRef, 1).Range.Text) < 5 Then NbRef = NbRef + 1
        If NbRef <= NbRefTot Then
            If TabRef.Columns.count = 1 Then
                If Reference = "" Or Reference = "Doc. Reference" Or Reference = "Doc Reference" Then NbRef = NbRef + 1
                TabRef.Cell(NbRef, 1).Range = Replace(TabRef.Cell(NbRef, 1).Range, Chr(150), "-")
                Reference = Trim(Left(TabRef.Cell(NbRef, 1).Range.Text, InStr(1, TabRef.Cell(NbRef, 1).Range.Text, "-") - 1))
                'GoTo errorCheckRefCol1
                If AllRef = "" Then AllRef = Reference Else AllRef = AllRef & ";" & Reference
                'AllRef = AllRef & ";" & Reference
            Else
                If Reference = "" Or Reference = "Doc. Reference" Or Reference = "Doc Reference" Then NbRef = NbRef + 1
                'Reference = NetText(TabRef.Cell(NbRef, 1).Range.Text)
                Reference = NetText(TabRef.Cell(NbRef, 1).Range.Text)
                'Reference = Trim(Left(TabRef.Cell(NbRef, 1).Range.Text, InStr(1, TabRef.Cell(NbRef, 1).Range.Text, "-") - 1))
                If AllRef = "" Then AllRef = Reference Else AllRef = AllRef & ";" & Reference
            End If
            i = 0
        End If
        
'Recherche les Reference de la table dans le text du doc entre la page 1 et ipage
'Si le résultat de la recherche est positif alors rien
'Sinon la reference de la table est commentée

        With rng.Find
            .ClearFormatting
            .Forward = True
            .MatchWholeWord = True
            .Text = Trim(Reference)
            .Wrap = wdFindStop
            .Execute
            Set CellRef = TabRef.Cell(NbRef, 1).Range
            If .Execute(findText:=Trim(Reference)) = False Then
                WordDoc.Comments.Add CellRef, "[Mandatory] All References in the document Reference shall be referred elsewhere in the document."
                CtTest1 = CtTest1 + 1
            End If
        End With
    Next
    

    If NetText(TabRef.Cell(NbRef, 1).Range.Text) = "" _
    Or NetText(TabRef.Cell(NbRef, 1).Range.Text) = "Doc. Reference" _
    Or NetText(TabRef.Cell(NbRef, 1).Range.Text) = "Doc.Reference" _
    Or NetText(TabRef.Cell(NbRef, 1).Range.Text) = "Doc Reference" Then NbRefTot = NbRefTot - 1
    
'Nous pouvons déjà remplir la case de résultat RQC2.Q09: "Over all references listed the table, how many were use in the text?"
'NbRefTot= Nombre D'abbriation total dans la table de reference
    Sheets("Doc_Check").RefTest1.Value = Format((NbRefTot - CtTest1) / NbRefTot, "0%")
    
    '
    '***************************************************************
    '*** - Recherche des docs ref non présent dans la table  - ***
    '***************************************************************
    
    wordApp.Selection.GoTo wdGoToPage, wdGoToAbsolute, 1
    wrdNameFull = Split(WordDoc.Name, "_")
    docName = wrdNameFull(0)

    iRef = 0
    RefFound = ""
    Set rng = WordDoc.Content
'La plage de recherche utilisée ici sera l'ensemble du document

    PSDlist = "UG;PR;RQ;RP;SP;TM;PL;SC;FM;PS;PSD;"
    Separator = RegKeyRead("HKEY_CURRENT_USER\Control Panel\International\sList")
    
    Com = Array( _
    "[Mandatory]" & Chr(10) & "All Doc References in the Document shall be referred elsewhere in the table.", _
    "[Mandatory]" & Chr(10) & "Procedural Support Document (PSD) shall not be refered in a procedural document, except for Tool User Guide (TUG) & List (LS) PSD." & _
    Chr(10) & "Please check that this document is not a PSD." & Chr(10) & "RQC2# 10" & "Source: A1000.1")
'La variable Com admets 2 arguments : un commentaire pour les PSD et un pour les autres...

        With rng.Find
            .ClearFormatting
            .Forward = True
            .MatchWildcards = True
            .IgnoreSpace = True
            .Text = "<[A-Z]{1" & Separator & "3}[0-9]{3" & Separator & "8}>"
            'To try : <[A-Z]{1,3}[0-9]{3,8}(.[0-9]{1,3})*>
            .Wrap = wdFindStop
            Do While .Execute = True
                i = 0
                Reference = rng.Text
                StartRg = rng.Start
                EndRg = rng.End
                If Left(Reference, 3) <> "ISO" And Left(Reference, 2) <> "EN" And Reference <> docName Then
                    Debug.Print rng
                    If rng.Start > sRgTb And rng.End < eRgTb Then 'check si la reference trouvée est déjà la dans table
                        rng.Collapse Word.WdCollapseDirection.wdCollapseEnd
                    Else
                        If rng.InRange(WordDoc.TablesOfContents(1).Range) = False Then
    ' les conditions suivantes permettent de recuperer la réference entière du document
                                If Right(WordDoc.Range(StartRg, EndRg + 2).Text, 2) = "." & Right(WordDoc.Range(StartRg, EndRg + 2).Text, 1) _
                                And IsNumeric(Right(WordDoc.Range(StartRg, EndRg + 2).Text, 1)) = True Then
                                
                                    If IsNumeric(Right(WordDoc.Range(StartRg, EndRg + 3).Text, 1)) = True Then
                                        If Right(WordDoc.Range(StartRg, EndRg + 4).Text, 1) = "." Then
                                            j = 5
                                            Do While IsNumeric(Right(WordDoc.Range(StartRg, EndRg + j).Text, 1)) = True
                                                j = j + 1
                                            Loop
                                            If j = 5 Then Reference = WordDoc.Range(StartRg, EndRg + 3).Text Else Reference = WordDoc.Range(StartRg, EndRg + j).Text
                                        Else
                                             Reference = WordDoc.Range(StartRg, EndRg + 3).Text
                                        End If
                                    Else
                                        If Right(WordDoc.Range(StartRg, EndRg + 3).Text, 1) = "." Then
                                            j = 4
                                            Do While IsNumeric(Right(WordDoc.Range(StartRg, EndRg + j).Text, 1)) = True
                                                j = j + 1
                                            Loop
                                            If j = 4 Then Reference = WordDoc.Range(StartRg, EndRg + 2).Text Else Reference = WordDoc.Range(StartRg, EndRg + j).Text
                                        Else
                                             Reference = WordDoc.Range(StartRg, EndRg + 2).Text
                                        End If
                                    End If
                                End If
                                
                                If IsNumeric(Right(Reference, 1)) = False Then Reference = Left(Reference, Len(Reference) - 1)
                                Reference = Trim(Reference)
                                
                                If IsRef(Reference) = True Then
                                    If InStr(1, AllRef, Reference) = 0 Then
                                        If AllRef = "" Then AllRef = Reference Else AllRef = AllRef & ";" & Reference
                                            If RefFound = "" Then RefFound = Reference Else RefFound = RefFound & ";" & Reference
                                            WordDoc.Bookmarks.Add Name:="Ref" & iRef, Range:=rng
    ' Chaque Référence trouvée on creer marque page sur sa position dans le Document afin d'utiliser ensuite des liens vers les résultats du check depuis la page d'accueil de l'outil
                                            iRef = iRef + 1
                                            If InStr(1, PSDlist, Left(Reference, 2)) Then
                                                WordDoc.Comments.Add rng, Com(1)
                                            Else
                                                WordDoc.Comments.Add rng, Com(0)
                                            End If
                                            NotFoundRef = NotFoundRef & Reference & ";"
                                                
    ' commente la reférence trouvée en function de son type
                                    End If
                                End If
                        End If
                        rng.Collapse Word.WdCollapseDirection.wdCollapseEnd
                    End If
                End If
           Loop
        End With
    
'Affichage des résultats dans la page d'acceuil de l'ouil:
'Table pour lister les références trouvées= "DocRef"
    If RefFound = "" Then
        Sheets("Doc_Check").ListObjects("DocRef").DataBodyRange(1, 1).Value = valErreurNA 'demande faite par QMI en Oct 2017, modifier "No Reference found" par N/A
        Sheets("Doc_Check").RefTest2.Value = valErreur100 'demande faite par QME - ALE 15/09/2017
    Else
        TbAllRef = Split(RefFound, ";")
        Dim nbCellInTbAllRef As Integer
        'BEFORE ALE
        'Set Rnge = Range("DocRef[#All]").Resize(UBound(TbAllRef) + 2, 2)
        
        'AFTER ALE 15/09/2017
        If (UBound(TbAllRef) = 0) Then
            nbCellInTbAllRef = 1
        Else
            nbCellInTbAllRef = UBound(TbAllRef)
        End If
        
        Set Rnge = Range("DocRef[#All]").Resize(nbCellInTbAllRef + 2, 2)
        
        Sheets("Doc_Check").ListObjects("DocRef").Resize Rnge
        
        Sheets("Doc_Check").ListObjects("DocRef").ShowTotals = False
        
        For y = 0 To UBound(TbAllRef)
        'Creation du lien hypertexte vers les différents marques page mis sur les références trouvées dans le document
            Sheets("Doc_Check").ListObjects("DocRef").DataBodyRange(y + 1, 1).Formula = "=HYPERLINK(" & """" & "[" & strpath & "]" & "Ref" & y & """" & "," & """" & TbAllRef(y) & """" & ")"
        Next y
        
        Sheets("Doc_Check").RefTest2.Value = Format(NbRefTot / (NbRefTot + y), "0%")
        
        ' Créé ou implémente le fichier de sortie contenant les résultats - Ancienne méthode d'appel
        ' insertResult = insert_Result(WordDoc.Name, "Doc.Ref", Format((NbRefTot - CtTest1) / NbRefTot, "0%"), Format(NbRefTot / (NbRefTot + y), "0%"), dateLancement, resultsBookName, resultsBookPath)
        
        Set Rnge = Range("DocRef[#All]").Resize(y + 1, 2)
        Sheets("Doc_Check").ListObjects("DocRef").Resize Rnge
    End If
    Exit Sub
    
errorCheckRefCol1:
    retMessage = MsgBox("Impossible to complete the 'Quality check'. The template of your input files is incorrect (only one column in 'Referenced Documents' chapter instead of 2 columns). Please update your file and try again.", vbExclamation, "Quality Check Error")
    Exit Sub
    
errorCheckRef:
    'retMessage = MsgBox("Impossible to complete the 'Quality check'. The template of your input files is incorrect. Please update your file and try again.", vbExclamation, "Quality Check Error")
    Sheets("Doc_Check").Range("G14").Value = "Invalid data or template, " & vbCrLf & "please verify your input file."
    Sheets("Doc_Check").Range("G14").Font.ColorIndex = 3
    Sheets("Doc_Check").Range("G14").Font.Bold = True
    Sheets("Doc_Check").Range("G14").Font.Size = 11
    Sheets("Doc_Check").Range("G14").Interior.ColorIndex = 6
    Resume Next
    
End Sub
Sub CheckForbiddenWord(wordApp As Word.Application, strpath As String, TabForbWords() As String)
On Error GoTo errorForbiddenWord
    Dim WordDoc As Word.Document
    Dim rng As Word.Range
    Dim WordFind, ForbWord As String
    Dim NbWord, iWrdFd, iFWrd As Integer
    
'   recupère le document word grace au lien strpath:
    Set WordDoc = GetObject(strpath)
    Set Wkb = ThisWorkbook
    Set rng = WordDoc.Content
' Rng = plage de recherche utilisée = l'ensemble du document
    
    Wkb.Activate
    StartRg = rng.Start
    EndRg = rng.End
    WordFind = ""
    iWrdFd = 0
    NbWord = Range("B65536").End(xlUp).row
    iFWrd = 0
    
'les 3 plages suivantes celles dans lesquelles on ne souhaite pas renvoyer de résultat
' Pour cela on récupère le debut et la fin de chacune des ces plages
' Et l'on conditionne le résultat pour qu'il soit hors des ces 3 plages:

' 1)Plage de la table des matière:
    wordApp.Selection.HomeKey Unit:=wdStory
'retour au debut du Doc
    wordApp.Selection.GoTo what:=wdGoToField, which:=GoToNext, Name:="TOC"
'Se déplace vers la table des matière "TOC"=Table of content

    SToC = wordApp.Selection.Range.Start
'SToC = Debut de la plage Table des matière
    wordApp.Selection.GoTo what:=wdGoToHeading, which:=GoToNext
'On se déplace vers le premier titre  après la TOC on l'on marque le debut de ce nouveau paragraphe comme la fin de la TOC
    EToC = wordApp.Selection.Range.Start
'EToC = Debut de la plage Table des matière

'    2)Plage de la section 3 du document (section avec les information sur le doc owner ect..)
    Set rngS3 = WordDoc.Sections(3).Range
    SRngS3 = rngS3.Start
    ERngS3 = rngS3.End
    
'    3)Plage de la table des contributeurs: cette fois-ci pour identifier cette partie du document il nous faut
'   utiliser une recherche par le titre "Contributors"
    wordApp.Selection.HomeKey Unit:=wdStory
    With wordApp.Selection.Find
        .Forward = False
        .ClearFormatting
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindContinue
        .Text = "Contributors"
        Do While .Execute
            If wordApp.Selection.Style.Font.Bold = True Then
                wordApp.Selection.GoTo what:=wdGoToTable, which:=wdGoToNext
                wordApp.Selection.Tables(1).Range.Select
                sRgTb = wordApp.Selection.Range.Start
                wordApp.Selection.GoTo what:=wdGoToHeading, which:=wdGoToNext
                eRgTb = wordApp.Selection.Range.End
                Exit Do
            End If
        Loop
    End With
    
    
    Dim copyrightPara As String
    Dim copyrightPage As Integer
    
' Cherche et supprime le paragraphe du Copyright Airbus pour l'exclure du check - Ancienne méthode
'    With wordApp.Selection.Find
'        .Forward = False
'        .ClearFormatting
'        .MatchWholeWord = True
'        .MatchCase = False
'        .Wrap = wdFindContinue
'        .Text = "This document and all information contained herein is the sole property of AIRBUS S.A.S."
'        Do While .Execute
'            wordApp.Selection.StartOf Unit:=wdParagraph
'            wordApp.Selection.MoveEnd Unit:=wdParagraph
'            wordApp.Selection.Select
'            copyrightPara = wordApp.Selection.Text
'            copyrightPage = wordApp.Selection.Information(wdActiveEndPageNumber)
'            'wordApp.Selection.Delete
'        Loop
'    End With
    
    
'TabForbWords: pour rappel cette variable est un paramètre de lors de l'appel de la macro et defini dans le UserForm "CheckType"
'elle liste l'ensemble des mots inderdits que l'ulisateur a selectionné pour son check
    For y = 0 To UBound(TabForbWords)
    
        ForbWord = TabForbWords(y)
        
            iWrdFd = 0
            rng.Start = StartRg
            rng.End = EndRg
' Lance la recherche du mot inderdit:
            With rng.Find
                .ClearFormatting
                .MatchWholeWord = True
                .MatchCase = False
                Do While .Execute(findText:=ForbWord) = True
' condition liés au plage interdites:
                    If Not ((rng.Start >= sRgTb And rng.End <= eRgTb) Or (rng.Start >= SRngS3 And rng.End <= ERngS3) Or (rng.Start >= SToC And rng.End <= EToC)) Then
                        rng.Select
                        If wordApp.Selection.Information(wdActiveEndPageNumber) <> 1 Then
                            iWrdFd = iWrdFd + 1
    'iWrdFd compte le nombre de fois où le mot a été trouvé dans le doc
                            i = 3
    'Afin de recuperer le commentaire associé au mot inderdit:
    'on boucle sur les cellules de la table "ForbWords" dans la Feuille "ForbiddenWords"
    ' Et l'on recupere le commentaire de la colonne E (#5)
                            Do While Sheets("ForbiddenWords").Cells(i, 2).Value <> ""
                                If ForbWord = Sheets("ForbiddenWords").Cells(i, 2).Value Then
                                    Debug.Print rng
                                    WordDoc.Comments.Add rng, Sheets("ForbiddenWords").Cells(i, 5).Value
                                End If
                                i = i + 1
                            Loop
                        End If
                    End If
                Loop
'Affiche le résultat de la recherche un par un dans la page d'acceuil:
                If iWrdFd > 0 Then
                    Wkb.Sheets("Doc_Check").ListObjects("ForWrd").DataBodyRange(iFWrd + 1, 1).Value = ForbWord & ": " & iWrdFd & " time(s)"
                    iFWrd = iFWrd + 1
                End If
                
            End With
    Next y
    
' Cherche le paragraphe du Copyright Airbus et y supprime tous les commentaires du check - Demande client
    With wordApp.Selection.Find
        .Forward = False
        .ClearFormatting
        .MatchWholeWord = True
        .MatchCase = False
        .Wrap = wdFindContinue
        .Text = "This document and all information contained herein is the sole property of AIRBUS S.A.S."
        .Execute
            wordApp.Selection.StartOf Unit:=wdParagraph
            wordApp.Selection.MoveEnd Unit:=wdParagraph
            wordApp.Selection.Select
            copyrightPara = wordApp.Selection.Text
            copyrightPage = wordApp.Selection.Information(wdActiveEndPageNumber)
            While wordApp.Selection.Comments.count <> 0
                wordApp.Selection.Comments.Item(1).Delete
            Wend
    End With
            
' Replace le paragraphe du Copyright Aribus à la fin du doc et le remet en forme - Ancienne méthode
'    wordApp.Selection.EndKey Unit:=wdStory
'    wordapp.Selection.GoTo What:=Page, which=copyrightPage
'    wordApp.Selection.Text = copyrightPara
'    wordApp.Selection.Font.Size = 8
'    wordApp.Selection.Borders(wdBorderLeft).Visible = True
'    wordApp.Selection.Borders(wdBorderTop).Visible = True
'    wordApp.Selection.Borders(wdBorderRight).Visible = True
'    wordApp.Selection.Borders(wdBorderBottom).Visible = True
    
    Exit Sub
errorForbiddenWord:
    'retMessage = MsgBox("Impossible to check the 'Forbidden Word'. Check the template of your input files.", vbExclamation, "Forbidden Words Error")
    Resume Next
End Sub
Sub CheckAbbreviation(strpath As String, wordApp As Word.Application, dateLancement As String, resultsBookName As String, resultsBookPath As String)
On Error GoTo errorCheckAbb
    Dim WordDoc As Word.Document
    Dim iPage, iAbbGloss, i, iAbbNotInGloss, iTotAbbGloss, iFullAbbNotSet, iTotFullAbbNotSe, iRgx As Integer
'iPage = recupère le numéro de la page dans laquelle se trouve le glossaire
'iTotAbbGloss = nombre d'abbréviation dans le glossaire
'iAbbGloss = abbréviation n°i dans le glossaire
'iAbbNotInGloss = compte le nombre d'abbrevations du glossaire qui n'ont pas été utilisé dans le texte du doc

    Dim iTotAbbFd, iAbbFd As Integer
'iAbbFd compte les abbréviation trouvé dans le texte qui répondent positivement au test

    Dim Abbreviation, Separator, FullAbb, Sigle, Liste, ListAbbFd, Gloss, TbAbbFd() As String
' Abbreviation = abbréviation listé dans la 1er colonne du glossaire
' FullAbb = Abbréviation en toute lettre listée dans la 2nd colonne du glossaire
' Sigle = chaine de caractère trouvé par la recherche RegEx
' Gloss= chaine de caractère qui liste toutes les abbréviation du glossaire séparé par ";"
' ListAbbFd = chaine de caractère qui liste chaque sigle considéré comme une abbréviation (separé par ";")

    Dim StartRg, EndRg, ERng, SRng, SRngS3, ERngS3 As Long
    Dim AllSigle, regEx, GlossName As Variant
'AllSigle= chaine de caractère qui liste chaque sigle trouvé par la recherche(par RegEx) dans le texte en les separant par ";"

    Dim TabGloss As Table
'TabGloss = table du glossaire
    Dim rng, Rng2, Rng3, rngS3 As Word.Range
    Dim siTemp As Word.SynonymInfo
    Dim isProcess, isInGlos As Boolean
    Dim Rnge As Range
    Dim Tbl As ListObject
    
    iAbbNotInGloss = 0
    iFullAbbNotSet = 0
    iAbbFd = 0
    iPage = 1
    Gloss = ""
    
    Application.ScreenUpdating = False
    
        Set WordDoc = GetObject(strpath)
        
'les 3 plages suivantes celles dans lesquelles on ne souhaite pas renvoyer de résultat
' Pour cela on récupère le debut et la fin de chacune des ces plages
' Et l'on conditionne le résultat pour qu'il soit hors des ces 3 plages:
'(Aussi utilisé dans CheckForbidden words)
        'Plage de la table des matière
        wordApp.Selection.HomeKey Unit:=wdStory
        wordApp.Selection.GoTo what:=wdGoToField, which:=GoToNext, Name:="TOC"
        SToC = wordApp.Selection.Range.Start
        wordApp.Selection.GoTo what:=wdGoToHeading, which:=GoToNext
        EToC = wordApp.Selection.Range.Start
        
        'Plage de la section 3 du document (section avec les information sur le doc owner ect..)
        Set rngS3 = WordDoc.Sections(3).Range
        SRngS3 = rngS3.Start
        ERngS3 = rngS3.End
        
        'Plage de la table des contributeurs
        wordApp.Selection.HomeKey Unit:=wdStory
        With wordApp.Selection.Find
            .Forward = False
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
            .Wrap = wdFindContinue
            .Text = "Contributors"
            Do While .Execute
                If wordApp.Selection.Style.Font.Bold = True Then
                    wordApp.Selection.GoTo what:=wdGoToTable, which:=wdGoToNext
                    wordApp.Selection.Tables(1).Range.Select
                    sRgTb = wordApp.Selection.Range.Start
                    wordApp.Selection.GoTo what:=wdGoToHeading, which:=wdGoToNext
                    eRgTb = wordApp.Selection.Range.End
                    Exit Do
                End If
            Loop
        End With
        
'Recherche la Plage du Glossaire: Boucle sur toutes les langues utilisées chez Airbus
        wordApp.Selection.HomeKey Unit:=wdStory
        
        GlossName = Array("Glossary", "Glossaire", "Glossar", "Glosario")
        For i = 0 To UBound(GlossName)
            wordApp.Selection.Collapse Direction:=wdCollapseStart
            With wordApp.Selection.Find
                    .ClearFormatting
                    .Forward = False
                    .MatchWholeWord = True
                    .MatchCase = False
                    .Wrap = wdFindContinue
                    .Text = GlossName(i)
                    If .Execute = True Then
                        iPage = wordApp.Selection.Information(wdActiveEndPageNumber)
'iPage -> recupère le numéro de la page dans laquelle se trouve le glossaire
                        Exit For
                    End If
            End With
        Next
        
'Si le glossaire n'est pas trouvé (la variable iPage n'a pas changé) affiche un message à l'utilisateur dans la table des abbreviations
        If iPage = 1 Then
            Sheets("Doc_Check").ListObjects("Abb").DataBodyRange(1, 1) = "Glossary could'nt be found in the document, please contact us"
            Application.Visible = True
            Application.Activate
            Exit Sub
        End If
        
        iPage = wordApp.Selection.Information(wdActiveEndPageNumber) ' Page du Glossary
        wordApp.Selection.GoTo what:=wdGoToTable, which:=wdGoToNext
        
        With wordApp.Selection.Tables(1).Columns(1)
            .PreferredWidthType = wdPreferredWidthPoints
            .PreferredWidth = 56.7
        End With
        
        'wordApp.Selection.Tables(1).Columns(1).SetWidth _
        ColumnWidth:=InchesToPoints(0.92), _
        RulerStyle:=wdAdjustFirstColumn
        
        wordApp.Selection.Tables(1).Columns(1).Select
        Set TabGloss = wordApp.Selection.Tables(1)
'TabGloss = table du glossaire
        iTotAbbGloss = TabGloss.Rows.count

'Defini le range (rng) de recherche: du debut du doc à la page du glossaire
        Set rng = WordDoc.Content
        rng.Start = WordDoc.Content.Start
        WordDoc.GoTo wdGoToPage, wdGoToAbsolute, iPage
        rng.End = wordApp.Selection.Start
        StartRg = rng.Start
        EndRg = rng.End
    
        For iAbbGloss = 1 To iTotAbbGloss
            rng.Start = StartRg
            rng.End = EndRg
            FullAbb = NetText(TabGloss.Cell(iAbbGloss, 2).Range.Text)
' FullAbb = Abbréviation en toute lettre listée dans la 2nd colonne du glossaire
            Abbreviation = NetText(TabGloss.Cell(iAbbGloss, 1).Range.Text)
            If Abbreviation = "" Then iAbbGloss = iAbbGloss + 1
            Abbreviation = NetText(TabGloss.Cell(iAbbGloss, 1).Range.Text)
            i = 0
            Gloss = Gloss & ";" & Abbreviation
            
'Recherche sur la plage rng:
            With rng.Find
                .ClearFormatting
                .Forward = True
                .MatchCase = True
                .MatchWholeWord = True
                .Text = Abbreviation
                .Wrap = wdFindStop
                Do While .Execute(findText:=Abbreviation) = True
                    i = i + 1
' une fois la recherche lancé: "rng" deviens la plage sur laquelle a été trouvé l'abbréviation
' on peut donc utiliser la condition des 3 plages à eviter:
' verifie si la nouvelle plage du résultat est dans l'une des 3 plages
                    If Not ((rng.Start >= sRgTb And rng.End <= eRgTb) Or (rng.Start >= SRngS3 And rng.End <= ERngS3) Or (rng.Start >= SToC And rng.End <= EToC)) Then
                        ERng = rng.End
                        SRng = rng.Start
                        
' la première fois où l'abbrévation est trouvée (i=1) on check si celle-ci a été écrite en toute lettre
                        If i = 1 Then
'La plage de recherche deviens alors le paragraphe dans lequel l'abbréviation est trouvé
                            rng.Expand Unit:=wdParagraph
                            With rng.Find
                                .ClearFormatting
                                .Forward = True
                                .MatchCase = False
                                .IgnorePunct = True
                                .IgnoreSpace = True
                                On Error Resume Next
                                If .Execute(findText:=FullAbb) = False Then
' Dans le cas l'abbreviation en toute lettre n'a pas été trouvé dans le paragraphe on commente l'abbreviation dans le texte:
                                    WordDoc.Comments.Add WordDoc.Range(SRng, ERng), "[Mandatory]" & Chr(10) & _
                                    "When an acronym/abbreviation is used for the 1st time in the document even in the titles, please:" _
                                    & Chr(10) & "-Write the full words;" & Chr(10) & "-Write the acronym/abbreviation in brackets, e.g.: Design Data Set (DDS)."
                                    iFullAbbNotSet = iFullAbbNotSet + 1
'iFullAbbNotSet = Compte le nombre d'abbréviation du glossaire qui n'ont pas été éecrite en toute lettre dans le texte
                                End If
                            End With
                            Exit Do
                        End If
                    End If
                Loop
                
' si i=0 c'est que la recherche n'a rien donné on peut donc commenter l'abbréviation directement dans le glossaire
                If i = 0 Then
                    Set CellAbb = TabGloss.Cell(iAbbGloss, 1).Range
                    WordDoc.Comments.Add CellAbb, "[Mandatory] All abbreviations in the Glossary shall be referred elsewhere in the document."
                    iAbbNotInGloss = iAbbNotInGloss + 1
' iAbbNotInGloss = compte le nombre d'abbrevations du glossaire qui n'ont pas été utilisé dans le texte du doc
                End If
            End With
        Next
        
    '    ----------------------------------------------------------------------------------------------------------------------------------------------------------------
    '    revue de l'ensemble des abbreviations dans le texte et check si est dans le Glossaire
        
        Set rng = WordDoc.Content
        StartRg = rng.Start
        EndRg = rng.End
    
        wordApp.Selection.HomeKey Unit:=wdStory
        
'RegEx pour retrouver toute chaine en MAJUSCULE: 2 types en fonction du séparateur utilisé par le user:
        Separator = RegKeyRead("HKEY_CURRENT_USER\Control Panel\International\sList")
        regEx = Array("<[A-Z]{2" & Separator & "7}>", _
                      "<[A-Z]{1" & Separator & "2}[&]{1}[A-Z]{1" & Separator & "2}>", _
                      "<[A-Z]{1" & Separator & "2}[.]{1}[A-Z]{1" & Separator & "2}>", _
                      "<[A-Z]{1}[a-z]{1}[A-Z]{1}>", _
                      "<[A-Z]{1}[/]{1}[A-Z]{1" & Separator & "2}>")
                      
        Sigle = ""
        AllSigle = ""
        ListAbbFd = ""
'On commence par boucler sur les 5 différentes RegEx:
        For iRgx = 0 To 4
            With rng.Find
                .ClearFormatting
                .Forward = True
                .Wrap = wdFindStop
                .MatchWildcards = True
'Pour utiliser un RegEx dans une recherche il faut utiliser la recherche par WildCards
                Do While .Execute(findText:=regEx(iRgx)) = True
                    'wordApp.Selection.GoTo What:=wdGoToTable, which:=wdGoToNext
                    rng.Select
                    If wordApp.Selection.Information(wdActiveEndPageNumber) <> 1 Then
                        isProcess = False
                        If (rng.Start >= sRgTb And rng.End <= eRgTb) Or (rng.Start >= SRngS3 And rng.End <= ERngS3) Or (rng.Start >= SToC And rng.End <= EToC) Then
                            rng.Collapse Word.WdCollapseDirection.wdCollapseEnd
                        Else
'                       check si Abbreviation est un nom de Process: Variable IsProcess de type booléen
' En effet les nom de process pose pb dans notre recherche (avec la RegEx n°1)car ils remontent de cette manière:
' Exemple: "MP.AC": la recherche va remonter "MP" dans une premier temps et "AC" dans un second tps
' pour eviter cela on analyse ce qui suit/précéde le resultat de la recherche
' s'il est suivi d'un "." alors on considère que c'est un process si après ce point il n'y pas d'espace
' même chose s'il est précédé d'un "." alors on considère que c'est un process
                            If rng.Start > 3 Or iRgx = 1 Then
                                Set Rng2 = WordDoc.Range(rng.Start, rng.End + 2)
                                Set Rng3 = WordDoc.Range(rng.Start - 1, rng.End)
                                If InStr(1, Rng2.Text, ".") Then
                                    isProcess = False
                                    If Right(Rng2.Text, 2) <> ". " Then isProcess = True
                                End If
                                If Left(Rng3.Text, 1) = "." Then isProcess = True
                            End If
                            
                            Sigle = rng.Text
                            If isProcess = False Then
' La recherche peut remonter le même résultat autant de fois qu'il est présent dans le texte
' pour éviter cela on stock chaque résultat dans la variable "AllSigle"
' Il suffit ensuite de vérifier rapidement si le sigle trouvé a déjà été analyser:
' De plus en utilisant la fonction "IsAbb" on peut rapidement regarder si le resultat correspond
' à la liste des NonAbb de la Feuille forbidden word
                                If InStr(1, AllSigle, Sigle) = 0 And IsAbb(Sigle) = True Then
                                    AllSigle = AllSigle & ";" & Sigle
                                    
' La condition suivante permet de verifier si le sigle trouvé est dans le glossaire (variable gloss)
                                    If InStr(1, Gloss, Sigle) = 0 Then
                                    
'Utilisation du Thesaurus pour enlever les mots en majuscule:
'Le test: Si le sigle trouvé est dans le dictionnaire alors on considère que c'est un mot et non une abbreviation
                                        Set siTemp = wordApp.SynonymInfo(Word:=Sigle, LanguageID:=wdEnglishUK)
                                        If siTemp.found = True Then
'                                        Application.Activate
'                                        If MsgBox("Is " & Sigle & " an abbreviation?", vbYesNo, "Abbreviation Check") = vbNo Then
'                                            rng.Collapse Word.WdCollapseDirection.wdCollapseEnd
                                                GoTo nextabb
'                                        End If
                                        End If
                                        WordDoc.Comments.Add rng, "[Mandatory]" & Chr(10) & _
                                        "All abbreviations in the Document shall be referred elsewhere in the Glossary."
'Set des marques pages pour chaque abb trouvée pour l'utilisation des liens Word/Excel:
                                        bkMark = "Abb" & iAbbFd
                                        WordDoc.Bookmarks.Add Name:=bkMark, Range:=rng
                                        If ListAbbFd = "" Then ListAbbFd = Sigle Else ListAbbFd = ListAbbFd & ";" & Sigle
                                        iAbbFd = iAbbFd + 1
                                    End If
                                End If
                            End If
                            rng.Collapse Word.WdCollapseDirection.wdCollapseEnd
                        End If
nextabb:
                    End If
               Loop
            End With
            Set rng = WordDoc.Content
        Next
        
        'Paramètre les stats de la page d'acceuil:
        iTotFullAbbNotSet = iTotAbbGloss - iAbbNotInGloss
        If iAbbNotInGloss <> 0 And iTotAbbGloss <> 0 _
        Then Sheets("Doc_Check").TextBox21.Value = Format(1 - (iAbbNotInGloss / TabGloss.Rows.count), "0%") _
        Else Sheets("Doc_Check").TextBox21.Value = "100%"
        
        If iFullAbbNotSet <> 0 And iTotFullAbbNotSet <> 0 _
        Then Sheets("Doc_Check").TextBox23.Value = Format(1 - (iFullAbbNotSet / iTotFullAbbNotSet), "0%") _
        Else Sheets("Doc_Check").TextBox23.Value = "100%"

        ' Créé ou implémente le fichier de sortie contenant les résultats - Ancienne méthode d'appel
        ' insertResult = insert_Result(WordDoc.Name, "Abbreviation", Format(1 - (iAbbNotInGloss / TabGloss.Rows.count), "0%"), Format(1 - (iFullAbbNotSet / iTotFullAbbNotSet), "0%"), dateLancement, resultsBookName, resultsBookPath)
        
        'Liste toutes les abbréviations trouvées qui ne sont pas dans le glossaire:
        TbAbbFd = Split(ListAbbFd, ";")
        
        Set Rnge = Range("Abb[#All]").Resize(iAbbFd + 1, 1)
        Sheets("Doc_Check").ListObjects("Abb").Resize Rnge
        
        For y = 0 To UBound(TbAbbFd)
            Sheets("Doc_Check").ListObjects("Abb").DataBodyRange(1 + y, 1).Formula = "=HYPERLINK(" & """" & "[" & strpath & _
            "]" & "Abb" & y & """" & "," & """" & TbAbbFd(y) & """" & ")"
        Next y
        
        'Sheets("Doc_Check").ListObjects("Abb").HeaderRowRange.Value = "RQC2.22: Are all abbreviations mentionned in the glossary?" _
        & Chr(10) & "List of " & iAbbFd & " Abbreviations NOT found in the Glossary"
        
        Sheets("Doc_Check").ListObjects("Abb").ShowTotals = False
        Sheets("Doc_Check").Range("D2").Select
        Application.ScreenUpdating = True
        Exit Sub
        
errorCheckAbb:
    'retMessage = MsgBox("Impossible to check the 'Abbreviation'. Check the template of your input files.", vbExclamation, "Abbreviation Error")
    Sheets("Doc_Check").Range("E14").Value = "Invalid data or template, " & vbCrLf & "please verify your input file."
    Sheets("Doc_Check").Range("E14").Font.ColorIndex = 3
    Sheets("Doc_Check").Range("E14").Font.Bold = True
    Sheets("Doc_Check").Range("E14").Font.Size = 11
    Sheets("Doc_Check").Range("E14").Interior.ColorIndex = 6
    Resume Next
    
End Sub
Sub VerticalLines()

On Error GoTo EndVerticalLines

    Reset_Click 'ALE: demande faite par QMI en Oct 2017, vider les résultats précédents
    Start = Timer    ' Set start time.
    Do While Timer < Start + 1
        DoEvents    ' Yield to other processes.
    Loop
'Rappel: cette macro a pour but de recuperer deux version d'un même doc et de les comparer.
' Pour réaliser cela nous utiliserons l'option "comparer" de word qui permet
' de signaler avec une barre verticale tout delta entre les 2 versions

Dim DocV1, DocV2, DocFinal As Word.Document
'DocV1 = 1er document selectionné qui doit correspondre à la version 1 du doc
'DocV2 = 2ieme doc selectionné
'DocFinal = Document de sortie après comparaison:
'ie 3ieme version du doc dans lequel sera marqué toutes les différence entre les 2iere versions
'
Dim sPathV1, sPathV2 As String
'sPathV1 = lien vers version 1
'sPathV2 = lien vers version 2
Dim wordApp As Word.Application

    Set wordApp = CreateObject("Word.Application")
    
    wordApp.ChangeFileOpenDirectory ("C:\")
    wordApp.FileDialog(msoFileDialogOpen).Title = "Select the Document 2 versions"
    wordApp.FileDialog(msoFileDialogOpen).Filters.Clear
    wordApp.FileDialog(msoFileDialogOpen).Filters.Add "Word Documents Only", "*.docx ; *.doc; *.docm"
    wordApp.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
'On authorise ici la selection multiple pour que l'utilisateur puisse selectionner les deux version en même temps dans l'explorer windows (en utilisant "Ctrl")
    
' cette condition: pour gerer le cas où l'on ferme la fenetrer d'explorer windows:
    If wordApp.FileDialog(msoFileDialogFilePicker).Show <> 0 Then
        sPathV1 = wordApp.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
        
' Condition suivante: si 2 fichiers sont selectionnés en même temps alors un on récupère les deux liens et l'on ferme l'explorer
' Si 1 seul document est selectionné, alors on ouvre une deuxième fenêtre afin de laisser l'utilisateur selectionner le deuxième doc
        If wordApp.FileDialog(msoFileDialogFilePicker).SelectedItems.count = 2 Then
            sPathV2 = wordApp.FileDialog(msoFileDialogFilePicker).SelectedItems(2)
        Else
            wordApp.ChangeFileOpenDirectory ("C:\")
            wordApp.FileDialog(msoFileDialogOpen).Title = "Select the Document 2nd version"
            wordApp.FileDialog(msoFileDialogOpen).Filters.Add "Word Documents Only", "*.docx ; *.doc; *.docm"
            wordApp.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
            If wordApp.FileDialog(msoFileDialogOpen).Show <> 0 Then
                sPathV2 = wordApp.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
            Else
                Application.Visible = True
                Exit Sub
            End If
        End If
    Else
        Application.Visible = True
        Exit Sub
    End If
    
    wordApp.Visible = True
    wordApp.Activate
    
' on ouvre les deux versions grâce aux liens récupérés plus haut:
    Set DocV1 = wordApp.Documents.Open(sPathV1)
    Set DocV2 = wordApp.Documents.Open(sPathV2)



' Utilisation de l'option "compare" dans word:
' paramétrage:
' V1 = document d'otigine
' V2 = Document corrigé
' Resultats affichés dans un nouveau document
' Granularity = niveau détail de la comparaison = au mot près
    wordApp.CompareDocuments originaldocument:=DocV1, reviseddocument:=DocV2, Destination:=wdCompareDestinationNew, _
    granularity:=wdGranularityWordLevel, compareformatting:=False, comparecasechanges:=False, comparewhitespace:=False, _
    comparetables:=True, compareheaders:=False, comparefootnotes:=False, comparetextboxes:=True, comparefields:=False, _
    comparecomments:=False, comparemoves:=False, ignoreallcomparisonwarnings:=False

    DocV1.Close
    DocV2.Close
    
' On change ici le paramétrage du "track changes" dans word:
' plus simplement on modifie la manière d'afficher les changements signalés
' La demande était de n'afficher que les barre vertical noirs pour signaler les modifications
    With wordApp.Options
        .InsertedTextMark = wdInsertedTextMarkColorOnly
        .InsertedTextColor = wdBlack
        .DeletedTextMark = wdDeletedTextMarkNone
        .DeletedTextColor = wdByAuthor
'        .RevisedPropertiesMark = wdRevisedPropertiesMarkColorOnly
'        .RevisedPropertiesColor = wdBrightGreen
        .RevisedLinesMark = wdRevisedLinesMarkOutsideBorder
        .RevisedLinesColor = wdAuto
        .CommentsColor = wdByAuthor
        .RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationPreserve
    End With
    'retMessage = MsgBox("'Vertical Lines' is COMPLETED")
    Application.StatusBar = "Vertical lines completed ! " & Now
    Range("E13").Value = "Vertical Lines is COMPLETED"
    Exit Sub
EndVerticalLines:
    retMessage = MsgBox("Impossible to perform the 'Vertical Lines'. Check the template of your input files.", vbExclamation, "Vertical Lines Error")
End Sub
Sub GetMydocInfo2()
Dim IE As InternetExplorer
Dim IEDoc As HTMLDocument
Dim Refbox, Search As Object
Dim winShell As New ShellWindows
Dim Reference As String
Dim Link As String
Dim Title, DocLinked, Process As String
Dim rng As Range
Dim ref As Range
Dim tabAllRefInDoc() As String
Dim taille As Long

'Deletes single cells that are blank located inside a designated range
On Error GoTo NoBlanksFound
'Range("List_Ref").Select
Set rng = Range("List_Ref").SpecialCells(xlCellTypeBlanks)
'Store blank cells inside a variable
On Error GoTo 0

'rng.Rows.Delete Shift:=xlShiftUp 'ALE 02102017: cas où il n'y a qu'une seule Doc Reference

'ERROR HANLDER
NoBlanksFound:

'Récupère l'item liste contenant les références non trouvées dans la table
listDoc = Sheets("Doc_Check").ListObjects("DocRef").DataBodyRange
tabAllRefInDoc = Split(AllRef, ";")
taille = 0

'Compte le nombre d'item dans la liste des ref non trouvées dans la table
For Each Item In listDoc
    If Item <> Empty Then
        taille = taille + 1
    End If
Next

TestValue = Sheets("Doc_Check").ListObjects("DocRef").DataBodyRange(1, 1).Value
If TestValue = valErreurNA Then
    taille = taille - 1
End If

If UBound(tabAllRefInDoc) <> 0 Then
    Sheets("Doc_Check").ListObjects("DocRef").DataBodyRange(taille + 1, 1).Value = "(Synchronize) List of ref found in table :"
    taille = taille + 1
End If

'Rajoute les références présentes dans la table pour le synchronize => Ces dernières ne sont pas des liens
For y = 0 To UBound(tabAllRefInDoc)
    'If (InStr(1, NotFoundRef, tabAllRefInDoc(y)) = 0) Then
        Sheets("Doc_Check").ListObjects("DocRef").DataBodyRange(taille + 1, 1).Value = tabAllRefInDoc(y)
        taille = taille + 1
    'End If
Next

On Error Resume Next
    For Each ref In Range("List_Ref")
        If ref.Value <> "(Synchronize) List of ref found in table :" Then
            Range("B1").ClearContents
            Set IE = New InternetExplorerMedium
            Reference = ref.Value
            If ref.Value = valErreurNA Or ref.Value = messageTableRefNotFound Then 'ALE 02102017: No Reference to be checked and No Synchronize to be runned
                GoTo RefNA
            End If
    ' Ouvre la page de recherche documentaire dans ECM(MyDoc: 02 - Procedural Documentation Advanced Search)
    ' Changer directement ici lien si besoin
            IE.navigate "http://ecm.eu.airbus.corp:1080/WorkplaceDMS/WcmObjectBookmark.jsp?vsId=%7B2A9EFD86-D036-469D-BB8D-E4055EA88500%7D&requestedWindowId=_1.T15adbdda105&objectType=searchtemplate&id={C8C40F88-B40F-4C64-AB04-CD3E860F32BD}&objectStoreName=Airbus"
    
    'cette ligne permet d'assurer la récuperation de la page Internet 'Cf. Tuto: "http://qwazerty.developpez.com/tutoriels/vba/ie-et-vba-excel/"
            'Set IE = winShell(winShell.count - 1)
    'WaitIE IE -> appel de la function qui permet d'attendre le chargement complet de la page
            WaitIE IE
            Set IEDoc = IE.Document
    
            'Recupere la textbox des references de la page IE:
            WaitIE IE
            Set Refbox = IEDoc.getElementsByName("prop_typestring_121_editable_document_eq").Item
    'prop_typestring_121_editable_document_eq: par lecture du code HTML de la page internet
    
            WaitIE IE
            Refbox.innerText = Reference
            WaitIE IE
            'Recupere Bouton search de la page IE pour lancer une recherche sur le document
            Set Search = IEDoc.getElementsByName("Search").Item
            WaitIE IE
            Search.Click
            WaitIE IE
            
            If objectHandler("itemRow", IEDoc) = True Then
                ' Si aucun l'objet "itemRow" n'est trouvé c'est qu'aucun résultat n'a été renvoyer: on peut considérer le document comme n'étant pas à jour
                'Sinon on va ci-dessous recuperer les résultat de la recherche
                WaitIE IE
                'recupere les lignes trouvées suite à la recherche sur le document
                ' et "MyDocInfo" Renvoie une chaine de caractère qui comprends l'ensemble des informations sur doc (Status/Iss/Process/...)
                MyDocInfo = IEDoc.getElementById("itemRow").innerText
                
                WaitIE IE
                '1ere condition: Vérifier que le document est publié "Released":
    ' Si "Released" est dans chaine de caractère alors le doc est tagué comme publié
                If InStr(1, MyDocInfo, "Released") = 0 Then
                                ref.Offset(0, 1).Value = "Not Released"
                                ref.Offset(0, 1).Font.Color = vbRed
                Else
    ' On peut aussi récuperer les informations sur le doc en recuperant la valeur des cellules qui composent cette ligne
    ' Exemple: IEDoc.getElementById("itemRow").Cells(10).innerText: on recupere ici le texte de la 10ieme cellule (correspondant la case Process)
                                
                                If IEDoc.getElementById("itemRow").Cells(10).innerText = " " Then
                                    Process = Left(IEDoc.getElementById("itemRow").Cells(11).innerText, 4)
                                Else
                                    Process = Left(IEDoc.getElementById("itemRow").Cells(10).innerText, 5)
                                End If
                                
                                If Process <> "" Then
                                    ref.Offset(0, 1).Value = "Released (" & Process & ")"
                                Else
                                    ref.Offset(0, 1).Value = "Released"
                                End If
                                'recupere le Titre du document:
                                Title = IEDoc.getElementById("itemRow").Cells(3).innerText
                                WaitIE IE
                                
    ' Dans le cas ou le document est lié ("Linked to" dans le titre) à un autre document
    ' on récupère aussi celui-ci et l'on effectue la même recherche que pour le 1er:
                                If InStr(1, Title, "linked to") <> 0 Then
                                    Link = Mid(Title, InStr(1, Title, "linked to"), Len(Title) - 1 - InStr(1, Title, "linked to"))
                                    DocLinked = Mid(Title, InStr(1, Title, "linked to") + 9, Len(Title) - 1 - InStr(1, Title, "linked to") - 9)
                                    DocLinked = Trim(DocLinked)
                                    Set Refbox = IEDoc.getElementsByName("prop_typestring_121_editable_document_eq").Item
                                    WaitIE IE
                                    Refbox.Value = DocLinked
                                    WaitIE IE
                                    Set Search = IEDoc.getElementsByName("Search").Item
                                    WaitIE IE
                                    Search.Click
                                    WaitIE IE
                                    If objectHandler("itemRow", IEDoc) = True Then
                                        WaitIE IE
                                        MyDocInfo = IEDoc.getElementById("itemRow").innerText
                                        WaitIE IE
                                        If InStr(1, MyDocInfo, "Released") = 0 Then
                                            ref.Offset(0, 1).Value = "Released " & "[ " & Link & "(Not Released)]"
                                        Else
                                            ref.Offset(0, 1).Value = "Released " & "[ " & Link & " (Released)]"
                                        End If
                                    Else
                                        ref.Offset(0, 1).Value = "Released " & "[ " & Link & "(Not Released) ]"
                                    End If
                                End If
                                ref.Offset(0, 1).Font.Color = vbBlack
                End If
            Else
    ' si la recherche n'a rien donné alors on considère que le document n'est pas publié dans la base MyDoc
                    ref.Offset(0, 1).Value = "Not Released"
                    ref.Offset(0, 1).Font.Color = vbRed
            End If
            ref.Offset(0, 1).Font.Underline = False
            
            IE.Quit
            Process = ""
            Link = ""
            Title = ""
            Set IE = Nothing
            Set IEDoc = Nothing
            Set Refbox = Nothing
            Set Search = Nothing
        End If
    Next ref
    retMessage = MsgBox("SYNCHRONIZE is COMPLETED")
Exit Sub
RefNA:
    ref.Offset(0, 1).Value = valErreurNA 'ALE 02102017: No Reference to be checked
    retMessage = MsgBox("No reference to be synchronized") 'ALE 02102017: No Reference to be checked
End Sub
Sub LastTemplate(Reference As String, Issue As String)

Dim IE As New InternetExplorer
Dim IEDoc As HTMLDocument
Dim winShell As New ShellWindows
Dim colTR As Object
Dim tr As Object

    Set IE = New InternetExplorer
    IE.navigate "http://ecm.eu.airbus.corp:1080/WorkplaceDMS/getContent?vsId=%7B2017E316-920E-46FC-A0A0-B60ACD025D9F%7D&objectStoreName=Airbus&objectType=storedsearch"
    Set IE = winShell(winShell.count - 1)
    WaitIE IE
    Set IEDoc = IE.Document
    Set colTR = IEDoc.getElementsByTagName("TR")
    
    For Each tr In colTR 'Boucle les balises <tr> présente dans la page
        If tr.ID = "itemRow" Then 'les lignes qui nous intéresse sont balisées par des <tr> avec l'ID "ItemRow"
            'il suffit ensuite de lire chaque cellule de la ligne selectionnée
            If Reference = Trim(tr.Cells(4).innerText) Then
                If Issue = Trim(tr.Cells(5).innerText) Then
                    Range("E13").Value = Range("E13").Value & Chr(10) & "Template Version (" & Trim(tr.Cells(5).innerText) & ")"
                Else
                    Range("E13").Value = Range("E13").Value & Chr(10) & "Wrong Template Version, please use V." & Trim(tr.Cells(5).innerText)
                End If
                Exit Sub
            End If
    '            Next c
        End If
    Next tr
    IE.Quit
End Sub

Public Sub Create_Result(file As String, dateNow As String, resultsBookName As String, resultsBookPath As String)
    
    ThisWorkbook.Activate
    Application.ScreenUpdating = False
    
    'Vérifie la présence du fichier, si il existe pas, on le créé en important la feuille template et en rempissant la date et le fichier traité
    'If Dir(resultsBookPath) = "" Then
    If resultsBookPath = "" Then
        Set resultsBook = Workbooks.Add
        resultsBook.SaveAs filename:=resultsBookName
        resultsBookPath = Workbooks(resultsBookName).FullName
        
        ThisWorkbook.Worksheets("Results_template").Visible = True
        'ThisWorkbook.Worksheets("Results_template").Copy After:=Workbooks(resultsBookName).Worksheets("Feuil1")
        ThisWorkbook.Worksheets("Results_template").Copy After:=Workbooks(resultsBookName).Worksheets("Sheet1")
        ThisWorkbook.Worksheets("Results_template").Visible = False
        Application.DisplayAlerts = False
        Workbooks(resultsBookName).Worksheets("Sheet1").Delete
        Workbooks(resultsBookName).Worksheets("Sheet2").Delete
        Workbooks(resultsBookName).Worksheets("Sheet3").Delete
        Application.DisplayAlerts = True
        Workbooks(resultsBookName).Worksheets("Results_template").Name = "Results"
        
        Workbooks(resultsBookName).Activate
        Worksheets("Results").Select
        
        Range("B3").Value = Replace(Replace(Replace(dateNow, ".", "/", 1, 2), "_", " "), ".", ":")
        Range("A3").Value = file
        
        'IsFileOpen ne fonctionne pas dans l'environnement Airbus, alors au moment de la création on ferme pour rouvrir
        Workbooks(resultsBookName).Save
        Workbooks(resultsBookName).Close
    End If
    
    'Récupère les résultats des derniers checks effectués
    abbCheck1 = ThisWorkbook.Worksheets("Doc_Check").TextBox21.Value
    abbCheck2 = ThisWorkbook.Worksheets("Doc_Check").TextBox23.Value
    docCheck1 = ThisWorkbook.Worksheets("Doc_Check").RefTest1.Value
    docCheck2 = ThisWorkbook.Worksheets("Doc_Check").RefTest2.Value
    
    If (Not ThisWorkbook.Worksheets("Doc_Check").ListObjects("Abb").DataBodyRange Is Nothing) Then
        abbCheckList = ThisWorkbook.Worksheets("Doc_Check").ListObjects("Abb").DataBodyRange
    End If
    If (Not ThisWorkbook.Worksheets("Doc_Check").ListObjects("DocRef").DataBodyRange Is Nothing) Then
        docCheckList = ThisWorkbook.Worksheets("Doc_Check").ListObjects("DocRef").DataBodyRange
    End If
    If (Not ThisWorkbook.Worksheets("Doc_Check").ListObjects("ForWrd").DataBodyRange Is Nothing) Then
        wordsCheckList = ThisWorkbook.Worksheets("Doc_Check").ListObjects("ForWrd").DataBodyRange
    End If
    
    'If Not IsFileOpen(resultsBookPath) Then
        Workbooks.Open filename:=resultsBookPath
    'End If
    Workbooks(resultsBookName).Activate
    Worksheets("Results").Select
    
    'Cible les bonnes cellules en fonction du traitement appelant cette fonction (Abbreviation ou Doc.Ref)
    Range("D3").Value = abbCheck1
    Range("E3").Value = abbCheck2

    Range("G3").Value = docCheck1
    Range("H3").Value = docCheck2
    
    On Error Resume Next
        For Each Item In abbCheckList
            If Item <> Empty And Item <> "" Then
                Range("F3").Value = Range("F3").Value & Item & "; "
            End If
        Next
'test1:
        
    'On Error GoTo test2
        For Each Item In docCheckList
            If Item <> Empty And Item <> "" Then
                Range("I3").Value = Range("I3").Value & Item & "; "
            End If
        Next
'test2:
        
    'On Error GoTo test3
        If (wordsCheckList <> Empty) Then
            For Each Item In wordsCheckList
                If (InStr(1, Item, "cf.") <> 0) Then
                    Range("J3").Value = Split(Item, ":")(1)
                ElseIf (InStr(1, Item, "Method") <> 0) Then
                    Range("K3").Value = Split(Item, ":")(1)
                ElseIf (InStr(1, Item, "Policy") <> 0) Then
                    Range("L3").Value = Split(Item, ":")(1)
                ElseIf (InStr(1, Item, "Process") <> 0) Then
                    Range("M3").Value = Split(Item, ":")(1)
                ElseIf (InStr(1, Item, "Requirement") <> 0) Then
                    Range("N3").Value = Split(Item, ":")(1)
                ElseIf (InStr(1, Item, "Shall") <> 0) Then
                    Range("O3").Value = Split(Item, ":")(1)
                ElseIf (InStr(1, Item, "Should") <> 0) Then
                    Range("P3").Value = Split(Item, ":")(1)
                ElseIf (InStr(1, Item, "Will") <> 0) Then
                    Range("Q3").Value = Split(Item, ":")(1)
                End If
            Next
        End If
'test3:
    
    Workbooks(resultsBookName).Save
    Workbooks(resultsBookName).Close
    
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets("Doc_Check").Select
    
End Sub

Function IsFileOpen(filename As String)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select

End Function



Attribute VB_Name = "Module6"
Sub Delete_Tab()
Attribute Delete_Tab.VB_Description = "Macro enregistrée le 20/01/2010 par to29305"
Attribute Delete_Tab.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_Tab Macro
' Macro enregistrée le 20/01/2010 par to29305
'

'
    Sheets("TSM Source").Select
    Cells.Select
    Selection.ClearContents
    Sheets("Dico Source").Select
    Cells.Select
    Selection.ClearContents
    Sheets("Macro").Select

End Sub
Sub Create_intermediaire_File()
Attribute Create_intermediaire_File.VB_Description = "Macro enregistrée le 20/01/2010 par to29305"
Attribute Create_intermediaire_File.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_intermediaire_File Macro
' Macro enregistrée le 20/01/2010 par to29305

MsgBox ("Delete folders: ATA, TSM intermediaire and Warning Index")
'
MsgBox ("Don't forget to put in folder SOURCE the new dicoMonAct.csv & TSM_FWS_SCADE.csv") '

    
' copy SM_FWS_SCADE.CSV to SM_FWS_SCADE.txt file
    currentPath = ActiveWorkbook.path
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.MoveFile currentPath & "\Source\TSM_FWS_SCADE.csv", currentPath & "\Source\TSM_FWS_SCADE.txt"
    Workbooks.OpenText Filename:=currentPath & "\Source\TSM_FWS_SCADE.txt", _
    DataType:=xlDelimited, Semicolon:=True
    ActiveWorkbook.SaveAs Filename:= _
        currentPath & "\TSM_FWS_SCADE.xls" _
        , FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWindow.SmallScroll Down:=6
    ActiveWindow.Close
    
' copy SM_FWS_SCADE.CSV to SM_FWS_SCADE.txt file
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.MoveFile currentPath & "\Source\dicoMonAct.csv", currentPath & "\Source\dicoMonAct.txt"
    Workbooks.OpenText Filename:=currentPath & "\Source\dicoMonAct.txt", _
    DataType:=xlDelimited, Semicolon:=True

    ActiveWorkbook.SaveAs Filename:= _
        currentPath & "\dicoMonAct.xls" _
        , FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWindow.SmallScroll Down:=6
    ActiveWindow.Close
 
currentPath = ActiveWorkbook.path

    Workbooks.Open Filename:= _
     currentPath & "\TSM_FWS_SCADE.xls"
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("Macro Warning Index.xls").Activate
    Sheets("TSM Source").Select
    'ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
  
    Workbooks.Open Filename:= _
    currentPath & "\dicoMonAct.xls"
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("Macro Warning Index.xls").Activate
    Sheets("Dico Source").Select
    'ActiveSheet.Paste
    Cells.Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    
    Worksheets("TSM Source").Activate
    Worksheets("TSM Source").Rows(5).Select
    'Sheets("TSM Source").Rows(5).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    ActiveWorkbook.SaveAs Filename:= _
        currentPath & "\TSM Intermediaire\TSM intermediaire.xls" _
        , FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
        Sheets("Feuil1").Name = "TSM Intermediaire"
        
       
    Workbooks("TSM Intermediaire.xls").Activate
    Worksheets("TSM Intermediaire").Cells(1, 1).Activate
    'ActiveSheet.Paste
    
    'Supprimer les lignes ayant STS_CST = "D" (colonne G)
    i = 5
    Do While Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i, 1).Value <> "" Or Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 1, 1).Value <> ""
        If Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i, 7).Value = "D" Or Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i, 6).Value = "IS_LINE" Then
            Workbooks("Macro Warning Index.xls").Activate
            Worksheets("TSM Source").Activate
            Worksheets("TSM Source").Rows(i).Delete
        Else
            i = i + 1
        End If
    Loop
    MsgBox ("Deletion of lines with STS_CST = D")
    
    'Mise en forme TSM
    i = 6
    compteur = 2
    Do While Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i, 1).Value <> "" Or Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 1, 1).Value <> ""
        If Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i, 5).Value = "STR_ALERTE" Then
            If Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 1, 5).Value = "STR_TITLE" And Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 2, 5).Value = "STR_TITLE" Then
            '   Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 1, 10).Value = Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 1, 10).Value & _
            '    "    " & Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 2, 10).Value
            Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 1, 10).Value = Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 1, 10).Value & _
                " " & Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 2, 10).Value
            '    j = 3
            End If
            Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i, 10).Value = Workbooks("Macro Warning Index.xls").Worksheets("TSM Source").Cells(i + 1, 10).Value
            
            'Copier la ligne i dans le fich
            Workbooks("Macro Warning Index.xls").Activate
            Worksheets("TSM Source").Activate
            Worksheets("TSM Source").Rows(i).Select
            Selection.Copy
            Workbooks("TSM Intermediaire.xls").Activate
            Worksheets("TSM Intermediaire").Cells(compteur, 1).Activate
            ActiveSheet.Paste
            
            ' Incrémenter une ligne dans le fichier TSM Intermediare
            compteur = compteur + 1
            
        End If
        i = i + 1
    Loop
  'Supprimer les colonnes B, D, E, G et K
    Workbooks("TSM Intermediaire.xls").Activate
    Worksheets("TSM Intermediaire").Activate
    Worksheets("TSM Intermediaire").Range("B:B,D:D,E:E,G:G,K:K").Delete
    

        ActiveWorkbook.Save
    
    MsgBox ("Creation OK and save in TSM intermédiaire")
   
    Workbooks("Macro Warning Index.xls").Activate
    Worksheets("Macro").Activate
        
    'CopierDicodansTSM.Enabled = True
        
    Workbooks("TSM_FWS_SCADE.xls").Activate
    ActiveWindow.Close
    Workbooks("dicoMonAct.xls").Activate
    ActiveWindow.Close
    
   End Sub
Sub Warningindex_File()
'
' Create_Warningindex_File Macro
' Macro enregistrée le 20/01/2010 par to29305
'

'
currentPath = ActiveWorkbook.path
    Windows("TSM intermediaire.xls").Activate
    Cells.Select
    'Range("A4072").Activate
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
       ActiveWorkbook.SaveAs Filename:= _
        currentPath & "\Warning Index\Warning Index.xls" _
        , FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    Sheets("Feuil1").Select
    Sheets("Feuil1").Name = "warning index"
    Range("A1").Select
    
'Copie la première ligne du Dico dans le Warning Index
Workbooks("TSM Intermediaire.xls").Activate
    Worksheets("TSM Intermediaire").Activate
    Worksheets("TSM Intermediaire").Range("A:A,B:B,C:C,D:D,E:E,F:F").Select
    Selection.Copy
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Cells(1, 1).Activate
    ActiveSheet.Paste
    
    Workbooks("Macro Warning Index.xls").Activate
    Worksheets("Dico Source").Activate
    Worksheets("Dico Source").Range("A2:V2").Select
    Selection.Copy
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Cells(1, 7).Activate
    ActiveSheet.Paste
    
    i = 2
    
    Do While Workbooks("Warning Index.xls").Worksheets("Warning Index").Cells(i, 1).Value <> ""
        j = 3
            Do While Workbooks("Macro Warning Index.xls").Worksheets("Dico Source").Cells(j, 1).Value <> "" And _
                    Workbooks("Macro Warning Index.xls").Worksheets("Dico Source").Cells(j + 1, 1).Value <> ""
                If Workbooks("Warning Index.xls").Worksheets("Warning Index").Cells(i, 2).Value = Workbooks("Macro Warning Index.xls").Worksheets("Dico Source").Cells(j, 2).Value Then
                    
                    'Copie la ligne j du Dico Dans Warning Index
                    Workbooks("Macro Warning Index.xls").Activate
                    Worksheets("Dico Source").Activate
                    Worksheets("Dico Source").Range("A" & j & ":" & "V" & j).Select
                
                    Selection.Copy
                
                    Workbooks("Warning Index.xls").Activate
                    Worksheets("Warning Index").Cells(i, 7).Activate
                    ActiveSheet.Paste
                
                    Exit Do
                Else
                    j = j + 1
                End If
            Loop
        i = i + 1
    Loop
    MsgBox ("Warning index file completed")
    
    Workbooks("TSM Intermediaire.xls").Close
    
    Workbooks("Macro Warning Index.xls").Activate
    Worksheets("Macro").Activate
    
    'CreerWarningIndex.Enabled = True


      End Sub
    
Sub Form()

currentPath = ActiveWorkbook.path
    ActiveFolder = ActiveWorkbook.path
    
    'Re-intituler des colonnes
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Cells(1, 1).Value = "IDENT"
    Worksheets("Warning Index").Cells(1, 6).Value = "WARNING TITLE"
    
    'Supprimer les colonnes non utiles
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Range("C:C,G:G,H:H,J:J,K:K,L:L,M:M,N:N,X:X,Y:Y,Z:Z,AA:AA,AB:AB").Delete
    
    'Mise en forme du Warning Index
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Range("A:A,B:B,C:C,D:D,E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O").Font.Size = 10
    Worksheets("Warning Index").Range("A:A,B:B,C:C,D:D,E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O").Font.Name = "Arial"
    
    'Intervertir colonne F "PRIORITY" et E "WARNING TITLE"
    'Copier colonne E vers colonne P
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Range("E:E").Select
    Selection.Copy
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Cells(1, 16).Activate
    ActiveSheet.Paste
    
    'Copier Colonne F vers Colonne E
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Range("F:F").Select
    Selection.Copy
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Cells(1, 5).Activate
    ActiveSheet.Paste
    
    'Copier Colonne P vers colonnes F
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Range("P:P").Select
    Selection.Copy
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Cells(1, 6).Activate
    ActiveSheet.Paste
    
    'Supprimer la colonne P
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Range("P:P").Delete
    
    'Changer l'intitulé PRIORITY en PRTY (colonne E)
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Cells(1, 5).Value = "PRTY"
    
    'Changer l'intitulé PRIORITY en PRTY (colonne B)
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Cells(1, 2).Value = "FAULT CODE"

    Worksheets("Warning Index").Range("A1:O1").Font.Bold = True
    Worksheets("Warning Index").Range("A1:O1").HorizontalAlignment = xlCenter
    Worksheets("Warning Index").Range("A1:O1").VerticalAlignment = xlCenter
    
    Worksheets("Warning Index").Columns("A:O").AutoFit
    Worksheets("Warning Index").Columns(1).ColumnWidth = 14
    Worksheets("Warning Index").Columns(2).ColumnWidth = 10.57
    Worksheets("Warning Index").Columns(6).ColumnWidth = 56
    Worksheets("Warning Index").Columns(9).ColumnWidth = 14
    Worksheets("Warning Index").Rows(1).RowHeight = 38.25
    Worksheets("Warning Index").Rows(1).WrapText = True
    
    'Définir le format Paysage
    Worksheets("Warning Index").PageSetup.Orientation = xlLandscape
    
    'Fixer la première ligne sur toute les pages
    Worksheets("Warning Index").PageSetup.PrintTitleRows = "$1:$1"
    'Fixer deux première colonnes INDENT et FAULT CODE toute les pages
    Worksheets("Warning Index").PageSetup.PrintTitleColumns = "$A:$B"
    'Imprimer en Z
    Worksheets("Warning Index").PageSetup.Order = xlOverThenDown
    
    'Supprimer les NO_SUBSTYPE
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Columns(4).Replace _
    What:="NO_SUBTYPE", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=False
    
    'Supprimer les NO_SOUND
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Columns(11).Replace _
    What:="NO_SOUND", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=False

    'Supprimer les NO_SYNTHVOICE
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Columns(12).Replace _
    What:="NO_SYNTHVOICE", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=False

    'Supprimer les NO_HYBRIDE
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Columns(13).Replace _
    What:="NO_HYBRIDE", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=False

    'Supprimer les NO_SYSTPAGE
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Columns(14).Replace _
    What:="NO_SYSTPAGE", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=False
    
    'Supprimer les NO_ML
    Workbooks("Warning Index.xls").Activate
    Worksheets("Warning Index").Activate
    Worksheets("Warning Index").Columns(15).Replace _
    What:="NO_ML", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByColumns, MatchCase:=False
    
    
    ActiveWorkbook.Save
    MsgBox ("The WARNING INDEX file is finished and saved on: " & ActiveFolder & "\Warning Index\")

    Workbooks("Macro Warning Index.xls").Activate
    Worksheets("Macro").Activate
    
    'IntitulerWarningIndex.Enabled = True
End Sub

   
Sub ATA()
currentPath = ActiveWorkbook.path
ActiveFolder = ActiveWorkbook.path
  
    i = 2

    With Workbooks("Warning Index.xls").Worksheets("Warning Index")
        Do While .Cells(i, 1).Value <> ""
            'Identifier le N° ATA
            ContenueCellule1 = Left(.Cells(i, 1).Value, 4)
            NumATA = Right(ContenueCellule1, 2)
        
            'ActiveFolder = ActiveWorkbook.Path
            'Créer un nouveau ATA
            With Application
                .SheetsInNewWorkbook = 1
            End With
            Workbooks.Add
            'Sheets("Feuil1").Select
            'Sheets("Feuil1").Name = "Warning Index ATA" & NumATA
        
            Filename = ActiveFolder & "\ATA\ATA " & NumATA & ".xls"
            ActiveWorkbook.SaveAs Filename:=Filename _
            , FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            
            'Copier la première ligne (nom des colonnes) dans les fichier ATA
            Workbooks("Warning Index.xls").Activate
            Worksheets("Warning Index").Activate
            Worksheets("Warning Index").Rows(1).Select
            Selection.Copy
            Workbooks("ATA " & NumATA & ".xls").Activate
            Worksheets("Feuil1").Cells(1, 1).Activate
            ActiveSheet.Paste
            
            'Copier la ligne i dans le fichier ATA
            Workbooks("Warning Index.xls").Activate
            Worksheets("Warning Index").Activate
            Worksheets("Warning Index").Rows(i).Select
            Selection.Copy
            Workbooks("ATA " & NumATA & ".xls").Activate
            Worksheets("Feuil1").Cells(2, 1).Activate
            ActiveSheet.Paste
            
            compteur = 1
            j = 3
            Do While Right(Left(.Cells(i + 1, 1).Value, 4), 2) = NumATA
                'Copier la ligne i+1 dans le fichier ATA
                Workbooks("Warning Index.xls").Activate
                Worksheets("Warning Index").Activate
                Worksheets("Warning Index").Rows(i + 1).Select
                Selection.Copy
                Workbooks("ATA " & NumATA & ".xls").Activate
                Worksheets("Feuil1").Cells(j, 1).Activate
                ActiveSheet.Paste
                j = j + 1
                compteur = compteur + 1
                i = i + 1
                
            Loop
            i = i + 1
            
            'Mise en forme
            Workbooks("ATA " & NumATA & ".xls").Activate
            Worksheets("Feuil1").Range("A1:O1").Font.Bold = True
            Worksheets("Feuil1").Range("A1:O1").HorizontalAlignment = xlCenter
            Worksheets("Feuil1").Range("A1:O1").VerticalAlignment = xlCenter

            Worksheets("Feuil1").Columns("A:O").AutoFit
            Worksheets("Feuil1").Columns(1).ColumnWidth = 12
            Worksheets("Feuil1").Columns(2).ColumnWidth = 8
            Worksheets("Feuil1").Columns(2).HorizontalAlignment = xlLeft
            Worksheets("Feuil1").Columns(3).ColumnWidth = 16
            Worksheets("Feuil1").Columns(4).ColumnWidth = 16
            Worksheets("Feuil1").Columns(5).ColumnWidth = 4.5
            Worksheets("Feuil1").Columns(6).ColumnWidth = 60
            Worksheets("Feuil1").Columns(7).ColumnWidth = 1.63
            Worksheets("Feuil1").Columns(8).ColumnWidth = 2.86
            Worksheets("Feuil1").Columns(9).ColumnWidth = 15
            Worksheets("Feuil1").Columns(10).ColumnWidth = 3.75
            Worksheets("Feuil1").Columns(11).ColumnWidth = 15
            Worksheets("Feuil1").Columns(12).ColumnWidth = 22
            Worksheets("Feuil1").Columns(13).ColumnWidth = 14
            Worksheets("Feuil1").Columns(14).ColumnWidth = 14
            Worksheets("Feuil1").Columns(15).ColumnWidth = 5
            Worksheets("Feuil1").Rows(1).RowHeight = 38.25
            Worksheets("Feuil1").Rows(1).WrapText = True
    
            'Définir le format Paysage
            Worksheets("Feuil1").PageSetup.Orientation = xlLandscape
    
            'Fixer la première ligne sur toute les pages
            Worksheets("Feuil1").PageSetup.PrintTitleRows = "$1:$1"
            'Fixer deux première colonnes INDENT et FAULT CODE toute les pages
            Worksheets("Feuil1").PageSetup.PrintTitleColumns = "$A:$B"
            'Imprimer en Z
            Worksheets("Feuil1").PageSetup.Order = xlOverThenDown
            ActiveWorkbook.Save
            Workbooks("ATA " & NumATA & ".xls").Close
            
            
            'Remonter un répertoire
            Workbooks("Macro Warning Index.xls").Activate
            Worksheets("Macro").Activate
        Loop

    End With
    

MsgBox (" All ATA XX excel file are finished and saved on:" & ActiveFolder & "\ATA\")
Workbooks("Warning Index.xls").Close
End Sub


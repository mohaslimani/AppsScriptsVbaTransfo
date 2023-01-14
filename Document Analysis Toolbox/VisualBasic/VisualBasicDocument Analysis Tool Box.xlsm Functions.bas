Attribute VB_Name = "Functions"
Function RegKeyRead(i_RegKey As String) As String
'Utilisé pour la récuperation du séparateur utilisé par le system de l'utilisateur
Dim myWS As Object
  On Error Resume Next
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'read key from registry
  RegKeyRead = myWS.RegRead(i_RegKey)
  
End Function
Public Function NetText(stTemp As String) As String
'Récupere les caractères de la chaîne sans les deux derniers
NetText = Left(stTemp, Len(stTemp) - 2)
End Function
Public Function IsAbb(ByVal Abb As String) As Boolean
' function pratique pour éviter de remonter des valeurs dans le check des abbréviations
Dim TabNonAbb As Variant
Dim Rnge As Range
Dim Cell As Range
'la liste des Non-Abbreviations est présente dans la feuille "ForbiddenWords" _ colonne J - Table "NonAbb"
    IsAbb = True
    Set Rnge = Sheets("ForbiddenWords").ListObjects("NonAbb").DataBodyRange
    For Each Cell In Rnge
        If Cell.Value = Abb Then
            IsAbb = False
            Exit For
        End If
    Next
    
'Ancienne utilisation:
'TabNonAbb = Array("OK", "NOT", "NO", "NEO", "FAL", "AIRBUS", _
'"SUMMARY", "SCOPE", "OF", "LEFT", "BLANK", "LEXINET", _
'"DD", "LR", "WB", "XW", "SA", "TABLE", "CONTENT", "PURPOSE", _
'"YES", "ROLES", "ACTOR", "BEGIN", "END", "DESIGN", "OPEN", "CLOSE", "N/A", "TITLE", "CLASS", "UK", "USA")
'
'IsAbb = True
'For y = 0 To UBound(TabNonAbb)
'    If Abb = CStr(TabNonAbb(y)) Then
'        IsAbb = False
'        Exit For
'    End If
'Next
End Function
Public Function IsRef(ByVal ref As String) As Boolean
' function pratique pour éviter de remonter des valeurs dans le check des référence

Dim TabNonRef As Variant
TabNonRef = Array("A380", "A340", "A350", "A300", "A310", "A320", "A330", "A400M", "A319", "A321", "A318", "A400")
'Variable avec les valeur à ne pas remonter lors de l'utilisation de la RegEx de recherche des référence... à completer si besoin.
IsRef = True
For y = 0 To UBound(TabNonRef)
    If ref = CStr(TabNonRef(y)) Then
        IsRef = False
        Exit For
    End If
Next
End Function
Public Function WaitIE(IE As InternetExplorer, Optional pTimeOut As Long = 0) As Boolean
'utilisée dans chaque macro avec des chargement de page internet
    Dim lTimer As Double
    lTimer = Timer
    'le timer est optionnel: son utilisation peut être pratique si l'on souhaite limiter l'attente de chargement de la page IE
    WaitingBar.Show False
    Do
       DoEvents
' 1er condition: regarde le status de la page IE
       If IE.readyState = READYSTATE_COMPLETE And Not IE.Busy Then
            Unload WaitingBar
            Exit Do
        End If
       If pTimeOut > 0 And Timer - lTimer > pTimeOut Then
           WaitIE = True
           
           Exit Do
       End If
    Loop
'    Next i
'    Call DeactivateProgressBar(ActiveWorkbook)
End Function
Public Function objectHandler(objID, IEDoc)
'utilisée dans chaque macro avec des chargement de page internet
' recuperation d'un object HTML par son ID
Dim TestObj As Object

On Error GoTo Handler:

Set TestObj = IEDoc.getElementById(objID)
'Variable qui recupere la classe avec l'ID "objID"
objectHandler = True
Exit Function

Handler:
objectHandler = False

End Function
Public Function ClassobjectHandler(objClass, IEDoc)
' recuperation d'un object HTML par son nom de Classe
'utilisée dans chaque macro avec des chargement de page internet
' fonction pour qui gère de manière booléene les erreurs liés à la recuperation d'une class/Object HTML
' True -> l'object existe et on peut l'uliser
' Fasle -> l'object en paramètre n'existe pas
Dim TestObj As Object

On Error GoTo Handler:
' Test: si TestObj renvoie une erreur alors on passe à l'obj suivant

Set TestObj = IEDoc.getElementsByClassName(objClass)
'Variable qui recupere la classe nommée avec l'element objClass
ClassobjectHandler = True
Exit Function

Handler:
ClassobjectHandler = False

End Function

Function DocType(Reference As String) As String
'NOT USE
Dim Gov As String
Dim Process As String
    Orga = ("A10")
    Gov = ("A1;M1;AP1;AM2;ABD")
    Process = ("A2;A5;M2;AP;AM;M5;M7;AP1-")
    PSD = ("UG;TUG;")
    If InStr(1, Gov, Left(Reference, 2)) = 0 Then DocType = "Governance"
    If InStr(1, Process, Left(Reference, 2)) = 0 Then DocType = "Process"
    If InStr(1, Legacy, Left(Reference, 3)) <> 0 Then DocType = "Legacy"
    If InStr(1, Orga, Left(Reference, 3)) <> 0 Then DocType = "Organization"
End Function


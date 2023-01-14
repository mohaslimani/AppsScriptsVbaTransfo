VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WaitingBar 
   Caption         =   "Please hold on"
   ClientHeight    =   720
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4872
   OleObjectBlob   =   "VisualBasicDocument Analysis Tool Box.xlsm WaitingBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WaitingBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' cette macro utilise des commandes HTML afin de re-creer une boite de dialogue dynamique
' seul chose à changer: LeTexte= texte affiché
' La couleur: code couleur du texte http://htmlcolorcodes.com/fr/
' v= la vitesse de diffelement du texte
Public LeTexte As String, LaCouleur As String
Dim v As Long
Sub ParametresHtml()
WaitingBar.WebBrowser1.navigate _
"about:<html><body BGCOLOR ='#FFFFFF' scroll='no'><font color= " & LaCouleur & _
" size='3.2' face='arial'><body topmargin='0'>" & _
"<marquee scrollamount=" & v & ">" & LeTexte & "</marquee></font></body><center></html>"
End Sub
Private Sub UserForm_Initialize()
LeTexte = "Please wait while MyDoc is loading"
v = 4: LaCouleur = "#000000": ParametresHtml
End Sub


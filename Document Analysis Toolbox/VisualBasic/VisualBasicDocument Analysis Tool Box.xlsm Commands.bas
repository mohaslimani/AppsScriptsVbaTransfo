Attribute VB_Name = "Commands"

Sub Macrofeedback()
On Error GoTo errorFeedback
    ActiveWorkbook.FollowHyperlink Address:="https://communities.intra.corp/sites/BMS/SitePages/Topic.aspx?RootFolder=%2Fsites%2FBMS%2FLists%2FCommunity%20Discussion%2FReferential%20Quality%20Check&FolderCTID=0x0120020051CED6CFD380C24694D615A6E2A96D14&SiteMapTitle=BMS%20General&SiteMapUrl=https%3A%2F%2Fcommunities%2Eintra%2Ecorp%2Fsites%2FBMS%2FSitePages%2FCategory%2Easpx%3FCategoryID%3D1%26SiteMapTitle%3DBMS%2520General"
    Exit Sub
errorFeedback:
    retMessage = MsgBox("Impossible to follow the link. Check your connection to the internet/Hub.", vbExclamation, "Feedback link")
End Sub

Sub Contact()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = "airbus.bms@airbus.com"
        .BCC = ""
        .Subject = "Question/Remark on Document Analysis Tool Box V5"
        .Body = strbody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        '.Send   'or use
        .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
Sub RoundedRectangle4_Click()
DocumentType.Show
End Sub
Sub RoundedRectangle2_Click()
ReadMe.Show
End Sub

Sub SearchIcon_Click()
Dim IE As New InternetExplorer
Dim IEDoc As HTMLDocument
Dim Refbox, Search As Object
Dim winShell As New ShellWindows
Dim Reference As String
    Range("E13").Select
    Set IE = New InternetExplorer
    Reference = ActiveSheet.MyDocSearch.Value
    'lien vers recherche MyDoc: 02 - Procedural Documentation Advanced Search
    IE.navigate "http://ecm.eu.airbus.corp:1080/WorkplaceDMS/WcmObjectBookmark.jsp?vsId=%7B2A9EFD86-D036-469D-BB8D-E4055EA88500%7D&requestedWindowId=_1.T15adbdda105&objectType=searchtemplate&id={C8C40F88-B40F-4C64-AB04-CD3E860F32BD}&objectStoreName=Airbus"
    Set IE = winShell(winShell.count - 1)
    WaitIE IE
    Set IEDoc = IE.Document
    'Recupere textbox des references de la page IE
    WaitIE IE
    Set Refbox = IEDoc.getElementsByName("prop_typestring_121_editable_document_eq").Item
'        IEDoc.getElementsByName ("prop_typestring_121_editable_document_eq")
    WaitIE IE
    Refbox.innerText = Reference
    WaitIE IE
    'Recupere Bouton search de la page IE pour lancer une recherche sur le document
    Set Search = IEDoc.getElementsByName("Search").Item
    WaitIE IE
    Search.Click
    WaitIE IE
End Sub
Sub Reset_Click()
    Dim Rnge As Range
    'Sheets("doc_check").Unprotect (pwdForTool)
    Sheets("doc_check").MyDocSearch = "Search in MyDoc"
    Sheets("Doc_Check").TextBox21.Value = ""
    Sheets("Doc_Check").TextBox23.Value = ""
    Sheets("Doc_Check").RefTest1.Value = ""
    Sheets("Doc_Check").RefTest2.Value = ""
    
    Sheets("Doc_Check").Range("E14").Value = ""
    Sheets("Doc_Check").Range("E14").Interior.Color = RGB(0, 32, 91)
    
    Sheets("Doc_Check").Range("G14").Value = ""
    Sheets("Doc_Check").Range("G14").Interior.Color = RGB(0, 32, 91)
    
    Application.StatusBar = ""
    
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
    'Sheets("doc_check").Protect (pwdForTool)
End Sub
Sub Like_us()
Dim IE As New InternetExplorer
Dim IEDoc As HTMLDocument
Dim Refbox, LikeUs As Object
Dim winShell As New ShellWindows
Dim Reference As String
    Set IE = New InternetExplorer
    IE.navigate "https://communities.intra.corp/sites/QE_Imp/Gazette/Lists/Posts/Post.aspx?List=fc94ec37%2Dc63c%2D426b%2D9d93%2D92af84329a35&ID=19&Web=8e9e8279%2D8c38%2D4a18%2D9e05%2D76af7519e010"
    Set IE = winShell(winShell.count - 1)
    WaitIE IE
    Set IEDoc = IE.Document
    WaitIE IE
    On Error GoTo Handler:
    
    If objectHandler("likesElement-19", IEDoc) = True Then
'    If ClassobjectHandler("ngActivityAction ngLikeLink ngActionbullet", IEDoc) = True Then
        Set LikeUs = IEDoc.getElementById("likesElement-19")
        LikeUs.Click
    End If
Handler:
End Sub



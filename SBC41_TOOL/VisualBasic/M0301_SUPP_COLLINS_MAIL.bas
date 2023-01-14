Attribute VB_Name = "M0301_SUPP_COLLINS_MAIL"
Sub SUPPLIER_COLLINS_EMAIL()

'requires adding the reference "Microsoft Outlook 12.0 Object Library"

Dim WSADDRESS As Worksheet
Dim WSTEMPLATES As Worksheet
Dim eMail_app As Outlook.Application
Dim eMail_item As Outlook.MailItem
Dim attached As Outlook.Attachments
Dim Message As String, messageAux As String, signature As String
Dim i As Integer, j As Integer
Dim address_list_TO As String, address_list_COPY As String
Dim aux As String
Dim valid As Boolean
Dim subjet As String
Dim filename As String
Dim folder As String
Dim notice As String
Dim PROGRAM As String, MSN As String, SITUATION As String, AIRLINE As String, user As String, tailN As String
Dim Pn() As String
Dim PNqty() As String
Dim PNnumber As Integer
Dim PNnumbertext As String
Dim AClocation As String, RTS As String
Dim PNnumbervalid As Boolean
Dim answer As Integer
Dim TRnumber As String

' parameter initialization ---------------------------------------------------------

Set WSADDRESS = Sheets("ADDRESS")
Set WSTEMPLATES = Sheets("TEMPLATES")
Set WSREF = Sheets("REF")
PNnumber = 0
'emailselection = 1
i = 0
j = 0


'DATA INPUT*******************************************************************

PROGRAM = WSTEMPLATES.Cells(6, 3).Value
If PROGRAM = Empty Then
    PROGRAM = "---"
End If

AIRLINE = WSTEMPLATES.Cells(9, 3).Value
If AIRLINE = Empty Then
    AIRLINE = "---"
End If

MSN = WSTEMPLATES.Cells(7, 3).Value
If MSN = Empty Then
    MSN = "---"
End If

tailN = WSTEMPLATES.Cells(8, 3).Value
If MSN = Empty Then
    tailN = "---"
End If


SITUATION = WSTEMPLATES.Cells(6, 5).Value
If SITUATION = Empty Then
    SITUATION = "---"
End If

AClocation = WSTEMPLATES.Cells(7, 5).Value
If AClocation = Empty Then
    AClocation = "---"
End If

RTS = WSTEMPLATES.Cells(8, 5).Value
If RTS = Empty Then
    RTS = "---"
End If

TRnumber = WSTEMPLATES.Cells(9, 5).Value
If TRnumber = Empty Then
    TRnumber = "---"
End If

'PN INPUT****************************************************************************************************

i = 15

Do

PNnumber = PNnumber + 1
i = i + 1
Loop While (WSTEMPLATES.Cells(i, 3).Value <> "")

ReDim Pn(1 To PNnumber)
ReDim PNqty(1 To PNnumber)
j = 0

For i = 1 To PNnumber
    j = i + 14
    Pn(i) = WSTEMPLATES.Cells(j, 3).Value
        If Pn(i) = Empty Then
            Pn(i) = "---"
        End If
    
    PNqty(i) = WSTEMPLATES.Cells(j, 5).Value
    
        If PNqty(i) = Empty Then
            PNqty(i) = "---"
        End If
Next
'*****************************************************************************



'******************EMAIL CREATION******************************

subjet = "Part availability request: " & SITUATION & " // " & AIRLINE & " // " & PROGRAM & " // TAIL " & tailN & " // MSN " & MSN & " // TR " & TRnumber

If SITUATION = "AOG" Then

    subjet = "AOG//AOG//AOG Part availability request: " & SITUATION & " // " & AIRLINE & " // " & PROGRAM & " // TAIL " & tailN & " // MSN " & MSN & " // TR " & TRnumber
    
End If



'signature*******************************************************
user = Application.UserName

If InStr(1, user, "IBA") > 0 Then
    signature = WSREF.Cells(15, 2).Value
End If
If InStr(1, user, "PASS") > 0 Then
    signature = WSREF.Cells(16, 2).Value
End If
If InStr(1, user, "KRAU") > 0 Then
    signature = WSREF.Cells(17, 2).Value
End If
If InStr(1, user, "LARTI") > 0 Then
    signature = WSREF.Cells(18, 2).Value
End If
If InStr(1, user, "ARTAU") > 0 Then
    signature = WSREF.Cells(19, 2).Value
End If
If InStr(1, user, "BEAN") > 0 Then
    signature = WSREF.Cells(20, 2).Value
End If
If InStr(1, user, "LAI") > 0 Then
    signature = WSREF.Cells(21, 2).Value
End If
If InStr(1, user, "LANTZ") > 0 Then
    signature = WSREF.Cells(22, 2).Value
End If
If InStr(1, user, "BLADA") > 0 Then
    signature = WSREF.Cells(23, 2).Value
End If

If InStr(1, user, "RODRIG") > 0 Then
    signature = WSREF.Cells(24, 2).Value
End If

If InStr(1, user, "LACOMB") > 0 Then
    signature = WSREF.Cells(25, 2).Value
End If

If InStr(1, user, "DURDIL") > 0 Then
    signature = WSREF.Cells(26, 2).Value
End If

If InStr(1, user, "DOUDI") > 0 Then
    signature = WSREF.Cells(27, 2).Value
End If



'outlook procedure created***************************************************************************************************

Set eMail_app = New Outlook.Application

'email addresses to ADDRESS_TO AND ADDRESS_COPY string files*****************************************************************

Call read_address

'mail content
Message = "Dear Collins AOG desk," & vbCrLf
Message = Message & vbCrLf
Message = Message & "This is an AIRTAC question."
Message = Message & vbCrLf



If PNnumber = 1 Then
Message = Message & "We have received a request for part investigation concerning the PN:" & vbCrLf
Else
Message = Message & "We have received a request for part investigation concerning the PNs:" & vbCrLf
End If

For i = 1 To PNnumber
    Message = Message & Pn(i) & "  Qty: " & PNqty(i) & vbCrLf
Next

Message = Message & vbCrLf


If PNnumber = 1 Then
Message = Message & "Could you please check if you could have the referred PN available in your stocks and where?" & vbCrLf
Else
Message = Message & "Could you please check if you could have the referred PNs available in your stocks and where?" & vbCrLf
End If


Message = Message & "If available, we will recommend to our customer to place a PO / EO directly to you." & vbCrLf
Message = Message & vbCrLf
Message = Message & "  Situation: " & SITUATION & vbCrLf
Message = Message & "  Program: " & PROGRAM & vbCrLf
Message = Message & "  Airline: " & AIRLINE & vbCrLf
Message = Message & "  TAIL Number: " & tailN & vbCrLf
Message = Message & "  MSN: " & MSN & vbCrLf
Message = Message & "  AC Location: " & AClocation & vbCrLf
Message = Message & "  RTS(Return to service): " & RTS & vbCrLf
Message = Message & vbCrLf
Message = Message & "Thank you very much in advance for your answer" & vbCrLf
Message = Message & vbCrLf
Message = Message & "Best regards / Cordialement / Saludos / Mit freundlichen Grüßen" & vbCrLf
Message = Message & vbCrLf
Message = Message & signature



'se envía el mensaje**********************************************************************************************************************

Set eMail_item = eMail_app.CreateItem(olMailItem)


eMail_item.SentOnBehalfOfName = "aogsupply@airbus.com"
eMail_item.To = ADDRESS_TO
eMail_item.CC = ADDRESS_COPY
eMail_item.Subject = subjet
eMail_item.Body = Message

notice = MsgBox("An email will be generated, please dont forget:" & vbCrLf & "- Attachements (if any)" & vbCrLf & "- In copy (if any)", 48)


eMail_item.Display


End Sub





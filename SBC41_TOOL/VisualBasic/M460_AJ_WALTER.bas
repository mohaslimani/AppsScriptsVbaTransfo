Attribute VB_Name = "M460_AJ_WALTER"
Sub mailbroker_AJWALTER()

''requires adding the reference "Microsoft Outlook 12.0 Object Library"
'
'Dim WSADDRESS As Worksheet
'Dim WSTEMPLATES As Worksheet
'Dim eMail_app As Outlook.Application
'Dim eMail_item As Outlook.MailItem
'Dim attached As Outlook.Attachments
'Dim message As String, messageAux As String, signature As String
'Dim i As Integer, j As Integer
'Dim address_list_TO As String, address_list_COPY As String
'Dim aux As String
'Dim valid As Boolean
'Dim subjet As String
'Dim filename As String
'Dim folder As String
'Dim notice As String
'Dim program As String, MSN As String, situation As String, airline As String, user As String
'Dim PN() As String
'Dim PNqty() As String
'Dim PNnumber As Integer
'Dim PNnumbertext As String
'Dim AClocation As String, RTS As String
'Dim PNnumbervalid As Boolean
'Dim answer As Integer
'Dim TRnumber As String
'
'
'Set WSADDRESS = Sheets("ADDRESS")
'Set WSTEMPLATES = Sheets("TEMPLATES")
'Set WSREF = Sheets("REF")
'PNnumber = 0
emailselection = 57
Call BROKER_EMAIL
'i = 0
'j = 0
'
'
''DATA INPUT*******************************************************************
'
'program = WSTEMPLATES.Cells(6, 3).Value
'If program = Empty Then
'    program = "---"
'End If
'
'airline = WSTEMPLATES.Cells(9, 3).Value
'If airline = Empty Then
'    airline = "---"
'End If
'
'MSN = WSTEMPLATES.Cells(7, 3).Value & " - " & WSTEMPLATES.Cells(8, 3).Value
'If MSN = Empty Then
'    MSN = "---"
'End If
'
'situation = WSTEMPLATES.Cells(6, 5).Value
'If situation = Empty Then
'    situation = "---"
'End If
'
'AClocation = WSTEMPLATES.Cells(7, 5).Value
'If AClocation = Empty Then
'    AClocation = "---"
'End If
'
'RTS = WSTEMPLATES.Cells(8, 5).Value
'If RTS = Empty Then
'    RTS = "---"
'End If
'
'TRnumber = WSTEMPLATES.Cells(9, 5).Value
'If TRnumber = Empty Then
'    TRnumber = "---"
'End If
'
'
''***************PN INPUT****************************************************
'
'i = 14
'
'Do
'
'PNnumber = PNnumber + 1
'i = i + 1
'Loop While (WSTEMPLATES.Cells(i, 3).Value <> "")
'
'ReDim PN(1 To PNnumber)
'ReDim PNqty(1 To PNnumber)
'j = 0
'
'For i = 1 To PNnumber
'    j = i + 13
'    PN(i) = WSTEMPLATES.Cells(j, 3).Value
'        If PN(i) = Empty Then
'            PN(i) = "---"
'        End If
'
'    PNqty(i) = WSTEMPLATES.Cells(j, 5).Value
'
'        If PNqty(i) = Empty Then
'            PNqty(i) = "---"
'        End If
'Next
''*****************************************************************************
'
'
''******************EMAIL CREATION******************************
'
'subjet = "Part availability request: " & situation & " // " & airline & " // " & program & " // MSN " & MSN & " // TR " & TRnumber
'
''signature*******************************************************
'user = Application.UserName
'
'If InStr(1, user, "IBA") > 0 Then
'    signature = WSREF.Cells(15, 2).Value
'End If
'If InStr(1, user, "PASS") > 0 Then
'    signature = WSREF.Cells(16, 2).Value
'End If
'If InStr(1, user, "KRAU") > 0 Then
'    signature = WSREF.Cells(17, 2).Value
'End If
'If InStr(1, user, "LARTI") > 0 Then
'    signature = WSREF.Cells(18, 2).Value
'End If
'If InStr(1, user, "ARTAU") > 0 Then
'    signature = WSREF.Cells(19, 2).Value
'End If
'If InStr(1, user, "BEAN") > 0 Then
'    signature = WSREF.Cells(20, 2).Value
'End If
'If InStr(1, user, "LAI") > 0 Then
'    signature = WSREF.Cells(21, 2).Value
'End If
'If InStr(1, user, "LANTZ") > 0 Then
'    signature = WSREF.Cells(22, 2).Value
'End If
'
'
'
''outlook procedure created***************************************************************************************************
'
'Set eMail_app = New Outlook.Application
'
''email addresses to ADDRESS_TO AND ADDRESS_COPY string files*****************************************************************
'
'Call read_address
'
''mail content********************************************************************************************************************************
'message = "Dear AOG Team: " & vbCrLf
'message = message & vbCrLf
'
'If PNnumber = 1 Then
'message = message & "We currently have the Customer mentioned below in AOG situation with the need of the following spare-part or interchangeabilities:" & vbCrLf
'Else
'message = message & "We currently have the Customer mentioned below in AOG situation with the need of the following spare-parts or interchangeabilities:" & vbCrLf
'End If
'
'
'For i = 1 To PNnumber
'    message = message & PN(i) & "  Qty: " & PNqty(i) & vbCrLf
'Next
'message = message & vbCrLf
'message = message & "We have seen on PartsBase/ILS that you may have several parts available." & vbCrLf
'
'If PNnumber = 1 Then
'message = message & "Could you please check if you could have the referred PN available in your stocks? If yes, would you please precise location? " & vbCrLf
'Else
'message = message & "Could you please check if you could have the referred PNs available in your stocks? If yes, would you please precise location? " & vbCrLf
'End If
'
'message = message & "In case of part availability, would you mind providing us with an ARC copy to confirm direct ownership?" & vbCrLf
'message = message & "No quotation needed, if you own physically the part, we'll refer the Customer directly to you for quotation and PO placement." & vbCrLf
'message = message & vbCrLf
'message = message & "  Situation: " & situation & vbCrLf
'message = message & "  Program: " & program & vbCrLf
'message = message & "  Airline: " & airline & vbCrLf
'message = message & "  MSN: " & MSN & vbCrLf
'message = message & "  AC Location: " & AClocation & vbCrLf
'message = message & "  RTS: " & RTS & vbCrLf
'message = message & vbCrLf
'message = message & "A prompt answer would be greatly appreciated." & vbCrLf
'message = message & vbCrLf
'message = message & "Best regards / Cordialement / Saludos / Mit freundlichen Grüßen" & vbCrLf
'message = message & vbCrLf
'message = message & signature
'
'
'
'
''se envía el mensaje*****************************************************************************************************************************
'
'Set eMail_item = eMail_app.CreateItem(olMailItem)
'
'
'eMail_item.SentOnBehalfOfName = "aogsupply@airbus.com"
'eMail_item.To = ADDRESS_TO
'eMail_item.CC = ADDRESS_COPY
'eMail_item.Subject = subjet
'eMail_item.Body = message
'
'notice = MsgBox("An email will be generated, please dont forget:" & vbCrLf & "- Attachements (if any)" & vbCrLf & "- In copy (if any)", 48)
'
'
'eMail_item.Display


End Sub





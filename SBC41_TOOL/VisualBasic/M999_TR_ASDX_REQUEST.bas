Attribute VB_Name = "M999_TR_ASDX_REQUEST"
Sub TR_ASDX_MAIL()


Dim WSADDRESS As Worksheet
Dim WSTEMPLATES As Worksheet
Dim txtfile, txtroute, Message As String
Dim obj As FileSystemObject
Dim tx As Scripting.TextStream
Dim Ht As Worksheet
Dim textfile As Integer
Dim messageAux As String, signature As String
Dim i As Integer, j As Integer
Dim address_list_TO As String, address_list_COPY As String
Dim aux As String
Dim valid As Boolean
Dim subjet As String
Dim filename As String
Dim folder As String
Dim notice As String
Dim PROGRAM As String, MSN As String, SITUATION As String, AIRLINE As String, user As String
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
i = 0
j = 0

txtfile = "ASDX_TR_TEXT"
'P:\\sfs.corp\Projects\CUSTOMERSERVICE\AIRTAC\Supply Engineering\010 Tools
txtroute = "\\sfs.corp\Projects\CUSTOMERSERVICE\AIRTAC\Supply Engineering\10 Tools\" & txtfile & ".txt"

'txt file creation***************************************************************
Set obj = New FileSystemObject
Set tx = obj.CreateTextFile(txtroute)
textfile = FreeFile

'DATA INPUT*******************************************************************

PROGRAM = WSTEMPLATES.Cells(6, 3).Value
If PROGRAM = Empty Then
    PROGRAM = "---"
End If

AIRLINE = WSTEMPLATES.Cells(9, 3).Value
If AIRLINE = Empty Then
    AIRLINE = "---"
End If

MSN = WSTEMPLATES.Cells(7, 3).Value & " - " & WSTEMPLATES.Cells(8, 3).Value
If MSN = Empty Then
    MSN = "---"
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

i = 14

Do

PNnumber = PNnumber + 1
i = i + 1
Loop While (WSTEMPLATES.Cells(i, 3).Value <> "")

ReDim Pn(1 To PNnumber)
ReDim PNqty(1 To PNnumber)
j = 0

For i = 1 To PNnumber
    j = i + 13
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

'MESSAGE COntent **********************************************************************
Message = "Dear colleagues: " & vbCrLf
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
Message = Message & "Could you please investigate possible Stock? " & vbCrLf
Message = Message & vbCrLf
Message = Message & "  Situation: " & SITUATION & vbCrLf
Message = Message & "  Program: " & PROGRAM & vbCrLf
Message = Message & "  Airline: " & AIRLINE & vbCrLf
Message = Message & "  MSN: " & MSN & vbCrLf
Message = Message & "  AC Location: " & AClocation & vbCrLf
Message = Message & "  RTS (Return to Service): " & RTS & vbCrLf
Message = Message & vbCrLf
Message = Message & "Thank you very much in advance for your answer" & vbCrLf
Message = Message & vbCrLf
Message = Message & "Best regards / Cordialement / Saludos / Mit freundlichen Grüßen" & vbCrLf
Message = Message & vbCrLf


'sow txt file******************************************************************

tx.Write (Message)
Shell "NotePad " & txtroute, vbMaximizedFocus

'Open ruta For Input As textfile

'Set tx = obj.OpenTextFile(ruta)





End Sub





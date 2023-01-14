Attribute VB_Name = "M9999_read_address"
Sub read_address()

Set WSADDRESS = Sheets("ADDRESS")
Set WSREF = Sheets("REF")


ADDRESS_TO = WSADDRESS.Cells(emailselection, 2)
ADDRESS_COPY = WSADDRESS.Cells(5, 2) & ";" & WSADDRESS.Cells(emailselection, 3)
   
If WSREF.Cells(3, 4) = True Then
    ADDRESS_COPY = ADDRESS_COPY & ";" & WSREF.Cells(3, 5)
End If

If WSREF.Cells(4, 4) = True Then
    ADDRESS_COPY = ADDRESS_COPY & ";" & WSREF.Cells(4, 5)
End If

If WSREF.Cells(5, 4) = True Then
    ADDRESS_COPY = ADDRESS_COPY & ";" & WSREF.Cells(5, 5)
End If

If WSREF.Cells(6, 4) = True Then
    ADDRESS_COPY = ADDRESS_COPY & ";" & WSREF.Cells(6, 5)
End If

If WSREF.Cells(7, 4) = True Then
    ADDRESS_COPY = ADDRESS_COPY & ";" & WSREF.Cells(7, 5)
End If

If emailselection = 49 Then
    ADDRESS_COPY = ADDRESS_COPY & ";" & WSADDRESS.Cells(emailselection, 3)
End If

If emailselection = 51 Then
    ADDRESS_COPY = ADDRESS_COPY & ";" & WSADDRESS.Cells(emailselection, 3)
End If
   
End Sub


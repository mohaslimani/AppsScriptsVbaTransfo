Attribute VB_Name = "Alert"
Sub Email_Alert()

'confirmation dialog box

    
    Worksheets(actionsheet).Activate

    should_continue = MsgBox("Are you sure you want to send an alert now ?", 1, "Super Confirm")
    
    If should_continue = 2 Then     'if cancelled
        MsgBox "Report Not Sent"
        Exit Sub
        
    ElseIf should_continue = 1 Then 'if confirmed
    

    Dim People As Worksheet
    Dim action As Worksheet
    Dim Menu As Worksheet
    Set action = ActiveWorkbook.Sheets(actionsheet)
    Set People = ActiveWorkbook.Sheets(peoplesheet)
    Set Menu = ActiveWorkbook.Sheets(menusheet)
    Dim OutApp As Outlook.Application
    Dim email As String
    Dim nom As String
    Dim message As String
    Dim has_action As Boolean
    
    n_pdist = 0     ' Init of Number of people in distribution
'
'   Title of mail
'
    Title = Menu.Cells(1, 1).Value
'
'
'   Build a message for each people in the list
'
    For i = 5 To 4 + People.Cells(4, 4).Value   ' Get the number of persons in people sheet
'
'   Header of the action
'
        nom = People.Cells(i, 4).Value  ' look in the people sheet the name
        email = People.Cells(i, 6).Value  ' look in the people shett the e-mail adress
        
        has_action = False
        
        message = "<b><i>You will find below your pending SCR actions on "
        message = message & Title & "."
        message = message & "<p>"
        message = message & " Please update you SCR activities </b></i>"
      
        
        For j = first_act To first_act - 1 + Menu.Cells(8, 1).Value 'look in Actions Sheet the number of actions
'
'   Looping on number of actions check for the name in Actions Sheet equal to selected name
'   if action is not Closed
'
        If ((action.Cells(j, jqr).Value = nom) And ((action.Cells(j, jscrstatus).Value) <> A_Status0)) Then
            has_action = True

 '   Color of the header depending on the action status
 '
            If (action.Cells(j, jscrstatus).Value = A_Status3) Then   ' Late-Red
                message = message & "<font color='darkred'>"
            ElseIf (action.Cells(j, jscrstatus).Value = A_Status2) Then  'Late
                message = message & "<font color='orangered'>"
'            Else
'                message = message & "<font color='goldenrod'>" 'On-Time Alert
            End If
            
            message = message & "<p>"
            message = message & "<b>ARP Id : " & "</b>" & action.Cells(j, jarp).Value
 '          message = message & "<p>"
            
            message = message & " - " & "<B>PG Level 2 : " & "</b>" & action.Cells(j, jpgl2).Value
            message = message & " - " & "<B>SCR Status : " & "</b>" & action.Cells(j, jscrstatus).Value
            message = message & " - " & "<B>Next SCR Date : " & action.Cells(j, jnextscr).Value
            
            message = message & "</b></font>"
  
             
            message = message & "<ul>"
            message = message & "<li>" & "<B>Supplier : " & "</b>" & action.Cells(j, jcie).Value
            message = message & " - " & "<B>City : " & "</b>" & action.Cells(j, jcity).Value
            message = message & " - " & "<B>Country : " & "</b>" & action.Cells(j, jcountry).Value
            
            message = message & "<li>" & "<B>Address : " & "</b>" & action.Cells(j, jstreet).Value
            message = message & " - " & "<B>Postal Code : " & "</b>" & action.Cells(j, jpcode).Value
            
        
            message = message & " </ul> "

            End If
        Next
        
     
        Dim OutMail As Outlook.MailItem
        
        If (email <> "" And has_action = True) Then
        
         
                      
          should_continue_n = MsgBox("Are you sure you want to send an alert to " & nom & " ?", 1, "Super Confirm")
    
            If should_continue_n = 2 Then     'if cancelled
                MsgBox "Report Not Sent to " & nom
            ElseIf should_continue_n = 1 Then 'if confirmed
                
                n_pdist = n_pdist + 1   ' Number of people getting a message
                Set OutApp = CreateObject("Outlook.Application")
                Set OutMail = OutApp.CreateItem(olMailItem)
                With OutMail
                    .To = email
                    .Subject = "ACTIONS on " & Title
                    .HTMLBody = message
                    .Send
                End With
                Set OutMail = Nothing
                MsgBox "Report Sent to  " & nom & " .Continue with other Action holders."
            End If
                    
        End If
        
        
    Next
    Application.ScreenUpdating = True
    Set OutApp = Nothing
            MsgBox "Report Sent to  " & n_pdist & "  person(s)"
            
        
    End If
    
    'Update of last alert date
    If n_pdist <> 0 Then
        ActiveWorkbook.Sheets(menusheet).Cells(4, 8).Value = Date
    End If
    

End Sub

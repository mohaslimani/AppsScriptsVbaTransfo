Attribute VB_Name = "People"
Public Function search_team(nom_people)

    Dim People As Worksheet
    Set People = ActiveWorkbook.Sheets(peoplesheet)

      nb_people = People.Cells(4, 4).Value

For i = 1 To nb_people
    
    If People.Cells(i + 4, 4).Value = nom_people Then
        search_team = People.Cells(i + 4, 5).Value
    End If

Next

End Function

Attribute VB_Name = "Menu"
Sub Menu_UpdateActions_Status()

    Call UpdateActions_Status
    Call Calcul_Indic
    Sheets(menusheet).Select
    Cells(1, 1).Select
    
End Sub
Sub Menu_ModActions()

    Sheets(actionsheet).Select
    Cells(1, 1).Select
    
End Sub
Sub Menu()

    Sheets(menusheet).Select
    Cells(1, 1).Select
    
End Sub
Sub Menu_People()

    Sheets(peoplesheet).Select
    Cells(1, 1).Select
    
End Sub
Sub Menu_Parameters()

    Sheets(parasheet).Select
    Cells(1, 1).Select
    
End Sub
Sub Menu_QMSAnal()

    Sheets("QMS Analysis").Select
    Cells(1, 1).Select
    
End Sub
Sub Menu_NAAAnal()

    Sheets("NAA Analysis").Select
    Cells(1, 1).Select
    
End Sub
Sub Menu_SyncBlock()

    Sheets("Synchro and Block").Select
    Cells(1, 1).Select
    
End Sub


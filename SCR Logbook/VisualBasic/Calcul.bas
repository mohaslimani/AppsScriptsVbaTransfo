Attribute VB_Name = "Calcul"

Sub Calcul_Indic()
'
' Clear All Previous Data
'
    Application.GoTo Reference:="Calc_Person_Data"
    Selection.ClearContents

    Application.GoTo Reference:="Calc_Origin_Data"
    Selection.ClearContents
'
'
nb_actions = Worksheets(menusheet).Cells(8, 1).Value
nb_pers = Worksheets(peoplesheet).Cells(4, 4).Value
nb_teams = Worksheets(peoplesheet).Cells(4, 2).Value
nb_severity = Worksheets(parasheet).Cells(2, 2).Value

For i = 1 To nb_pers
    nom = Worksheets(peoplesheet).Cells(i + 4, 4).Value
    nb_act_ongoing = 0
    nb_act_oga = 0
    nb_act_late = 0
    nb_act_latered = 0
   
       
    For j = 1 To nb_actions
        If Worksheets(actionsheet).Cells(j + first_act - 1, jqr).Value = nom Then
            If Worksheets(actionsheet).Cells(j + first_act - 1, jscrstatus).Value = A_Status0 Then
                nb_act_ongoing = nb_act_ongoing + 1
            ElseIf Worksheets(actionsheet).Cells(j + first_act - 1, jscrstatus).Value = A_Status1 Then
                nb_act_oga = nb_act_oga + 1
            ElseIf Worksheets(actionsheet).Cells(j + first_act - 1, jscrstatus).Value = A_Status2 Then
                nb_act_late = nb_act_late + 1
            ElseIf Worksheets(actionsheet).Cells(j + first_act - 1, jscrstatus).Value = A_Status3 Then
                nb_act_latered = nb_act_latered + 1
            End If
        End If
    Next


Application.GoTo Reference:="Calc_Person"
ActiveCell.Offset(i, 0).Value = nom
ActiveCell.Offset(i, 1).Value = nb_act_ongoing
ActiveCell.Offset(i, 2).Value = nb_act_oga
ActiveCell.Offset(i, 3).Value = nb_act_late
ActiveCell.Offset(i, 4).Value = nb_act_latered

Next
'
' Actions per severity type
'
For i = 1 To nb_severity
    nom = Worksheets(parasheet).Cells(i + 3, 2).Value
    nb_act_ongoing = 0
    nb_act_oga = 0
    nb_act_late = 0
    nb_act_latered = 0
   
    
    For j = 1 To nb_actions
        If Worksheets(actionsheet).Cells(j + first_act - 1, jpgl2).Value = nom Then
              If Worksheets(actionsheet).Cells(j + first_act - 1, jscrstatus).Value = A_Status0 Then
                nb_act_ongoing = nb_act_ongoing + 1
            ElseIf Worksheets(actionsheet).Cells(j + first_act - 1, jscrstatus).Value = A_Status1 Then
                nb_act_oga = nb_act_oga + 1
            ElseIf Worksheets(actionsheet).Cells(j + first_act - 1, jscrstatus).Value = A_Status2 Then
                nb_act_late = nb_act_late + 1
            ElseIf Worksheets(actionsheet).Cells(j + first_act - 1, jscrstatus).Value = A_Status3 Then
                nb_act_latered = nb_act_latered + 1
            End If
        End If
    Next

Application.GoTo Reference:="Calc_Origin"
ActiveCell.Offset(i, 0).Value = nom
ActiveCell.Offset(i, 1).Value = nb_act_ongoing
ActiveCell.Offset(i, 2).Value = nb_act_oga
ActiveCell.Offset(i, 3).Value = nb_act_late
ActiveCell.Offset(i, 4).Value = nb_act_latered


Next
'
' Fill in Log Update if necessary
'
    Sheets(logsheet).Select
    range_max = col_log_first & log_first_line & ":" & col_log_first & log_last_line
    nb_date = 0
    Dim myRange_logu As Range
    Set myRange_logu = Worksheets(logsheet).Range(range_max)
    nb_date = Application.WorksheetFunction.CountA(myRange_logu)
    nb_date_last = nb_date + 1
   
'    If Cells(nb_date_last, collogflag).Value = "" Then
        Sheets(menusheet).Select
        Range(Range_Copy_H).Select
        Selection.Copy
        Sheets(logsheet).Select
        Range_to_copy = col_log_copy & nb_date_last
        Range(Range_to_copy).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Sheets(menusheet).Select
        Application.CutCopyMode = False
        Cells(1, 1).Select
 '   End If
    Cells(1, 1).Select
'
'


End Sub

Sub calc_block()
'
'----------------------------------------------------------------
' Calculation of PO Block and Synchronization
'----------------------------------------------------------------
'
' Init of Parameters
'
Dim book0 As Workbook
Dim book0Name As String
Dim book0NamePath As String
    book0Name = SCR_Log_File
    Set book0 = Workbooks(book0Name)
'
Dim sht0Name As String
sht0Name = scrptsheet ' Raw Supplier Control Review Report

Dim sht1Name As String
sht1Name = "Synchro and Block" ' Result Sheet for Synchonisation and Block issues
'
' Creating and cleaning result sheet for Issues
'
Dim WS As Worksheet
Dim SHEX As Boolean
SHEX = False ' Flag to check if sheet already exists
'
    For Each WS In book0.Worksheets
       If WS.Name = sht1Name Then
         SHEX = True
       End If
    Next WS
    If SHEX = False Then  ' Add sheet
        book0.Sheets.Add.Name = sht1Name
    End If
    
    book0.Sheets(sht1Name).Cells.ClearContents  ' clearing sheet
'
'  Data are on Supplier Control Reviw Report sheet with Headers
'   "Synchronised (Y/N)" ' in Column AD
'   "PO block" in Column AE
'   "Blocking reason" in Column AF
'   "Blocked by" in Column AG
'   "Approval Responsible" in Column X
'
'   Calculate Last Col in Origin Data in SCR Report sheet
    LastCol = book0.Sheets(scrptsheet).UsedRange.Columns.count
'
'   Loop on Column Headers to find columns to be used
'
    For i = LastCol To 1 Step -1
        
        If InStr(1, book0.Sheets(sht0Name).Cells(1, i).Value, "Approval Responsible", vbTextCompare) > 0 Then
            RespCol = i
        End If
        If InStr(1, book0.Sheets(sht0Name).Cells(1, i).Value, "Synchronised (Y/N)", vbTextCompare) > 0 Then
            SyncCol = i
        End If
        If InStr(1, book0.Sheets(sht0Name).Cells(1, i).Value, "PO block", vbTextCompare) > 0 Then
            BlockCol = i
        End If
        If InStr(1, book0.Sheets(sht0Name).Cells(1, i).Value, "Blocking reason", vbTextCompare) > 0 Then
            BlReasonCol = i
        End If
        If InStr(1, book0.Sheets(sht0Name).Cells(1, i).Value, "Blocked by", vbTextCompare) > 0 Then
            BlPersonCol = i
        End If
          
    Next i
'
' Place Headers in Result sheet
'
' K = 11 col
' Then work with offest
'
LColCopy = 11
F_Col = "A"
L_Col = "K"

book0.Sheets(sht0Name).Range("A1:K1").Copy book0.Sheets(sht1Name).Range(F_Col & 1)
book0.Sheets(sht1Name).Cells(1, LColCopy).Offset(0, 1) = "Approval Responsible"
book0.Sheets(sht1Name).Cells(1, LColCopy).Offset(0, 2) = "Desynchronised"
book0.Sheets(sht1Name).Cells(1, LColCopy).Offset(0, 3) = "PO block"
book0.Sheets(sht1Name).Cells(1, LColCopy).Offset(0, 4) = "Blocking reason"
book0.Sheets(sht1Name).Cells(1, LColCopy).Offset(0, 5) = "Blocked by"
'
' Rows init
LSearchRow = 2 'for looping in SCR Report Sheet
LResulRow = 2
'
'   Loop on all lines of Origin sheet
'
Do While Len(book0.Sheets(sht0Name).Range(F_Col & CStr(LSearchRow)).Value) > 0
'
        SyncFlag = ""
        BlockFlag = ""
        FlagFind = False
        
        If book0.Sheets(sht0Name).Cells(LSearchRow, SyncCol).Value = "N" Then
        '
        ' Supplier is desynchronozied
                SyncFlag = "Y"
                FlagFind = True
        End If
        
        If book0.Sheets(sht0Name).Cells(LSearchRow, BlockCol).Value = "X" Then
        '
        ' Supplier PO is blocked
                BlockFlag = "Y"
                FlagFind = True
        End If
'
        If FlagFind Then
        '
            RangetoCopy = F_Col & CStr(LSearchRow) & ":" & L_Col & CStr(LSearchRow)
            RangetoPaste = F_Col & CStr(LResulRow)
            book0.Sheets(sht0Name).Range(RangetoCopy).Copy book0.Sheets(sht1Name).Range(RangetoPaste)
            '
            App_Resp = book0.Sheets(sht0Name).Cells(LSearchRow, RespCol).Value
            BockReas = book0.Sheets(sht0Name).Cells(LSearchRow, BlReasonCol).Value
            BlockPers = book0.Sheets(sht0Name).Cells(LSearchRow, BlPersonCol).Value
            
            book0.Sheets(sht1Name).Cells(LResulRow, LColCopy).Offset(0, 1).Value = App_Resp
            book0.Sheets(sht1Name).Cells(LResulRow, LColCopy).Offset(0, 2).Value = SyncFlag
            book0.Sheets(sht1Name).Cells(LResulRow, LColCopy).Offset(0, 3).Value = BlockFlag
            book0.Sheets(sht1Name).Cells(LResulRow, LColCopy).Offset(0, 4).Value = BlockReas
            book0.Sheets(sht1Name).Cells(LResulRow, LColCopy).Offset(0, 5).Value = BlockPers
        
            LResulRow = LResulRow + 1
        '
        End If
        
LSearchRow = LSearchRow + 1

Loop ' end of loop on ARP ID
'
End Sub




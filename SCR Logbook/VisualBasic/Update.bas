Attribute VB_Name = "Update"
Sub UpdateActions_Status()
'
' Retrieve Delta_amber days from Parameters sheet.
'
    Application.GoTo Reference:="Alert_LD"
    Delta_amber = ActiveCell.Value
    
    Application.GoTo Reference:="Red_Overdue"
    Delta_red = ActiveCell.Value
    Delta_red = -Delta_red
'
' Get number of logged actions in Open Actions sheet
'
    Sheets(actionsheet).Select
'
' nb_actions from number of logged actions
'
    range_max = Lcol_first & first_act & ":" & Lcol_first & last_act
    
    Dim myRange As Range
    Set myRange = Worksheets(actionsheet).Range(range_max)
    nb_actions = Application.WorksheetFunction.CountA(myRange)
    
    
    
    Sheets(actionsheet).Select

nb_actions_og = 0    ' On Time
nb_actions_oga = 0   ' On-Time Alert
nb_actions_late = 0  ' Late by 1 day
nb_actions_latered = 0 ' Late by 2 months
' nb_actions_na = 0 ' SCR date not avail

'
'   Copy of the action status
'
' First action is logged in line "first_act" line
' i is the index of lines of logged actions
'
' Action is Late if Closure Date is blank
' and if Target date is less than today

' Action is Ongoing Green if Closure Date is blank
' and if Target date is greater than today
'
' Action is Ongoing Ammber if Closure Date is blank and
' if Target Date is within the next Delta_amber days.

' Action is Closed if Closure date is filled and if Action Progress is 100%
'
colRed = 0
colAmb = 0
colYel = 0

For i = first_act To nb_actions + first_act - 1
'
'   SCR
'

        Delta = Cells(i, jnextscr).Value - Date
    
          
        If Delta < Delta_red Then
            Cells(i, jarp).Interior.ColorIndex = 3
            Cells(i, jscrstatus).Value = A_Status3    ' "Late-Red"
            Cells(i, jscrstatus).Interior.ColorIndex = 3
            Cells(i, jscrstatus).Font.ColorIndex = 1
            nb_actions_latered = nb_actions_latered + 1
            colRed = 1
        
        ElseIf Delta < 0 Then
            Cells(i, jarp).Interior.ColorIndex = 45
            Cells(i, jscrstatus).Value = A_Status2    ' "Late"
            Cells(i, jscrstatus).Interior.ColorIndex = 45
            Cells(i, jscrstatus).Font.ColorIndex = 1
            nb_actions_late = nb_actions_late + 1
            colAmb = 1
            
        ElseIf Delta <= Delta_amber Then
            Cells(i, jarp).Interior.ColorIndex = 6
            Cells(i, jscrstatus).Value = A_Status1    ' "On Time Alert"
            Cells(i, jscrstatus).Interior.ColorIndex = 6
            Cells(i, jscrstatus).Font.ColorIndex = 1
            nb_actions_oga = nb_actions_oga + 1
            colYel = 1
            
        ElseIf Delta > Delta_amber Then
            Cells(i, jarp).Interior.ColorIndex = 4
            Cells(i, jscrstatus).Value = A_Status0    '"On Time "
            Cells(i, jscrstatus).Interior.ColorIndex = 4
            Cells(i, jscrstatus).Font.ColorIndex = 1
            nb_actions_og = nb_actions_og + 1
        End If
'
Skip_SCR:
'
Next
'
'
'   Update Actions Count in Actionsheet
'

    Sheets(menusheet).Select
    Cells(3, 1).Value = nb_actions_og
    Cells(3, 2).Value = A_Status0
    Cells(3, 1).Interior.ColorIndex = 4
    Cells(3, 2).Interior.ColorIndex = 4
    
    Cells(4, 1).Value = nb_actions_oga
    Cells(4, 2).Value = A_Status1 & " - " & Range("Alert_LD").Value & " days"
    Cells(4, 1).Interior.ColorIndex = 6 * colYel + 15 * (1 - colYel)
    Cells(4, 2).Interior.ColorIndex = 6 * colYel + 15 * (1 - colYel)
              
    Cells(5, 1).Value = nb_actions_late
    Cells(5, 2).Value = A_Status2 & " - Less than " & Range("Red_Overdue").Value & " days"
    Cells(5, 1).Interior.ColorIndex = 45 * colAmb + 15 * (1 - colAmb)
    Cells(5, 2).Interior.ColorIndex = 45 * colAmb + 15 * (1 - colAmb)

    Cells(6, 1).Value = nb_actions_latered
    Cells(6, 2).Value = A_Status3 & " - more than " & Range("Red_Overdue").Value & " days"
    Cells(6, 1).Interior.ColorIndex = 3 * colRed + 15 * (1 - colRed)
    Cells(6, 2).Interior.ColorIndex = 3 * colRed + 15 * (1 - colRed)
    
    Cells(8, 1).Value = nb_actions
    Cells(8, 2).Value = "Total"
    Cells(8, 1).Interior.ColorIndex = 2
    Cells(8, 2).Interior.ColorIndex = 2
      
    Sheets(menusheet).Select
    Cells(1, 1).Select
    

End Sub

Sub UpdateNAA_Status()
'
'----------------------------------------------------------------
' Calculation of NAA Expiry Dates
'---------------------------------------------------------------
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
sht0Name = "NAA Analysis" ' NAA Analysis sheet

Dim sht1Name As String
sht1Name = "NAA Format" ' QMS Analysis sheet
F_Col = "A"

For rw = 1 To 2
book0.Sheets(sht1Name).Rows(rw & ":" & rw).Copy
book0.Sheets(sht0Name).Rows(rw & ":" & rw).PasteSpecial Paste:=xlPasteFormats
Next rw

rw = 3
Do While Len(book0.Sheets(sht0Name).Range(F_Col & rw).Value) > 0

book0.Sheets(sht1Name).Rows("2:2").Copy
book0.Sheets(sht0Name).Rows(rw & ":" & rw).PasteSpecial Paste:=xlPasteFormats
rw = rw + 1
Loop

nb_col = book0.Sheets(sht1Name).UsedRange.Columns.count

Dim icol As Integer
Dim wcol As Double

For icol = 1 To nb_col

wcol = book0.Sheets(sht1Name).Columns(icol).Width
wcol = wcol / 5.5
book0.Sheets(sht0Name).Columns(icol).ColumnWidth = wcol

Next icol

'
' Columns for Expiration interval
'
     ExpRedCol = 11         'K
     ExpAmberCol = 12      'L
     ExpYellowCol = 13       'M
'
     naaCol = 10 ' J
'
    F_Col = "A"
    LSearchRow = 2
    
    
nb_Red = 0
nb_efRed = 0
nb_Amber = 0
nb_Yellow = 0
    
colRed = 0
colefRed = 0
colAmb = 0
colYel = 0

Do While Len(book0.Sheets(sht0Name).Range(F_Col & CStr(LSearchRow)).Value) > 0
   
    If book0.Sheets(sht0Name).Cells(LSearchRow, ExpYellowCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 6
        book0.Sheets(sht0Name).Cells(LSearchRow, ExpYellowCol).Interior.ColorIndex = 6
        nb_Yellow = nb_Yellow + 1
        colYel = 1
    End If
    If book0.Sheets(sht0Name).Cells(LSearchRow, ExpAmberCol).Value <> "" Then
        If InStr(book0.Sheets(sht0Name).Cells(LSearchRow, naaCol).Value, "EASA") <> 0 Or _
           InStr(book0.Sheets(sht0Name).Cells(LSearchRow, naaCol).Value, "FAR") <> 0 Then
            book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 3
            book0.Sheets(sht0Name).Cells(LSearchRow, ExpAmberCol).Interior.ColorIndex = 3
        nb_efRed = nb_efRed + 1
        colefRed = 1
        Else
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 45
        book0.Sheets(sht0Name).Cells(LSearchRow, ExpAmberCol).Interior.ColorIndex = 45
        nb_Amber = nb_Amber + 1
        colAmb = 1
        End If
    End If
    If book0.Sheets(sht0Name).Cells(LSearchRow, ExpRedCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 3
        book0.Sheets(sht0Name).Cells(LSearchRow, ExpRedCol).Interior.ColorIndex = 3
        nb_Red = nb_Red + 1
        colRed = 1
    End If
          
LSearchRow = LSearchRow + 1

Loop
'
'
'   Update Actions Count in Menusheet
'
    book0.Sheets(menusheet).Cells(4, 3).Value = nb_Yellow
    book0.Sheets(menusheet).Cells(4, 3).Interior.ColorIndex = 6 * colYel + 15 * (1 - colYel)
    book0.Sheets(menusheet).Cells(5, 3).Value = nb_Amber
    book0.Sheets(menusheet).Cells(5, 3).Interior.ColorIndex = 45 * colAmb + 15 * (1 - colAmb)
    book0.Sheets(menusheet).Cells(6, 3).Value = nb_efRed
    book0.Sheets(menusheet).Cells(6, 3).Interior.ColorIndex = 3 * colefRed + 15 * (1 - colefRed)
    book0.Sheets(menusheet).Cells(7, 3).Value = nb_Red
    book0.Sheets(menusheet).Cells(7, 3).Interior.ColorIndex = 3 * colRed + 15 * (1 - colRed)
    
    book0.Sheets(menusheet).Cells(4, 4).Value = book0.Sheets(sht0Name).Cells(1, ExpYellowCol).Value
    book0.Sheets(menusheet).Cells(4, 4).Interior.ColorIndex = 6 * colYel + 15 * (1 - colYel)
    book0.Sheets(menusheet).Cells(5, 4).Value = book0.Sheets(sht0Name).Cells(1, ExpAmberCol).Value
    book0.Sheets(menusheet).Cells(5, 4).Interior.ColorIndex = 45 * colAmb + 15 * (1 - colAmb)
    book0.Sheets(menusheet).Cells(6, 4).Value = book0.Sheets(sht0Name).Cells(1, ExpAmberCol).Value & " EASA or FAA"
    book0.Sheets(menusheet).Cells(6, 4).Interior.ColorIndex = 3 * colefRed + 15 * (1 - colefRed)
    book0.Sheets(menusheet).Cells(7, 4).Value = book0.Sheets(sht0Name).Cells(1, ExpRedCol).Value
    book0.Sheets(menusheet).Cells(7, 4).Interior.ColorIndex = 3 * colRed + 15 * (1 - colRed)
'
' Calculation of Global Indicator
'
    
    naaKPI = 2  ' Green
    nb_KPIred = nb_efRed + nb_Red
       
    If nb_KPIred <> 0 Then
    naaKPI = 0          ' Red
    ElseIf nb_Amber <> 0 Then
    naaKPI = 1          'Amber
    End If
 '
    book0.Sheets(menusheet).Cells(11, 3).Value = naaKPI
    book0.Sheets(menusheet).Select
    

End Sub

Sub UpdateQMS_Status()
'
'----------------------------------------------------------------
' Calculation of QMS Expiry Dates
'---------------------------------------------------------------
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
sht0Name = "QMS Analysis" ' QMS Analysis sheet

Dim sht1Name As String
sht1Name = "QMS Format" ' QMS Analysis sheet Format
F_Col = "A"

For rw = 1 To 2
book0.Sheets(sht1Name).Rows(rw & ":" & rw).Copy
book0.Sheets(sht0Name).Rows(rw & ":" & rw).PasteSpecial Paste:=xlPasteFormats
Next rw

rw = 3
Do While Len(book0.Sheets(sht0Name).Range(F_Col & rw).Value) > 0

book0.Sheets(sht1Name).Rows("2:2").Copy
book0.Sheets(sht0Name).Rows(rw & ":" & rw).PasteSpecial Paste:=xlPasteFormats
rw = rw + 1
Loop

nb_col = book0.Sheets(sht1Name).UsedRange.Columns.count

Dim icol As Integer
Dim wcol As Double

For icol = 1 To nb_col

wcol = book0.Sheets(sht1Name).Columns(icol).Width
wcol = wcol / 5.5
book0.Sheets(sht0Name).Columns(icol).ColumnWidth = wcol

Next icol
'
' Columns for Expiration interval
'
     ExpRedCol = 10         'J
     ExpAmberCol = 11      'K
     ExpYellowCol = 12        'L
     
     ComQMSCol = 15 'O      'Amber
     OASISStatCol = 14 ' N  'Red
     DateDisCol = 13 ' M    'Yellow
     '
'
LSearchRow = 2
    
    
nb_Red = 0
nb_Amber = 0
nb_Yellow = 0

colRed = 0
colAmber = 0
colYellow = 0

nb_ComRed = 0  ' Red for ComQMSCol Commitment
nb_OASISRed = 0   ' Red for OASISStatCol OASIS status
nb_DisYellow = 0  ' Yellow for ARP / Discrepancy
    
colComRed = 0
colOasRed = 0
colDisYel = 0

Do While Len(book0.Sheets(sht0Name).Range(F_Col & CStr(LSearchRow)).Value) > 0
      
    If book0.Sheets(sht0Name).Cells(LSearchRow, ExpYellowCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 6
        book0.Sheets(sht0Name).Cells(LSearchRow, ExpYellowCol).Interior.ColorIndex = 6
        nb_Yellow = nb_Yellow + 1
        colYel = 1
    End If
    If book0.Sheets(sht0Name).Cells(LSearchRow, DateDisCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 6
        book0.Sheets(sht0Name).Cells(LSearchRow, DateDisCol).Interior.ColorIndex = 6
        nb_DisYellow = nb_DisYellow + 1
        colDisYel = 1
     End If
     If book0.Sheets(sht0Name).Cells(LSearchRow, ExpAmberCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 45
        book0.Sheets(sht0Name).Cells(LSearchRow, ExpAmberCol).Interior.ColorIndex = 45
        nb_Amber = nb_Amber + 1
        colAmb = 1
    End If
    If book0.Sheets(sht0Name).Cells(LSearchRow, ComQMSCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 3
        book0.Sheets(sht0Name).Cells(LSearchRow, ComQMSCol).Interior.ColorIndex = 3
        nb_ComRed = nb_ComRed + 1
        colComRed = 1
        End If
    If book0.Sheets(sht0Name).Cells(LSearchRow, ExpRedCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 3
        book0.Sheets(sht0Name).Cells(LSearchRow, ExpRedCol).Interior.ColorIndex = 3
        nb_Red = nb_Red + 1
        colRed = 1
    End If
       If book0.Sheets(sht0Name).Cells(LSearchRow, OASISStatCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 3
        book0.Sheets(sht0Name).Cells(LSearchRow, OASISStatCol).Interior.ColorIndex = 3
        nb_OASISRed = nb_OASISRed + 1
        colOasRed = 1
    End If
    

Skip_SCR:

LSearchRow = LSearchRow + 1

Loop
'
'
'   Update Actions Count in Menusheet
'
    book0.Sheets(menusheet).Cells(4, 5).Value = nb_Yellow
    book0.Sheets(menusheet).Cells(4, 5).Interior.ColorIndex = 6 * colYel + 15 * (1 - colYel)
    book0.Sheets(menusheet).Cells(5, 5).Value = nb_Amber
    book0.Sheets(menusheet).Cells(5, 5).Interior.ColorIndex = 45 * colAmb + 15 * (1 - colAmb)
    book0.Sheets(menusheet).Cells(6, 5).Value = nb_Red
    book0.Sheets(menusheet).Cells(6, 5).Interior.ColorIndex = 3 * colRed + 15 * (1 - colRed)
      
    
    book0.Sheets(menusheet).Cells(4, 6).Value = book0.Sheets(sht0Name).Cells(1, ExpYellowCol).Value
    book0.Sheets(menusheet).Cells(4, 6).Interior.ColorIndex = 6 * colYel + 15 * (1 - colYel)
    book0.Sheets(menusheet).Cells(5, 6).Value = book0.Sheets(sht0Name).Cells(1, ExpAmberCol).Value
    book0.Sheets(menusheet).Cells(5, 6).Interior.ColorIndex = 45 * colAmb + 15 * (1 - colAmb)
    book0.Sheets(menusheet).Cells(6, 6).Value = book0.Sheets(sht0Name).Cells(1, ExpRedCol).Value
    book0.Sheets(menusheet).Cells(6, 6).Interior.ColorIndex = 3 * colRed + 15 * (1 - colRed)
    
    book0.Sheets(menusheet).Cells(7, 5).Value = nb_DisYellow
    book0.Sheets(menusheet).Cells(7, 5).Interior.ColorIndex = 6 * colDisYel + 15 * (1 - colDisYel)
    book0.Sheets(menusheet).Cells(8, 5).Value = nb_ComRed
    book0.Sheets(menusheet).Cells(8, 5).Interior.ColorIndex = 3 * colComRed + 15 * (1 - colComRed)
    book0.Sheets(menusheet).Cells(9, 5).Value = nb_OASISRed
    book0.Sheets(menusheet).Cells(9, 5).Interior.ColorIndex = 3 * colOasRed + 15 * (1 - colOasRed)
    
    book0.Sheets(menusheet).Cells(7, 6).Value = book0.Sheets(sht0Name).Cells(1, DateDisCol).Value
    book0.Sheets(menusheet).Cells(7, 6).Interior.ColorIndex = 6 * colDisYel + 15 * (1 - colDisYel)
    book0.Sheets(menusheet).Cells(8, 6).Value = book0.Sheets(sht0Name).Cells(1, ComQMSCol).Value
    book0.Sheets(menusheet).Cells(8, 6).Interior.ColorIndex = 3 * colComRed + 15 * (1 - colComRed)
    book0.Sheets(menusheet).Cells(9, 6).Value = book0.Sheets(sht0Name).Cells(1, OASISStatCol).Value
    book0.Sheets(menusheet).Cells(9, 6).Interior.ColorIndex = 3 * colOasRed + 15 * (1 - colOasRed)
'
' Calculation of Global Indicator
'
    QMSRed = 0
    QMSAmber = 0
    QMSKPI = 2  ' Green
    
    QMSRed = nb_Red + nb_OASISRed + nb_ComRed
    QMSAmber = nb_Amber + nb_ComAmber
        
    If QMSRed <> 0 Then
    QMSKPI = 0          ' Red
    ElseIf QMSAmber <> 0 Then
    QMSKPI = 1          'Amber
    End If
 '

    book0.Sheets(menusheet).Cells(11, 5).Value = QMSKPI
    book0.Sheets(menusheet).Select



End Sub

Sub UpdateBlock_Status()
'
'----------------------------------------------------------------
' Calculation of Bloc Sync Issues
'---------------------------------------------------------------
'
' Init of Parameters
'
Call calc_block
'
'
Dim book0 As Workbook
Dim book0Name As String
Dim book0NamePath As String
    book0Name = SCR_Log_File
    Set book0 = Workbooks(book0Name)
'
Dim sht0Name As String
sht0Name = "Synchro and Block" ' Issues Sheeet

Dim sht1Name As String
sht1Name = "Block Format" ' Issues Analysis sheet Format
F_Col = "A"


For rw = 1 To 2
book0.Sheets(sht1Name).Rows(rw & ":" & rw).Copy
book0.Sheets(sht0Name).Rows(rw & ":" & rw).PasteSpecial Paste:=xlPasteFormats
Next rw

rw = 3
Do While Len(book0.Sheets(sht0Name).Range(F_Col & rw).Value) > 0

book0.Sheets(sht1Name).Rows("2:2").Copy
book0.Sheets(sht0Name).Rows(rw & ":" & rw).PasteSpecial Paste:=xlPasteFormats
rw = rw + 1
Loop

nb_col = book0.Sheets(sht1Name).UsedRange.Columns.count

Dim icol As Integer
Dim wcol As Double

For icol = 1 To nb_col

wcol = book0.Sheets(sht1Name).Columns(icol).Width
wcol = wcol / 5.5
book0.Sheets(sht0Name).Columns(icol).ColumnWidth = wcol

Next icol


'
' Columns for Expiration interval
'
     DesyncCol = 13         'M
     POBlockCol = 14     'L
     
'
    LSearchRow = 2
    
    
nb_Sync = 0
nb_Block = 0

colAmb = 0
colRed = 0
    
Do While Len(book0.Sheets(sht0Name).Range(F_Col & CStr(LSearchRow)).Value) > 0
   
    If book0.Sheets(sht0Name).Cells(LSearchRow, DesyncCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 45
        book0.Sheets(sht0Name).Cells(LSearchRow, DesyncCol).Interior.ColorIndex = 45
        nb_Sync = nb_Sync + 1
        colAmb = 1
    End If
   If book0.Sheets(sht0Name).Cells(LSearchRow, POBlockCol).Value <> "" Then
        book0.Sheets(sht0Name).Cells(LSearchRow, 1).Interior.ColorIndex = 3
        book0.Sheets(sht0Name).Cells(LSearchRow, POBlockCol).Interior.ColorIndex = 3
        nb_Block = nb_Block + 1
        colRed = 1
    End If
          
Skip_SCR:

LSearchRow = LSearchRow + 1

Loop
'
'
'   Update Actions SynCount in Menusheet
'
    book0.Sheets(menusheet).Cells(5, 7).Value = nb_Sync
    book0.Sheets(menusheet).Cells(5, 7).Interior.ColorIndex = 45 * colAmb + 15 * (1 - colAmb)
    book0.Sheets(menusheet).Cells(6, 7).Value = nb_Block
    book0.Sheets(menusheet).Cells(6, 7).Interior.ColorIndex = 3 * colRed + 15 * (1 - colRed)
'
    book0.Sheets(menusheet).Cells(5, 8).Value = book0.Sheets(sht0Name).Cells(1, DesyncCol).Value
    book0.Sheets(menusheet).Cells(5, 8).Interior.ColorIndex = 45 * colAmb + 15 * (1 - colAmb)
    book0.Sheets(menusheet).Cells(6, 8).Value = book0.Sheets(sht0Name).Cells(1, POBlockCol).Value
    book0.Sheets(menusheet).Cells(6, 8).Interior.ColorIndex = 3 * colRed + 15 * (1 - colRed)
'
' Calculation of Global Indicator
'
    
    blockKPI = 2  ' Green
       
    If nb_Block <> 0 Then
    blockKPI = 0          ' Red
    ElseIf nb_Sync <> 0 Then
    blockKPI = 1          'Amber
    End If
 '
    book0.Sheets(menusheet).Cells(11, 7).Value = blockKPI
    book0.Sheets(menusheet).Select
    
    
    
End Sub

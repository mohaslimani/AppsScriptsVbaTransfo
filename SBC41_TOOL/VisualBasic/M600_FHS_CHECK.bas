Attribute VB_Name = "M600_FHS_CHECK"
Sub FHS_Check()

Set WSTEMPLATES = Sheets("TEMPLATES")
Set WSENG8 = Sheets("ENG8")
Set WSREF = Sheets("REF")

Dim j As Integer
Dim firstrowPN As Long, lastrowPN As Long, PNmatchrow As Long
Dim i As Long
Dim PNrange As Range
Dim targetAIRLINE As String, targetProgram As String, Message As String

Dim FHSmatch As Boolean, PNOK As Boolean, programOK As Boolean, airlineOK As Boolean

' check fields ok********************************************************************************
PNOK = False
programOK = False
airlineOK = False

Do

    If WSTEMPLATES.Cells(15, 3) <> "" Then
        PNOK = True
    Else: Message = MsgBox("Please insert PN to check", 48)
        Exit Sub
    End If

Loop While (PNOK <> True)

Do

    If WSTEMPLATES.Cells(6, 3) <> "" Then
        programOK = True
    Else: Message = MsgBox("Please insert Program", 48)
        Exit Sub
    End If

Loop While (programOK <> True)

Do

    If WSTEMPLATES.Cells(9, 3) <> "" Then
        airlineOK = True
    Else: Message = MsgBox("Please insert AIRLINE code", 48)
        Exit Sub
    End If

Loop While (airlineOK <> True)



' **************************************************************************************************



targetAIRLINE = WSTEMPLATES.Cells(9, 3)
targetProgram = WSTEMPLATES.Cells(6, 3)
j = 15

Do

    WSENG8.Cells(59999, 1).Value = WSTEMPLATES.Cells(j, 3)
    FHSmatch = False
    Set PNrange = WSENG8.Range("A1:A60000")
    
    firstrowPN = Application.WorksheetFunction.Match(WSTEMPLATES.Cells(j, 3), PNrange, 0)
       
    If firstrowPN = 59999 Then
        WSTEMPLATES.Cells(j, 4) = "NO FHS"
        WSTEMPLATES.Cells(j, 3).Select
        '    Range("M27:P28").Select
        '    With Selection
        '        .HorizontalAlignment = xlCenter
        '        .VerticalAlignment = xlCenter
        '        .WrapText = False
        '        .Orientation = 0
        '        .AddIndent = False
        '        .IndentLevel = 0
        '        .ShrinkToFit = False
        '        .ReadingOrder = xlContext
        '        .MergeCells = True
        '    End With
        '
        '    Selection.Font.Size = 14
        '    Selection.Font.Bold = True
            'With Selection.Interior
             '   .Pattern = xlSolid
              '  .PatternColorIndex = xlAutomatic
               ' .Color = 255
                '.TintAndShade = 0
                '.PatternTintAndShade = 0
            'End With
    
    End If

    If firstrowPN <> 59999 Then
   
        i = firstrowPN
        lastrowPN = firstrowPN
        PNmatchrow = 0
            Do
            
                lastrowPN = lastrowPN + 1
             
            Loop While WSENG8.Cells(lastrowPN, 1) = WSTEMPLATES.Cells(j, 3)
        
        lastrowPN = lastrowPN - 1
    
    
        For i = firstrowPN To lastrowPN
        
            If InStr(1, WSENG8.Cells(i, 24), targetAIRLINE) > 0 And InStr(1, WSENG8.Cells(i, 22), targetProgram) > 0 And FHSmatch = False And WSENG8.Cells(i, 23) <> "" Then
                    WSTEMPLATES.Cells(j, 4) = "FHS"
                    WSTEMPLATES.Cells(j, 3).Select
                    PNmatchrow = i
                '    Range("M27:P28").Select
                '    With Selection
                '        .HorizontalAlignment = xlCenter
                '        .VerticalAlignment = xlCenter
                '        .WrapText = False
                '        .Orientation = 0
                '        .AddIndent = False
                '        .IndentLevel = 0
                '        .ShrinkToFit = False
                '        .ReadingOrder = xlContext
                '        .MergeCells = True
                '    End With
                '
                '    Selection.Font.Size = 14
                '    Selection.Font.Bold = True
                 '   With Selection.Interior
                  '      .Pattern = xlSolid
                   '     .PatternColorIndex = xlAutomatic
                    '    .Color = 5287936
                     '   .TintAndShade = 0
                      '  .PatternTintAndShade = 0
                    'End With
                    'FHSmatch = True
                    
                '    WSTEMPLATES.Cells(29, 11).Value = "ENG8 ROW:"
                '    WSTEMPLATES.Cells(30, 11).Value = "CONTRACT:"
                '    WSTEMPLATES.Cells(30, 15).Value = "EXP DATE:"
                '    WSTEMPLATES.Cells(29, 16).Value = PNmatchrow
                '    WSTEMPLATES.Cells(30, 14).Value = WSENG8.Cells(PNmatchrow, 23)
                '    WSTEMPLATES.Cells(30, 16).Value = WSENG8.Cells(PNmatchrow, 26)
            
            End If
        
        Next
    
        If FHSmatch = False Then
            WSTEMPLATES.Cells(j, 5) = "NO FHS"
            WSTEMPLATES.Cells(j, 3).Select
        '    Range("M27:P28").Select
        '    With Selection
        '        .HorizontalAlignment = xlCenter
        '        .VerticalAlignment = xlCenter
        '        .WrapText = False
        '        .Orientation = 0
        '        .AddIndent = False
        '        .IndentLevel = 0
        '        .ShrinkToFit = False
        '        .ReadingOrder = xlContext
        '        .MergeCells = True
        '    End With
        '
        '    Selection.Font.Size = 14
        '    Selection.Font.Bold = True
            'With Selection.Interior
             '   .Pattern = xlSolid
              '  .PatternColorIndex = xlAutomatic
               ' .Color = 255
                '.TintAndShade = 0
               ' .PatternTintAndShade = 0
           ' End With
        End If
    
    
    End If
    
    j = j + 1
    
Loop While (WSTEMPLATES.Cells(j, 3) <> "")

End Sub


Attribute VB_Name = "MM0100_TEMPLATE_GEN"
Sub TEMPLATE_GEN()

'**************parameter creation***********************

Dim WBTARGET As Workbook
Dim WBSOURCE As Workbook
Dim WSSOURCE As Worksheet
'Dim WSTEMPLATES As Worksheet
Dim WBTEMPLATES As Workbook
Dim WSAIBPN As Worksheet, WSAIBPNQTR As Worksheet, WSEQP As Worksheet, WSEQPA350 As Worksheet, WSEQPQTR As Worksheet
Dim WSAVIONICS As Worksheet, WSAIBtool As Worksheet, WSOEMtool As Worksheet, WSCONSUM As Worksheet, WSSTDpart As Worksheet, WSSTDpartQTR As Worksheet
Dim WSRAWmat As Worksheet

Dim i As Integer, j As Integer
Dim Pn As String
Dim PNqty As String
Dim PNtype As String, PNref As String
Dim AIRLINE As String, PROGRAM As String, MSN As String, TAIL As String, SITUATION As String, LOCATION As String, RTS As String, FHS As String, TR As String
Dim filename As String



Application.ScreenUpdating = False

'**********FHS CHECK***************

Call FHS_Check

'**************parameter INITIALIZATION***********************

path = ActiveWorkbook.path
Set WBTARGET = Workbooks.Add
Set WBSOURCE = ThisWorkbook
Set WSSOURCE = WBSOURCE.Sheets("TEMPLATES")
Set WBTEMPLATES = Workbooks.Open(path & "\PN_TEMPLATES.xlsx")
'Set WBTEMPLATES = Workbooks.Open("\\sfs.corp\Projects\CUSTOMERSERVICE\AIRTAC\Supply Engineering\10 Tools\SOA4_TOOL\PN_TEMPLATES.xlsx")
'P:\\sfs.corp\Projects\CUSTOMERSERVICE\AIRTAC\Supply Engineering\10 Tools\SOA4_TOOL


Set WSAIBPN = WBTEMPLATES.Sheets("AIB PN")
Set WSAIBPNQTR = WBTEMPLATES.Sheets("AIB PN QTR")

Set WSEQP = WBTEMPLATES.Sheets("EQUIPMENT")
Set WSEQPA350 = WBTEMPLATES.Sheets("EQUIPMENT A350")
Set WSEQPQTR = WBTEMPLATES.Sheets("EQUIPMENT QTR")

Set WSAVIONICS = WBTEMPLATES.Sheets("AIB Avionics")

Set WSAIBtool = WBTEMPLATES.Sheets("AIB Tool")
Set WSOEMtool = WBTEMPLATES.Sheets("OEM Tool")

Set WSSTDpart = WBTEMPLATES.Sheets("STD part")
Set WSSTDpartQTR = WBTEMPLATES.Sheets("STD part QTR")

Set WSCONSUM = WBTEMPLATES.Sheets("CONSUMABLES")
Set WSRAWmat = WBTEMPLATES.Sheets("RAW MATERIAL")




'****************CHECK LIST CREATION************************************************

            'PN INPUT****************************************************************************************************
            
            i = 15
            
            Do
            
            PNnumber = PNnumber + 1
            i = i + 1
            Loop While (WSSOURCE.Cells(i, 3).Value <> "")
            
                  j = 0
            
            '*****************************************************************************

AIRLINE = WSSOURCE.Cells(9, 3).Value
PROGRAM = WSSOURCE.Cells(6, 3).Value

For i = 1 To PNnumber
 j = i + 14
 PNtype = WSSOURCE.Cells(j, 2).Value
 PNref = WSSOURCE.Cells(j, 3).Value

'******AIB PN*************************

 If PNtype = "AIB PN" And AIRLINE = "QTR" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSAIBPN.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
 
  If PNtype = "AIB PN" And AIRLINE <> "QTR" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSAIBPNQTR.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
 
 
'****** EQUIPMENT *************************

 If PNtype = "Equipment" And PROGRAM = "A350" And AIRLINE <> "QTR" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSEQPA350.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
 
  If PNtype = "Equipment" And AIRLINE = "QTR" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSEQPQTR.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
 
 If PNtype = "Equipment" And PROGRAM <> "A350" And AIRLINE <> "QTR" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSEQP.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
 
 '****** AVIONICS *************************

 If PNtype = "AIB Avionics" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSAVIONICS.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
 
  '****** TOOLS *************************

 If PNtype = "AIB tool" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSAIBtool.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
  
  If PNtype = "OEM tool" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSOEMtool.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
 
 '****** STD PARTS *************************

 If PNtype = "STD" And AIRLINE <> "QTR" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSSTDpart.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
 
  If PNtype = "STD" And AIRLINE = "QTR" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSSTDpartQTR.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If

'****** CONSUM AND RAW PARTS *************************

 If PNtype = "Consumible" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSCONSUM.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If
 
  If PNtype = "Raw Material" Then
   WBTARGET.Activate
   WBTARGET.Sheets.Add(After:=Worksheets(Worksheets.count)).Name = PNref
   WSRAWmat.Cells.Copy
   WBTARGET.Sheets(PNref).Range("A1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 End If

Next

Application.DisplayAlerts = False
WBTARGET.Activate
WBTARGET.Sheets(1).Delete
WBTARGET.Sheets(1).Delete
WBTARGET.Sheets(1).Delete
WBTEMPLATES.Close
Application.DisplayAlerts = True

'*******FILL AUTO FIELDS*******************
AIRLINE = WSSOURCE.Cells(9, 3).Value
PROGRAM = WSSOURCE.Cells(6, 3).Value
MSN = WSSOURCE.Cells(7, 3).Value
TAIL = WSSOURCE.Cells(8, 3).Value
SITUATION = WSSOURCE.Cells(6, 5).Value
LOCATION = WSSOURCE.Cells(7, 5).Value
RTS = WSSOURCE.Cells(8, 5).Value
TR = WSSOURCE.Cells(9, 5).Value


For i = 1 To PNnumber

Pn = WSSOURCE.Cells(i + 14, 3).Value
PNqty = WSSOURCE.Cells(i + 14, 5).Value
PNtype = WSSOURCE.Cells(i + 14, 2).Value
FHS = WSSOURCE.Cells(i + 14, 4).Value
WBTARGET.Activate
WBTARGET.Sheets(i).Cells(2, 3) = TR
WBTARGET.Sheets(i).Cells(4, 2) = AIRLINE
WBTARGET.Sheets(i).Cells(5, 2) = PROGRAM
WBTARGET.Sheets(i).Cells(6, 2) = MSN
WBTARGET.Sheets(i).Cells(7, 2) = TAIL
WBTARGET.Sheets(i).Cells(4, 4) = SITUATION
WBTARGET.Sheets(i).Cells(5, 4) = LOCATION
WBTARGET.Sheets(i).Cells(6, 4) = RTS

WBTARGET.Sheets(i).Cells(9, 2) = Pn
WBTARGET.Sheets(i).Cells(9, 4) = PNqty

If PNtype = "AIB PN" Or PNtype = "Equipment" Or PNtype = "STD" Then
 WBTARGET.Sheets(i).Cells(9, 7) = FHS
End If

Next

'WSTEMPLATES.Range("C6").Select

WBTARGET.Activate
Sheets(1).Activate
Sheets(1).Range("A1").Select


filename = TR & "_" & PROGRAM & "_" & AIRLINE & "_MSN" & MSN & ".xlsx"

'C:\Users\c82360\Desktop\AUTO_CHECK_LIST\CHECK_LISTS_CREATED

WBTARGET.SaveAs "\\sfs.corp\Projects\CUSTOMERSERVICE\AIRTAC\Supply Engineering\10 Tools\SOA4_TOOL\CHECK_LISTS_CREATED\" & filename, xlOpenXMLStrictWorkbook
   
End Sub



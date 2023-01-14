Attribute VB_Name = "Parameters"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'               DATA INITIALIZATION
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'============================================================
'   Files
'============================================================
'
Public Const SCR_Log_File       As String = "SCR_LogBook v8-Draft.xlsm"
'
'============================================================
'   Sheet Names
'============================================================
'
Public Const menusheet = "Menu"
Public Const indicsheet = "Indicators"
Public Const actionsheet = "Supplier Control Review"
Public Const clossheet = "Closed Actions"
Public Const calcsheet = "Calculation"
Public Const peoplesheet = "People"
Public Const parasheet = "Fixed Parameters"
Public Const logsheet = "Update Log"
Public Const scrptsheet = "Supplier Control Review Report"
'
'============================================================
'   Name of the 4 different Status of the actions
'============================================================
'
Public Const A_Status0 = "On Time"
Public Const A_Status1 = "On Time - Alert"
Public Const A_Status2 = "Late"
Public Const A_Status3 = "Late - Red"
Public Const A_Status4 = "Not Available"

'============================================================
'   Column Numbers of the actions sheet
'============================================================
Public Const jarp = 1           ' ARP Id
Public Const jcie = 2           '  Company Name
Public Const jstreet = 3        ' Company Street
Public Const jpcode = 4         ' Company Postal Code
Public Const jcity = 5          ' Company city
Public Const jcountry = 6       ' Company Country
Public Const jpgl2 = 9          ' Product Group Level 2
Public Const jqr = 10           ' Product Group QR
'
'        Action Contributors
'
Public Const jnextscr = 13         ' Next SCR
Public Const jscrstatus = 12       ' Status
'
'============================================================
'   Action Sheet parameters
'============================================================
'
Public Const first_act = 2      ' Line of the first action logged in Actions sheet
Public Const last_act = 1000 ' Line of the last action logged in Actions sheet
Public Const colref = 1         ' Column reference to be taken into consideration
Public Const Lcol_first = "A"   ' First letter column where a field exist
Public Const Lcol_last = "AH"    ' Last  letter column where a field exist
'
'============================================================
'   Update Log parameters
'============================================================
Public Const Range_Copy_H = "I33:O33"   ' Range to copy from Menu to Log Update
Public Const col_log_first = "A"    ' First column
Public Const col_log_copy = "B"    ' Column where to copy
Public Const log_first_line = 2     ' Line first date
Public Const log_last_line = "1000" ' Line last date
Public Const collogflag = 8     ' Column of FLAG in LogUpdate


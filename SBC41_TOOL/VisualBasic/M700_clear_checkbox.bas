Attribute VB_Name = "M700_clear_checkbox"
Sub ClearCheckBoxes()
'Updateby Extendoffice 20161129
    Dim chkBox As Excel.CheckBox
    Application.ScreenUpdating = False
    For Each chkBox In ActiveSheet.CheckBoxes
            chkBox.Value = xlOff
    Next chkBox
    Application.ScreenUpdating = True
End Sub

Attribute VB_Name = "MM0101_OPEN_CHECKLIST_FOLDER"
Sub open_checklist_folder()

'"\\sfs.corp\Projects\CUSTOMERSERVICE\AIRTAC\Supply Engineering\10 Tools\SOA4_TOOL\CHECK_LISTS_CREATED\"

Dim foldername As String
Dim FileSystem

foldername = "\\sfs.corp\Projects\CUSTOMERSERVICE\AIRTAC\Supply Engineering\10 Tools\SOA4_TOOL\CHECK_LISTS_CREATED"

Set FileSystem = CreateObject("Scripting.FileSystemObject")

Call Shell("explorer.exe " & foldername, vbNormalFocus)

End Sub

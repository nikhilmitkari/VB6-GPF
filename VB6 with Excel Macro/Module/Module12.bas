Attribute VB_Name = "Module12"
Sub ProtectAllWorksheets()

'Create a variable to hold worksheets
Dim ws As Worksheet

'Loop through each worksheet in the active workbook
For Each ws In ActiveWorkbook.Worksheets

'Protect each worksheet
ws.Protect Password:="Gpf@vba", UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=True, _
    AllowInsertingRows:=True, _
    AllowInsertingHyperlinks:=True, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=False, _
    AllowSorting:=False, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False

Next ws

End Sub
Sub UnprotectAllWorksheets()

'Create a variable to hold worksheets
Dim ws As Worksheet

'Loop through each worksheet in the active workbook
For Each ws In ActiveWorkbook.Worksheets

'Unprotect each worksheet
ws.Unprotect Password:="Gpf@vba"

Next ws

End Sub

Private Sub SleepDemo()
Sleep 500 'milliseconds (pause for 0.5 second)
'resume macro
End Sub



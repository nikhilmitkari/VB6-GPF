Attribute VB_Name = "Module13"
Sub Save_pdf()
Attribute Save_pdf.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Save_pdf Macro
'

'
    Range("I1:P48").Select
    ActiveSheet.PageSetup.PrintArea = "$I$1:$P$48"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub
Sub Save_pdf2()
Attribute Save_pdf2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Save_pdf2 Macro
'

'
    Range("I2:O23").Select
    ActiveSheet.PageSetup.PrintArea = "$I$2:$O$23"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub

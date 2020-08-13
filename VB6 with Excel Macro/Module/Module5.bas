Attribute VB_Name = "Module5"
Sub clickbutton()
Call Find_data
End Sub
Sub nomineede()
Attribute nomineede.VB_ProcData.VB_Invoke_Func = " \n14"
'
' nomineede Macro
'

    Sheets("Employeed_details").Select
    Range("M24").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Sheets("Employeed_details").Select
    Range("C24").Select
    Selection.Copy
    Sheets("Nominee").Select
    Range("V2").Select
    ActiveSheet.Paste
    Call Find_data
End Sub

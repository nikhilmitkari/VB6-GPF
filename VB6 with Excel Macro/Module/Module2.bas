Attribute VB_Name = "Module2"
Option Explicit
Sub Find_data()
Dim emp_name As String
Dim finalrow As Integer
Dim i As Integer


'Sheets("Nominee").Range("Emp_table[#All]").ClearContents
Sheets("Nominee").Range("G8:I30").ClearContents
'emp_name = DropDown2.Value
Application.ScreenUpdating = False
ActiveWindow.Zoom = 100
Sheet4.Select
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 4")).Visible = msoFalse
Application.ScreenUpdating = True

MsgBox "Please Enter the employee name to view Nominee details"
emp_name = Application.InputBox("Enter Employee Name", "Nominee View")
'emp_name = Range("V2").value

'Set dd = ActiveSheet.Shapes("Drop Down 2").OLEFormat.Object
'emp_name = dd.List(dd.ListIndex)

If emp_name = "" Then
MsgBox "Please Enter the employee name!"
End If
finalrow = Sheets("Nominee").Range("A1000").End(xlUp).Row

For i = 2 To finalrow
  If Cells(i, 1) = emp_name Then
  Range(Cells(i, 1), Cells(i, 3)).Copy
  Range("G100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  'Else MsgBox "Employee is not available"
  
  End If
Next i
    Application.ScreenUpdating = False
ActiveWindow.Zoom = 100
Sheet4.Select
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 4")).Visible = msoTrue
Application.ScreenUpdating = True
End Sub


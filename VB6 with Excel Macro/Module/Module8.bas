Attribute VB_Name = "Module8"
Sub Button2_Click()
Dim scl_name As String
Dim finalrow As Integer
Dim i As Integer


'Sheets("Nominee").Range("Emp_table[#All]").ClearContents
Sheets("Employeed_details").Range("Y4:Z100").ClearContents
'emp_name = DropDown2.Value
'MsgBox "Please Enter the employee name to view Nominee details"
scl_name = Application.InputBox("Enter Employee Name", "Nominee View")
'emp_name = Range("V2").value

'Set dd = ActiveSheet.Shapes("Drop Down 2").OLEFormat.Object
'emp_name = dd.List(dd.ListIndex)


finalrow = Sheets("Employeed_details").Range("B1000").End(xlUp).Row

For i = 8 To finalrow
  If Sheets("Employeed_details").Cells(i, 2) = scl_name Then
  Range(Cells(i, 2), Cells(i, 3)).Copy
  Range("Y100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  'Else MsgBox "Employee is not available"
  
  End If
Next i

End Sub
Sub Button3_Click()

Dim Employee_name As String
Dim finalrow As Integer
Dim i As Integer
'ComboBox_calDeg.Clear

Sheets("Employeed_details").Range("AB4:AC100").ClearContents

'Employee_name = ComboBox_calEmpName.value

 Employee_name = "Ashutosh"



finalrow = Sheets("Employeed_details").Range("C1000").End(xlUp).Row

For i = 8 To finalrow
  If Sheets("Employeed_details").Cells(i, 3) = Employee_name Then
  Range(Cells(i, 3), Cells(i, 4)).Copy
  Range("AB100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  End If
Next i




End Sub

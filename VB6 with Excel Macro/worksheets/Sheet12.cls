VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ComboBox1_change()
On Error Resume Next



Dim scl_name As String
Dim finalrow As Integer
Dim i As Integer
 'Worksheets("School Report").OLEObjects("ComboBox1").Object.value
 ComboBox1 = Worksheets("School Report").OLEObjects("ComboBox1").Object.value
 ComboBox2 = Worksheets("School Report").OLEObjects("ComboBox2").Object
ComboBox2.Clear

'Sheets("Nominee").Range("Emp_table[#All]").ClearContents
Sheets("School_Data").Range("T2:V50").ClearContents
'emp_name = DropDown2.Value
'MsgBox "Please Enter the employee name to view Nominee details"
scl_name = ComboBox1.value
'scl_name = "J D "

'emp_name = Range("V2").value

'Set dd = ActiveSheet.Shapes("Drop Down 2").OLEFormat.Object
'emp_name = dd.List(dd.ListIndex)


finalrow = Sheets("Data").Range("C100").End(xlUp).Row

For i = 3 To finalrow
  If Sheet7.Cells(i, 3) = scl_name Then
  Sheet7.Cells(i, 3).Copy
  Sheets("School_Data").Range("T50").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
Sheet7.Cells(i, 5).Copy
  Sheets("School_Data").Range("U50").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  'Else MsgBox "Employee is not available"
  
  End If
Next i



For Each acell In Range("syear")
    If acell.value <> "" Then
     Me.ComboBox2.AddItem acell.value
    End If
 Next


End Sub

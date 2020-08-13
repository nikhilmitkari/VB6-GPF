Attribute VB_Name = "Module3"
Sub calculation_Button1_Click()

Dim emp_name As String
Dim finalrow As Integer
Dim i As Integer


'Sheets("Nominee").Range("Emp_table[#All]").ClearContents
Sheets("calculation").Range("A2:D50").ClearContents
'emp_name = DropDown2.Value



emp_name = Sheets("Pay_Slip").Range("K4").value
'emp_name = Range("V2").value

'Set dd = ActiveSheet.Shapes("Drop Down 2").OLEFormat.Object
'emp_name = dd.List(dd.ListIndex)


finalrow = Sheets("Data").Range("B10000").End(xlUp).Row

For i = 3 To finalrow
  If Sheets("Data").Cells(i, 2) = emp_name Then
  'Application.Worksheets.Select
  
 Sheets("Data").Cells(i, 2).Copy
  Sheets("calculation").Range("A100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  'Else MsgBox "Employee is not available"
      Sheets("Data").Cells(i, 5).Copy
      Sheets("calculation").Range("B100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
     
        Sheets("Data").Cells(i, 24).Copy
   Sheets("calculation").Range("C100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
    Sheets("Data").Cells(i, 25).Copy
   Sheets("calculation").Range("D100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  
  End If
Next i

   
  
End Sub

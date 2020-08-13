VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Refund_Update 
   ClientHeight    =   10572
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   22524
   OleObjectBlob   =   "Refund_Update.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Refund_Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Addrefund_Click()
Dim deta As Long

deta = Sheets("Refund_Details").Cells(Rows.Count, "A").End(xlUp).Row

If deta = 6 And Sheets("Refund_Details").Cells(deta, "A").value = "" Then
deta = deta
Else
deta = deta + 1
End If

If ComboBox_refundScl.value = "" Or ComboBox_refundEmp.value = "" Or tv_refundGPF.value = "" Or ComboBox_refundYear.value = "" Then
MsgBox "Please enter details in all fields", vbOKOnly + vbCritical
Else
'On Error GoTo ErrorHandler
Sheets("Refund_Details").Cells(deta, "A").value = Application.WorksheetFunction.Max(Sheets("Refund_Details").Range("A:A")) + 1
Sheets("Refund_Details").Cells(deta, "B").value = ComboBox_refundScl.value
Sheets("Refund_Details").Cells(deta, "C").value = ComboBox_refundEmp.value
Sheets("Refund_Details").Cells(deta, "D").value = tv_refundGPF.value
nxt_year = Right(ComboBox_refundYear + 1, 2)
Sheets("Refund_Details").Cells(deta, "E").value = ComboBox_refundYear.value & " - " & nxt_year
Sheets("Refund_Details").Cells(deta, "F").value = tv_refundApr.value
Sheets("Refund_Details").Cells(deta, "G").value = tv_refundMay.value
Sheets("Refund_Details").Cells(deta, "H").value = tv_refundJun.value
Sheets("Refund_Details").Cells(deta, "I").value = tv_refundJuly.value
Sheets("Refund_Details").Cells(deta, "J").value = tv_refundAug.value
Sheets("Refund_Details").Cells(deta, "K").value = tv_refundSep.value
Sheets("Refund_Details").Cells(deta, "L").value = tv_refundOct.value
Sheets("Refund_Details").Cells(deta, "M").value = tv_refundNov.value
Sheets("Refund_Details").Cells(deta, "N").value = tv_refundDec.value
Sheets("Refund_Details").Cells(deta, "O").value = tv_refundJan.value
Sheets("Refund_Details").Cells(deta, "P").value = tv_refundFeb.value
Sheets("Refund_Details").Cells(deta, "Q").value = tv_refundMar.value


ActiveWorkbook.Save

MsgBox "Details updated Successfully", vbOKOnly + vbExclamation

'Exit Sub
'ErrorHandler: ActiveWorkbook.Save

ComboBox_refundScl.value = ""
ComboBox_refundEmp.value = ""
tv_refundGPF.value = ""
ComboBox_refundYear.value = ""
tv_refundApr.value = ""
tv_refundMay.value = ""
tv_refundJun.value = ""
tv_refundJuly.value = ""
tv_refundAug.value = ""
tv_refundSep.value = ""
tv_refundOct.value = ""
tv_refundNov.value = ""
tv_refundDec.value = ""
tv_refundJan.value = ""
tv_refundFeb.value = ""
tv_refundMar.value = ""
End If

ListBox1.RowSource = "Refund_Data"
End Sub


Private Sub btn_refundHome_Click()
Unload Me
MainWindow.Show

End Sub

Private Sub btn_refundSearch_Click()
'Dim x As Long
'Dim i As Long
' x = Sheet11.Range("A" & Rows.Count).End(xlUp).Row
'
'For i = 5 To x
'If Sheets("Refund_Details").Cells(i, 4) = tv_refundSearchBox.Text Then
'ComboBox_refundScl = Sheet11.Cells(i, 2).value
'ComboBox_refundEmp = Sheet11.Cells(i, 3).value
'tv_refundGPF = Sheet11.Cells(i, 4).value
'ComboBox_refundYear = Sheet11.Cells(i, 5).value
'tv_refundApr = Sheet11.Cells(i, 6).value
'tv_refundMay = Sheet11.Cells(i, 7).value
'tv_refundJun = Sheet11.Cells(i, 8).value
'tv_refundJuly = Sheet11.Cells(i, 9).value
'tv_refundAug = Sheet11.Cells(i, 10).value
'tv_refundSep = Sheet11.Cells(i, 11).value
'tv_refundOct = Sheet11.Cells(i, 12).value
'tv_refundNov = Sheet11.Cells(i, 13).value
'tv_refundDec = Sheet11.Cells(i, 14).value
'tv_refundJan = Sheet11.Cells(i, 15).value
'tv_refundFeb = Sheet11.Cells(i, 16).value
'tv_refundMar = Sheet11.Cells(i, 17).value
'End If
'Next i


Dim database(1 To 100, 1 To 17)
Dim My_range As Integer
Dim colum As Byte

Me.ListBox1.RowSource = ""

Lastrow = Sheet11.Range("A" & Rows.Count).End(xlUp).Row
On Error Resume Next 'GoTo ErrorHandler
Sheet11.Range("D").AutoFilter field:=1, Criteria1:=Me.tv_refundSearchBox

For i = 5 To Lastrow
If Sheets("Refund_Details").Cells(i, 4).value = tv_refundSearchBox.Text Then
My_range = My_range + 1
For colum = 1 To 17
database(My_range, colum) = Sheets("Refund_Details").Cells(i, colum)
Next colum
End If
Next i

Me.ListBox1.List = database
'Exit Sub
'ErrorHandler:
'Call btn_reset_Click


End Sub

Private Sub btn_reset_Click()
ListBox1.RowSource = "Refund_Data"
tv_refundSearchBox.Text = ""
ComboBox_refundScl = ""
ComboBox_refundEmp = ""
tv_refundGPF = ""
ComboBox_refundYear = ""
tv_refundApr = ""
tv_refundMay = ""
tv_refundJun = ""
tv_refundJuly = ""
tv_refundAug = ""
tv_refundSep = ""
tv_refundOct = ""
tv_refundNov = ""
tv_refundDec = ""
tv_refundJan = ""
tv_refundFeb = ""
tv_refundMar = ""

End Sub

Private Sub ComboBox_refundEmp_AfterUpdate()
Set tblrange = Worksheets("Employeed_details").Range("C2:H500")
On Error GoTo ErrorHandler
tv_refundGPF.value = Application.WorksheetFunction.VLookup(ComboBox_refundEmp.value, tblrange, 6, False)
'ComboBox_refundYear.value = Application.WorksheetFunction.VLookup(ComboBox_refundScl.value, tblrange, 4, False)
Exit Sub ' Exit to avoid handler.
ErrorHandler: ComboBox_refundEmp.Clear
tv_refundGPF.value = ""
Call ComboBox_refundScl_AfterUpdate
End Sub

Private Sub ComboBox_refundScl_AfterUpdate()
On Error Resume Next

Dim scl_name As String
Dim finalrow As Integer
Dim i As Integer

ComboBox_refundEmp.Clear

Sheets("Employeed_details").Range("Y4:Z100").ClearContents

scl_name = ComboBox_refundScl.value

finalrow = Sheets("Employeed_details").Range("B1000").End(xlUp).Row

For i = 8 To finalrow
  If Sheets("Employeed_details").Cells(i, 2) = scl_name Then
  Sheets("Employeed_details").Range(Cells(i, 2), Cells(i, 3)).Copy
  Sheets("Employeed_details").Range("Y100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  'Else MsgBox "Employee is not available"
  
  End If
Next i

For Each acell In Range("sclname")
    If acell.value <> "" Then
     Me.ComboBox_refundEmp.AddItem acell.value
    End If
 Next

End Sub
'Exit(ByVal Cancel As MSForms.ReturnBoolean)


Private Sub Label23_Click()

End Sub

Private Sub ListBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim Lastrow As Long
Lastrow = Sheet11.Range("A" & Rows.Count).End(xlUp).Row
'Set rng = Sheet11.Range("A1:A" & Lastrow)
'
For i = 5 To Lastrow
If Sheet11.Cells(i, 1) = ListBox1.List(ListBox1.ListIndex) Then
Rows(i).Select
ComboBox_refundScl = Sheet11.Cells(i, 2).value
ComboBox_refundEmp = Sheet11.Cells(i, 3).value
tv_refundGPF = Sheet11.Cells(i, 4).value
ComboBox_refundYear = Sheet11.Cells(i, 5).value
tv_refundApr = Sheet11.Cells(i, 6).value
tv_refundMay = Sheet11.Cells(i, 7).value
tv_refundJun = Sheet11.Cells(i, 8).value
tv_refundJuly = Sheet11.Cells(i, 9).value
tv_refundAug = Sheet11.Cells(i, 10).value
tv_refundSep = Sheet11.Cells(i, 11).value
tv_refundOct = Sheet11.Cells(i, 12).value
tv_refundNov = Sheet11.Cells(i, 13).value
tv_refundDec = Sheet11.Cells(i, 14).value
tv_refundJan = Sheet11.Cells(i, 15).value
tv_refundFeb = Sheet11.Cells(i, 16).value
tv_refundMar = Sheet11.Cells(i, 17).value
End If
Next i
End Sub

Private Sub tv_refundDelete_Click()
Dim Lastrow As Long
On Error GoTo ErrorHandler
Lastrow = Sheet11.Range("A" & Rows.Count).End(xlUp).Row
'Set rng = Sheet11.Range("A1:A" & Lastrow)
'
For i = 5 To Lastrow
If Sheet11.Cells(i, 1) = ListBox1.List(ListBox1.ListIndex) Then
Sheet11.Select
Rows(i).Select
Selection.Delete

End If
Next i

MsgBox "Data Deleted successfully", vbInformation, ""
ActiveWorkbook.Save
ComboBox_refundScl.value = ""
ComboBox_refundEmp.value = ""
tv_refundGPF.value = ""
ComboBox_refundYear.value = ""
tv_refundApr.value = ""
tv_refundMay.value = ""
tv_refundJun.value = ""
tv_refundJuly.value = ""
tv_refundAug.value = ""
tv_refundSep.value = ""
tv_refundOct.value = ""
tv_refundNov.value = ""
tv_refundDec.value = ""
tv_refundJan.value = ""
tv_refundFeb.value = ""
tv_refundMar.value = ""
ListBox1.RowSource = "Refund_Data"

ErrorHandler: MsgBox "Invalid Details! To delete please select existing employee", vbOKOnly + vbExclamation, ""
ListBox1.BackColor = &H8080FF
Application.Wait (Now + 0.000009)
'Application.Wait (Now + TimeValue("0:00:01"))
ListBox1.BackColor = &HFFC0C0
Call btn_reset_Click
End Sub

Private Sub tv_refundUpdate_Click()
Dim Lastrow As Long
On Error GoTo ErrorHandler
Lastrow = Sheet11.Range("A" & Rows.Count).End(xlUp).Row
'Set rng = Sheet11.Range("A1:A" & Lastrow)
'
For i = 5 To Lastrow
If Sheet11.Cells(i, 1) = ListBox1.List(ListBox1.ListIndex) Then
'Rows(i).Select
Sheets("Refund_Details").Cells(i, "B").value = ComboBox_refundScl.value
Sheets("Refund_Details").Cells(i, "C").value = ComboBox_refundEmp.value
Sheets("Refund_Details").Cells(i, "D").value = tv_refundGPF.value
Sheets("Refund_Details").Cells(i, "E").value = ComboBox_refundYear.value
Sheets("Refund_Details").Cells(i, "F").value = tv_refundApr.value
Sheets("Refund_Details").Cells(i, "G").value = tv_refundMay.value
Sheets("Refund_Details").Cells(i, "H").value = tv_refundJun.value
Sheets("Refund_Details").Cells(i, "I").value = tv_refundJuly.value
Sheets("Refund_Details").Cells(i, "J").value = tv_refundAug.value
Sheets("Refund_Details").Cells(i, "K").value = tv_refundSep.value
Sheets("Refund_Details").Cells(i, "L").value = tv_refundOct.value
Sheets("Refund_Details").Cells(i, "M").value = tv_refundNov.value
Sheets("Refund_Details").Cells(i, "N").value = tv_refundDec.value
Sheets("Refund_Details").Cells(i, "O").value = tv_refundJan.value
Sheets("Refund_Details").Cells(i, "P").value = tv_refundFeb.value
Sheets("Refund_Details").Cells(i, "Q").value = tv_refundMar.value
End If
Next i

MsgBox "Data updated successfully", vbInformation, ""
ActiveWorkbook.Save
Exit Sub
ErrorHandler: MsgBox "Invalid Details", vbExclamation, ""
Call btn_reset_Click
'ListBox1.RowSource = "Refund_Data"

End Sub

Private Sub UserForm_Activate()

Application.WindowState = xlMaximized
Me.Height = Application.Height
Me.Width = Application.Width

End Sub

Private Sub UserForm_Initialize()
Me.ListBox1.RowSource = "Refund_Data"
End Sub

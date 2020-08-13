VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Add_Scl 
   ClientHeight    =   10572
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   18924
   OleObjectBlob   =   "Add_Scl.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Add_Scl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_AsclHome_Click()
Unload Me
MainWindow.Show
End Sub

Private Sub UserForm_Activate()
Application.WindowState = xlMaximized
Me.Height = Application.Height
Me.Width = Application.Width

End Sub

Private Sub UserForm_Initialize()
Me.tv_SrNo = Application.WorksheetFunction.Max(Sheets("School_Details").Range("A:A")) + 1
Me.label_Date = Now()
End Sub
Private Sub btn_SclUpdate_Click()
Dim Data As Long

Data = Sheets("School_Details").Cells(Rows.Count, "A").End(xlUp).Row

If Data = 2 And Cells(Data, "A").value = "" Then
Data = Data
Else
Data = Data + 1
End If

If tv_SclName.value = "" Or tv_SclAddress.value = "" Or tv_SclDistrict.value = "" Or tv_SclPayUnitNo.value = "" Or tv_SclHmName.value = "" Or tv_SclContact.value = "" Or tv_SclPanchayatSamiti.value = "" Then
MsgBox "Please enter details in all fields", vbOKOnly + vbQuestion, ""
Else
On Error GoTo ErrorHandler
Sheets("School_Details").Cells(Data, "A").value = tv_SrNo.value
Sheets("School_Details").Cells(Data, "B").value = tv_SclName.value
Sheets("School_Details").Cells(Data, "C").value = tv_SclAddress.value
Sheets("School_Details").Cells(Data, "D").value = tv_SclDistrict.value
Sheets("School_Details").Cells(Data, "E").value = tv_SclPayUnitNo.value
Sheets("School_Details").Cells(Data, "F").value = tv_SclHmName.value
Sheets("School_Details").Cells(Data, "G").value = tv_SclContact.value
Sheets("School_Details").Cells(Data, "I").value = tv_SclPanchayatSamiti.value
Sheets("School_Details").Cells(Data, "H").value = "=COUNTIF(Table3[School_Name],[@[School_name]])"
 
ActiveWorkbook.Save

MsgBox "School Details updated Successfully", vbOKOnly, ""
tv_SrNo = Application.WorksheetFunction.Max(Sheets("School_Details").Range("A:A")) + 1
tv_SclName.value = ""
tv_SclAddress.value = ""
tv_SclDistrict.value = ""
tv_SclPayUnitNo.value = ""
tv_SclHmName.value = ""
tv_SclContact.value = ""
tv_SclPanchayatSamiti.value = ""
Exit Sub ' Exit to avoid handler.
ErrorHandler: ActiveWorkbook.Save

End If
End Sub

Private Sub btn_SclReset_Click()
tv_SclName.value = ""
tv_SclAddress.value = ""
tv_SclDistrict.value = ""
tv_SclPayUnitNo.value = ""
tv_SclHmName.value = ""
tv_SclContact.value = ""
tv_SclPanchayatSamiti.value = ""
End Sub





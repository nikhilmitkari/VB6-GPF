VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Add_Teacher 
   ClientHeight    =   10572
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   25932
   OleObjectBlob   =   "Add_Teacher.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Add_Teacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_empAddNominee_Click()
Dim Data As Long

Data = Sheets("Nominee").Cells(Rows.Count, "A").End(xlUp).Row

If Data = 2 And Cells(Data, "A").value = "" Then
Data = Data
Else
Data = Data + 1
End If

If tv_empNominee.value = "" Or ComboBox_empnr.value = "" Then
MsgBox "Please enter details in all fields", vbOKOnly + vbExclamation, ""
Else
With Me
Sheets("Nominee").Cells(Data, "A").value = tv_empName.value
Sheets("Nominee").Cells(Data, "B").value = tv_empNominee.value
Sheets("Nominee").Cells(Data, "C").value = ComboBox_empnr.value
ActiveWorkbook.Save

End With
End If
tv_empNominee.value = ""
ComboBox_empnr.value = ""
End Sub

Private Sub btnDateCal_Click()
cal.lblCtrlName = "tv_empDob"
cal.lblUF = "Add_Teacher"
cal.Show
End Sub

Private Sub btndorcal_Click()
cal.lblCtrlName = "tv_empDoR"
cal.lblUF = "Add_Teacher"
cal.Show
End Sub
Private Sub btn_AempHome_Click()
Unload Me
MainWindow.Show
End Sub

Private Sub btn_empUpdate_Click()
Dim Data As Long
Dim LValue As String

Data = Sheets("Employeed_details").Cells(Rows.Count, "A").End(xlUp).Row

If Data = 8 And Cells(Data, "A").value = "" Then
Data = Data
Else
Data = Data + 1
End If

If ComboBox_empscl.value = "" Or tv_EmpAddress.value = "" Or ComboBox_empdeg.value = "" Or tv_empContact.value = "" Or tv_emGPF.value = "" Or tv_empDistrict.value = "" Or tv_empCader.value = "" Or tv_empDob.value = "" Or tv_empName.value = "" Or tv_empShalarthId.value = "" Or tv_empDoR.value = "" Then
MsgBox "Please enter details in all fields", vbOKOnly + vbExclamation, ""
Else
With Me
Sheets("Employeed_details").Cells(Data, "A").value = tv_empSrNo.value
Sheets("Employeed_details").Cells(Data, "B").value = ComboBox_empscl.value
Sheets("Employeed_details").Cells(Data, "C").value = tv_empName.value
Sheets("Employeed_details").Cells(Data, "D").value = ComboBox_empdeg.value
Sheets("Employeed_details").Cells(Data, "E").value = tv_EmpAddress.value
Sheets("Employeed_details").Cells(Data, "F").value = tv_empDistrict.value
Sheets("Employeed_details").Cells(Data, "G").value = tv_empDob.value
Sheets("Employeed_details").Cells(Data, "H").value = tv_emGPF.value
Sheets("Employeed_details").Cells(Data, "I").value = tv_empDoR.value
Sheets("Employeed_details").Cells(Data, "J").value = tv_empCader.value
Sheets("Employeed_details").Cells(Data, "K").value = tv_empShalarthId.value
Sheets("Employeed_details").Cells(Data, "L").value = tv_empContact.value
Sheets("Employeed_details").Cells(Data, "P").value = tv_empDoJ.value
Sheets("Employeed_details").Cells(Data, "Q").value = tv_empBankName.value
Sheets("Employeed_details").Cells(Data, "R").value = tv_empBankAc.value
Sheets("Employeed_details").Cells(Data, "S").value = tv_empBranch.value
Sheets("Employeed_details").Cells(Data, "T").value = tv_empIfsc.value
Sheets("Employeed_details").Cells(Data, "M").Hyperlinks.Add Anchor:=Sheets("Employeed_details").Cells(Data, "M"), Address:="", SubAddress:= _
        "'Nominee'!A1", TextToDisplay:="NavigatetoNominee"
Sheets("Employeed_details").Cells(Data, "N").value = "=COUNTIF(tbl_Nominee[[#All],[Emp_Name]],[@[Emp_Name]])"

label_empDate = Now()
LValue = Format(Date, "yyyy/mm/dd")
Sheets("Employeed_details").Cells(Data, "O").value = LValue

End With

MsgBox "Employee Details updated Successfully", vbOKOnly + vbInformation, ""

tv_empSrNo = Application.WorksheetFunction.Max(Sheets("Employeed_details").Range("A:A")) + 1
ComboBox_empscl.value = ""
tv_EmpAddress.value = ""
ComboBox_empdeg.value = ""
tv_empContact.value = ""
tv_emGPF.value = ""
tv_empDistrict.value = ""
tv_empCader.value = ""
tv_empDob.value = ""
tv_empName.value = ""
tv_empShalarthId.value = ""
tv_empDoR.value = ""
tv_empBankName.value = ""
tv_empDoJ.value = ""
tv_empBankAc.value = ""
tv_empBranch.value = ""
tv_empIfsc.value = ""
End If
End Sub

Private Sub btn_EmpReset_Click()
ComboBox_empscl.value = ""
tv_EmpAddress.value = ""
ComboBox_empdeg.value = ""
tv_empContact.value = ""
tv_emGPF.value = ""
tv_empDistrict.value = ""
tv_empCader.value = ""
tv_empDob.value = ""
tv_empName.value = ""
tv_empShalarthId.value = ""
tv_empDoR.value = ""
tv_empBankName.value = ""
tv_empDoJ.value = ""
tv_empBankAc.value = ""
tv_empBranch.value = ""
tv_empIfsc.value = ""
End Sub

Private Sub ComboBox_empdeg_AfterUpdate()
Set tblrange = Worksheets("DesignationSheet").ListObjects("Table2").Range
tv_empCader.value = Application.WorksheetFunction.VLookup(ComboBox_empdeg.value, tblrange, 2, False)
End Sub

Private Sub backpage()
Dim ibackPage As Long
With Me.MultiPage1
    ibackPage = .value - 1
    If ibackPage < .Pages.Count Then
       .Pages(ibackPage).Visible = True
       .value = ibackPage
    End If
 End With

End Sub

Private Sub CommandButton1_Click()
Dim iNextPage As Long
With Me.MultiPage1
    iNextPage = .value + 1
    If iNextPage < .Pages.Count Then
       .Pages(iNextPage).Visible = True
       .value = iNextPage
    End If
 End With

End Sub

Private Sub CommandButton2_Click()
Call CommandButton1_Click
End Sub

Private Sub CommandButton3_Click()
Call CommandButton1_Click
End Sub

Private Sub CommandButton4_Click()
Call backpage
End Sub

Private Sub CommandButton5_Click()
Call backpage
End Sub

Private Sub CommandButton6_Click()
Call backpage
End Sub

Private Sub tv_empDob_Change()
year1 = tv_empDob
If tv_empCader = "A" Or tv_empCader = "B" Or tv_empCader = "C" Then
year2 = DateAdd("yyyy", 58, year1)
Else
If tv_empCader = "D" Then
year2 = DateAdd("yyyy", 60, year1)
End If
End If
tv_empDoR.value = year2

End Sub

Private Sub UserForm_Activate()
Application.WindowState = xlMaximized
Me.Height = Application.Height
Me.Width = Application.Width

End Sub

Private Sub UserForm_Initialize()
Me.tv_empSrNo = Application.WorksheetFunction.Max(Sheets("Employeed_details").Range("A:A")) + 1
Me.label_empDate = Now()
ComboBox_empnr.AddItem "Husband"
ComboBox_empnr.AddItem "Wife"
ComboBox_empnr.AddItem "Parents"
ComboBox_empnr.AddItem "Children"
ComboBox_empnr.AddItem "Minor Brothers"
ComboBox_empnr.AddItem "Unmarried Sisters"
ComboBox_empnr.AddItem "Deceased Son’S Widow"
ComboBox_empnr.AddItem "Deceased Son’S Children"
ComboBox_empnr.AddItem "Paternal Grandparent"
End Sub

Private Sub ComboBox_empscl_AfterUpdate()
'check to see if school value exist IFs
If WorksheetFunction.CountIf(Sheets("School_Details").Range("B:B"), Me.ComboBox_empscl.value) = 0 Then
 MsgBox "This School is not present in the list", vbCritical, ""
 Me.ComboBox_empscl.value = ""
 Exit Sub
 End If

End Sub

Private Sub Image_empDor_Click()
 dateVariable = CalendarForm.GetDate( _
        SelectedDate:=Now(), _
        FirstDayOfWeek:=Monday, _
        DateFontSize:=12, _
        TodayButton:=True, _
        OkayButton:=True, _
        ShowWeekNumbers:=True, _
        BackgroundColor:=RGB(243, 249, 251), _
        HeaderColor:=RGB(147, 205, 2221), _
        HeaderFontColor:=RGB(255, 255, 255), _
        SubHeaderColor:=RGB(223, 240, 245), _
        SubHeaderFontColor:=RGB(31, 78, 120), _
        DateColor:=RGB(243, 249, 251), _
        DateFontColor:=RGB(31, 78, 120), _
        TrailingMonthFontColor:=RGB(155, 194, 230), _
        DateHoverColor:=RGB(223, 240, 245), _
        DateSelectedColor:=RGB(202, 223, 242), _
        SaturdayFontColor:=RGB(0, 176, 240), _
        SundayFontColor:=RGB(0, 176, 240), _
        TodayFontColor:=RGB(0, 176, 80))
    If dateVariable <> 0 Then tv_empDoR = dateVariable

End Sub

Private Sub Image1_Click()
 dateVariable = CalendarForm.GetDate( _
        SelectedDate:=Now(), _
        FirstDayOfWeek:=Monday, _
        DateFontSize:=12, _
        TodayButton:=True, _
        OkayButton:=True, _
        ShowWeekNumbers:=True, _
        BackgroundColor:=RGB(243, 249, 251), _
        HeaderColor:=RGB(147, 205, 2221), _
        HeaderFontColor:=RGB(255, 255, 255), _
        SubHeaderColor:=RGB(223, 240, 245), _
        SubHeaderFontColor:=RGB(31, 78, 120), _
        DateColor:=RGB(243, 249, 251), _
        DateFontColor:=RGB(31, 78, 120), _
        TrailingMonthFontColor:=RGB(155, 194, 230), _
        DateHoverColor:=RGB(223, 240, 245), _
        DateSelectedColor:=RGB(202, 223, 242), _
        SaturdayFontColor:=RGB(0, 176, 240), _
        SundayFontColor:=RGB(0, 176, 240), _
        TodayFontColor:=RGB(0, 176, 80))
    If dateVariable <> 0 Then tv_empDob = dateVariable
End Sub

Private Sub Image_empDoJ_Click()
dateVariable = CalendarForm.GetDate( _
        SelectedDate:=Now(), _
        FirstDayOfWeek:=Monday, _
        DateFontSize:=12, _
        TodayButton:=True, _
        OkayButton:=True, _
        ShowWeekNumbers:=True, _
        BackgroundColor:=RGB(243, 249, 251), _
        HeaderColor:=RGB(147, 205, 2221), _
        HeaderFontColor:=RGB(255, 255, 255), _
        SubHeaderColor:=RGB(223, 240, 245), _
        SubHeaderFontColor:=RGB(31, 78, 120), _
        DateColor:=RGB(243, 249, 251), _
        DateFontColor:=RGB(31, 78, 120), _
        TrailingMonthFontColor:=RGB(155, 194, 230), _
        DateHoverColor:=RGB(223, 240, 245), _
        DateSelectedColor:=RGB(202, 223, 242), _
        SaturdayFontColor:=RGB(0, 176, 240), _
        SundayFontColor:=RGB(0, 176, 240), _
        TodayFontColor:=RGB(0, 176, 80))
    If dateVariable <> 0 Then tv_empDoJ = dateVariable
End Sub

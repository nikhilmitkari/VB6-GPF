VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calculation_Sheet 
   ClientHeight    =   10572
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   18924
   OleObjectBlob   =   "Calculation_Sheet.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Calculation_Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_calGenerate_Click()

'----------------------------------------------------Put the input data in the Pay slip---------------------------------------------------------------------------
With Me
Sheets("Pay_Slip").Range("A1").value = "TRUE"
Sheets("Pay_Slip").Range("K4").value = ComboBox_calEmpName.value
Sheets("Pay_Slip").Range("K5").value = ComboBox_calSclName.value
nxt_year = Right(ComboBox_calFYear + 1, 2)
Sheets("Pay_Slip").Range("N3").value = ComboBox_calFYear.value & " - " & nxt_year
Sheets("Pay_Slip").Range("O7").value = ComboBox_calDeg.value
Sheets("Pay_Slip").Range("M9").value = tv_calNonRefund.value

Sheets("Pay_Slip").Range("O10").value = tv_calSubciption.value

Sheets("Pay_Slip").Range("P9").value = tv_calNonRefundDate.value
Sheets("Pay_Slip").Range("K6").value = "=VLOOKUP(K4,Table3[[#All],[Emp_Name]:[GPF No]],3,0)"
Sheets("Pay_Slip").Range("K7").value = "=VLOOKUP(K4,Table3[[Emp_Name]:[GPF No]],6,0)"

Sheets("Pay_Slip").Range("J26").value = "=SUM(J13:J24)"
Sheets("Pay_Slip").Range("K26").value = "=SUM(K13:K24)"
Sheets("Pay_Slip").Range("L26").value = "=SUM(L13:L24)"
Sheets("Pay_Slip").Range("M26").value = "=SUM(M13:M24)"
Sheets("Pay_Slip").Range("N26").value = "=SUM(N13:N24)"
Sheets("Pay_Slip").Range("O26").value = "=SUM(O13:O24)"
Sheets("Pay_Slip").Range("P26").value = "=SUM(P13:P24)"
Sheets("Pay_Slip").Range("N31").value = "=SUM(N27:N29)"

End With

'----------------------------------------------Refund value code----------------------------------------------------------------

'refunddate = Sheets("Pay_Slip").Range("P8").value
'refunddate = Month(refunddate)
'If refunddate = "4" Then
'Sheets("Pay_Slip").Range("K13").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "5" Then
'Sheets("Pay_Slip").Range("K14").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "6" Then
'Sheets("Pay_Slip").Range("K15").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "7" Then
'Sheets("Pay_Slip").Range("K16").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "8" Then
'Sheets("Pay_Slip").Range("K17").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "9" Then
'Sheets("Pay_Slip").Range("K18").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "10" Then
'Sheets("Pay_Slip").Range("K19").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "11" Then
'Sheets("Pay_Slip").Range("K20").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "12" Then
'Sheets("Pay_Slip").Range("K21").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "1" Then
'Sheets("Pay_Slip").Range("K22").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "2" Then
'Sheets("Pay_Slip").Range("K23").value = Sheets("Pay_Slip").Range("M8").value
'Else
'If refunddate = "3" Then
'Sheets("Pay_Slip").Range("K24").value = Sheets("Pay_Slip").Range("M8").value
'End If
'End If
'End If
'End If
'End If
'End If
'End If
'End If
'End If
'End If
'End If
'End If

'*****************************************Non-Refund value code

refunddate = Sheets("Pay_Slip").Range("P9").value
refunddate = Month(refunddate)
If refunddate = "4" Then
Sheets("Pay_Slip").Range("M13").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "5" Then
Sheets("Pay_Slip").Range("M14").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "6" Then
Sheets("Pay_Slip").Range("M15").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "7" Then
Sheets("Pay_Slip").Range("M16").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "8" Then
Sheets("Pay_Slip").Range("M17").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "9" Then
Sheets("Pay_Slip").Range("M18").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "10" Then
Sheets("Pay_Slip").Range("M19").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "11" Then
Sheets("Pay_Slip").Range("M20").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "12" Then
Sheets("Pay_Slip").Range("M21").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "1" Then
Sheets("Pay_Slip").Range("M22").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "2" Then
Sheets("Pay_Slip").Range("M23").value = Sheets("Pay_Slip").Range("M9").value
Else
If refunddate = "3" Then
Sheets("Pay_Slip").Range("M24").value = Sheets("Pay_Slip").Range("M9").value
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If

' opening balance & opening 6th7thpay value population
Dim emp_name As String
Dim finalrow As Integer
Dim i As Integer

'Sheets("Nominee").Range("Emp_table[#All]").ClearContents
Sheets("calculation").Range("A2:D50").ClearContents
'emp_name = DropDown2.Value

emp_name = Sheets("Pay_Slip").Range("K4").value

finalrow = Sheets("Data").Range("B10000").End(xlUp).Row

For i = 3 To finalrow
  If Sheets("Data").Cells(i, 2) = emp_name Then
  'Application.Worksheets.Select
  
 Sheets("Data").Cells(i, 2).Copy
  Sheets("calculation").Range("A100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  
        Sheets("Data").Cells(i, 5).Copy
      Sheets("calculation").Range("B100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
     
        Sheets("Data").Cells(i, 24).Copy
      Sheets("calculation").Range("C100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
   
       Sheets("Data").Cells(i, 25).Copy
      Sheets("calculation").Range("D100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  
  End If
Next i

year1 = Sheets("Pay_Slip").Range("N3").value
year2 = Left(year1, 4) - 1 & " " & "-" & " " & Right(year1, 2) - 1

Sheets("Pay_Slip").Range("M12").value = Application.VLookup(year2, Sheet5.Range("Shodh"), 3, 0)
Sheets("Pay_Slip").Range("P12").value = Application.VLookup(year2, Sheet5.Range("Shodh"), 2, 0)

'*********************************************************************************************************************************************************************************************************************************************
Dim employee As String
Dim X As Integer
Dim Y As Integer

Sheets("calculation").Range("I2:U50").ClearContents
employee = Sheets("Pay_Slip").Range("K4").value
X = Sheets("Refund_Details").Range("C10000").End(xlUp).Row

For Y = 5 To X
  If Sheets("Refund_Details").Cells(Y, 3) = employee Then
   
      Sheets("Refund_Details").Cells(Y, 5).Copy
          Sheets("calculation").Range("I100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
     
      Sheets("Refund_Details").Cells(Y, 6).Copy
        Sheets("calculation").Range("J100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
        
      Sheets("Refund_Details").Cells(Y, 7).Copy
        Sheets("calculation").Range("K100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 8).Copy
        Sheets("calculation").Range("L100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 9).Copy
        Sheets("calculation").Range("M100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 10).Copy
        Sheets("calculation").Range("N100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 11).Copy
        Sheets("calculation").Range("O100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 12).Copy
        Sheets("calculation").Range("P100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 13).Copy
        Sheets("calculation").Range("Q100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 14).Copy
        Sheets("calculation").Range("R100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 15).Copy
        Sheets("calculation").Range("S100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 16).Copy
        Sheets("calculation").Range("T100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
      Sheets("Refund_Details").Cells(Y, 17).Copy
        Sheets("calculation").Range("U100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      
  End If
Next Y

yr1 = Sheets("Pay_Slip").Range("N3").value

Sheets("Pay_Slip").Range("K13").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 2, False)
Sheets("Pay_Slip").Range("K14").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 3, False)
Sheets("Pay_Slip").Range("K15").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 4, False)
Sheets("Pay_Slip").Range("K16").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 5, False)
Sheets("Pay_Slip").Range("K17").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 6, False)
Sheets("Pay_Slip").Range("K18").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 7, False)
Sheets("Pay_Slip").Range("K19").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 8, False)
Sheets("Pay_Slip").Range("K20").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 9, False)
Sheets("Pay_Slip").Range("K21").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 10, False)
Sheets("Pay_Slip").Range("K22").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 11, False)
Sheets("Pay_Slip").Range("K23").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 12, False)
Sheets("Pay_Slip").Range("K24").value = Application.VLookup(yr1, Sheet5.Range("I2:U20"), 13, False)

For Each cell In Worksheets("Pay_Slip").Range("K13:K24")
    If Application.WorksheetFunction.IsNA(cell) Then
        'If cell contains #N/A, then set the value to 0
        cell.value = 0
    End If
Next

' ******************************************INTEREST CALCULATION******************************************************************************************

Dim Year As String
Set tblrange = Worksheets("Interest_Rate").ListObjects("Table7").Range
On Error Resume Next
'Set range to whatever you like
Dim rng As Range
Dim rngP As Range
Set rng = Worksheets("Pay_Slip").Range("M12")
Set rngP = Worksheets("Pay_Slip").Range("P12")

'Loop all the cells in range M12
For Each cell In rng
    If Application.WorksheetFunction.IsNA(cell) Then
        'If cell contains #N/A, then set the value to 0
        cell.value = 0
    End If
Next

For Each cell In rngP
    If Application.WorksheetFunction.IsNA(cell) Then
        cell.value = 0
    End If
Next

apr = (Sheets("Pay_Slip").Range("N13").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 2, False))
may = (Sheets("Pay_Slip").Range("N14").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 3, False))
jun = (Sheets("Pay_Slip").Range("N15").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 4, False))
jul = (Sheets("Pay_Slip").Range("N16").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 5, False))
aug = (Sheets("Pay_Slip").Range("N17").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 6, False))
sep = (Sheets("Pay_Slip").Range("N18").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 7, False))
octo = (Sheets("Pay_Slip").Range("N19").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 8, False))
nov = (Sheets("Pay_Slip").Range("N20").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 9, False))
dec = (Sheets("Pay_Slip").Range("N21").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 10, False))
Year = Int(ComboBox_calFYear.value) + 1
jan = (Sheets("Pay_Slip").Range("N22").value) * (Application.VLookup(Int(Year), tblrange, 11, False))
feb = (Sheets("Pay_Slip").Range("N23").value) * (Application.VLookup(Int(Year), tblrange, 12, False))
mar = (Sheets("Pay_Slip").Range("N24").value) * (Application.VLookup(Int(Year), tblrange, 13, False))

Sheets("Pay_Slip").Range("N29") = Round((apr + may + jun + jul + aug + sep + octo + nov + dec + jan + feb + mar) / 1200, 0)
'p = Sheets("Pay_Slip").Range("N3").value
'sal = Left(p, 4)
'apr = (Sheets("Pay_Slip").Range("P13").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 2, False))
'may = (Sheets("Pay_Slip").Range("P14").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 3, False))
'jun = (Sheets("Pay_Slip").Range("P15").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 4, False))
'jul = (Sheets("Pay_Slip").Range("P16").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 5, False))
'aug = (Sheets("Pay_Slip").Range("P17").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 6, False))
'sep = (Sheets("Pay_Slip").Range("P18").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 7, False))
'octo = (Sheets("Pay_Slip").Range("P19").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 8, False))
'nov = (Sheets("Pay_Slip").Range("P20").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 9, False))
'dec = (Sheets("Pay_Slip").Range("P21").value) * (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 10, False))
'Year = Int(ComboBox_calFYear.value) + 1
'jan = (Sheets("Pay_Slip").Range("P22").value) * (Application.VLookup(Int(Year), tblrange, 11, False))
'feb = (Sheets("Pay_Slip").Range("P23").value) * (Application.VLookup(Int(Year), tblrange, 12, False))
'mar = (Sheets("Pay_Slip").Range("P24").value) * (Application.VLookup(Int(Year), tblrange, 13, False))
'
'Sheets("Pay_Slip").Range("P29") = Round((apr + may + jun + jul + aug + sep + octo + nov + dec + jan + feb + mar) / 1200, 0)

'----------------------------------------------------Display interest rate in cell I29-----------------------------------------------------------------------------
Dim Varsh As String
Varsh = Int(ComboBox_calFYear.value) + 1
Set tblrange = Worksheets("Interest_Rate").ListObjects("Table7").Range
'ComboBox_calFYear.value & " - " & nxt_year
Sheets("Pay_Slip").Range("I29").value = "                                Interest Rates for                                " _
                                         & "Apr - " & (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 2, False)) & "% " _
                                         & "May - " & (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 3, False)) & "% " _
                                         & "Jun - " & (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 4, False)) & "% " _
                                         & "Jul - " & (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 5, False)) & "% " _
                                         & "Aug - " & (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 6, False)) & "% " _
                                         & "Sep - " & (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 7, False)) & "% " _
                                         & "Oct - " & (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 8, False)) & "% " _
                                         & "Nov - " & (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 9, False)) & "% " _
                                         & "Dec - " & (Application.VLookup(Int(ComboBox_calFYear.value), tblrange, 10, False)) & "% " _
                                         & "Jan - " & (Application.VLookup(Int(Varsh), tblrange, 11, False)) & "% " _
                                         & "Feb - " & (Application.VLookup(Int(Varsh), tblrange, 12, False)) & "% " _
                                         & "Mar - " & (Application.VLookup(Int(Varsh), tblrange, 13, False)) & "% "

'clearing input fields values on userform
ComboBox_calEmpName.value = ""
ComboBox_calSclName.value = ""
ComboBox_calFYear.value = ""
ComboBox_calDeg.value = ""

tv_calNonRefund.value = ""

tv_calSubciption.value = ""

tv_calNonRefundDate.value = ""
Sheets("Pay_Slip").Range("J45").value = ""

Application.Visible = True
Call Sheets("Pay_Slip").Select
ActiveSheet.Range("I1").Select
Calculation_Sheet.Hide

End Sub

Private Sub btn_calHome_Click()
Unload Me
MainWindow.Show
End Sub

Private Sub btn_calReset_Click()

ComboBox_calEmpName.value = ""
ComboBox_calSclName.value = ""
ComboBox_calFYear.value = ""
ComboBox_calDeg.value = ""
tv_calNonRefund.value = ""
tv_calSubciption.value = ""
tv_calNonRefundDate.value = ""

End Sub

Private Sub ComboBox_calEmpName_AfterUpdate()

On Error Resume Next

Dim Employee_name As String
Dim finalrow As Integer
Dim j As Integer

ComboBox_calDeg.Clear

Sheets("Employeed_details").Range("AB4:AC100").ClearContents

Employee_name = ComboBox_calEmpName.value

' Employee_name =Vicky Mitkari


finalrow = Sheets("Employeed_details").Range("C1000").End(xlUp).Row

For j = 8 To finalrow
  If Sheets("Employeed_details").Cells(j, 3) = Employee_name Then
  Sheets("Employeed_details").Range(Cells(j, 3), Cells(j, 4)).Copy
  Sheets("Employeed_details").Range("AB100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  End If
Next j

For Each acell In Range("DegName")
    If acell.value <> "" Then
     Me.ComboBox_calDeg.AddItem acell.value
    End If
 Next

End Sub

Private Sub ComboBox_calSclName_AfterUpdate()

On Error Resume Next



Dim scl_name As String
Dim finalrow As Integer
Dim i As Integer

ComboBox_calEmpName.Clear

Sheets("Employeed_details").Range("Y4:Z100").ClearContents
scl_name = ComboBox_calSclName.value

finalrow = Sheets("Employeed_details").Range("B1000").End(xlUp).Row

For i = 8 To finalrow
  If Sheets("Employeed_details").Cells(i, 2) = scl_name Then
  Sheets("Employeed_details").Range(Cells(i, 2), Cells(i, 3)).Copy
  Sheets("Employeed_details").Range("Y100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  
  End If
Next i



For Each acell In Range("sclname")
    If acell.value <> "" Then
     Me.ComboBox_calEmpName.AddItem acell.value
    End If
 Next

'Me.ComboBox_calEmpName.RowSource = "sclname"
End Sub

Private Sub Image_calNonRefund_Click()
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
    If dateVariable <> 0 Then tv_calNonRefundDate = dateVariable
End Sub

Private Sub UserForm_Initialize()

Application.WindowState = xlMaximized
Me.Height = Application.Height
Me.Width = Application.Width


 Sheets("Pay_Slip").Range("N3") = ""
 Sheets("Pay_Slip").Range("K4") = ""
 Sheets("Pay_Slip").Range("K5") = ""
 Sheets("Pay_Slip").Range("K6") = ""
 Sheets("Pay_Slip").Range("K7") = ""
 Sheets("Pay_Slip").Range("O7") = ""
 Sheets("Pay_Slip").Range("P8") = ""
 Sheets("Pay_Slip").Range("M9") = ""
 Sheets("Pay_Slip").Range("O10") = ""
 Sheets("Pay_Slip").Range("M12") = ""
 Sheets("Pay_Slip").Range("P12") = ""
 Sheets("Pay_Slip").Range("P9") = ""
 Sheets("Pay_Slip").Range("K13:K24") = ""
 Sheets("Pay_Slip").Range("M13:M24") = ""
 
End Sub



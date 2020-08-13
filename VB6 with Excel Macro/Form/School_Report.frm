VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} School_Report 
   Caption         =   "School Annual Report"
   ClientHeight    =   3984
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   6276
   OleObjectBlob   =   "School_Report.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "School_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_Scl_Change()

On Error Resume Next

Dim scl_name As String
Dim finalrow As Integer
Dim i As Integer

cmb_Year.Clear

Sheets("School_Data").Range("Z2:Z100").ClearContents

scl_name = cmb_Scl.value

finalrow = Sheets("Data").Range("C1000").End(xlUp).Row

For i = 3 To finalrow
  If Sheet7.Cells(i, 3) = scl_name Then
  Sheets("Data").Cells(i, 5).Copy
  Sheets("School_Data").Range("Z100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
  
  
  End If
Next i



For Each acell In Range("syear")
    If acell.value <> "" Then
     Me.cmb_Year.AddItem acell.value
    End If
 Next

End Sub

Private Sub Generate_Click()
Dim scl_name As String
Dim Year As String
Dim X As Integer
Dim Y As Integer

scl_name = cmb_Scl.value
Year = cmb_Year.value
X = Sheets("Data").Range("C10000").End(xlUp).Row

For Y = 3 To X
  If Sheet7.Cells(Y, 3) = scl_name Then
   If Sheet7.Cells(Y, 5) = Year Then
   
      Sheets("Data").Cells(Y, 3).Copy
         Sheets("School_Data").Range("A100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
     
      Sheets("Data").Cells(Y, 5).Copy
       Sheets("School_Data").Range("B100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
        'Sheets("School Report").OLEObjects("ComboBox2").Object.AddItem = Sheets("Data").Cells(y, 5)
        
      Sheets("Data").Cells(Y, 14).Copy
       Sheets("School_Data").Range("C100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      opBalance = opBalance + Sheets("Data").Cells(Y, 14)
      
      Sheets("Data").Cells(Y, 24).Copy
       Sheets("School_Data").Range("P100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
       interest = interest + Sheets("Data").Cells(Y, 24)
       
      Sheets("Data").Cells(Y, 26).Copy
       Sheets("School_Data").Range("D100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
       apr = apr + Sheet7.Cells(Y, 26).value
       
      Sheets("Data").Cells(Y, 27).Copy
       Sheets("School_Data").Range("E100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
       may = may + Sheets("Data").Cells(Y, 27)
       
      Sheets("Data").Cells(Y, 28).Copy
       Sheets("School_Data").Range("F100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      jun = jun + Sheets("Data").Cells(Y, 28)
    
      Sheets("Data").Cells(Y, 29).Copy
       Sheets("School_Data").Range("G100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      jul = jul + Sheets("Data").Cells(Y, 29)
      
      Sheets("Data").Cells(Y, 30).Copy
       Sheets("School_Data").Range("H100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      aug = aug + Sheets("Data").Cells(Y, 30)
      
      Sheets("Data").Cells(Y, 31).Copy
       Sheets("School_Data").Range("I100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      sep = sep + Sheets("Data").Cells(Y, 31)
      
      Sheets("Data").Cells(Y, 32).Copy
       Sheets("School_Data").Range("J100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
       octo = octo + Sheets("Data").Cells(Y, 32)
       
      Sheets("Data").Cells(Y, 33).Copy
       Sheets("School_Data").Range("K100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      nov = nov + Sheets("Data").Cells(Y, 33)
      
      Sheets("Data").Cells(Y, 34).Copy
       Sheets("School_Data").Range("L100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      dec = dec + Sheets("Data").Cells(Y, 34)
      
      Sheets("Data").Cells(Y, 35).Copy
       Sheets("School_Data").Range("M100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
       jan = jan + Sheets("Data").Cells(Y, 35)
       
      Sheets("Data").Cells(Y, 36).Copy
       Sheets("School_Data").Range("N100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
       feb = feb + Sheets("Data").Cells(Y, 36)
       
      Sheets("Data").Cells(Y, 37).Copy
       Sheets("School_Data").Range("O100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
      mar = mar + Sheets("Data").Cells(Y, 37)
      
      withdrawals = withdrawals + Sheets("Data").Cells(Y, 19)
      
      End If
  End If
Next Y


Sheet12.Range("L4").value = Year
yr1 = Sheet12.Range("L4")
Sheet12.Range("K5") = scl_name
Sheet12.Range("K6").value = "=VLOOKUP(K5,Table1[[School_name]:[HM_NAME]],5,0)"
Sheet12.Range("K7").value = "=VLOOKUP(K5,Table1[[School_name]:[Address]],2,0)"
Sheet12.Range("K8").value = "=VLOOKUP(K5,Table1[[School_name]:[PanchayatSamiti]],8,0)"
Sheet12.Range("N8").value = "=VLOOKUP(K5,Table1[[School_name]:[District]],3,0)"
Sheet12.Range("K9").value = "=VLOOKUP(K5,Table1[[School_name]:[PayUnit No]],4,0)"
Sheet12.Range("N9").value = "=VLOOKUP(K5,Table1[[School_name]:[Contact_No]],6,0)"
Sheet12.Range("J12").value = apr
Sheet12.Range("J13").value = may
Sheet12.Range("J14").value = jun
Sheet12.Range("J15").value = jul
Sheet12.Range("J16").value = aug
Sheet12.Range("J17").value = sep
Sheet12.Range("N12").value = octo
Sheet12.Range("N13").value = nov
Sheet12.Range("N14").value = dec
Sheet12.Range("N15").value = jan
Sheet12.Range("N16").value = feb
Sheet12.Range("N17").value = mar
Sheet12.Range("M18").value = opBalance
Sheet12.Range("M20").value = interest
Sheet12.Range("M22").value = withdrawals

For Each cell In Worksheets("School Report").Range("J12:N23")
    If Application.WorksheetFunction.IsNA(cell) Then
        'If cell contains #N/A, then set the value to 0
        cell.value = 0
    End If
Next

Application.Visible = True
Call Sheets("School Report").Activate
School_Report.Hide
MainWindow.Hide

End Sub

Private Sub UserForm_Initialize()
Call ProtectAllWorksheets
Sheets("School_Data").Range("Z2:Z100").ClearContents
Sheet14.Range("A3:Q100").Clear

'*********************************Clear All field in report********************************************************************************
Sheet12.Range("L4") = ""
Sheet12.Range("K5") = ""
Sheet12.Range("K6").value = ""
Sheet12.Range("K7").value = ""
Sheet12.Range("K8").value = ""
Sheet12.Range("N8").value = ""
Sheet12.Range("K9").value = ""
Sheet12.Range("N9").value = ""
Sheet12.Range("J12").value = ""
Sheet12.Range("J13").value = ""
Sheet12.Range("J14").value = ""
Sheet12.Range("J15").value = ""
Sheet12.Range("J16").value = ""
Sheet12.Range("J17").value = ""
Sheet12.Range("N12").value = ""
Sheet12.Range("N13").value = ""
Sheet12.Range("N14").value = ""
Sheet12.Range("N15").value = ""
Sheet12.Range("N16").value = ""
Sheet12.Range("N17").value = ""
Sheet12.Range("M18").value = ""
Sheet12.Range("M20").value = ""
Sheet12.Range("M22").value = ""

End Sub


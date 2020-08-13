Attribute VB_Name = "Module1"
Private Sub refund()

refunddate = "19/04/2019" 'Sheets("Pay_Slip").Range("P8").value
refunddate = Month(refunddate)
If refunddate = "4" Then
Sheets("Pay_Slip").Range("K13").value = Sheets("Pay_Slip").Range("M8").value
End If
End Sub

Sub SclReport_Click()

Dim scl_name As String
Dim Year As String
Dim X As Integer
Dim Y As Integer

scl_name = Worksheets("School Report").OLEObjects("ComboBox1").Object.value

Year = Worksheets("School Report").OLEObjects("ComboBox2").Object.value
X = Sheets("Data").Range("C10000").End(xlUp).Row

For Y = 3 To X
  If Sheet7.Cells(Y, 3) = scl_name Then
   If Sheet7.Cells(Y, 5) = Year Then
   
      Sheets("Data").Cells(Y, 3).Copy
         Sheets("School_Data").Range("A100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
     
      Sheets("Data").Cells(Y, 5).Copy
       Sheets("School_Data").Range("B100").End(xlUp).Offset(1, 0).PasteSpecial xlPasteFormulasAndNumberFormats
        Sheets("School Report").OLEObjects("ComboBox2").Object.AddItem = Sheets("Data").Cells(Y, 5)
        
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


box = Sheet1.Range("Table1")
'yr1 = Sheets("School Report").Range("N3")
'
'Sheets("School Report").Range("K6").value = Application.VLookup(scl_name, box, 3, False)
'Sheets("School Report").Range("K7").value = Application.VLookup(scl_name, box, 6, False)
'Sheets("School Report").Range("K8").value = Application.VLookup(yr1, box, 9, False)
'Sheets("School Report").Range("N8").value = Application.VLookup(yr1, box, 4, False)
'Sheets("School Report").Range("K9").value = Application.VLookup(yr1, box, 5, False)
'Sheets("School Report").Range("N9").value = Application.VLookup(yr1, box, 7, False)
'Sheets("School Report").Range("J12").value = apr.value
'Sheets("School Report").Range("J13").value = may.value
'Sheets("School Report").Range("J14").value = jun.value
'Sheets("School Report").Range("J15").value = jul.value
'Sheets("School Report").Range("J16").value = aug.value
'Sheets("School Report").Range("J17").value = sep.value
'Sheets("School Report").Range("N12").value = octo.value
'Sheets("School Report").Range("N13").value = nov.value
'Sheets("School Report").Range("N14").value = dec.value
'Sheets("School Report").Range("N15").value = jan.value
'Sheets("School Report").Range("N16").value = feb.value
'Sheets("School Report").Range("N17").value = mar.value
'Sheets("School Report").Range("M18").value = opBalance.value
'Sheets("School Report").Range("M20").value = interest.value
'Sheets("School Report").Range("M22").value = withdrawals.value
'
'For Each cell In Worksheets("School Report").Range("J12:N23")
'    If Application.WorksheetFunction.IsNA(cell) Then
'        'If cell contains #N/A, then set the value to 0
'        cell.value = 0
'    End If
'Next



End Sub


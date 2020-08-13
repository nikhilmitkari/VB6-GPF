Attribute VB_Name = "Module10"
Sub SavePaySlip_Click()

If Sheets("Pay_Slip").Range("A1").value = "" Then
MsgBox "Already Updated!!!"
Else

Data = Sheets("Data").Cells(Rows.Count, "A").End(xlUp).Row

If Data = 3 And Sheets("Data").Cells(Data, "A").value = "" Then
Data = Data
Else
Data = Data + 1
End If

Sheets("Data").Cells(Data, "A").value = Application.WorksheetFunction.Max(Sheets("Data").Range("A:A")) + 1
Sheets("Data").Cells(Data, "B").value = Sheets("Pay_Slip").Range("K4").value
Sheets("Data").Cells(Data, "C").value = Sheets("Pay_Slip").Range("K5").value
Sheets("Data").Cells(Data, "D").value = Sheets("Pay_Slip").Range("K7").value
Sheets("Data").Cells(Data, "E").value = Sheets("Pay_Slip").Range("N3").value
Sheets("Data").Cells(Data, "F").value = Sheets("Pay_Slip").Range("K6").value
Sheets("Data").Cells(Data, "G").value = Sheets("Pay_Slip").Range("O7").value
Sheets("Data").Cells(Data, "H").value = Sheets("Pay_Slip").Range("M8").value
Sheets("Data").Cells(Data, "I").value = Sheets("Pay_Slip").Range("P8").value
Sheets("Data").Cells(Data, "J").value = Sheets("Pay_Slip").Range("M9").value
Sheets("Data").Cells(Data, "K").value = Sheets("Pay_Slip").Range("P9").value
Sheets("Data").Cells(Data, "L").value = Sheets("Pay_Slip").Range("K10").value
Sheets("Data").Cells(Data, "M").value = Sheets("Pay_Slip").Range("O10").value
Sheets("Data").Cells(Data, "N").value = Sheets("Pay_Slip").Range("M12").value
Sheets("Data").Cells(Data, "O").value = Sheets("Pay_Slip").Range("P12").value
Sheets("Data").Cells(Data, "P").value = Sheets("Pay_Slip").Range("J26").value
Sheets("Data").Cells(Data, "Q").value = Sheets("Pay_Slip").Range("K26").value
Sheets("Data").Cells(Data, "R").value = Sheets("Pay_Slip").Range("L26").value
Sheets("Data").Cells(Data, "S").value = Sheets("Pay_Slip").Range("M26").value
Sheets("Data").Cells(Data, "T").value = Sheets("Pay_Slip").Range("N26").value
Sheets("Data").Cells(Data, "U").value = Sheets("Pay_Slip").Range("N29").value
Sheets("Data").Cells(Data, "V").value = Sheets("Pay_Slip").Range("P29").value
Sheets("Data").Cells(Data, "W").value = Sheets("Pay_Slip").Range("N33").value
Sheets("Data").Cells(Data, "X").value = Sheets("Pay_Slip").Range("P33").value
Sheets("Data").Cells(Data, "Y").value = Sheets("Pay_Slip").Range("N34").value
Sheets("Data").Cells(Data, "Z").value = Sheets("Pay_Slip").Range("J13").value
Sheets("Data").Cells(Data, "AA").value = Sheets("Pay_Slip").Range("J14").value
Sheets("Data").Cells(Data, "AB").value = Sheets("Pay_Slip").Range("J15").value
Sheets("Data").Cells(Data, "AC").value = Sheets("Pay_Slip").Range("J16").value
Sheets("Data").Cells(Data, "AD").value = Sheets("Pay_Slip").Range("J17").value
Sheets("Data").Cells(Data, "AE").value = Sheets("Pay_Slip").Range("J18").value
Sheets("Data").Cells(Data, "AF").value = Sheets("Pay_Slip").Range("J19").value
Sheets("Data").Cells(Data, "AG").value = Sheets("Pay_Slip").Range("J20").value
Sheets("Data").Cells(Data, "AH").value = Sheets("Pay_Slip").Range("J21").value
Sheets("Data").Cells(Data, "AI").value = Sheets("Pay_Slip").Range("J22").value
Sheets("Data").Cells(Data, "AJ").value = Sheets("Pay_Slip").Range("J23").value
Sheets("Data").Cells(Data, "AK").value = Sheets("Pay_Slip").Range("J24").value
Sheets("Pay_Slip").Range("A1").value = ""

MsgBox "Pay Slip Data Saved Successfully"

End If

End Sub

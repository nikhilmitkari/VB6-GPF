Attribute VB_Name = "Module11"
Private Sub interestrate()
Dim Year As String
Set tblrange = Worksheets("Interest_Rate").ListObjects("Table7").Range
apr = (Sheets("Pay_Slip").Range("N13").value) * (Application.VLookup(2019, tblrange, 2, False))
MsgBox apr
may = (Sheets("Pay_Slip").Range("N14").value) * (Application.VLookup("2019", tblrange, 2, False))
jun = (Sheets("Pay_Slip").Range("N15").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 3, False))
jul = (Sheets("Pay_Slip").Range("N16").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 4, False))
aug = (Sheets("Pay_Slip").Range("N17").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 5, False))
sep = (Sheets("Pay_Slip").Range("N18").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 6, False))
octo = (Sheets("Pay_Slip").Range("N19").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 7, False))
nov = (Sheets("Pay_Slip").Range("N20").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 8, False))
dec = (Sheets("Pay_Slip").Range("N21").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 9, False))
Year = "2019" + 1
jan = (Sheets("Pay_Slip").Range("N22").value) * (Application.WorksheetFunction.VLookup(Year, TbleRange, 10, False))
feb = (Sheets("Pay_Slip").Range("N23").value) * (Application.WorksheetFunction.VLookup(Year, TbleRange, 11, False))
mar = (Sheets("Pay_Slip").Range("N24").value) * (Application.WorksheetFunction.VLookup(Year, TbleRange, 12, False))

Sheets("Pay_Slip").Range("N29") = "=Round((APR+MAY+JUN+JUL+AUG+SEP+OCTO+NOV+DEC+JAN+FEB+MAR)/1200,0)"

apr = (Sheets("Pay_Slip").Range("P13").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 1, False))
may = (Sheets("Pay_Slip").Range("P14").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 2, False))
jun = (Sheets("Pay_Slip").Range("P15").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 3, False))
jul = (Sheets("Pay_Slip").Range("P16").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 4, False))
aug = (Sheets("Pay_Slip").Range("P17").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 5, False))
sep = (Sheets("Pay_Slip").Range("P18").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 6, False))
octo = (Sheets("Pay_Slip").Range("P19").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 7, False))
nov = (Sheets("Pay_Slip").Range("P20").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 8, False))
dec = (Sheets("Pay_Slip").Range("P21").value) * (Application.WorksheetFunction.VLookup("2019", TbleRange, 9, False))
Year = "2019" + 1
jan = (Sheets("Pay_Slip").Range("P22").value) * (Application.WorksheetFunction.VLookup(Year, TbleRange, 10, False))
feb = (Sheets("Pay_Slip").Range("P23").value) * (Application.WorksheetFunction.VLookup(Year, TbleRange, 11, False))
mar = (Sheets("Pay_Slip").Range("P24").value) * (Application.WorksheetFunction.VLookup(Year, TbleRange, 12, False))

Sheets("Pay_Slip").Range("P29") = "=Round((APR+MAY+JUN+JUL+AUG+SEP+OCTO+NOV+DEC+JAN+FEB+MAR)/1200,0)"
End Sub



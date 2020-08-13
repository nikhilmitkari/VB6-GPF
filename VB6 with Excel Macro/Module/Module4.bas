Attribute VB_Name = "Module4"
Sub Button1_Click()
Application.Visible = False
MainWindow.Show
End Sub

Sub School()
RowSource = "School_name"
End Sub


Function pay_interest(ByVal p, opbal)

Dim Year As String
Set tblrange = Worksheets("Interest_Rate").ListObjects("Table7").Range
On Error Resume Next

Dim rngP As Range
Set rngP = Worksheets("Pay_Slip").Range("P12")


For Each cell In rngP
    If Application.WorksheetFunction.IsNA(cell) Then
        cell.value = 0
    End If
Next

opbal = Sheets("Pay_Slip").Range("P12").value
sal = Left(p, 4)
apr = (Sheets("Pay_Slip").Range("P13").value) * (Application.VLookup(Int(sal), tblrange, 2, False))
may = (Sheets("Pay_Slip").Range("P14").value) * (Application.VLookup(Int(sal), tblrange, 3, False))
jun = (Sheets("Pay_Slip").Range("P15").value) * (Application.VLookup(Int(sal), tblrange, 4, False))
jul = (Sheets("Pay_Slip").Range("P16").value) * (Application.VLookup(Int(sal), tblrange, 5, False))
aug = (Sheets("Pay_Slip").Range("P17").value) * (Application.VLookup(Int(sal), tblrange, 6, False))
sep = (Sheets("Pay_Slip").Range("P18").value) * (Application.VLookup(Int(sal), tblrange, 7, False))
octo = (Sheets("Pay_Slip").Range("P19").value) * (Application.VLookup(Int(sal), tblrange, 8, False))
nov = (Sheets("Pay_Slip").Range("P20").value) * (Application.VLookup(Int(sal), tblrange, 9, False))
dec = (Sheets("Pay_Slip").Range("P21").value) * (Application.VLookup(Int(sal), tblrange, 10, False))
Year = Int(sal) + 1
jan = (Sheets("Pay_Slip").Range("P22").value) * (Application.VLookup(Int(Year), tblrange, 11, False))
feb = (Sheets("Pay_Slip").Range("P23").value) * (Application.VLookup(Int(Year), tblrange, 12, False))
mar = (Sheets("Pay_Slip").Range("P24").value) * (Application.VLookup(Int(Year), tblrange, 13, False))

pay_interest = Round((apr + may + jun + jul + aug + sep + octo + nov + dec + jan + feb + mar) / 1200, 0)

End Function


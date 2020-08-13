VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainWindow 
   ClientHeight    =   13584
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   22560
   OleObjectBlob   =   "MainWindow.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_Aemp_Click()
Unload Me
Call UnprotectAllWorksheets
Add_Teacher.Show
End Sub

Private Sub btn_Ascl_Click()
Unload Me
Add_Scl.Show
End Sub

Private Sub btn_logout_Click()
Application.DisplayAlerts = False
Application.ThisWorkbook.Save
Call ProtectAllWorksheets
Application.Workbooks.Close

End Sub



Private Sub btn_ReportScl_Click()
Set ws = ThisWorkbook.Sheets("Pay_Slip")
Unload Me
Call disable
Call UnprotectAllWorksheets
Calculation_Sheet.Show
End Sub

Private Sub btn_updateRefund_Click()
Unload Me
Refund_Update.Show

End Sub

Private Sub btn_Vdeg_Click()
Application.Visible = True
Call disable
Call ProtectAllWorksheets
Call Sheets("DesignationSheet").Activate
MainWindow.Hide
End Sub

Private Sub btn_Vemp_Click()
Application.Visible = True
Call disable
Call ProtectAllWorksheets
Call openemp
MainWindow.Hide
End Sub
Private Sub openemp()
Sheets("Employeed_details").Select
ActiveSheet.Range("A1").Select
End Sub

Private Sub btn_VinterestRate_Click()
Application.Visible = True
Call disable
Call ProtectAllWorksheets
Call Sheets("Interest_Rate").Activate
MainWindow.Hide
End Sub

Private Sub btn_Vnominee_Click()
Application.Visible = True
Call disable
Call ProtectAllWorksheets
Call Sheets("Nominee").Activate
MainWindow.Hide
End Sub

Private Sub btn_Vscl_Click()
Application.Visible = True
Call disable
Call ProtectAllWorksheets
Call openscl
MainWindow.Hide
End Sub
Private Sub openscl()
Sheets("School_Details").Activate
Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
End Sub

Private Sub Image2_Click()
Application.Visible = True
School_Report.Show
End Sub

Private Sub Label4_Click()
Set ws = ThisWorkbook.Sheets("Pay_Slip")
Unload Me
'ws.Unprotect Password:="Gpf@vba", UserInterfaceonly:=True
Calculation_Sheet.Show
End Sub

Private Sub UserForm_Activate()

Application.WindowState = xlMaximized
Zoom = Int(Application.Width / Me.Width * 100)
Me.Height = Application.Height
Me.Width = Application.Width

End Sub



Private Sub UserForm_Initialize()
Call UnprotectAllWorksheets
End Sub

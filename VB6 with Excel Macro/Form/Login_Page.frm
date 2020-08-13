VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login_Page 
   Caption         =   " Copyright © 2020 "
   ClientHeight    =   11580
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   22560
   OleObjectBlob   =   "Login_Page.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Login_Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnLogin_Click()
Application.Visible = False
Dim UserId, Password As String
UserId = "admin"
Password = "admin"

If (UserId = userid_txt.value And Password = pass_txt.value) Then

 Unload Me
MainWindow.Show

 Else
 MsgBox " LogIn failed !", vbCritical, ""
 End If

End Sub


Private Sub UserForm_Activate()

Application.WindowState = xlMaximized
Me.Height = Application.Height
Me.Width = Application.Width

End Sub


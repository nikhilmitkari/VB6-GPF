Attribute VB_Name = "Module7"
Sub RoundedRectangle3_Click()
Application.ScreenUpdating = False
ActiveWindow.Zoom = 100
Sheet1.Select
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 3")).Visible = msoFalse
Application.ScreenUpdating = True

End Sub
Sub RoundedRectangle7_Click()
Application.ScreenUpdating = False
ActiveWindow.Zoom = 100
Sheet2.Select
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 2")).Visible = msoFalse
Application.ScreenUpdating = True

End Sub


Sub Nav_Interest_Rates()

Application.ScreenUpdating = False
ActiveWindow.Zoom = 100
Sheet6.Select
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 11")).Visible = msoFalse
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 12")).Visible = msoTrue
Application.ScreenUpdating = True

End Sub


Sub save_data()
'
' save_data Macro
'

'
    ActiveWorkbook.Save
    ActiveWindow.SmallScroll down:=-3
    
    
Application.ScreenUpdating = False
ActiveWindow.Zoom = 100
Sheet6.Select
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 12")).Visible = msoFalse
Application.ScreenUpdating = True

End Sub


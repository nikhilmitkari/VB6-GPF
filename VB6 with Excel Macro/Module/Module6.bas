Attribute VB_Name = "Module6"
Sub Print_Sheet()
Application.ScreenUpdating = False
ActiveWindow.Zoom = 100
Sheet9.Select
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 7")).Visible = msoFalse
Application.ScreenUpdating = True
'------------------------------------------------------------------------------

                               Selection.PrintOut

'---------------------------------------------------------------------------------

End Sub


Sub disable()
Application.DisplayFormulaBar = False
Application.DisplayScrollBars = False
Application.DisplayStatusBar = False
ActiveWindow.DisplayFormulas = False
ActiveWindow.DisplayGridlines = False
ActiveWindow.DisplayHeadings = False
End Sub

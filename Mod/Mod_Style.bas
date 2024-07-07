Attribute VB_Name = "Mod_Style"
Option Explicit
Public Sub AutoFit_()
Call Init
ws.UsedRange.Columns.AutoFit
End Sub
Public Sub Aling_()
Call Init
ws.UsedRange.HorizontalAlignment = xlLeft
End Sub
Sub BackgroundColor_(Optional Color As Variant = xlColorIndexNone)
 ws.Cells.Interior.ColorIndex = Color
End Sub

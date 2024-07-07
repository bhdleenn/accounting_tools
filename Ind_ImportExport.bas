Attribute VB_Name = "Ind_ImportExport"
Option Explicit
Private chemin As String
Public tw As Workbook
Public Sub exportModule(chemi1 As String, Name As String, Optional export_ As Integer = 1)
On Error Resume Next
chemin = "G:\macro\VBA\" & chemi1 & "\" & Name & ".bas"
Set tw = ThisWorkbook 'Workbooks("TE.xlam")
If export_ = 1 Then tw.VBProject.VBComponents(Name).export chemin
If export_ = 2 Then tw.VBProject.VBComponents.Import chemin
If export_ = 3 Then tw.VBProject.VBComponents.Remove tw.VBProject.VBComponents(Name)
End Sub
Sub save(i As Integer, Optional j As String = "")

Call exportModule("TE_app\Mod", "Mod_CalibrageTab" & j, i)
Call exportModule("TE_app\Mod", "Mod_Compta" & j, i)
Call exportModule("TE_app\Mod", "Mod_SheetConvert" & j, i)
Call exportModule("TE_app\Mod", "Mod_StringConvert" & j, i)
Call exportModule("TE_app\Mod", "Mod_Style" & j, i)

Call exportModule("TE_app", "Index" & j, i)
Call exportModule("TE_app", "Ind_errors" & j, i)
Call exportModule("TE_app", "Ind_Buttons" & j, i)

Call exportModule("TE_app\App", "App_Vente" & j, i)
Call exportModule("TE_app\App", "App_Paie" & j, i)
End Sub
Sub But_export()
Call save(1)
End Sub
Sub But_Import()
Call save(2)
End Sub
Sub But_Remove()
Call save(3)
End Sub



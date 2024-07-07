Attribute VB_Name = "Ind_Buttons"
Option Explicit
Private S1 As Range
Private x As Range
Sub But_DuplicateSheet()
Call DuplicateSheet_
End Sub
Sub But_Replace()
Call replaceName_
End Sub
Sub But_compta()
Set S1 = Selection
Call FindValue_(S1.Address)
If S1.Columns.Count > 1 Or S1.Areas.Count > 1 Then Call Error_(121)
Call normalize_Compte_(S1.Rows(1).Row, S1.Columns(1).Column, S1.Rows.Count + S1.Rows(1).Row)
End Sub
Sub But_UCase()
Set S1 = Selection
For Each x In Range(S1.Address)
x.Value = UCase(x.Value)
Next
End Sub
Sub But_LCase()
Set S1 = Selection
For Each x In Range(S1.Address)
x.Value = LCase(x.Value)
Next
End Sub
Sub But_Proper_Case()
Set S1 = Selection
For Each x In Range(S1.Address)
x.Value = Application.Proper(x.Value)
'x.Value = LCase(Left(x.Value, 1)) & UCase(Right(x.Value, Len(x.Value) - 1))
Next
End Sub
Sub selectin()
Set S1 = Selection
MsgBox S1.Address
End Sub


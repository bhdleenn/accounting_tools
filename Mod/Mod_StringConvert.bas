Attribute VB_Name = "Mod_StringConvert"
Option Explicit
Public Sub replaceName_(Optional selected As Boolean = True)
Call Init

Dim selected2 As Variant
If selected = False Then Set selected2 = ws.UsedRange Else Set selected2 = Selection

With selected2
    .Replace What:="à", Replacement:="a", LookAt:=xlPart
    .Replace What:="a", Replacement:="a", LookAt:=xlPart
    .Replace What:="@", Replacement:="a", LookAt:=xlPart
    
    .Replace What:="é", Replacement:="e", LookAt:=xlPart
    .Replace What:="ê", Replacement:="e", LookAt:=xlPart
    .Replace What:="è", Replacement:="e", LookAt:=xlPart
    .Replace What:="è", Replacement:="e", LookAt:=xlPart
    .Replace What:="ë", Replacement:="e", LookAt:=xlPart
    .Replace What:="ê", Replacement:="e", LookAt:=xlPart
    .Replace What:="Ã©", Replacement:="e", LookAt:=xlPart
    .Replace What:="œ", Replacement:="e", LookAt:=xlPart
    .Replace What:="ô", Replacement:="o", LookAt:=xlPart
    .Replace What:="î", Replacement:="i", LookAt:=xlPart
    .Replace What:="ö", Replacement:="o", LookAt:=xlPart
    .Replace What:="ï", Replacement:="i", LookAt:=xlPart
    .Replace What:="û", Replacement:="u", LookAt:=xlPart
    .Replace What:="ñ", Replacement:="n", LookAt:=xlPart
    
    .Replace What:="ç", Replacement:="c", LookAt:=xlPart
    
    .Replace What:="Mr Mme", Replacement:="", LookAt:=xlPart
    .Replace What:="Mme Mme", Replacement:="", LookAt:=xlPart
    .Replace What:="Mr Mr", Replacement:="", LookAt:=xlPart
    .Replace What:="M M", Replacement:="", LookAt:=xlPart
    .Replace What:="M MME M MME", Replacement:="", LookAt:=xlPart
    .Replace What:="ME ME", Replacement:="", LookAt:=xlPart
    
    .Replace What:="!", Replacement:="", LookAt:=xlPart
    .Replace What:="-", Replacement:="", LookAt:=xlPart
    .Replace What:="&", Replacement:="", LookAt:=xlPart
    .Replace What:="ø", Replacement:="", LookAt:=xlPart
    .Replace What:="°", Replacement:="", LookAt:=xlPart

    .Replace What:="'", Replacement:="", LookAt:=xlPart, MatchCase:=False, _
            FormulaVersion:=xlReplaceFormula2
    .Replace What:="Mr  Mme Mr  Mme", Replacement:="", LookAt:=xlPart, _
            MatchCase:=False, FormulaVersion:=xlReplaceFormula2
End With

' SearchFormat:=False, ReplaceFormat:=False,
End Sub
Sub EnleveApostrophe_()
Call Init
Dim Cel As Range

On Error Resume Next
For Each Cel In ws.Cells.SpecialCells(xlCellTypeConstants, 23)
    If Cel.PrefixCharacter <> "" Then
        Cel.Formula = Cel.Formula
    End If
Next Cel

End Sub


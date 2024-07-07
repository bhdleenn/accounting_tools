Attribute VB_Name = "Mod_Compta"
Option Explicit
Private MaxPlage&, j%, c As Range
Public Sub normalize_Compte_(ligne As Long, colonne As Integer, Plage As Long)
On Error Resume Next
Call Init

MaxPlage = 1000
Dim ncell As Long: ncell = ws.Cells(ligne, colonne).Value
If Plage > MaxPlage Then Call Error_(201): Plage = 1000

Do While ligne <= Plage
    If ncell > 0 Then
        Do While ncell < 9999999
        ncell = ncell * 10
        ws.Cells(ligne, colonne).Value = ncell
        Loop
    End If
    ligne = ligne + 1
    ncell = ws.Cells(ligne, colonne).Value
Loop
End Sub
Sub FindValue_(rt As String)
Call Init
    For j = 1 To 9
        Set c = ws.Range(rt).Find(j, LookIn:=xlValues)
        If Not c Is Nothing Then Exit Sub
    Next
    End
End Sub

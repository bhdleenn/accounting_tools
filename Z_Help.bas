Attribute VB_Name = "Z_Help"
'le raccourci pour les types est le suivant : % -Integer ; & -long ; @ -devise ; # -double ; ! -unique; $ -chaîne
Private VBC As Object, i%, Zqd As String
Sub Help()
On Error Resume Next
Retry:
Call Verif1
If i > 0 Then GoTo dev Else Zqd = InputBox("Version 1.0.1 - Vérouillé", "Saisie de la mot de passe")
If Zqd = "" Then Exit Sub
If Zqd <> "0000" Then GoTo Invalide Else GoTo Valide
Exit Sub
Invalide:
If MsgBox("Mot de passe incorrect", vbRetryCancel, "Ressayer") = vbRetry Then GoTo Retry
Exit Sub
Valide:
Call But_Import
Call Verif1
If i <= 0 Then i = MsgBox("Import échoué", vbMsgBoxHelpButton, "Import échoué") Else MsgBox ("Succès!")
Exit Sub
dev: MsgBox ("Version 1.0.1 - Dévérouillé")
End Sub
Sub Verif1()
i = 0
With ThisWorkbook.VBProject
For Each VBC In .VBComponents
If VBC.Name = "Index" Then i = i + 1
Next VBC
End With
End Sub


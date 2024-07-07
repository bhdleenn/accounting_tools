Attribute VB_Name = "Ind_errors"
Option Explicit

Public Sub Error_(Optional h As Integer = 0)
'Globals errors

If h = 1 Then err = MsgBox("Veuiller selectionner une feuille de calcul", 0, "Erreur")
If h = 2 Then err = MsgBox("Veuiller selectionner une plage de cellules", 0, "Erreur")
If h = 3 Then err = MsgBox("Caractère Spécial", 0, "Erreur"): Exit Sub

'app errors
If h = 100 Then err = MsgBox("Convertion annulée", 0, "Terminé")
If h = 101 Then err = MsgBox("Mise en forme dejà effectuée", 0, "Erreur")

If h = 121 Then err = MsgBox("veuillez selectionner une seule colonne", 0, "Erreur")

If h = 201 Then err = MsgBox("La plage ne peut exceder 1000 lignes", 1, "Attention"): If err = vbOK Then Exit Sub

Call Stop_
End
End Sub

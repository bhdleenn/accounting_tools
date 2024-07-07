Attribute VB_Name = "App_Paie"
Option Explicit
Private c%, i&, k%, l%, mois%, annee%, nom$, j&
Public Sub Paie_()
Attribute Paie_.VB_ProcData.VB_Invoke_Func = " \n14"
Call Init

'conditions
If InStr(1, ws.Cells(3, 1).Value, "/") > 0 Then Call Error_(101)
j = WorksheetFunction.CountA(Range("A:A")): If j = 0 Or InStr(1, ws.[A1].Value, "Ecritures comptables") = False Then Call Error_
Call Verif3("Convertion de en OD de paie")

Call Start_(True)
Call DuplicateSheet_(True)
Call EnleveApostrophe_

j = WorksheetFunction.CountA(Range("A:A")): j = j - 2: k = j + 1: l = j + 2

Nna:
nom = InputBox("Entrez le nom :", "Name", "Unknow")
If InStr(1, nom, "/") Or InStr(1, nom, ",") Or InStr(1, nom, ";") Or InStr(1, nom, ":") _
Or InStr(1, nom, "!") Or InStr(1, nom, ".") Or InStr(1, nom, "\") _
Or InStr(1, nom, "'") Or InStr(1, nom, ")") Or InStr(1, nom, "(") _
Or InStr(1, nom, "+") Or InStr(1, nom, "=") Or InStr(1, nom, "<") _
Or InStr(1, nom, ">") Or InStr(1, nom, "%") Or InStr(1, nom, "?") _
> 0 Then Call Error_(3): GoTo Nna
mois = InputBox("Entrez le mois :", "Mois", Format(Date, "mm"))
annee = InputBox("Entrez l'année :", "Année", Format(Date, "yyyy"))
nom = Trim(nom) & "paie" & mois & annee

Dim dt As Date: dt = DateSerial(annee, mois + 1, 1 - 1)

'Arangement des colonnes
ws.Rows("1:3").Delete
ws.Rows(l).Delete
ws.Rows(k).Delete
ws.Columns(1).Insert
ws.Columns(6).Delete

'Ajouter les dates sur la premiere colonne
i = 1: c = 1
Do While i <= j
    ws.Cells(i, c).Value = dt
    i = i + 1
Loop

'Normaliser la longeur des comptes
i = 1: c = 2
Call normalize_Compte_(i, c, j)

'Modifier le libélé à la date indiqué en msgbox
i = 1: c = 3
Do While i <= j
    ws.Cells(i, c).Value = "SALAIRES " & mois & " " & annee
    i = i + 1
Loop

'Mise en forme obligatoire
ws.Columns("D:E").NumberFormat = "0.00"
ws.Range("A:A").NumberFormat = "dd/mm/yyyy"
'Mise en forme stylistique
Call AutoFit_
Call Aling_
Call BackgroundColor_

Call saveAsCVS_(nom)
Call Stop_(True)
Application.DisplayAlerts = False: ws.Delete: Application.DisplayAlerts = True
Application.Quit
End Sub

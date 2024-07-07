Attribute VB_Name = "App_Vente"
Option Explicit
Sub FactVente_()
Call Error_(0)
Call Init
Call Start_(True)
    
'conditions
Call Verif3("Convertir la mise en forme des Ventes")
Call DuplicateSheet_(False)
Call CalibrageTabTop_

'Date;Numéro;Client;Libellé;Montant HT;Montant TVA;Montant TTC
        
'Date
Dim date1 As Range
Set date1 = ws.UsedRange.Find("/")

If date1.Column <> 1 Then
date1.EntireColumn.Cut
Columns(1).Insert Shift:=xlToRight
End If

'numéro
Dim numero1 As Range
Set numero1 = ws.UsedRange.Find("FA-")

If numero1.Column <> 2 Then
numero1.EntireColumn.Cut
Columns(2).Insert Shift:=xlToRight
End If

'client
Dim client1 As Range
Set client1 = ws.UsedRange.Find("client")

If client1.Column <> 3 Then
client1.EntireColumn.Cut
Columns(3).Insert Shift:=xlToRight
End If

'libélé
Dim lib1 As Range
Set lib1 = ws.UsedRange.Find("Désignation")

If lib1.Column <> 4 Then
lib1.EntireColumn.Cut
Columns(4).Insert Shift:=xlToRight
End If

'HT
Columns(5).Insert Shift:=xlToRight

'TVA
Dim TVA1 As Range
Set TVA1 = ws.UsedRange.Find("TVA")

If TVA1.Column <> 6 Then
TVA1.EntireColumn.Cut
Columns(6).Insert Shift:=xlToRight
End If

'TTC
Dim ttc1 As Range
Set ttc1 = ws.UsedRange.Find("TTC")

If ttc1.Column <> 7 Then
ttc1.EntireColumn.Cut
Columns(7).Insert Shift:=xlToRight
End If

Call CalibrageTabBottom_
Call replaceName_
Call AutoFit_
    
'Debug

'Debug.Print "numero1: " & numero1.Column
'Debug.Print "date1: " & date1.Column
'ws.UsedRange.Select
Call Stop_(True)

End Sub




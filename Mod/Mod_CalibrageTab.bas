Attribute VB_Name = "Mod_CalibrageTab"
Option Explicit
Private i%, j%, g%
Private countcol&, countlin&, maxcol&, mincol&, maxlin&, minlin&
Public Sub CalibrageTabTop_()
Call Init

i = 1
j = 1

countcol = WorksheetFunction.CountA(Columns(i))

If countcol < 1 Then
    Do While countcol < 7
        ws.Columns(i).Delete
        countcol = WorksheetFunction.CountA(Columns(i))
    Loop
End If

countlin = WorksheetFunction.CountA(Rows(j))

If countlin < 1 Then
    Do While countlin < 5
        ws.Rows(j).Delete
        countlin = WorksheetFunction.CountA(Rows(j))
    Loop
End If

End Sub

Private Sub calibrageMatrix_()
Call Init

i = 1
j = 1

'récupérer la longeur des colonnes
countcol = WorksheetFunction.CountA(Columns(i))
maxcol = countcol
mincol = countcol

Do While i < countlin

    If countcol > maxcol Then
        maxcol = countcol
    End If
    
    If countcol < mincol Then
        mincol = countcol
    End If
    
    i = i + 1
    countcol = WorksheetFunction.CountA(Columns(i))
Loop


'récupérer la longeur des lignes
countlin = WorksheetFunction.CountA(Rows(j))
maxlin = countlin
minlin = countlin


Do While j < countcol

    If countlin > maxlin Then
        maxlin = countlin
    End If
    
    If countlin < minlin Then
        minlin = countlin
    End If
    
    j = j + 1
    countlin = WorksheetFunction.CountA(Rows(j))
Loop

'Definir la plage du tableau
Dim collet As String: collet = Split(Cells(1, minlin).Address, "$")(1)
Dim tab1 As String: tab1 = "A1:" & collet & mincol

'debug
'Debug.Print "countcol: " & countcol
'Debug.Print "maxcol: " & maxcol
'Debug.Print "mincol: " & mincol
'
'Debug.Print "countlin: " & countlin
'Debug.Print "maxlin: " & maxlin
'Debug.Print "minlin: " & minlin
'
'Debug.Print "i: " & i
'Debug.Print "j: " & j
End Sub

Public Sub CalibrageTabBottom_()
Call Init
Call calibrageMatrix_

g = 6

Do While g < maxlin
    g = g + 1
    ws.Columns(8).Delete
Loop

End Sub



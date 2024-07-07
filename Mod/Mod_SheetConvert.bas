Attribute VB_Name = "Mod_SheetConvert"
Option Explicit
Private chemin$, chem$, final$, Rep As FileDialog
Public Sub DuplicateSheet_(Optional SelectNewPage As Boolean = True)
Call Init
ws.Copy after:=Sheets(Sheets.Count)
If SelectNewPage = False Then wb.Sheets(wsn).Select Else wb.ActiveSheet.Select
End Sub
Sub FileFolder_(Optional OpenFolder As Boolean = False)
Set Rep = Application.FileDialog(msoFileDialogFolderPicker)

If chemin <> "0" Then chem = chemin

With Rep
.Title = "Choix du répertoire d'enregistrement ..."
.ButtonName = "Enregistrer"
.InitialFileName = chem
.Show
End With

If Rep.SelectedItems.Count = 0 Then: chemin = "0": Exit Sub

chemin = Rep.SelectedItems(1)
If OpenFolder = True Then Shell Environ("WINDIR") & "\explorer.exe " & chemin, vbNormalFocus
End Sub
Sub saveAsCVS_(Optional Name As String = "", Optional FileF As String = xlCSV)
Call Init
Set app = New Application
app.Visible = True

Set nwb = app.Workbooks.Add
Set nws = nwb.Worksheets(1)

Call FileFolder_(True)
'chemin = "D:\Paie"
If chemin = "0" Then: app.Quit: Exit Sub
If Name = "" Then Name = "Save_" & Format(Now, "yyyymmddhhmmss")
If Len(chemin) <= 3 Then final = chemin & Name Else final = chemin & "\" & Name & ".csv"

ws.UsedRange.Copy
nws.Paste
nws.SaveAs Filename:=final, FileFormat:=xlCSV, CreateBackup:=False, local:=True

app.Quit
'chdrive
End Sub
Sub But_Savecvs()
Call Init
Dim nom As String
On Error Resume Next
nom = Left(wbn, InStr(wbn, ".") - 1)
Call saveAsCVS_(nom)

End Sub

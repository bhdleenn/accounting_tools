Attribute VB_Name = "Index"
Option Explicit
Global app As Application:
Global wb As Workbook
Global nwb As Workbook
Global ws As Worksheet
Global nws As Worksheet
Global wsn$, wbn$
Public Sub Init()
On Error GoTo err
Set wb = ActiveWorkbook: wbn = wb.Name: Set ws = wb.ActiveSheet: wsn = ws.Name: Exit Sub
err: Call Error_(1)
End Sub
Public Sub Start_(Optional RangeSelect_Start As Boolean = False)
Application.ScreenUpdating = False
'Application.DisplayAlerts = False
On Error Resume Next: If RangeSelect_Start = True Then ws.[A1].Select: Exit Sub
End Sub
Public Sub Stop_(Optional RangeSelect_Stop As Boolean = False)
Application.ScreenUpdating = True
'Application.DisplayAlerts = True
On Error Resume Next: If RangeSelect_Stop = True Then ws.[A1].Select: Exit Sub
End Sub
Public Sub Verif3(ContenuText As String)
If MsgBox(ContenuText, vbYesNo, "Demande de confirmation") = vbNo Then Call Error_(100)
End Sub

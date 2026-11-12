Attribute VB_Name = "ModInicio"
Option Explicit
Public Sub Main()
    On Error GoTo ErrorInicio
    frmMain.Show vbModal
    If Not frmMain.Listo Then
        Unload frmMain
        Exit Sub
    End If
    Unload frmMain
    Load MdiPri
    Exit Sub
ErrorInicio:
    MsgBox "No se pudo iniciar Clever SoftPro 6.0." & vbCrLf & Err.Description, vbExclamation, App.Title
    Unload frmMain
    Unload MdiPri
End Sub

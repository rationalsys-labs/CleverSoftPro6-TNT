VERSION 5.00
Begin VB.Form FrmHistorialRP
 Caption="Historial de gestión RP"
 ClientHeight=6200
 ClientWidth=9500
 StartUpPosition=1
 Begin VB.TextBox TxtHistorial
  Left=240
  Top=240
  Width=9000
  Height=5100
  MultiLine=-1
  ScrollBars=3
  Locked=-1
 End
 Begin VB.CommandButton CmdCerrar
  Caption="Cerrar"
  Left=7300
  Top=5500
  Width=1900
  Height=450
  Cancel=-1
 End
End
Attribute VB_Name = "FrmHistorialRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Sub CmdCerrar_Click()
 Unload Me
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

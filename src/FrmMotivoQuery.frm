VERSION 5.00
Begin VB.Form FrmMotivoQuery
 Caption="Anular solicitud Query"
 BorderStyle=3
 ClientWidth=6900
 ClientHeight=2500
 MaxButton=0
 MinButton=0
 ShowInTaskbar=0
 StartUpPosition=1
 BeginProperty Font
  Name="Tahoma"
  Size=8.25
 EndProperty
 Begin VB.Label LblMotivo
  Caption="Motivo de anulación de la solicitud:"
  Left=240
  Top=240
  Width=6300
  Height=300
 End
 Begin VB.TextBox TxtMotivo
  Left=240
  Top=600
  Width=6360
  Height=1050
  MultiLine=-1
  MaxLength=1000
  ScrollBars=2
 End
 Begin VB.CommandButton CmdAceptar
  Caption="Aceptar"
  Left=3240
  Top=1860
  Width=1560
  Height=375
  Default=-1
 End
 Begin VB.CommandButton CmdCancelar
  Caption="Cancelar"
  Left=5040
  Top=1860
  Width=1560
  Height=375
  Cancel=-1
 End
End
Attribute VB_Name = "FrmMotivoQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Public Aceptado As Boolean
Private Sub CmdAceptar_Click()
 If Len(Trim$(TxtMotivo.Text)) = 0 Then
  MsgBox "Escriba el motivo de anulación.", vbExclamation, App.Title
  TxtMotivo.SetFocus: Exit Sub
 End If
 Aceptado = True: Me.Hide
End Sub
Private Sub CmdCancelar_Click()
 Aceptado = False: Me.Hide
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = vbFormControlMenu Then Cancel = 1: Aceptado = False: Me.Hide
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

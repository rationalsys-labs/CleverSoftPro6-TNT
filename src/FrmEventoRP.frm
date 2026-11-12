VERSION 5.00
Begin VB.Form FrmEventoRP
 BorderStyle=3
 Caption="Registrar evento de gestión RP"
 ClientHeight=3300
 ClientWidth=6900
 StartUpPosition=1
 MaxButton=0
 MinButton=0
 BeginProperty Font
  Name="Tahoma"
  Size=8.25
  Charset=0
  Weight=400
 EndProperty
 Begin VB.Frame FraEvento
  Caption="Estado y comentario de la gestión"
  Left=180
  Top=120
  Width=6540
  Height=2460
  Begin VB.Label LblEstado
   Caption="Estado:"
   Left=180
   Top=360
   Width=900
   Height=255
  End
  Begin VB.ComboBox CmbEstado
   Style=2
   Left=1140
   Top=300
   Width=5160
   Height=315
  End
  Begin VB.Label LblComentario
   Caption="Comentario:"
   Left=180
   Top=840
   Width=1200
   Height=255
  End
  Begin VB.TextBox TxtComentario
   Left=180
   Top=1140
   Width=6120
   Height=1080
   MultiLine=-1
   MaxLength=1000
   ScrollBars=2
  End
 End
 Begin VB.CommandButton CmdAceptar
  Caption="Aceptar"
  Left=3780
  Top=2760
  Width=1380
  Height=375
  Default=-1
 End
 Begin VB.CommandButton CmdCancelar
  Caption="Cancelar"
  Left=5340
  Top=2760
  Width=1380
  Height=375
  Cancel=-1
 End
End
Attribute VB_Name = "FrmEventoRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Public Aceptado As Boolean
Public CodigoEstado As Long
Public Comentario As String
Public Sub Preparar()
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 Aceptado = False
 CmbEstado.Clear
 TxtComentario.Text = ""
 Set R = ConsultaOperacion("SELECT CodEst,Status,TitSta FROM dbo.Int_StatusRP WHERE ActEst=1 ORDER BY CodEst")
 Do While Not R.EOF
  CmbEstado.AddItem CStr(R!Status) & " - " & CStr(R!TitSta)
  CmbEstado.ItemData(CmbEstado.NewIndex) = CLng(R!CodEst)
  R.MoveNext
 Loop
 CerrarRecordset R
 Exit Sub
Fallo:
 CerrarRecordset R
 Err.Raise vbObjectError + 3400, , Err.Description
End Sub
Private Sub CmdAceptar_Click()
 If CmbEstado.ListIndex < 0 Or Len(Trim$(TxtComentario.Text)) = 0 Then
  MsgBox "Seleccione el estado y escriba el comentario del evento.", vbInformation, App.Title
  Exit Sub
 End If
 CodigoEstado = CmbEstado.ItemData(CmbEstado.ListIndex)
 Comentario = Trim$(TxtComentario.Text)
 Aceptado = True
 Me.Hide
End Sub
Private Sub CmdCancelar_Click()
 Aceptado = False
 Me.Hide
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = vbFormControlMenu Then
  Cancel = True
  Aceptado = False
  Me.Hide
 End If
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

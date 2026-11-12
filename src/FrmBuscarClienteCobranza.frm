VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBuscarClienteCobranza
 BorderStyle=3
 Caption="Buscar cliente existente"
 ClientHeight=5400
 ClientWidth=10950
 StartUpPosition=1
 KeyPreview=-1
 MaxButton=0
 MinButton=0
 BeginProperty Font
  Name="Tahoma"
  Size=8.25
  Charset=0
  Weight=400
 EndProperty
 Begin VB.Label LblNombre
  Caption="Nombre del cliente:"
  Left=240
  Top=300
  Width=1800
  Height=255
 End
 Begin VB.TextBox TxtBuscar
  Left=2040
  Top=240
  Width=8610
  Height=315
  MaxLength=100
 End
 Begin MSFlexGridLib.MSFlexGrid GridClientes
  Left=240
  Top=840
  Width=10425
  Height=3540
  Rows=1
  Cols=4
  FixedCols=0
  SelectionMode=1
  AllowUserResizing=1
 End
 Begin VB.Label LblResumen
  Left=240
  Top=4530
  Width=10425
  Height=255
 End
 Begin VB.CommandButton CmdSeleccionar
  Caption="Seleccionar"
  Left=7260
  Top=4860
  Width=1590
  Height=375
  Default=-1
 End
 Begin VB.CommandButton CmdCancelar
  Caption="Cancelar"
  Left=9060
  Top=4860
  Width=1590
  Height=375
  Cancel=-1
 End
End
Attribute VB_Name = "FrmBuscarClienteCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Public IdElegido As Long
Private Modulo As Integer, Datos As ADODB.Recordset, Preparando As Boolean, Seleccionando As Boolean
Public Sub Preparar(ByVal Codigo As Integer, ByVal Buscar As String)
 Modulo = Codigo: IdElegido = 0: Preparando = True
 TxtBuscar.Text = Left$(Buscar, 100): Preparando = False
 Cargar
End Sub
Private Sub Cargar()
 On Error GoTo Fallo
 CerrarRecordset Datos: Set Datos = CobClientes(Modulo, Trim$(TxtBuscar.Text))
 CobGrilla GridClientes, Datos, "Código|Nombre o razón social|RUC / DNI|Dirección", "Asient|CLIENT|RUCDNI|DIRECC"
 GridClientes.ColWidth(0) = 1600: GridClientes.ColWidth(1) = 4800: GridClientes.ColWidth(2) = 1600: GridClientes.ColWidth(3) = 5000
 CmdSeleccionar.Enabled = GridClientes.Rows > 1
 LblResumen.Caption = CStr(Datos.RecordCount) & " coincidencias (máximo 100). Enter o doble clic selecciona toda la fila."
 Exit Sub
Fallo:
 CmdSeleccionar.Enabled = False: LblResumen.Caption = Err.Description
End Sub
Private Sub TxtBuscar_Change()
 If Not Preparando And Modulo > 0 Then Cargar
End Sub
Private Sub CmdSeleccionar_Click()
 On Error GoTo Fallo
 IdElegido = CobFila(Datos, GridClientes): Me.Hide
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdCancelar_Click()
 IdElegido = 0: Me.Hide
End Sub
Private Sub GridClientes_Click()
 SeleccionFilaQuery GridClientes, Seleccionando
End Sub
Private Sub GridClientes_RowColChange()
 SeleccionFilaQuery GridClientes, Seleccionando
End Sub
Private Sub GridClientes_DblClick()
 CmdSeleccionar_Click
End Sub
Private Sub GridClientes_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then CmdSeleccionar_Click: KeyCode = 0
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = vbFormControlMenu Then Cancel = 1: IdElegido = 0: Me.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

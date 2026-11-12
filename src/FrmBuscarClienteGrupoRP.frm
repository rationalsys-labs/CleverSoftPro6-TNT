VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBuscarClienteGrupoRP
 BorderStyle=3
 Caption="Seleccionar cliente para el grupo RP"
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
Attribute VB_Name = "FrmBuscarClienteGrupoRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Public IdElegido As Long
Public CodigoElegido As String
Public NombreElegido As String
Private Datos As ADODB.Recordset
Private Preparando As Boolean
Public Sub Preparar(ByVal Nombre As String)
 Preparando = True
 IdElegido = 0
 CodigoElegido = "": NombreElegido = ""
 TxtBuscar.Text = Left$(Nombre, 100)
 Preparando = False
 CargarClientes
End Sub
Private Sub Form_Load()
 If Not AccesoPermitido(9) Then Unload Me
End Sub
Private Sub CargarClientes()
 Dim F As Long
 On Error GoTo Fallo
 IdElegido = 0
 CmdSeleccionar.Enabled = False
 GridClientes.Rows = 1
 GridClientes.FormatString = "Código|Nombre o razón social|RUC / DNI|En grupo"
 GridClientes.ColWidth(0) = 1650: GridClientes.ColWidth(1) = 6100
 GridClientes.ColWidth(2) = 1550: GridClientes.ColWidth(3) = 900
 CerrarRecordset Datos
 Set Datos = BuscarClientesParaGrupo(TxtBuscar.Text)
 Do While Not Datos.EOF
  GridClientes.AddItem CStr(Datos!Asient) & vbTab & CStr(Datos!CLIENT) & vbTab & CStr(Datos!RUCDNI) & vbTab & CStr(Datos!EnGrupo)
  Datos.MoveNext
 Loop
 If GridClientes.Rows > 1 Then GridClientes.Row = 1
 CmdSeleccionar.Enabled = (GridClientes.Rows > 1)
 LblResumen.Caption = CStr(GridClientes.Rows - 1) & " coincidencias. Máximo 100; escriba más caracteres para precisar. Enter o doble clic selecciona la fila."
 Exit Sub
Fallo:
 CerrarRecordset Datos
 GridClientes.Rows = 1
 LblResumen.Caption = Err.Description
End Sub
Private Sub TxtBuscar_Change()
 If Not Preparando Then CargarClientes
End Sub
Private Sub CmdSeleccionar_Click()
 If Datos Is Nothing Then Exit Sub
 If GridClientes.Row < 1 Or GridClientes.Row >= GridClientes.Rows Then Exit Sub
 If Datos.RecordCount = 0 Then Exit Sub
 Datos.MoveFirst: Datos.Move GridClientes.Row - 1
 IdElegido = CLng(Datos!Item)
 CodigoElegido = CStr(Datos!Asient)
 NombreElegido = CStr(Datos!CLIENT)
 Me.Hide
End Sub
Private Sub GridClientes_DblClick()
 CmdSeleccionar_Click
End Sub
Private Sub GridClientes_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then CmdSeleccionar_Click: KeyCode = 0
End Sub
Private Sub CmdCancelar_Click()
 IdElegido = 0
 Me.Hide
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = vbFormControlMenu Then
  Cancel = True
  IdElegido = 0
  Me.Hide
 Else
  CerrarRecordset Datos
 End If
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

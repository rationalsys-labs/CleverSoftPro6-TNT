VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmGruposRP
 BorderStyle=3
 Caption="Administrar grupos del Diario RP"
 ClientHeight=6840
 ClientWidth=12000
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
 Begin VB.Frame FraHerramientas
  Left=120
  Top=60
  Width=2100
  Height=6420
  Begin VB.Label LblHerramientas
   Caption="GRUPOS RP"
   Left=180
   Top=240
   Width=1740
   Height=300
  End
  Begin Clever.Boton CmdAgregar
   TX="Agregar grupo"
   BTYPE=6
   Left=120
   Top=720
   Width=1860
   Height=510
  End
  Begin Clever.Boton CmdEliminar
   TX="Eliminar grupo"
   BTYPE=6
   Left=120
   Top=1380
   Width=1860
   Height=450
  End
  Begin Clever.Boton CmdGuardar
   TX="Guardar cambios"
   BTYPE=6
   Left=120
   Top=2520
   Width=1860
   Height=450
  End
  Begin Clever.Boton CmdCancelar
   TX="Cancelar cambios"
   BTYPE=6
   Left=120
   Top=3120
   Width=1860
   Height=450
  End
  Begin Clever.Boton CmdCerrar
   TX="Cerrar"
   BTYPE=6
   Left=120
   Top=5580
   Width=1860
   Height=450
  End
 End
 Begin VB.Frame FraCliente
  Caption="Seleccionar un cliente existente"
  Left=2400
  Top=120
  Width=9360
  Height=1320
  Begin VB.Label LblNombre
   Caption="Nombre:"
   Left=180
   Top=360
   Width=900
   Height=255
  End
  Begin VB.TextBox TxtNombre
   Left=1140
   Top=300
   Width=5940
   Height=330
   MaxLength=100
  End
  Begin VB.CommandButton CmdBuscar
   Caption="Buscar (F9)"
   Left=7320
   Top=300
   Width=1800
   Height=375
  End
  Begin VB.Label LblCliente
   Caption="Escriba parte del nombre y pulse F9 para seleccionar la empresa."
   Left=180
   Top=840
   Width=8940
   Height=300
  End
 End
 Begin VB.Frame FraGrupos
  Caption="Clientes separados del grupo general"
  Left=2400
  Top=1560
  Width=9360
  Height=4920
  Begin MSFlexGridLib.MSFlexGrid GridGrupos
   Left=120
   Top=300
   Width=9120
   Height=4080
   Rows=1
   Cols=4
   FixedCols=0
   SelectionMode=1
   AllowUserResizing=1
  End
  Begin VB.Label LblAviso
   Caption="Los cambios se aplican únicamente al pulsar Guardar. El grupo general se conserva siempre."
   Left=180
   Top=4500
   Width=9000
   Height=300
  End
 End
 Begin VB.Label LblEstado
  Caption="Consulta"
  Left=2400
  Top=6540
  Width=9360
  Height=255
 End
End
Attribute VB_Name = "FrmGruposRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Datos As ADODB.Recordset
Private ClienteElegido As Long
Private CodigoElegido As String
Private NombreElegido As String
Private HayCambios As Boolean
Private Cargando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(9) Then Unload Me: Exit Sub
 CargarGrupos
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
 Unload Me
End Sub
Private Sub CargarGrupos()
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 Cargando = True
 CerrarRecordset Datos
 Set Datos = New ADODB.Recordset
 Datos.Fields.Append "Item", adInteger
 Datos.Fields.Append "IdCliente", adInteger
 Datos.Fields.Append "Codigo", adVarWChar, 20
 Datos.Fields.Append "Nombre", adVarWChar, 255
 Datos.Fields.Append "Guias", adInteger
 Datos.Fields.Append "Version", adBinary, 8
 Datos.Fields.Append "Cambio", adVarWChar, 10
 Datos.Open
 Set R = ConsultarGruposDiario()
 Do While Not R.EOF
  Datos.AddNew
  Datos!Item = R!Item: Datos!IdCliente = R!IdCliente
  Datos!Codigo = R!CodCli: Datos!Nombre = R!Nombre
  Datos!Guias = R!Guias: Datos!Version = R!VersionFila
  Datos!Cambio = ""
  Datos.Update
  R.MoveNext
 Loop
 CerrarRecordset R
 HayCambios = False
 Cargando = False
 PintarGrupos
 Exit Sub
Fallo:
 Cargando = False
 CerrarRecordset R
 Err.Raise vbObjectError + 3500, , Err.Description
End Sub
Private Sub PintarGrupos(Optional ByVal SeleccionarCliente As Long = 0)
 Dim F As Long, FilaElegida As Long, Estado As String
 Cargando = True
 GridGrupos.Rows = 1
 GridGrupos.FormatString = "Código|Cliente / grupo|Guías en Diario|Cambio pendiente"
 GridGrupos.ColWidth(0) = 1500: GridGrupos.ColWidth(1) = 4620
 GridGrupos.ColWidth(2) = 1300: GridGrupos.ColWidth(3) = 1600
 HayCambios = False
 If Datos.RecordCount > 0 Then Datos.MoveFirst
 Do While Not Datos.EOF
  Estado = ValorTexto(Datos!Cambio)
  If Estado <> "" Then HayCambios = True
  GridGrupos.AddItem CStr(Datos!Codigo) & vbTab & CStr(Datos!Nombre) & vbTab & CStr(Datos!Guias) & vbTab & Estado
  If CLng(Datos!IdCliente) = SeleccionarCliente Then FilaElegida = GridGrupos.Rows - 1
  Datos.MoveNext
 Loop
 If GridGrupos.Rows > 1 Then
  If FilaElegida = 0 Then FilaElegida = 1
  GridGrupos.Row = FilaElegida
 End If
 Cargando = False
 CmdGuardar.Enabled = HayCambios
 CmdCancelar.Enabled = HayCambios
 CmdEliminar.Enabled = (GridGrupos.Rows > 1)
 ActualizarAccionEliminar
 LblEstado.Caption = IIf(HayCambios, "Cambios pendientes de guardar. ", "Consulta. ") & CStr(GridGrupos.Rows - 1) & " grupos separados."
End Sub
Private Sub TxtNombre_Change()
 ClienteElegido = 0
 CodigoElegido = "": NombreElegido = ""
 LblCliente.Caption = "Pulse F9 para buscar y seleccionar la empresa."
End Sub
Private Sub CmdBuscar_Click()
 On Error GoTo Fallo
 FrmBuscarClienteGrupoRP.Preparar TxtNombre.Text
 FrmBuscarClienteGrupoRP.Show vbModal, Me
 If FrmBuscarClienteGrupoRP.IdElegido > 0 Then
  TxtNombre.Text = FrmBuscarClienteGrupoRP.NombreElegido
  ClienteElegido = FrmBuscarClienteGrupoRP.IdElegido
  CodigoElegido = FrmBuscarClienteGrupoRP.CodigoElegido
  NombreElegido = FrmBuscarClienteGrupoRP.NombreElegido
  LblCliente.Caption = CodigoElegido & " - " & NombreElegido
 End If
 Unload FrmBuscarClienteGrupoRP
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdAgregar_Click()
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If ClienteElegido = 0 Then CmdBuscar_Click
 If ClienteElegido = 0 Then Exit Sub
 If Datos.RecordCount > 0 Then Datos.MoveFirst
 Do While Not Datos.EOF
  If CLng(Datos!IdCliente) = ClienteElegido Then
   If ValorTexto(Datos!Cambio) = "Quitar" Then
    Datos!Cambio = "": Datos.Update
    PintarGrupos ClienteElegido
   Else
    MsgBox "El cliente ya está en la lista de grupos.", vbInformation, App.Title
   End If
   Exit Sub
  End If
  Datos.MoveNext
 Loop
 Set R = ConsultaOperacion("SELECT COUNT(*) Cantidad FROM dbo.DbsDiarRp WHERE IdCliente=?", Array(Array("Cliente", adInteger, 0, ClienteElegido)))
 Datos.AddNew
 Datos!Item = 0: Datos!IdCliente = ClienteElegido
 Datos!Codigo = CodigoElegido: Datos!Nombre = NombreElegido
 Datos!Guias = CLng(R!Cantidad): Datos!Cambio = "Agregar"
 CerrarRecordset R
 Datos.Update
 PintarGrupos ClienteElegido
 Exit Sub
Fallo:
 CerrarRecordset R
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdEliminar_Click()
 Dim Cliente As Long
 On Error GoTo Fallo
 If GridGrupos.Row < 1 Or GridGrupos.Row >= GridGrupos.Rows Then Exit Sub
 Datos.MoveFirst: Datos.Move GridGrupos.Row - 1
 Cliente = CLng(Datos!IdCliente)
 If ValorTexto(Datos!Cambio) = "Quitar" Then
  Datos!Cambio = "": Datos.Update
 Else
  If MsgBox("¿Eliminar el grupo " & CStr(Datos!Codigo) & " - " & CStr(Datos!Nombre) & "?" & vbCrLf & "Al guardar, todas sus guías del Diario pasarán al grupo general (actualmente " & CStr(Datos!Guias) & ")." & vbCrLf & "No se borrarán guías ni se cambiarán su cliente, estado o historial.", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
  If ValorTexto(Datos!Cambio) = "Agregar" Then
   Datos.Delete
  Else
   Datos!Cambio = "Quitar": Datos.Update
  End If
 End If
 PintarGrupos Cliente
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdGuardar_Click()
 Dim XML As String, R As ADODB.Recordset, Confirmado As Boolean, Mensaje As String
 On Error GoTo Fallo
 If Not HayCambios Then Exit Sub
 XML = "<cambios>"
 If Datos.RecordCount > 0 Then Datos.MoveFirst
 Do While Not Datos.EOF
  If ValorTexto(Datos!Cambio) = "Agregar" Then XML = XML & "<agregar" & AtributoXML("Cliente", CStr(Datos!IdCliente)) & "/>"
  If ValorTexto(Datos!Cambio) = "Quitar" Then XML = XML & "<quitar" & AtributoXML("Item", CStr(Datos!Item)) & AtributoXML("Version", VersionXML(Datos.Fields("Version").Value)) & "/>"
  Datos.MoveNext
 Loop
 XML = XML & "</cambios>"
 If MsgBox("¿Guardar los cambios en los grupos del Diario RP?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
 CmdGuardar.Enabled = False
 Set R = ConfirmarGruposDiario(XML)
 Confirmado = True
 HayCambios = False
 CerrarRecordset R
 CargarGrupos
 MsgBox "Grupos guardados. Al volver al Diario se actualizarán las agrupaciones.", vbInformation, App.Title
 Exit Sub
Fallo:
 Mensaje = Err.Description
 CerrarRecordset R
 If Confirmado Then
  HayCambios = False
  CmdCancelar.Enabled = False
  Mensaje = "Los grupos se guardaron. Cierre y vuelva a abrir para actualizar; no repita la escritura. " & Mensaje
 Else
  CmdGuardar.Enabled = HayCambios
 End If
 MsgBox Mensaje, vbExclamation, App.Title
End Sub
Private Function Descartar() As Boolean
 Descartar = True
 If HayCambios Then Descartar = (MsgBox("¿Descartar los cambios pendientes en los grupos?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) = vbYes)
End Function
Private Sub CmdCancelar_Click()
 On Error GoTo Fallo
 If Descartar Then CargarGrupos
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdCerrar_Click()
 Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If Not Descartar Then Cancel = True: Exit Sub
 CerrarRecordset Datos
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF9 Then CmdBuscar_Click: KeyCode = 0
 If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub GridGrupos_RowColChange()
 If Not Cargando Then ActualizarAccionEliminar
End Sub
Private Sub ActualizarAccionEliminar()
 CmdEliminar.Caption = "Eliminar grupo"
 If Datos Is Nothing Then Exit Sub
 If GridGrupos.Row < 1 Or GridGrupos.Row >= GridGrupos.Rows Then Exit Sub
 If Datos.RecordCount = 0 Then Exit Sub
 Datos.MoveFirst: Datos.Move GridGrupos.Row - 1
 If ValorTexto(Datos!Cambio) = "Quitar" Then CmdEliminar.Caption = "Deshacer retiro"
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

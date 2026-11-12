VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCatalogos
   Caption = "Mantenimiento de Catálogos"
   BorderStyle = 1
   ClientHeight = 6480
   ClientWidth = 10080
   MaxButton = 0
   MinButton = 0
   MDIChild = -1
   KeyPreview = -1
   Begin VB.Label LblCatalogo
      Left = 120
      Top = 120
      Width = 1050
      Height = 240
      Caption = "Catálogo:"
   End
   Begin VB.ComboBox CmbCatalogo
      Left = 1200
      Top = 90
      Width = 2880
      Height = 315
      Style = 2
      TabIndex = 0
   End
   Begin VB.CheckBox ChkEliminados
      Left = 4320
      Top = 120
      Width = 2400
      Height = 285
      Caption = "Incluir eliminados"
      TabIndex = 1
   End
   Begin VB.Label LblBuscar
      Left = 120
      Top = 540
      Width = 1050
      Height = 240
      Caption = "Buscar:"
   End
   Begin VB.TextBox TxtBuscar
      Left = 1200
      Top = 510
      Width = 6480
      Height = 285
      Appearance = 0
      MaxLength = 150
      TabIndex = 2
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCatalogos
      Left = 120
      Top = 960
      Width = 7560
      Height = 3180
      Rows = 2
      Cols = 4
      FixedCols = 0
      TabIndex = 3
   End
   Begin VB.Frame FraFicha
      Left = 120
      Top = 4260
      Width = 7560
      Height = 1740
      Caption = "Ficha del registro"
   Begin VB.Label LblCodigo
      Left = 150
      Top = 330
      Width = 900
      Height = 240
      Caption = "Código:"
   End
   Begin VB.TextBox TxtCodigo
      Left = 1110
      Top = 300
      Width = 1050
      Height = 285
      Appearance = 0
      Locked = -1
      TabStop = 0
   End
   Begin VB.Label LblAbreviatura
      Left = 2640
      Top = 330
      Width = 1440
      Height = 240
      Caption = "Abreviatura:"
   End
   Begin VB.TextBox TxtAbreviatura
      Left = 4110
      Top = 300
      Width = 840
      Height = 285
      Appearance = 0
      MaxLength = 2
      TabIndex = 4
   End
   Begin VB.Label LblNombre
      Left = 150
      Top = 900
      Width = 1200
      Height = 240
      Caption = "Descripción:"
   End
   Begin VB.TextBox TxtNombre
      Left = 150
      Top = 1200
      Width = 7200
      Height = 285
      Appearance = 0
      MaxLength = 150
      TabIndex = 5
   End
   End
   Begin VB.Frame FraAcciones
      Left = 7920
      Top = 90
      Width = 2040
      Height = 5100
      Caption = ""
   Begin Clever.Boton CmdNuevo
      Left = 150
      Top = 300
      Width = 1740
      Height = 480
      TX = "Crear"
      BTYPE = 8
      COLTYPE = 2
      BCOL = 14606046
      BCOLO = 8612156
      FCOLO = 16777215
      FOCUSR = -1
      TabIndex = 6
   End
   Begin Clever.Boton CmdModificar
      Left = 150
      Top = 990
      Width = 1740
      Height = 480
      TX = "Modificar"
      BTYPE = 8
      COLTYPE = 2
      BCOL = 14606046
      BCOLO = 8612156
      FCOLO = 16777215
      FOCUSR = -1
      TabIndex = 7
   End
   Begin Clever.Boton CmdEliminar
      Left = 150
      Top = 1680
      Width = 1740
      Height = 480
      TX = "Eliminar"
      BTYPE = 8
      COLTYPE = 2
      BCOL = 14606046
      BCOLO = 8612156
      FCOLO = 16777215
      FOCUSR = -1
      TabIndex = 8
   End
   Begin Clever.Boton CmdGuardar
      Left = 150
      Top = 2370
      Width = 1740
      Height = 480
      TX = "Guardar"
      BTYPE = 8
      COLTYPE = 2
      BCOL = 14606046
      BCOLO = 8612156
      FCOLO = 16777215
      FOCUSR = -1
      TabIndex = 9
   End
   Begin Clever.Boton CmdCancelar
      Left = 150
      Top = 3060
      Width = 1740
      Height = 480
      TX = "Cancelar"
      BTYPE = 8
      COLTYPE = 2
      BCOL = 14606046
      BCOLO = 8612156
      FCOLO = 16777215
      FOCUSR = -1
      TabIndex = 10
   End
   End
   Begin Clever.Boton CmdSalir
      Left = 8070
      Top = 5460
      Width = 1740
      Height = 480
      TX = "Salir"
      BTYPE = 8
      COLTYPE = 2
      BCOL = 14606046
      BCOLO = 8612156
      FCOLO = 16777215
      TabIndex = 11
   End
   Begin MSComctlLib.StatusBar StatusBar1
      Align = 2
      Height = 300
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628}
         NumPanels = 3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628}
            MinWidth = 2200
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628}
            MinWidth = 3500
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628}
            AutoSize = 1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Datos As ADODB.Recordset
Private Modo As Integer
Private Seleccion As Long
Private Version As Variant
Private Cargando As Boolean
Private CatalogoActual As Integer

Private Sub Form_Load()
    If Not AccesoPermitido(39) Then
        MsgBox "No tiene acceso al mantenimiento de catálogos.", vbExclamation, App.Title
        Unload Me
        Exit Sub
    End If
    Cargando = True
    CmbCatalogo.AddItem "Territorios"
    CmbCatalogo.AddItem "Niveles de cliente"
    CmbCatalogo.AddItem "Tipos de cobranza"
    CmbCatalogo.AddItem "Monedas"
    CmbCatalogo.AddItem "Tipos de documento"
    CmbCatalogo.ListIndex = 0
    CatalogoActual = 1
    GrdCatalogos.SelectionMode = flexSelectionByRow
    GrdCatalogos.HighLight = flexHighlightAlways
    GrdCatalogos.FocusRect = flexFocusNone
    GrdCatalogos.AllowUserResizing = flexResizeNone
    GrdCatalogos.FormatString = "Código|Descripción|Abrev.|Estado"
    GrdCatalogos.ColWidth(0) = 900
    GrdCatalogos.ColWidth(1) = 4500
    GrdCatalogos.ColWidth(2) = 750
    GrdCatalogos.ColWidth(3) = 1050
    Cargando = False
    CargarDatos
End Sub

Private Sub ActualizarControles()
    Dim Activo As Boolean
    If Seleccion <> 0 Then Activo = CBool(Datos!Activo)
    CmbCatalogo.Enabled = (Modo = 0)
    TxtBuscar.Enabled = (Modo = 0)
    ChkEliminados.Enabled = (Modo = 0)
    GrdCatalogos.Enabled = (Modo = 0)
    CmdNuevo.Enabled = (Modo = 0 And AccesoPermitido(39))
    CmdModificar.Enabled = (Modo = 0 And Activo And AccesoPermitido(39))
    CmdEliminar.Enabled = CmdModificar.Enabled
    CmdGuardar.Visible = (Modo <> 0)
    CmdCancelar.Visible = (Modo <> 0)
    TxtNombre.Locked = (Modo = 0)
    TxtAbreviatura.Locked = (Modo <> 1)
    TxtAbreviatura.Visible = (CatalogoActual = 5)
    LblAbreviatura.Visible = TxtAbreviatura.Visible
    If CatalogoActual = 1 Then TxtNombre.MaxLength = 150 Else TxtNombre.MaxLength = 50
    If CatalogoActual = 4 Then LblNombre.Caption = "Símbolo:" Else LblNombre.Caption = "Descripción:"
    StatusBar1.Panels(1).Text = Choose(Modo + 1, "Consulta", "Crear", "Modificar")
    If Seleccion = 0 Then
        StatusBar1.Panels(2).Text = "Sin selección"
    Else
        StatusBar1.Panels(2).Text = "Código: " & CStr(Seleccion)
    End If
End Sub

Private Sub LimpiarFicha()
    Seleccion = 0
    Version = Null
    TxtCodigo.Text = ""
    TxtNombre.Text = ""
    TxtAbreviatura.Text = ""
End Sub

Private Sub CargarDatos(Optional ByVal CodigoGuardado As Long = 0)
    Dim Cn As ADODB.Connection
    Dim Orden As ADODB.Command
    Dim Nuevos As ADODB.Recordset
    Dim Fila As Long, Elegida As Long
    On Error GoTo ErrorConsulta
    Cargando = True
    LimpiarFicha
    GrdCatalogos.Rows = 2
    GrdCatalogos.Row = 1
    GrdCatalogos.Col = 0
    For Fila = 0 To GrdCatalogos.Cols - 1
        GrdCatalogos.TextMatrix(1, Fila) = ""
    Next Fila
    ActualizarControles
    AbrirConexion Cn
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandType = adCmdStoredProc
    Orden.CommandText = "dbo.ConsultarCatalogo"
    Orden.Parameters.Append Orden.CreateParameter("@Catalogo", adUnsignedTinyInt, adParamInput, , CatalogoActual)
    Orden.Parameters.Append Orden.CreateParameter("@IncluyeEliminados", adBoolean, adParamInput, , ChkEliminados.Value = vbChecked)
    Orden.Parameters.Append Orden.CreateParameter("@Filtro", adVarWChar, adParamInput, 150, TxtBuscar.Text)
    Set Nuevos = New ADODB.Recordset
    Nuevos.CursorLocation = adUseClient
    Nuevos.Open Orden, , adOpenStatic, adLockReadOnly
    Set Nuevos.ActiveConnection = Nothing
    Cn.Close
    CerrarRecordset Datos
    Set Datos = Nuevos
    If Datos.RecordCount > 0 Then
        GrdCatalogos.Rows = Datos.RecordCount + 1
        Fila = 1
        Do While Not Datos.EOF
            GrdCatalogos.TextMatrix(Fila, 0) = CStr(Datos!Codigo)
            GrdCatalogos.TextMatrix(Fila, 1) = CStr(Datos!Nombre)
            GrdCatalogos.TextMatrix(Fila, 2) = CStr(Datos!Abreviatura)
            GrdCatalogos.TextMatrix(Fila, 3) = IIf(CBool(Datos!Activo), "Activo", "Eliminado")
            If CLng(Datos!Codigo) = CodigoGuardado Then Elegida = Fila
            Fila = Fila + 1
            Datos.MoveNext
        Loop
        If Elegida = 0 Then Elegida = 1
        GrdCatalogos.Row = Elegida
    End If
    Cargando = False
    SeleccionarFila
    StatusBar1.Panels(3).Text = CStr(Datos.RecordCount) & " registros. F5: actualizar."
    Exit Sub
ErrorConsulta:
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    CerrarRecordset Nuevos
    CerrarRecordset Datos
    Cargando = False
    LimpiarFicha
    ActualizarControles
    StatusBar1.Panels(3).Text = "No se pudo consultar. F5 permite reintentar."
    If CodigoGuardado <> 0 Then
        MsgBox "Se guardó el código " & CStr(CodigoGuardado) & ". Falló la recarga; pulse F5 sin repetir el guardado.", vbExclamation, App.Title
    Else
        MsgBox "No se pudo consultar: " & Err.Description, vbExclamation, App.Title
    End If
End Sub

Private Sub SeleccionarFila()
    If Cargando Or Modo <> 0 Then Exit Sub
    LimpiarFicha
    If Datos Is Nothing Then GoTo SinDatos
    If Datos.State <> adStateOpen Then GoTo SinDatos
    If Datos.RecordCount = 0 Or GrdCatalogos.Row < 1 Then GoTo SinDatos
    Datos.MoveFirst
    Datos.Move GrdCatalogos.Row - 1
    If Datos.EOF Then GoTo SinDatos
    Seleccion = CLng(Datos!Codigo)
    Version = Datos!VersionFila
    TxtCodigo.Text = CStr(Seleccion)
    TxtNombre.Text = CStr(Datos!Nombre)
    TxtAbreviatura.Text = CStr(Datos!Abreviatura)
    Cargando = True
    GrdCatalogos.Col = 0
    GrdCatalogos.ColSel = GrdCatalogos.Cols - 1
    GrdCatalogos.RowSel = GrdCatalogos.Row
    Cargando = False
SinDatos:
    ActualizarControles
End Sub

Private Sub CmdNuevo_Click()
    If Modo <> 0 Or Not AccesoPermitido(39) Then Exit Sub
    Modo = 1
    LimpiarFicha
    ActualizarControles
    StatusBar1.Panels(3).Text = "Nuevo registro. Guardar o Cancelar."
    TxtNombre.SetFocus
End Sub

Private Sub CmdModificar_Click()
    If Modo <> 0 Or Seleccion = 0 Or Not AccesoPermitido(39) Then Exit Sub
    If Not CBool(Datos!Activo) Then Exit Sub
    Modo = 2
    ActualizarControles
    StatusBar1.Panels(3).Text = "Edición. Guardar o Cancelar."
    TxtNombre.SetFocus
End Sub

Private Sub CmdGuardar_Click()
    Dim Cn As ADODB.Connection, Orden As ADODB.Command, Resultado As ADODB.Recordset
    Dim NuevoCodigo As Long, Confirmado As Boolean, Mensaje As String
    On Error GoTo ErrorGuardado
    If Modo = 0 Or Not AccesoPermitido(39) Then Exit Sub
    If Len(Trim$(TxtNombre.Text)) = 0 Then
        MsgBox "Ingrese la descripción o símbolo.", vbExclamation, App.Title
        TxtNombre.SetFocus
        Exit Sub
    End If
    If CatalogoActual = 5 And Len(Trim$(TxtAbreviatura.Text)) <> 2 Then
        MsgBox "Ingrese una abreviatura de dos letras o números.", vbExclamation, App.Title
        TxtAbreviatura.SetFocus
        Exit Sub
    End If
    If MsgBox("¿Guardar en " & CmbCatalogo.Text & ": " & Trim$(TxtNombre.Text) & "?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
    AbrirConexion Cn
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandType = adCmdStoredProc
    Orden.CommandText = "dbo.GuardarCatalogo"
    Orden.Parameters.Append Orden.CreateParameter("@Catalogo", adUnsignedTinyInt, adParamInput, , CatalogoActual)
    Orden.Parameters.Append Orden.CreateParameter("@Codigo", adInteger, adParamInput, , Seleccion)
    Orden.Parameters.Append Orden.CreateParameter("@Version", adBinary, adParamInput, 8, Version)
    Orden.Parameters.Append Orden.CreateParameter("@Nombre", adVarWChar, adParamInput, 150, Trim$(TxtNombre.Text))
    Orden.Parameters.Append Orden.CreateParameter("@Abreviatura", adVarWChar, adParamInput, 2, UCase$(Trim$(TxtAbreviatura.Text)))
    Set Resultado = Orden.Execute
    NuevoCodigo = CLng(Resultado!Codigo)
    Confirmado = True
    Modo = 0
    Resultado.Close
    Cn.Close
    Cargando = True
    TxtBuscar.Text = ""
    Cargando = False
    CargarDatos NuevoCodigo
    Exit Sub
ErrorGuardado:
    Mensaje = Err.Description
    If Not Resultado Is Nothing Then
        If Resultado.State = adStateOpen Then Resultado.Close
    End If
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    Cargando = False
    If Confirmado Then
        Modo = 0
        LimpiarFicha
        ActualizarControles
        MsgBox "Se guardó el código " & CStr(NuevoCodigo) & ". Pulse F5 para recargar; no repita el guardado.", vbExclamation, App.Title
    Else
        MsgBox "No se pudo confirmar el guardado. Los datos se conservan. Consulte antes de reintentar: " & Mensaje, vbExclamation, App.Title
    End If
End Sub

Private Sub CmdEliminar_Click()
    Dim Cn As ADODB.Connection, Orden As ADODB.Command
    On Error GoTo ErrorBaja
    If Modo <> 0 Or Seleccion = 0 Or Not AccesoPermitido(39) Then Exit Sub
    If Not CBool(Datos!Activo) Then Exit Sub
    If MsgBox("¿Eliminar " & CmbCatalogo.Text & ", código " & CStr(Seleccion) & " - " & TxtNombre.Text & "?" & vbCrLf & "La baja será lógica. Sus referencias se conservan. No podrá reactivarse.", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
    AbrirConexion Cn
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandType = adCmdStoredProc
    Orden.CommandText = "dbo.EliminarCatalogo"
    Orden.Parameters.Append Orden.CreateParameter("@Catalogo", adUnsignedTinyInt, adParamInput, , CatalogoActual)
    Orden.Parameters.Append Orden.CreateParameter("@Codigo", adInteger, adParamInput, , Seleccion)
    Orden.Parameters.Append Orden.CreateParameter("@Version", adBinary, adParamInput, 8, Version)
    Orden.Execute , , adExecuteNoRecords
    Cn.Close
    CargarDatos
    StatusBar1.Panels(3).Text = "Baja lógica realizada. Las referencias se conservan."
    Exit Sub
ErrorBaja:
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    MsgBox "No se pudo confirmar la baja. Recargue antes de reintentar: " & Err.Description, vbExclamation, App.Title
End Sub

Private Function ConfirmarDescarte() As Boolean
    ConfirmarDescarte = True
    If Modo <> 0 Then ConfirmarDescarte = (MsgBox("¿Descartar los cambios del catálogo?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) = vbYes)
End Function
Private Sub CmdCancelar_Click()
    If Not ConfirmarDescarte Then Exit Sub
    Modo = 0
    SeleccionarFila
    StatusBar1.Panels(3).Text = "Edición cancelada."
End Sub
Private Sub CmbCatalogo_Click()
    If Cargando Or Modo <> 0 Then Exit Sub
    CatalogoActual = CmbCatalogo.ListIndex + 1
    Cargando = True
    TxtBuscar.Text = ""
    Cargando = False
    CargarDatos
End Sub
Private Sub TxtBuscar_Change()
    If Cargando Or Modo <> 0 Then Exit Sub
    CargarDatos
End Sub
Private Sub ChkEliminados_Click()
    If Cargando Or Modo <> 0 Then Exit Sub
    CargarDatos
End Sub
Private Sub GrdCatalogos_RowColChange()
    SeleccionarFila
End Sub
Private Sub GrdCatalogos_Click()
    SeleccionarFila
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not ConfirmarDescarte Then Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CerrarRecordset Datos
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        If Modo <> 0 Then CmdCancelar_Click Else Unload Me
    ElseIf KeyCode = vbKeyF5 And Modo = 0 Then
        KeyCode = 0
        CargarDatos
    End If
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

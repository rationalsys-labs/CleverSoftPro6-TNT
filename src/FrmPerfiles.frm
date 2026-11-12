VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPerfiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfiles de acceso"
   ClientHeight = 9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12495
   Icon            =   "FrmUsuarios.frx":0000
   KeyPreview = -1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12495
   Begin VB.Frame Panel
      Left = 0
      Top = 0
      Width = 2490
      Height = 8010
      Caption = ""
   Begin VB.Label LblOpciones
      Left = 45
      Top = 45
      Width = 2400
      Height = 240
      Caption = "Seleccionar opción"
   End
   Begin Clever.Boton CmdListado 
      TabIndex = 10
      Height          =   420
      Left            =   90
      Top             =   3900
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Volver al listado"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUsuarios.frx":7CA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton CmdNuevo 
      TabIndex = 11
      Height          =   420
      Left            =   90
      Top             =   1080
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Crear perfil"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUsuarios.frx":7CA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton CmdModificar 
      TabIndex = 12
      Height          =   420
      Left            =   90
      Top             =   1710
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Modificar perfil"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUsuarios.frx":7CA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton CmdEliminar 
      TabIndex = 14
      Height          =   420
      Left            =   90
      Top             =   2340
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Eliminar perfil"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUsuarios.frx":7CA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   End
   Begin VB.Label LblSeleccion
      Left = 2700
      Top = 90
      Width = 5700
      Height = 240
      Caption = "Sin perfil seleccionado"
   End
   Begin VB.Frame FraFicha
      Left = 2565
      Top = 360
      Width = 9735
      Height = 7650
      Caption = "Ficha del perfil"
      Visible = 0
   Begin VB.Frame FraGeneral
      Left = 90
      Top = 240
      Width = 9540
      Height = 1920
      Caption = "Datos generales"
   Begin VB.Label LblNombre
      Left = 240
      Top = 330
      Width = 1680
      Height = 240
      Caption = "Perfil:"
   End
   Begin VB.TextBox TxtNombre
      TabIndex = 1
      Left = 2100
      Top = 300
      Width = 7170
      Height = 285
      Appearance = 0
      MaxLength = 50
   End
   Begin VB.Label LblCodigo
      Left = 240
      Top = 870
      Width = 1680
      Height = 240
      Caption = "Código:"
   End
   Begin VB.TextBox TxtCodigo
      Left = 2100
      Top = 840
      Width = 1500
      Height = 285
      Locked = -1
      Appearance = 0
      TabStop = 0
   End
   Begin VB.Label LblVinculados
      Left = 240
      Top = 1350
      Width = 7500
      Height = 240
      Caption = "Usuarios vinculados: 0"
   End
   Begin VB.Label LblAyuda
      Left = 240
      Top = 1650
      Width = 9150
      Height = 240
      Caption = "Los cambios del perfil se aplican a sus usuarios al iniciar sesión nuevamente."
   End
   End
   Begin VB.Frame FraAccess
      Left = 90
      Top = 2250
      Width = 9540
      Height = 5250
      Caption = "Accesos de usuario"
      Visible = 0
   Begin VB.Label LblPermisos
      Left = 135
      Top = 120
      Width = 5700
      Height = 240
      Caption = "Permisos de menús y módulos"
   End
   Begin MSComctlLib.TreeView ViewMenus
      TabIndex = 7
      Left = 135
      Top = 450
      Width = 9255
      Height = 4650
      Checkboxes = -1
      Style = 7
      HideSelection = 0
   End
   End
   End
   Begin VB.Frame FraListado
      Left = 2565
      Top = 360
      Width = 9735
      Height = 7650
      Caption = "Perfiles de acceso"
   Begin VB.Label LblBuscar
      Left = 135
      Top = 330
      Width = 900
      Height = 240
      Caption = "Buscar:"
   End
   Begin VB.TextBox TxtBuscar
      Left = 1110
      Top = 300
      Width = 5400
      Height = 285
      Appearance = 0
      MaxLength = 100
   End
   Begin VB.CheckBox ChkEliminados
      Left = 6750
      Top = 300
      Width = 2700
      Height = 285
      Caption = "Incluir eliminados"
   End
   Begin MSFlexGridLib.MSFlexGrid GridPerfiles
      Left = 135
      Top = 900
      Width = 9450
      Height = 6465
      Rows = 2
      Cols = 4
      FixedCols = 0
   End
   End
   Begin Clever.Boton CmdGuardar 
      TabIndex = 8
      Height          =   420
      Left            =   4950
      Top             =   8190
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Guardar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUsuarios.frx":7CA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton CmdCancelar 
      TabIndex = 9
      Height          =   420
      Left            =   6795
      Top             =   8190
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUsuarios.frx":7CA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton CmdSalir 
      TabIndex = 17
      Height          =   420
      Left            =   90
      Top             =   8190
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUsuarios.frx":7CA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1
      Align = 2
      Height = 300
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628}
         NumPanels = 3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628}
            MinWidth = 1800
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628}
            MinWidth = 2800
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628}
            AutoSize = 1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPerfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Modo As Integer
Private Seleccion As Long
Private Version As Variant
Private Activo As Boolean
Private Vinculados As Long
Private Cargando As Boolean
Private MostrandoListado As Boolean
Private Padres(1 To 40) As Integer
Private Vigentes(1 To 40) As Boolean
Private Consulta(1 To 40) As Boolean

Private Sub Form_Load()
    If Not PuedeAdministrarPerfiles Then
        MsgBox "Solo el administrador autorizado puede gestionar perfiles.", vbExclamation, App.Title
        Unload Me
        Exit Sub
    End If
    GridPerfiles.SelectionMode = flexSelectionByRow
    GridPerfiles.HighLight = flexHighlightAlways
    GridPerfiles.FocusRect = flexFocusNone
    GridPerfiles.AllowUserResizing = flexResizeNone
    GridPerfiles.FormatString = "Código|Perfil|Estado|Usuarios vinculados"
    GridPerfiles.ColWidth(0) = 900
    GridPerfiles.ColWidth(1) = 5550
    GridPerfiles.ColWidth(2) = 1200
    GridPerfiles.ColWidth(3) = 1500
    MostrandoListado = True
    CargarListado
    MostrarListado
End Sub
Private Sub LimpiarFicha()
    Seleccion = 0
    Version = Null
    Activo = False
    Vinculados = 0
    TxtNombre.Text = ""
    TxtCodigo.Text = ""
    LblVinculados.Caption = "Usuarios vinculados: 0"
    LblSeleccion.Caption = "Sin perfil seleccionado"
    Cargando = True
    ViewMenus.Nodes.Clear
    Cargando = False
End Sub
Private Sub ActualizarControles()
    CmdListado.Enabled = (Modo = 0 And Not MostrandoListado)
    CmdNuevo.Enabled = (Modo = 0)
    CmdModificar.Enabled = (Modo = 0 And Seleccion > 0 And Activo)
    CmdEliminar.Enabled = CmdModificar.Enabled
    CmdGuardar.Visible = (Modo <> 0)
    CmdCancelar.Visible = (Modo <> 0)
    TxtNombre.Locked = (Modo = 0)
    TxtBuscar.Enabled = (Modo = 0 And MostrandoListado)
    ChkEliminados.Enabled = TxtBuscar.Enabled
    GridPerfiles.Enabled = (Modo = 0)
    ViewMenus.Enabled = True
    StatusBar1.Panels(1).Text = Choose(Modo + 1, "Consulta", "Crear", "Modificar")
    If Seleccion = 0 Then StatusBar1.Panels(2).Text = "Sin selección" Else StatusBar1.Panels(2).Text = "Perfil: " & TxtNombre.Text
End Sub
Private Sub CargarArbol(ByVal Codigo As Long)
    Dim Datos As ADODB.Recordset, Nodo As MSComctlLib.Node, I As Integer
    Set Datos = ListarPermisosPerfil(Codigo)
    Cargando = True
    ViewMenus.Nodes.Clear
    Do While Not Datos.EOF
        I = CInt(Datos!Item)
        Padres(I) = 0
        If Not IsNull(Datos!IdPadre) Then Padres(I) = CInt(Datos!IdPadre)
        Vigentes(I) = CBool(Datos!Vigente)
        Consulta(I) = CBool(Datos!Acceso)
        If Vigentes(I) Then
            If Padres(I) = 0 Then Set Nodo = ViewMenus.Nodes.Add(, , "M" & CStr(I), CStr(Datos!Nombre)) Else Set Nodo = ViewMenus.Nodes.Add("M" & CStr(Padres(I)), tvwChild, "M" & CStr(I), CStr(Datos!Nombre))
            If Not FuncionDisponible(I) Then Nodo.Text = Nodo.Text & " (pendiente)"
            Nodo.Checked = Consulta(I)
            Nodo.Expanded = True
        End If
        Datos.MoveNext
    Loop
    Datos.Close
    Cargando = False
End Sub
Private Sub CargarFicha(ByVal Codigo As Long)
    Dim Datos As ADODB.Recordset
    On Error GoTo ErrorFicha
    Set Datos = ListarPerfilesAcceso(True, "", Codigo)
    If Datos.EOF Then Err.Raise vbObjectError + 3010, , "No se encontró el perfil."
    LimpiarFicha
    Seleccion = CLng(Datos!Item)
    Version = Datos!VersionFila
    Activo = CBool(Datos!Activo)
    Vinculados = CLng(Datos!UsuariosVinculados)
    TxtNombre.Text = CStr(Datos!Nombre)
    TxtCodigo.Text = CStr(Seleccion)
    LblSeleccion.Caption = TxtNombre.Text & IIf(Activo, "", " (eliminado)")
    LblVinculados.Caption = "Usuarios vinculados: " & CStr(Vinculados)
    Datos.Close
    CargarArbol Seleccion
    ActualizarControles
    Exit Sub
ErrorFicha:
    CerrarRecordset Datos
    Cargando = False
    LimpiarFicha
    ActualizarControles
    MsgBox "No se pudo cargar el perfil: " & Err.Description, vbExclamation, App.Title
End Sub
Private Sub CargarListado(Optional ByVal Codigo As Long = 0, Optional ByVal Guardado As Boolean = False)
    Dim Datos As ADODB.Recordset, Fila As Long, Elegida As Long, Mensaje As String
    On Error GoTo ErrorListado
    LimpiarFicha
    Cargando = True
    GridPerfiles.Rows = 2
    For Fila = 0 To 3
        GridPerfiles.TextMatrix(1, Fila) = ""
    Next Fila
    ActualizarControles
    Set Datos = ListarPerfilesAcceso(ChkEliminados.Value = vbChecked, TxtBuscar.Text)
    If Datos.RecordCount > 0 Then GridPerfiles.Rows = Datos.RecordCount + 1
    Fila = 1
    Do While Not Datos.EOF
        GridPerfiles.TextMatrix(Fila, 0) = CStr(Datos!Item)
        GridPerfiles.TextMatrix(Fila, 1) = CStr(Datos!Nombre)
        GridPerfiles.TextMatrix(Fila, 2) = IIf(CBool(Datos!Activo), "Activo", "Eliminado")
        GridPerfiles.TextMatrix(Fila, 3) = CStr(Datos!UsuariosVinculados)
        If CLng(Datos!Item) = Codigo Then Elegida = Fila
        Fila = Fila + 1
        Datos.MoveNext
    Loop
    Estado CStr(Datos.RecordCount) & " perfiles. Enter: detalle. F5: actualizar."
    Datos.Close
    If Elegida = 0 Then Elegida = 1
    GridPerfiles.Row = Elegida
    Cargando = False
    SeleccionarFila
    If Guardado Then Estado "Perfil guardado. Código: " & CStr(Codigo)
    Exit Sub
ErrorListado:
    Mensaje = Err.Description
    CerrarRecordset Datos
    Cargando = False
    LimpiarFicha
    ActualizarControles
    If Guardado Then MsgBox "Se guardó el perfil " & CStr(Codigo) & ". Falló la recarga; pulse F5 sin repetir el guardado.", vbExclamation, App.Title Else MsgBox "No se pudo consultar: " & Mensaje, vbExclamation, App.Title
End Sub
Private Sub SeleccionarFila()
    If Cargando Or Modo <> 0 Then Exit Sub
    If GridPerfiles.Row < 1 Then Exit Sub
    If Len(GridPerfiles.TextMatrix(GridPerfiles.Row, 0)) = 0 Then LimpiarFicha: ActualizarControles: Exit Sub
    CargarFicha CLng(GridPerfiles.TextMatrix(GridPerfiles.Row, 0))
    Cargando = True
    GridPerfiles.Col = 0
    GridPerfiles.ColSel = 3
    GridPerfiles.RowSel = GridPerfiles.Row
    Cargando = False
End Sub
Private Sub CmdNuevo_Click()
    On Error GoTo ErrorNuevo
    If Modo <> 0 Then Exit Sub
    LimpiarFicha
    CargarArbol 0
    Modo = 1
    MostrarDetalle
    TxtNombre.SetFocus
    Estado "Complete nombre y permisos. Guardar o Cancelar."
    Exit Sub
ErrorNuevo:
    Cargando = False
    MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdModificar_Click()
    If Modo <> 0 Or Seleccion = 0 Or Not Activo Then Exit Sub
    Modo = 2
    MostrarDetalle
    TxtNombre.SetFocus
    Estado "La modificación afectará a los usuarios vinculados."
End Sub
Private Function TextoPermisos() As String
    Dim I As Integer
    For I = 1 To 40
        If Vigentes(I) Then
            If ViewMenus.Nodes("M" & CStr(I)).Checked Then TextoPermisos = TextoPermisos & "1" Else TextoPermisos = TextoPermisos & "0"
        Else
            TextoPermisos = TextoPermisos & "0"
        End If
    Next I
End Function
Private Sub CmdGuardar_Click()
    Dim Codigo As Long, Confirmado As Boolean, Mensaje As String
    On Error GoTo ErrorGuardar
    If Modo = 0 Then Exit Sub
    If Len(Trim$(TxtNombre.Text)) = 0 Then MsgBox "Ingrese el nombre del perfil.", vbExclamation, App.Title: Exit Sub
    If MsgBox("¿Guardar el perfil " & Trim$(TxtNombre.Text) & "?" & vbCrLf & "Usuarios vinculados actualmente: " & CStr(Vinculados) & ". Los cambios se aplicarán al iniciar sesión nuevamente.", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
    Codigo = GuardarDatosPerfil(Seleccion, Version, TxtNombre.Text, TextoPermisos())
    Confirmado = True
    Modo = 0
    Cargando = True
    TxtBuscar.Text = ""
    Cargando = False
    CargarListado Codigo, True
    MostrarDetalle
    Exit Sub
ErrorGuardar:
    Mensaje = Err.Description
    If Confirmado Then
        Modo = 0
        LimpiarFicha
        ActualizarControles
        MsgBox "Se guardó el perfil " & CStr(Codigo) & ". Pulse F5 sin repetir el guardado.", vbExclamation, App.Title
    Else
        MsgBox "No se pudo confirmar el guardado. El borrador se conserva; consulte antes de reintentar: " & Mensaje, vbExclamation, App.Title
    End If
End Sub
Private Sub CmdEliminar_Click()
    Dim Codigo As Long
    On Error GoTo ErrorBaja
    If Modo <> 0 Or Seleccion = 0 Or Not Activo Then Exit Sub
    If MsgBox("¿Eliminar el perfil " & CStr(Seleccion) & " - " & TxtNombre.Text & "?" & vbCrLf & "Usuarios vinculados: " & CStr(Vinculados) & ". La baja será lógica y conservará sus usuarios y permisos. No podrá reactivarse.", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
    Codigo = GuardarDatosPerfil(Seleccion, Version, "", "", True)
    CargarListado
    MostrarListado
    Estado "Perfil eliminado lógicamente. Sus referencias se conservan."
    Exit Sub
ErrorBaja:
    MsgBox "No se pudo confirmar la baja. Recargue antes de reintentar: " & Err.Description, vbExclamation, App.Title
End Sub
Private Function ConfirmarDescarte() As Boolean
    ConfirmarDescarte = True
    If Modo <> 0 Then ConfirmarDescarte = (MsgBox("¿Descartar los cambios del perfil?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) = vbYes)
End Function
Private Sub CmdCancelar_Click()
    If Not ConfirmarDescarte Then Exit Sub
    Modo = 0
    SeleccionarFila
    MostrarListado
    Estado "Edición cancelada."
End Sub
Private Sub ViewMenus_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim Codigo As Integer, I As Integer, Padre As Integer
    If Cargando Then Exit Sub
    Codigo = CInt(Mid$(Node.Key, 2))
    Cargando = True
    If Modo = 0 Then
        Node.Checked = Consulta(Codigo)
    ElseIf Codigo = 1 Or Codigo = 2 Then
        Node.Checked = True
    ElseIf Codigo = 27 Or Codigo = 33 Or Codigo = 34 Then
        Node.Checked = False
    ElseIf Node.Checked Then
        Padre = Padres(Codigo)
        Do While Padre <> 0
            ViewMenus.Nodes("M" & CStr(Padre)).Checked = True
            Padre = Padres(Padre)
        Loop
    Else
        For I = 1 To 40
            If Vigentes(I) Then
                Padre = Padres(I)
                Do While Padre <> 0
                    If Padre = Codigo Then ViewMenus.Nodes("M" & CStr(I)).Checked = False: Exit Do
                    Padre = Padres(Padre)
                Loop
            End If
        Next I
    End If
    Cargando = False
End Sub
Private Sub Estado(ByVal Texto As String)
    StatusBar1.Panels(3).Text = Texto
End Sub
Private Sub MostrarListado()
    MostrandoListado = True
    FraListado.Visible = True
    FraFicha.Visible = False
    ActualizarControles
End Sub
Private Sub MostrarDetalle()
    MostrandoListado = False
    FraListado.Visible = False
    FraFicha.Visible = True
    FraGeneral.Visible = True
    FraAccess.Visible = True
    ActualizarControles
End Sub
Private Sub CmdListado_Click()
    If Modo = 0 Then MostrarListado
End Sub
Private Sub AbrirFichaSeleccionada()
    If Modo = 0 And Seleccion > 0 Then MostrarDetalle
End Sub
Private Sub GridPerfiles_RowColChange()
    SeleccionarFila
End Sub
Private Sub GridPerfiles_Click()
    SeleccionarFila
End Sub
Private Sub GridPerfiles_DblClick()
    AbrirFichaSeleccionada
End Sub
Private Sub GridPerfiles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 0 Then KeyCode = 0: SeleccionarFila: AbrirFichaSeleccionada
End Sub
Private Sub TxtBuscar_Change()
    If Not Cargando And Modo = 0 Then CargarListado
End Sub
Private Sub ChkEliminados_Click()
    If Not Cargando And Modo = 0 Then CargarListado
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not ConfirmarDescarte Then Cancel = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        If Modo <> 0 Then CmdCancelar_Click Else Unload Me
    ElseIf KeyCode = vbKeyF5 And Modo = 0 Then
        KeyCode = 0
        CargarListado Seleccion
    End If
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

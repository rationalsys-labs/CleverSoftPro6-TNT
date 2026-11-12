VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Clientes"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
   KeyPreview = -1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12495
   Begin VB.CheckBox ChkEliminados
      Caption = "Incluir eliminados"
      Left = 5520
      Top = 120
      Width = 1935
      Height = 285
   End
   Begin VB.ComboBox CmbFil 
      Height          =   315
      ItemData        =   "FrmClientes.frx":0000
      Left            =   7560
      List            =   "FrmClientes.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin Clever.Boton Cmd_Inf 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   8
      TX              =   "General"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14606046
      BCOLO           =   8612156
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmClientes.frx":004F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton Cmd_Det 
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      BTYPE           =   8
      TX              =   "Detalle del Cliente"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14606046
      BCOLO           =   8612156
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmClientes.frx":006B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtBusCli 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9600
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      BorderStyle = 1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   8430
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton CmdNuevo
         Caption = "Crear"
         Left = 120
         Top = 360
         Width = 1935
         Height = 375
      End
      Begin VB.CommandButton CmdModificar
         Caption = "Modificar"
         Left = 120
         Top = 840
         Width = 1935
         Height = 375
      End
      Begin VB.CommandButton CmdEliminar
         Caption = "Eliminar"
         Left = 120
         Top = 1320
         Width = 1935
         Height = 375
      End
      Begin VB.CommandButton CmdGuardar
         Caption = "Guardar"
         Left = 120
         Top = 1800
         Width = 1935
         Height = 375
      End
      Begin VB.CommandButton CmdCancelar
         Caption = "Cancelar"
         Left = 120
         Top = 2280
         Width = 1935
         Height = 375
      End
      Begin VB.CommandButton CmdSalir
         Caption = "Salir"
         Left = 120
         Top = 7920
         Width = 1935
         Height = 375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "22/03/2011"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraInf 
      Height          =   8130
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid GridCli 
         Height          =   7980
         Left            =   15
         TabIndex        =   3
         Top             =   120
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   9922
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   16185077
         BackColorBkg    =   16777215
         GridColor       =   15790320
         GridColorFixed  =   15790320
         GridLinesFixed  =   1
         MergeCells      =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
      Begin VB.Frame FraDet
         Left = 2400
         Top = 480
         Width = 9975
         Height = 8130
         Visible = 0
      Begin VB.Label LblCampo0
         Left = 120
         Top = 180
         Width = 1500
         Height = 195
         Caption = "Código:"
      End
      Begin VB.TextBox Regist
         Index = 0
         Left = 120
         Top = 420
         Width = 1500
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 20
      End
      Begin VB.Label CmbPersonaLbl
         Left = 1800
         Top = 180
         Width = 1800
         Height = 195
         Caption = "Tipo de persona:"
      End
      Begin VB.ComboBox CmbPersona
         Left = 1800
         Top = 420
         Width = 1800
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbIdentificacionLbl
         Left = 3720
         Top = 180
         Width = 2400
         Height = 195
         Caption = "Identificación:"
      End
      Begin VB.ComboBox CmbIdentificacion
         Left = 3720
         Top = 420
         Width = 2400
         Height = 315
         Style = 2
      End
      Begin VB.Label LblCampo2
         Left = 6240
         Top = 180
         Width = 3300
         Height = 195
         Caption = "Número:"
      End
      Begin VB.TextBox Regist
         Index = 2
         Left = 6240
         Top = 420
         Width = 3300
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 20
      End
      Begin VB.Label LblCampo1
         Left = 120
         Top = 780
         Width = 6480
         Height = 195
         Caption = "Nombre / razón social:"
      End
      Begin VB.TextBox Regist
         Index = 1
         Left = 120
         Top = 1020
         Width = 6480
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 255
      End
      Begin VB.Label LblCampo5
         Left = 6840
         Top = 780
         Width = 2700
         Height = 195
         Caption = "Teléfono:"
      End
      Begin VB.TextBox Regist
         Index = 5
         Left = 6840
         Top = 1020
         Width = 2700
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 50
      End
      Begin VB.Frame FraDireccion0
         Left = 120
         Top = 1440
         Width = 9735
         Height = 1920
         Caption = "Dirección principal"
      Begin VB.OptionButton OptNacional
         Index = 0
         Left = 120
         Top = 210
         Width = 1500
         Height = 285
         Caption = "Nacional"
      End
      Begin VB.OptionButton OptInternacional
         Index = 0
         Left = 1800
         Top = 210
         Width = 1800
         Height = 285
         Caption = "Internacional"
      End
      Begin VB.Label CmbPaisLbl0
         Left = 120
         Top = 630
         Width = 1680
         Height = 195
         Caption = "País:"
      End
      Begin VB.ComboBox CmbPais
         Index = 0
         Left = 120
         Top = 870
         Width = 1680
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbDepartamentoLbl0
         Left = 1980
         Top = 630
         Width = 1920
         Height = 195
         Caption = "Departamento:"
      End
      Begin VB.ComboBox CmbDepartamento
         Index = 0
         Left = 1980
         Top = 870
         Width = 1920
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbProvinciaLbl0
         Left = 4080
         Top = 630
         Width = 2040
         Height = 195
         Caption = "Provincia:"
      End
      Begin VB.ComboBox CmbProvincia
         Index = 0
         Left = 4080
         Top = 870
         Width = 2040
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbDistritoLbl0
         Left = 6300
         Top = 630
         Width = 3300
         Height = 195
         Caption = "Distrito:"
      End
      Begin VB.ComboBox CmbDistrito
         Index = 0
         Left = 6300
         Top = 870
         Width = 3300
         Height = 315
         Style = 2
      End
      Begin VB.Label LblCampo3
         Left = 120
         Top = 1230
         Width = 9450
         Height = 195
         Caption = "Calle / avenida / número / referencia:"
      End
      Begin VB.TextBox Regist
         Index = 3
         Left = 120
         Top = 1470
         Width = 9450
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 255
      End
      End
      Begin VB.Frame FraDireccion1
         Left = 120
         Top = 3540
         Width = 9735
         Height = 2520
         Caption = "Dirección de entrega (Perú)"
      Begin VB.Label CmbDepartamentoLbl1
         Left = 120
         Top = 630
         Width = 2700
         Height = 195
         Caption = "Departamento:"
      End
      Begin VB.ComboBox CmbDepartamento
         Index = 1
         Left = 120
         Top = 870
         Width = 2700
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbProvinciaLbl1
         Left = 3000
         Top = 630
         Width = 2700
         Height = 195
         Caption = "Provincia:"
      End
      Begin VB.ComboBox CmbProvincia
         Index = 1
         Left = 3000
         Top = 870
         Width = 2700
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbDistritoLbl1
         Left = 5880
         Top = 630
         Width = 3720
         Height = 195
         Caption = "Distrito:"
      End
      Begin VB.ComboBox CmbDistrito
         Index = 1
         Left = 5880
         Top = 870
         Width = 3720
         Height = 315
         Style = 2
      End
      Begin VB.Label LblCampo9
         Left = 120
         Top = 1230
         Width = 9450
         Height = 195
         Caption = "Calle / avenida / número / referencia:"
      End
      Begin VB.TextBox Regist
         Index = 9
         Left = 120
         Top = 1470
         Width = 9450
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 255
      End
      Begin VB.Label LblCampo8
         Left = 120
         Top = 1890
         Width = 6000
         Height = 195
         Caption = "Contacto de entrega:"
      End
      Begin VB.TextBox Regist
         Index = 8
         Left = 120
         Top = 2130
         Width = 6000
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 100
      End
      Begin VB.Label LblCampo11
         Left = 6600
         Top = 1890
         Width = 2970
         Height = 195
         Caption = "Teléfono de entrega:"
      End
      Begin VB.TextBox Regist
         Index = 11
         Left = 6600
         Top = 2130
         Width = 2970
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 50
      End
      End
      Begin VB.Label CmbNivelLbl
         Left = 120
         Top = 6360
         Width = 1800
         Height = 195
         Caption = "Nivel del cliente:"
      End
      Begin VB.ComboBox CmbNivel
         Left = 120
         Top = 6600
         Width = 1800
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbCobranzaLbl
         Left = 2100
         Top = 6360
         Width = 1800
         Height = 195
         Caption = "Cobranza:"
      End
      Begin VB.ComboBox CmbCobranza
         Left = 2100
         Top = 6600
         Width = 1800
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbDocumentoLbl
         Left = 4080
         Top = 6360
         Width = 1800
         Height = 195
         Caption = "Comprobante:"
      End
      Begin VB.ComboBox CmbDocumento
         Left = 4080
         Top = 6600
         Width = 1800
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbMonedaLbl
         Left = 6060
         Top = 6360
         Width = 1500
         Height = 195
         Caption = "Moneda:"
      End
      Begin VB.ComboBox CmbMoneda
         Left = 6060
         Top = 6600
         Width = 1500
         Height = 315
         Style = 2
      End
      Begin VB.Label CmbTerritorioLbl
         Left = 7740
         Top = 6360
         Width = 1800
         Height = 195
         Caption = "Territorio:"
      End
      Begin VB.ComboBox CmbTerritorio
         Left = 7740
         Top = 6600
         Width = 1800
         Height = 315
         Style = 2
      End
      Begin VB.Label LblCampo7
         Left = 120
         Top = 7005
         Width = 6000
         Height = 195
         Caption = "Contacto principal:"
      End
      Begin VB.TextBox Regist
         Index = 7
         Left = 120
         Top = 7245
         Width = 6000
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 100
      End
      Begin VB.Label LblCampo12
         Left = 6600
         Top = 7005
         Width = 1500
         Height = 195
         Caption = "Recargo (%):"
      End
      Begin VB.TextBox Regist
         Index = 12
         Left = 6600
         Top = 7245
         Width = 1500
         Height = 285
         Appearance = 0
         BorderStyle = 1
         MaxLength = 8
      End
      Begin VB.TextBox Regist
         Index = 4
         Left = 0
         Top = 0
         Width = 120
         Height = 285
         Visible = 0
      End
      Begin VB.TextBox Regist
         Index = 10
         Left = 0
         Top = 0
         Width = 120
         Height = 285
         Visible = 0
      End
      Begin VB.TextBox Regist
         Index = 13
         Left = 0
         Top = 0
         Width = 120
         Height = 285
         Visible = 0
      End
      Begin VB.TextBox Regist
         Index = 14
         Left = 0
         Top = 0
         Width = 120
         Height = 285
         Visible = 0
      End
      Begin VB.TextBox Regist
         Index = 15
         Left = 0
         Top = 0
         Width = 120
         Height = 285
         Visible = 0
      End
      Begin VB.TextBox Regist
         Index = 16
         Left = 0
         Top = 0
         Width = 120
         Height = 285
         Visible = 0
      End
   End
End
Attribute VB_Name = "FrmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Type RectanguloCliente
    Izquierda As Long
    Arriba As Long
    Derecha As Long
    Abajo As Long
End Type
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef Rectangulo As RectanguloCliente) As Long
Private CargandoUbicaciones As Boolean
Private Datos As ADODB.Recordset
Private Territorios As ADODB.Recordset
Private Seleccion As Long
Private FilaResaltada As Long
Private Version As Variant
Private Modo As Integer
Private Total As Long
Private Cargando As Boolean
Private Guardado As Boolean

Private Sub Form_Load()
    If Not AccesoPermitido(20) Then
        MsgBox "No tiene permiso para clientes.", vbExclamation, App.Title
        Unload Me
        Exit Sub
    End If
    On Error GoTo ErrorCarga
    Cargando = True
    CmbFil.ListIndex = 0
    GridCli.Cols = 5
    GridCli.FixedCols = 0
    GridCli.SelectionMode = flexSelectionByRow
    GridCli.HighLight = flexHighlightAlways
    GridCli.FocusRect = flexFocusNone
    GridCli.BackColorSel = &H8000000D
    GridCli.ForeColorSel = &H8000000E
    GridCli.AllowUserResizing = flexResizeNone
    GridCli.Rows = 2
    GridCli.FormatString = "Código|Razón social|RUC / DNI|Dirección|Estado"
    StatusBar1.Panels(1).Style = sbrText
    Cargando = False
    CargarTerritorios
    CargarCatalogos
    CargarClientes
    Exit Sub
ErrorCarga:
    Cargando = False
    Estado "No se pudo cargar. F5 permite reintentar."
End Sub

Private Sub CargarTerritorios(Optional ByVal Conservar As Boolean = False)
    Dim TerritorioAnterior As Long
    If Conservar Then TerritorioAnterior = CLng(Datos!IdTerritorio)
    Dim Cn As ADODB.Connection
    Dim Nuevos As ADODB.Recordset
    On Error GoTo ErrorCarga
    AbrirConexion Cn
    Set Nuevos = New ADODB.Recordset
    Nuevos.CursorLocation = adUseClient
    Nuevos.Open "SELECT Item,Descrip,Active FROM dbo.IntTerritorios ORDER BY Descrip", Cn, adOpenStatic, adLockReadOnly
    Set Nuevos.ActiveConnection = Nothing
    Cn.Close
    CerrarRecordset Territorios
    Set Territorios = Nuevos
    CmbTerritorio.Clear
    Do While Not Territorios.EOF
        If Territorios!Active = 1 Or CLng(Territorios!Item) = TerritorioAnterior Then
            CmbTerritorio.AddItem CStr(Territorios!Descrip) & IIf(Territorios!Active = 0, " (eliminado)", "")
            CmbTerritorio.ItemData(CmbTerritorio.NewIndex) = CLng(Territorios!Item)
        End If
        Territorios.MoveNext
    Loop
    Exit Sub
ErrorCarga:
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    Err.Raise vbObjectError + 2101, , "No se pudieron cargar los territorios."
End Sub

Private Function TextoBusqueda(ByVal Texto As String) As String
    Texto = Replace(Texto, "\", "\\")
    Texto = Replace(Texto, "%", "\%")
    Texto = Replace(Texto, "_", "\_")
    TextoBusqueda = "%" & Replace(Texto, "[", "\[") & "%"
End Function

Private Function CargarClientes() As Boolean
    Dim Cn As ADODB.Connection
    Dim Orden As ADODB.Command
    Dim Nuevos As ADODB.Recordset
    Dim Campo As String
    Dim Fila As Long
    Dim Columna As Integer
    Dim Campos As Variant
    On Error GoTo ErrorCarga
    If Modo <> 0 Then Exit Function
    Seleccion = 0
    LimpiarFicha
    ActualizarBotones
    AbrirConexion Cn
    Campos = Array("CLIENT", "RUCDNI", "DIRECC", "TELEF1", "Contac")
    Campo = CStr(Campos(CmbFil.ListIndex))
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandType = adCmdText
    Orden.CommandText = "SELECT TOP 100 c.*, t.Descrip AS Territorio, d.IdProvincia AS ProvinciaPrincipal, pr.IdDepartamento AS DepartamentoPrincipal, de.IdProvincia AS ProvinciaEntrega, pe.IdDepartamento AS DepartamentoEntrega, COUNT(*) OVER() AS Cantidad FROM dbo.DbClientes c JOIN dbo.IntTerritorios t ON c.IdTerritorio=t.Item LEFT JOIN dbo.Distritos d ON c.IdDistrito=d.Item LEFT JOIN dbo.Provincias pr ON d.IdProvincia=pr.Item LEFT JOIN dbo.Distritos de ON c.IdDistritoEntrega=de.Item LEFT JOIN dbo.Provincias pe ON de.IdProvincia=pe.Item WHERE (?=1 OR c.Activo=1) AND c.[" & Campo & "] LIKE ? ESCAPE '\' ORDER BY c.CLIENT,c.Item"
    Orden.Parameters.Append Orden.CreateParameter("Eliminados", adInteger, adParamInput, , ChkEliminados.Value)
    Orden.Parameters.Append Orden.CreateParameter("Texto", adVarWChar, adParamInput, 520, TextoBusqueda(TxtBusCli.Text))
    Set Nuevos = New ADODB.Recordset
    Nuevos.CursorLocation = adUseClient
    Nuevos.Open Orden, , adOpenStatic, adLockReadOnly
    Set Nuevos.ActiveConnection = Nothing
    Cn.Close
    CerrarRecordset Datos
    Set Datos = Nuevos
    Cargando = True
    GridCli.Clear
    GridCli.Rows = 2
    GridCli.FormatString = "Código|Razón social|RUC / DNI|Dirección|Estado"
    GridCli.TextMatrix(1, 0) = ""
    Total = 0
    If Not Datos.EOF Then
        Total = CLng(Datos!Cantidad)
        GridCli.Rows = Datos.RecordCount + 1
        Fila = 1
        Do While Not Datos.EOF
            GridCli.TextMatrix(Fila, 0) = CStr(Datos!Asient)
            GridCli.TextMatrix(Fila, 1) = CStr(Datos!CLIENT)
            GridCli.TextMatrix(Fila, 2) = CStr(Datos!RUCDNI)
            GridCli.TextMatrix(Fila, 3) = CStr(Datos!DIRECC)
            GridCli.TextMatrix(Fila, 4) = IIf(Datos!Activo, "Activo", "Eliminado")
            Fila = Fila + 1
            Datos.MoveNext
        Loop
        GridCli.Row = 1
    End If
    AjustarColumnas
    Cargando = False
    SeleccionarFila
    ActualizarBotones
    CargarClientes = True
    Estado CStr(Datos.RecordCount) & " mostrados / " & CStr(Total) & " encontrados."
    Exit Function
ErrorCarga:
    Cargando = False
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    CerrarRecordset Datos
    GridCli.Clear
    Seleccion = 0
    LimpiarFicha
    ActualizarBotones
    Estado "No se pudo consultar. F5 permite reintentar."
End Function

Private Sub SeleccionarFila()
    Dim I As Integer
    Dim N As Integer
    Dim Campos As Variant
    Dim FilaActual As Long
    If Cargando Or Modo <> 0 Then Exit Sub
    Seleccion = 0
    LimpiarFicha
    If Datos Is Nothing Then Exit Sub
    If Datos.State <> adStateOpen Then Exit Sub
    If Datos.RecordCount = 0 Or GridCli.Row < 1 Then Exit Sub
    Datos.MoveFirst
    Datos.Move GridCli.Row - 1
    If Datos.EOF Then Exit Sub
    Seleccion = CLng(Datos!Item)
    Version = Datos!VersionFila
    Campos = Array("Asient", "CLIENT", "RUCDNI", "DIRECC", "IdDistrito", "TELEF1", "", "Contac", "ContacEnt", "DireccEnt", "IdDistritoEntrega", "TelefoEnt", "Recarg", "IdComprobante", "IdMoneda", "IdNivel", "IdCobranza")
    For I = 0 To 16
        If I <> 6 Then
            If Not IsNull(Datos.Fields(CStr(Campos(I))).Value) Then Regist(I).Text = CStr(Datos.Fields(CStr(Campos(I))).Value)
        End If
    Next I
    CargarTerritorios True
    For N = 0 To CmbTerritorio.ListCount - 1
        If CmbTerritorio.ItemData(N) = CLng(Datos!IdTerritorio) Then CmbTerritorio.ListIndex = N
    Next N
    MostrarCatalogos
    Cargando = True
    FilaActual = GridCli.Row
    GridCli.FillStyle = flexFillRepeat
    If FilaResaltada > 0 And FilaResaltada < GridCli.Rows Then
        GridCli.Row = FilaResaltada
        GridCli.RowSel = FilaResaltada
        GridCli.Col = 0
        GridCli.ColSel = GridCli.Cols - 1
        GridCli.CellBackColor = GridCli.BackColor
        GridCli.CellForeColor = GridCli.ForeColor
    End If
    GridCli.Row = FilaActual
    GridCli.RowSel = FilaActual
    GridCli.Col = 0
    GridCli.ColSel = GridCli.Cols - 1
    GridCli.CellBackColor = &H83693C
    GridCli.CellForeColor = &HFFFFFF
    GridCli.FillStyle = flexFillSingle
    FilaResaltada = FilaActual
    Cargando = False
    ActualizarBotones
End Sub

Private Sub LimpiarFicha()
    Dim I As Integer
    For I = 0 To 16
        If I <> 6 Then Regist(I).Text = ""
    Next I
    CmbTerritorio.ListIndex = -1
    CargandoUbicaciones = True
    CmbDocumento.ListIndex = -1
    CmbMoneda.ListIndex = -1
    CmbNivel.ListIndex = -1
    CmbCobranza.ListIndex = -1
    CmbPersona.ListIndex = -1
    CmbIdentificacion.Clear
    CmbPais(0).Clear
    CmbDepartamento(0).Clear
    CmbDepartamento(1).Clear
    CmbProvincia(0).Clear
    CmbProvincia(1).Clear
    CmbDistrito(0).Clear
    CmbDistrito(1).Clear
    CargandoUbicaciones = False
End Sub

Private Sub ActualizarBotones()
    Dim I As Integer
    Dim Activo As Boolean
    If Seleccion <> 0 Then Activo = CBool(Datos!Activo)
    CmdNuevo.Enabled = (Modo = 0)
    CmdModificar.Enabled = (Modo = 0 And Seleccion <> 0 And Activo)
    If Seleccion <> 0 Then CmdModificar.Enabled = CmdModificar.Enabled And CBool(Datos!EsLocal)
    CmdEliminar.Enabled = CmdModificar.Enabled
    CmdGuardar.Visible = (Modo <> 0)
    CmdCancelar.Visible = (Modo <> 0)
    Cmd_Det.Enabled = (Seleccion <> 0 Or Modo <> 0)
    Cmd_Inf.Enabled = (Modo = 0)
    If GridCli.Enabled <> (Modo = 0) Then GridCli.Enabled = (Modo = 0)
    TxtBusCli.Enabled = (Modo = 0 And Not FraDet.Visible)
    CmbFil.Enabled = (Modo = 0 And Not FraDet.Visible)
    ChkEliminados.Enabled = (Modo = 0 And Not FraDet.Visible)
    For I = 0 To 16
        If I <> 6 Then Regist(I).Locked = (Modo = 0)
    Next I
    Regist(0).Locked = True
    CmbTerritorio.Enabled = (Modo <> 0)
    HabilitarUbicaciones
    StatusBar1.Panels(1).Text = Choose(Modo + 1, "Consulta", "Crear", "Modificar")
    StatusBar1.Panels(2).Text = IIf(Seleccion = 0, "Sin selección", "Cliente: " & Regist(0).Text)
    If Seleccion <> 0 Then StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & IIf(CBool(Datos!EsLocal), " - Local", " - Importado (consulta)")
End Sub

Private Sub Estado(ByVal Mensaje As String)
    StatusBar1.Panels(3).Text = Mensaje
End Sub

Private Sub CmdNuevo_Click()
    If Not AccesoPermitido(20) Then Exit Sub
    CargarTerritorios
    CargarCatalogos
    Modo = 1
    LimpiarFicha
    Regist(12).Text = "0"
    PredeterminarLima 0
    PredeterminarLima 1
    ElegirId CmbNivel, 1
    ElegirId CmbCobranza, 1
    ElegirId CmbMoneda, 1
    ElegirId CmbPersona, 1
    CargarIdentificaciones
    ElegirId CmbIdentificacion, 1
    MostrarDetalle
    ActualizarBotones
    Estado "Código local automático al guardar. Complete la ficha del cliente."
    Regist(1).SetFocus
End Sub

Private Sub CmdModificar_Click()
    If Not AccesoPermitido(20) Or Seleccion = 0 Then Exit Sub
    SeleccionarFila
    If Not CBool(Datos!EsLocal) Then
        MsgBox "El cliente importado solo puede actualizarse mediante carga.", vbInformation, App.Title
        Exit Sub
    End If
    Modo = 2
    MostrarDetalle
    ActualizarBotones
    Regist(1).SetFocus
End Sub

Private Function ConfirmarDescarte() As Boolean
    ConfirmarDescarte = True
    If Modo <> 0 Then ConfirmarDescarte = (MsgBox("¿Desea descartar los datos de la edición?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) = vbYes)
End Function

Private Sub CmdCancelar_Click()
    If Not ConfirmarDescarte Then Exit Sub
    Modo = 0
    SeleccionarFila
    FraDet.Visible = False
    FraInf.Visible = True
    ActualizarBotones
    Estado "Edición cancelada."
End Sub

Private Sub CmdGuardar_Click()
    Dim Cn As ADODB.Connection
    Dim Orden As ADODB.Command
    Dim Resultado As ADODB.Recordset
    Dim Campos As Variant
    Dim Limites As Variant
    Dim I As Integer
    Dim NuevoItem As Long
    On Error GoTo ErrorGuardado
    If Modo = 0 Or Not AccesoPermitido(20) Then Exit Sub
    If CmbTerritorio.ListIndex < 0 Then
        MsgBox "Seleccione el territorio.", vbExclamation, App.Title
        CmbTerritorio.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Regist(12).Text) Then
        MsgBox "El recargo debe ser numérico.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("¿Guardar el cliente " & Trim$(Regist(0).Text) & " - " & Trim$(Regist(1).Text) & "?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
    Guardado = False
    CmdGuardar.Enabled = False
    AbrirConexion Cn
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandType = adCmdStoredProc
    Orden.CommandText = "dbo.GuardarCliente"
    Orden.Parameters.Append Orden.CreateParameter("@Item", adInteger, adParamInput, , IIf(Modo = 1, 0, Seleccion))
    Orden.Parameters.Append Orden.CreateParameter("@Version", adBinary, adParamInput, 8, IIf(Modo = 1, Null, Version))
    Campos = Array("Asient", "CLIENT", "RUCDNI", "DIRECC", "IdDistrito", "TELEF1", "IdTerritorio", "Contac", "ContacEnt", "DireccEnt", "IdDistritoEntrega", "TelefoEnt", "Recarg", "IdComprobante", "IdMoneda", "IdNivel", "IdCobranza")
    Limites = Array(20, 255, 20, 255, 50, 50, 0, 100, 100, 255, 50, 50, 0, 10, 5, 10, 10)
    For I = 0 To 16
        Select Case I
            Case 4, 10: ' La ubicación se guarda mediante FK, no como texto.
            Case 13: Orden.Parameters.Append Orden.CreateParameter("@IdComprobante", adInteger, adParamInput, , IdCombo(CmbDocumento))
            Case 14: Orden.Parameters.Append Orden.CreateParameter("@IdMoneda", adInteger, adParamInput, , IdCombo(CmbMoneda))
            Case 15: Orden.Parameters.Append Orden.CreateParameter("@IdNivel", adInteger, adParamInput, , IdCombo(CmbNivel))
            Case 16: Orden.Parameters.Append Orden.CreateParameter("@IdCobranza", adInteger, adParamInput, , IdCombo(CmbCobranza))
            Case 6: Orden.Parameters.Append Orden.CreateParameter("@IdTerritorio", adInteger, adParamInput, , CmbTerritorio.ItemData(CmbTerritorio.ListIndex))
            Case 12: Orden.Parameters.Append Orden.CreateParameter("@Recarg", adCurrency, adParamInput, , CCur(Regist(12).Text))
            Case Else: Orden.Parameters.Append Orden.CreateParameter("@" & CStr(Campos(I)), adVarWChar, adParamInput, CLng(Limites(I)), Trim$(Regist(I).Text))
        End Select
    Next I
    AgregarUbicacion Orden, 0
    AgregarUbicacion Orden, 1
    Orden.Parameters.Append Orden.CreateParameter("@IdTipoPersona", adInteger, adParamInput, , IdCombo(CmbPersona))
    Orden.Parameters.Append Orden.CreateParameter("@IdTipoIdentificacion", adInteger, adParamInput, , IdCombo(CmbIdentificacion))
    Set Resultado = Orden.Execute
    Guardado = True
    NuevoItem = CLng(Resultado!Item)
    CerrarRecordset Resultado
    Cn.Close
    Modo = 0
    CmdGuardar.Enabled = True
    If CargarClientes Then
        Estado "Guardado el cliente " & CStr(NuevoItem) & ". Consulta actualizada."
    Else
        Estado "Guardado el cliente " & CStr(NuevoItem) & ". Pulse F5 para recargar; no repita el guardado."
    End If
    Exit Sub
ErrorGuardado:
    Dim Mensaje As String
    Mensaje = Err.Description
    CerrarRecordset Resultado
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    CmdGuardar.Enabled = True
    If Guardado Then
        Modo = 0
        Seleccion = 0
        ActualizarBotones
        MsgBox "El cliente se guardó. No repita el guardado; pulse F5 para recargar.", vbExclamation, App.Title
    Else
        MsgBox "No se pudo confirmar el guardado. Revise la consulta antes de reintentar: " & Mensaje & vbCrLf & "Los datos de la edición se conservan.", vbExclamation, App.Title
    End If
End Sub

Private Sub CmdEliminar_Click()
    Dim Cn As ADODB.Connection
    Dim Orden As ADODB.Command
    On Error GoTo ErrorBaja
    If Modo <> 0 Or Seleccion = 0 Or Not AccesoPermitido(20) Then Exit Sub
    If MsgBox("¿Eliminar el cliente " & Regist(0).Text & " - " & Regist(1).Text & "?" & vbCrLf & "La baja será lógica. Sus documentos y referencias se conservarán.", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
    AbrirConexion Cn
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandType = adCmdStoredProc
    Orden.CommandText = "dbo.EliminarCliente"
    Orden.Parameters.Append Orden.CreateParameter("@Item", adInteger, adParamInput, , Seleccion)
    Orden.Parameters.Append Orden.CreateParameter("@Version", adBinary, adParamInput, 8, Version)
    Orden.Execute , , adExecuteNoRecords
    Cn.Close
    CargarClientes
    Estado "Cliente eliminado lógicamente. Sus referencias se conservan."
    Exit Sub
ErrorBaja:
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    MsgBox "No se pudo confirmar la baja. Recargue la consulta: " & Err.Description, vbExclamation, App.Title
End Sub

Private Sub MostrarDetalle()
    FraInf.Visible = False
    FraDet.Visible = True
    ActualizarBotones
End Sub
Private Sub Cmd_Det_Click()
    If Seleccion <> 0 Then MostrarDetalle
End Sub
Private Sub Cmd_Inf_Click()
    If Modo <> 0 Then Exit Sub
    FraInf.Visible = True
    FraDet.Visible = False
    ActualizarBotones
End Sub
Private Sub GridCli_RowColChange()
    SeleccionarFila
End Sub
Private Sub GridCli_Click()
    SeleccionarFila
End Sub
Private Sub GridCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Or Shift <> 0 Then Exit Sub
    KeyCode = 0
    If Cargando Or Modo <> 0 Then Exit Sub
    SeleccionarFila
    Cmd_Det_Click
End Sub
Private Sub TxtBusCli_Change()
    If Cargando Or Modo <> 0 Then Exit Sub
    CargarClientes
End Sub
Private Sub CmbFil_Click()
    If Cargando Or Modo <> 0 Then Exit Sub
    CargarClientes
End Sub
Private Sub ChkEliminados_Click()
    CargarClientes
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not ConfirmarDescarte Then Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CerrarRecordset Datos
    CerrarRecordset Territorios
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        If Modo <> 0 Then CmdCancelar_Click Else Unload Me
    ElseIf KeyCode = vbKeyF5 And Modo = 0 Then
        KeyCode = 0
        On Error GoTo ErrorRecarga
        CargarTerritorios
        CargarCatalogos
        CargarClientes
    End If
    Exit Sub
ErrorRecarga:
    Estado "No se pudo recargar. Puede reintentar con F5."
End Sub

Private Sub AjustarColumnas()
    Dim Rectangulo As RectanguloCliente
    Dim Disponible As Long
    Dim AnchoCodigo As Long
    Dim AnchoDocumento As Long
    Dim AnchoEstado As Long
    Dim Fila As Long
    Dim Ancho As Long
    Dim Resto As Long
    GetClientRect GridCli.hWnd, Rectangulo
    Disponible = Rectangulo.Derecha * Screen.TwipsPerPixelX - 30
    AnchoCodigo = TextWidth("Código") + 300
    AnchoDocumento = TextWidth("00000000000") + 300
    AnchoEstado = TextWidth("Eliminado") + 300
    For Fila = 1 To GridCli.Rows - 1
        Ancho = TextWidth(GridCli.TextMatrix(Fila, 0)) + 300
        If Ancho > AnchoCodigo Then AnchoCodigo = Ancho
        Ancho = TextWidth(GridCli.TextMatrix(Fila, 2)) + 300
        If Ancho > AnchoDocumento Then AnchoDocumento = Ancho
    Next Fila
    Resto = Disponible - AnchoCodigo - AnchoDocumento - AnchoEstado
    GridCli.ColWidth(0) = AnchoCodigo
    GridCli.ColWidth(2) = AnchoDocumento
    GridCli.ColWidth(4) = AnchoEstado
    GridCli.ColWidth(1) = Resto \ 2
    GridCli.ColWidth(3) = Resto - GridCli.ColWidth(1)
End Sub
Private Sub CargarCombo(ByVal Consulta As String, ByVal Lista As Object, Optional ByVal Filtro As Long = -1)
    Dim Cn As ADODB.Connection
    Dim Orden As ADODB.Command
    Dim R As ADODB.Recordset
    On Error GoTo ErrorCombo
    AbrirConexion Cn
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandText = Consulta
    Orden.CommandType = adCmdText
    If Filtro >= 0 Then Orden.Parameters.Append Orden.CreateParameter("Filtro", adInteger, adParamInput, , Filtro)
    Set R = Orden.Execute
    Lista.Clear
    Do While Not R.EOF
        Lista.AddItem CStr(R!Nombre)
        Lista.ItemData(Lista.NewIndex) = CLng(R!Item)
        R.MoveNext
    Loop
    CerrarRecordset R
    Cn.Close
    Exit Sub
ErrorCombo:
    CerrarRecordset R
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    Err.Raise vbObjectError + 2401, , "No se pudo cargar el catálogo. F5 permite reintentar."
End Sub

Private Function IdCombo(ByVal Lista As Object) As Variant
    IdCombo = Null
    If Lista.ListIndex >= 0 Then IdCombo = CLng(Lista.ItemData(Lista.ListIndex))
End Function

Private Sub ElegirId(ByVal Lista As Object, ByVal Valor As Variant)
    Dim I As Integer
    Lista.ListIndex = -1
    If IsNull(Valor) Then Exit Sub
    For I = 0 To Lista.ListCount - 1
        If Lista.ItemData(I) = CLng(Valor) Then Lista.ListIndex = I: Exit For
    Next I
End Sub

Private Sub CargarCatalogos(Optional ByVal Conservar As Boolean = False)
    Dim Nivel As Long, Cobranza As Long, Moneda As Long, Documento As Long
    If Conservar Then
        Nivel = CLng(Datos!IdNivel)
        Cobranza = CLng(Datos!IdCobranza)
        Moneda = CLng(Datos!IdMoneda)
        Documento = CLng(Datos!IdComprobante)
    End If
    CargandoUbicaciones = True
    CargarCombo "SELECT Item,Nombre FROM NivelesCliente WHERE Activo=1 OR Item=? ORDER BY Item", CmbNivel, Nivel
    CargarCombo "SELECT Item,Nombre FROM TiposCobranza WHERE Activo=1 OR Item=? ORDER BY Item", CmbCobranza, Cobranza
    CargarCombo "SELECT CodIte AS Item,NamVar AS Nombre FROM Int_TipDoc WHERE CodIte IN(1,2)AND(Activo=1 OR CodIte=?)ORDER BY CodIte", CmbDocumento, Documento
    CargarCombo "SELECT CodEst AS Item,Moneda AS Nombre FROM Int_Moneda WHERE ActEst=1 OR CodEst=? ORDER BY CodEst", CmbMoneda, Moneda
    CargarCombo "SELECT Item,Nombre FROM TiposPersona WHERE Activo=1 ORDER BY Item", CmbPersona
    CargandoUbicaciones = False
End Sub

Private Sub CargarIdentificaciones()
    If CmbPersona.ListIndex < 0 Then CmbIdentificacion.Clear: Exit Sub
    If CLng(IdCombo(CmbPersona)) = 1 Then
        CargarCombo "SELECT Item,Nombre FROM TiposIdentificacion WHERE Activo=1 AND Item IN(1,2,3,4)ORDER BY Item", CmbIdentificacion
    Else
        CargarCombo "SELECT Item,Nombre FROM TiposIdentificacion WHERE Activo=1 AND Item IN(2,5)ORDER BY Item", CmbIdentificacion
    End If
End Sub

Private Sub PrepararUbicacion(ByVal Indice As Integer, ByVal Internacional As Boolean)
    CargandoUbicaciones = True
    If Indice = 1 Then Internacional = False
    If Indice = 0 Then OptNacional(0).Value = Not Internacional
    If Indice = 0 Then OptInternacional(0).Value = Internacional
    If Internacional Then
        CargarCombo "SELECT Item,Nombre FROM Paises WHERE Activo=1 AND Item<>1 ORDER BY Nombre", CmbPais(Indice)
        CmbDepartamento(Indice).Clear
        CmbProvincia(Indice).Clear
        CmbDistrito(Indice).Clear
    Else
        If Indice = 0 Then CargarCombo "SELECT Item,Nombre FROM Paises WHERE Activo=1 AND Item=1", CmbPais(0)
        If Indice = 0 Then ElegirId CmbPais(0), 1
        CargarCombo "SELECT Item,Nombre FROM Departamentos WHERE Activo=1 AND IdPais=1 ORDER BY Nombre", CmbDepartamento(Indice)
        CmbProvincia(Indice).Clear
        CmbDistrito(Indice).Clear
    End If
    CargandoUbicaciones = False
    HabilitarUbicaciones
End Sub

Private Sub CargarProvincias(ByVal Indice As Integer)
    CmbProvincia(Indice).Clear
    CmbDistrito(Indice).Clear
    If CmbDepartamento(Indice).ListIndex < 0 Then Exit Sub
    CargarCombo "SELECT Item,Nombre FROM Provincias WHERE Activo=1 AND IdDepartamento=? ORDER BY Nombre", CmbProvincia(Indice), CLng(IdCombo(CmbDepartamento(Indice)))
End Sub

Private Sub CargarDistritos(ByVal Indice As Integer)
    CmbDistrito(Indice).Clear
    If CmbProvincia(Indice).ListIndex < 0 Then Exit Sub
    CargarCombo "SELECT Item,Nombre FROM Distritos WHERE Activo=1 AND IdProvincia=? ORDER BY Nombre", CmbDistrito(Indice), CLng(IdCombo(CmbProvincia(Indice)))
End Sub

Private Sub PredeterminarLima(ByVal Indice As Integer)
    PrepararUbicacion Indice, False
    CargandoUbicaciones = True
    ElegirId CmbDepartamento(Indice), 15
    CargarProvincias Indice
    ElegirId CmbProvincia(Indice), 1501
    CargarDistritos Indice
    ElegirId CmbDistrito(Indice), 150101
    CargandoUbicaciones = False
End Sub

Private Sub MostrarUbicacion(ByVal Indice As Integer)
    Dim Sufijo As String
    Dim Internacional As Boolean
    Sufijo = IIf(Indice = 0, "", "Entrega")
    If Indice = 0 Then Internacional = CBool(Datos!EsInternacional)
    PrepararUbicacion Indice, Internacional
    CargandoUbicaciones = True
    If Internacional Then
        ElegirId CmbPais(Indice), Datos.Fields("IdPais" & Sufijo).Value
    Else
        ElegirId CmbDepartamento(Indice), Datos.Fields(IIf(Indice = 0, "DepartamentoPrincipal", "DepartamentoEntrega")).Value
        CargarProvincias Indice
        ElegirId CmbProvincia(Indice), Datos.Fields(IIf(Indice = 0, "ProvinciaPrincipal", "ProvinciaEntrega")).Value
        CargarDistritos Indice
        ElegirId CmbDistrito(Indice), Datos.Fields("IdDistrito" & Sufijo).Value
    End If
    CargandoUbicaciones = False
End Sub

Private Sub MostrarCatalogos()
    CargarCatalogos True
    CargandoUbicaciones = True
    ElegirId CmbNivel, Datos!IdNivel
    ElegirId CmbCobranza, Datos!IdCobranza
    ElegirId CmbDocumento, Datos!IdComprobante
    ElegirId CmbMoneda, Datos!IdMoneda
    ElegirId CmbPersona, Datos!IdTipoPersona
    CargarIdentificaciones
    ElegirId CmbIdentificacion, Datos!IdTipoIdentificacion
    CargandoUbicaciones = False
    MostrarUbicacion 0
    MostrarUbicacion 1
End Sub

Private Sub HabilitarUbicaciones()
    Dim I As Integer
    Dim Nacional As Boolean
    For I = 0 To 1
        Nacional = True
        If I = 0 Then
            Nacional = OptNacional(0).Value
            OptNacional(0).Enabled = (Modo <> 0)
            OptInternacional(0).Enabled = (Modo <> 0)
            CmbPais(0).Enabled = (Modo <> 0)
        End If
        CmbDepartamento(I).Enabled = (Modo <> 0 And Nacional)
        CmbProvincia(I).Enabled = (Modo <> 0 And Nacional)
        CmbDistrito(I).Enabled = (Modo <> 0 And Nacional)
    Next I
    CmbNivel.Enabled = (Modo <> 0)
    CmbCobranza.Enabled = (Modo <> 0)
    CmbDocumento.Enabled = (Modo <> 0)
    CmbMoneda.Enabled = (Modo <> 0)
    CmbPersona.Enabled = (Modo <> 0)
    CmbIdentificacion.Enabled = (Modo <> 0)
End Sub

Private Sub AgregarUbicacion(ByVal Orden As ADODB.Command, ByVal Indice As Integer)
    Dim Pais As Variant
    Dim Distrito As Variant
    If Indice = 1 Then
        Orden.Parameters.Append Orden.CreateParameter("@IdDistritoEntrega", adInteger, adParamInput, , IdCombo(CmbDistrito(1)))
        Exit Sub
    End If
    Pais = Null
    Distrito = Null
    If OptInternacional(0).Value Then Pais = IdCombo(CmbPais(0)) Else Distrito = IdCombo(CmbDistrito(0))
    Orden.Parameters.Append Orden.CreateParameter("@EsInternacional", adBoolean, adParamInput, , OptInternacional(0).Value)
    Orden.Parameters.Append Orden.CreateParameter("@IdPais", adInteger, adParamInput, , Pais)
    Orden.Parameters.Append Orden.CreateParameter("@IdDistrito", adInteger, adParamInput, , Distrito)
End Sub
Private Sub OptNacional_Click(Index As Integer)
    If CargandoUbicaciones Then Exit Sub
    On Error GoTo ErrorUbicacion
    PredeterminarLima Index
    Exit Sub
ErrorUbicacion:
    CargandoUbicaciones = False
    MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub OptInternacional_Click(Index As Integer)
    If CargandoUbicaciones Then Exit Sub
    On Error GoTo ErrorUbicacion
    PrepararUbicacion Index, True
    Exit Sub
ErrorUbicacion:
    CargandoUbicaciones = False
    MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmbDepartamento_Click(Index As Integer)
    If CargandoUbicaciones Then Exit Sub
    On Error GoTo ErrorUbicacion
    CargandoUbicaciones = True
    CargarProvincias Index
    CargandoUbicaciones = False
    Exit Sub
ErrorUbicacion:
    CargandoUbicaciones = False
    MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmbProvincia_Click(Index As Integer)
    If CargandoUbicaciones Then Exit Sub
    On Error GoTo ErrorUbicacion
    CargandoUbicaciones = True
    CargarDistritos Index
    CargandoUbicaciones = False
    Exit Sub
ErrorUbicacion:
    CargandoUbicaciones = False
    MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmbPersona_Click()
    If CargandoUbicaciones Then Exit Sub
    On Error GoTo ErrorPersona
    CargarIdentificaciones
    If CmbIdentificacion.ListCount > 0 Then CmbIdentificacion.ListIndex = 0
    Exit Sub
ErrorPersona:
    MsgBox Err.Description, vbExclamation, App.Title
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de Usuario"
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
   Begin Clever.Boton CmdDetalle 
      TabIndex = 10
      Height          =   420
      Left            =   90
      Top             =   450
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Detalle del usuario"
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
      TX = "Crear usuario"
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
      TX = "Modificar usuario"
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
   Begin Clever.Boton CmdClave 
      TabIndex = 13
      Height          =   420
      Left            =   90
      Top             =   2340
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Cambiar clave"
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
      Top             =   2970
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Eliminar usuario"
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
      Caption = "Sin usuario seleccionado"
   End
   Begin VB.Frame FraFicha
      Left = 2565
      Top = 360
      Width = 9735
      Height = 7650
      Caption = "Ficha del usuario"
      Visible = 0
   Begin Clever.Boton CmdGeneral 
      TabIndex = 15
      Height          =   285
      Left            =   135
      Top             =   240
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Datos Generales"
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
   Begin Clever.Boton CmdAccesos 
      TabIndex = 16
      Height          =   285
      Left            =   2580
      Top             =   240
      Width           =   2310
      _ExtentX        =   3122
      _ExtentY        =   741
      BTYPE           =   5
      TX = "Accesos de Usuario"
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
   Begin VB.Frame FraGeneral
      Left = 90
      Top = 750
      Width = 9540
      Height = 6750
      Caption = "Datos generales"
   Begin VB.Label LblUsuario
      Left = 240
      Top = 540
      Width = 1680
      Height = 240
      Caption = "Usuario de acceso:"
   End
   Begin VB.TextBox TxtUsuario
      TabIndex = 0
      Left = 2100
      Top = 510
      Width = 7170
      Height = 285
      Appearance = 0
      MaxLength = 5
   End
   Begin VB.Label LblNombre
      Left = 240
      Top = 1140
      Width = 1680
      Height = 240
      Caption = "Nombres:"
   End
   Begin VB.TextBox TxtNombre
      TabIndex = 1
      Left = 2100
      Top = 1110
      Width = 7170
      Height = 285
      Appearance = 0
      MaxLength = 50
   End
   Begin VB.Label LblApellido
      Left = 240
      Top = 1740
      Width = 1680
      Height = 240
      Caption = "Apellidos:"
   End
   Begin VB.TextBox TxtApellido
      TabIndex = 2
      Left = 2100
      Top = 1710
      Width = 7170
      Height = 285
      Appearance = 0
      MaxLength = 50
   End
   Begin VB.Label LblCorreo
      Left = 240
      Top = 2340
      Width = 1680
      Height = 240
      Caption = "Correo:"
   End
   Begin VB.TextBox TxtCorreo
      TabIndex = 3
      Left = 2100
      Top = 2310
      Width = 7170
      Height = 285
      Appearance = 0
      MaxLength = 100
   End
   Begin VB.Label LblPerfil
      Left = 240
      Top = 2940
      Width = 1550
      Height = 240
      Caption = "Tipo de usuario:"
   End
   Begin VB.ComboBox CmbPerfil
      TabIndex = 4
      Left = 2100
      Top = 2910
      Width = 7170
      Height = 315
      Style = 2
   End
   Begin VB.Label LblPerfilAcceso
      Left = 240
      Top = 3480
      Width = 1680
      Height = 240
      Caption = "Perfil de acceso:"
   End
   Begin VB.ComboBox CmbPerfilAcceso
      Left = 2100
      Top = 3450
      Width = 7170
      Height = 315
      Style = 2
   End
   Begin VB.Label LblEstado
      Left = 240
      Top = 4020
      Width = 1550
      Height = 240
      Caption = "Estado:"
   End
   Begin VB.Label LblEstadoUsuario
      Left = 2100
      Top = 4020
      Width = 3570
      Height = 240
      Caption = "Sin selección"
   End
   Begin VB.Label LblClave
      Left = 240
      Top = 4560
      Width = 1550
      Height = 240
      Caption = "Clave nueva:"
   End
   Begin VB.TextBox TxtClave
      TabIndex = 5
      Left = 2100
      Top = 4530
      Width = 7170
      Height = 285
      Appearance = 0
      MaxLength = 64
      PasswordChar = "*"
   End
   Begin VB.Label LblConfirmacion
      Left = 240
      Top = 5100
      Width = 1550
      Height = 240
      Caption = "Confirmar clave:"
   End
   Begin VB.TextBox TxtConfirmacion
      TabIndex = 6
      Left = 2100
      Top = 5070
      Width = 7170
      Height = 285
      Appearance = 0
      MaxLength = 64
      PasswordChar = "*"
   End
   Begin VB.Label LblAyuda
      Left = 240
      Top = 5880
      Width = 5460
      Height = 240
      Caption = "El usuario de acceso se conserva al modificar."
   End
   End
   Begin VB.Frame FraAccess
      Left = 90
      Top = 750
      Width = 9540
      Height = 6750
      Caption = "Accesos de usuario"
      Visible = 0
   Begin VB.Label LblPermisos
      Left = 135
      Top = 315
      Width = 5700
      Height = 240
      Caption = "Permisos de menús y módulos"
   End
   Begin MSComctlLib.TreeView ViewMenus
      TabIndex = 7
      Left = 135
      Top = 750
      Width = 9255
      Height = 5700
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
      Caption = "Usuarios del Sistema"
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
   Begin MSFlexGridLib.MSFlexGrid GridUsuarios
      Left = 135
      Top = 900
      Width = 9450
      Height = 6465
      Rows = 2
      Cols = 5
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
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Modo As Integer
Private Seleccion As Long
Private Version As Variant
Private UsuarioActivo As Boolean
Private Cargando As Boolean
Private Padres(1 To 40) As Integer
Private PermisosConsulta(1 To 40) As Boolean
Private ListaUsuarios As ADODB.Recordset
Private MostrandoListado As Boolean
Private ModulosVigentes(1 To 40) As Boolean
Private VersionPerfilElegido As Variant

Private Sub Form_Load()
    If Not PuedeAdministrarUsuarios Then
        MsgBox "Solo el administrador autorizado puede gestionar usuarios.", vbExclamation, App.Title
        Unload Me
        Exit Sub
    End If
    On Error GoTo ErrorCarga
    Cargando = True
    CmbPerfil.AddItem "Administrador"
    CmbPerfil.AddItem "Operador"
    CmbPerfil.ListIndex = 1
    Cargando = False
    CargarPerfilesUsuario
    LimpiarFicha
    MostrandoListado = True
    ConfigurarGrilla
    CargarListado
    MostrarListado
    Exit Sub
ErrorCarga:
    MsgBox "No se pudo abrir Usuarios: " & Err.Description, vbExclamation, App.Title
    Unload Me
End Sub

Private Sub LimpiarFicha()
    Cargando = True
    Seleccion = 0
    ViewMenus.Nodes.Clear
    Version = Null
    VersionPerfilElegido = Null
    If CmbPerfilAcceso.ListCount > 0 Then CmbPerfilAcceso.ListIndex = 0
    UsuarioActivo = False
    TxtUsuario.Text = ""
    TxtNombre.Text = ""
    TxtApellido.Text = ""
    TxtCorreo.Text = ""
    TxtClave.Text = ""
    TxtConfirmacion.Text = ""
    CmbPerfil.ListIndex = 1
    LblEstadoUsuario.Caption = "Sin selección"
    LblSeleccion.Caption = "Sin usuario seleccionado"
    Cargando = False
End Sub

Private Sub CargarArbol(ByVal Codigo As Long, Optional ByVal CodigoPerfil As Long = -1)
    Dim Datos As ADODB.Recordset, Nodo As MSComctlLib.Node, I As Integer, Padre As Integer, Texto As String
    If CodigoPerfil >= 0 Then Set Datos = ListarPermisosPerfil(CodigoPerfil) Else Set Datos = ConsultarPermisosUsuario(Codigo)
    Cargando = True
    ViewMenus.Nodes.Clear
    VersionPerfilElegido = Null
    Do While Not Datos.EOF
        I = CInt(Datos!Item)
        Padre = 0
        If Not IsNull(Datos!IdPadre) Then Padre = CInt(Datos!IdPadre)
        Padres(I) = Padre
        ModulosVigentes(I) = CBool(Datos!Vigente)
        PermisosConsulta(I) = CBool(Datos!Acceso)
        If Not IsNull(Datos!VersionPerfil) Then VersionPerfilElegido = Datos!VersionPerfil
        If ModulosVigentes(I) Then
            Texto = CStr(Datos!Nombre)
            If Not FuncionDisponible(I) Then Texto = Texto & " (pendiente)"
            If Padre = 0 Then Set Nodo = ViewMenus.Nodes.Add(, , "M" & CStr(I), Texto) Else Set Nodo = ViewMenus.Nodes.Add("M" & CStr(Padre), tvwChild, "M" & CStr(I), Texto)
            Nodo.Checked = PermisosConsulta(I)
            Nodo.Expanded = True
        End If
        Datos.MoveNext
    Loop
    Datos.Close
    Cargando = False
End Sub

Private Sub CargarPerfilesUsuario(Optional ByVal Actual As Long = 0)
    Dim Datos As ADODB.Recordset
    Cargando = True
    CmbPerfilAcceso.Clear
    CmbPerfilAcceso.AddItem "Personalizado"
    CmbPerfilAcceso.ItemData(0) = 0
    Set Datos = ListarPerfilesAcceso(True)
    Do While Not Datos.EOF
        If CBool(Datos!Activo) Or CLng(Datos!Item) = Actual Then
            CmbPerfilAcceso.AddItem CStr(Datos!Nombre) & IIf(CBool(Datos!Activo), "", " (eliminado)")
            CmbPerfilAcceso.ItemData(CmbPerfilAcceso.NewIndex) = CLng(Datos!Item)
            If CLng(Datos!Item) = Actual Then CmbPerfilAcceso.ListIndex = CmbPerfilAcceso.NewIndex
        End If
        Datos.MoveNext
    Loop
    Datos.Close
    If Actual = 0 Then CmbPerfilAcceso.ListIndex = 0
    Cargando = False
End Sub
Private Function PerfilElegido() As Long
    If CmbPerfilAcceso.ListIndex >= 0 Then PerfilElegido = CmbPerfilAcceso.ItemData(CmbPerfilAcceso.ListIndex)
End Function
Private Sub CmbPerfilAcceso_Click()
    If Cargando Or (Modo <> 1 And Modo <> 2) Then Exit Sub
    On Error GoTo ErrorPerfil
    If PerfilElegido > 0 Then CargarArbol 0, PerfilElegido Else VersionPerfilElegido = Null
    Exit Sub
ErrorPerfil:
    Cargando = False
    MsgBox "No se pudo cargar el perfil: " & Err.Description, vbExclamation, App.Title
End Sub


Public Sub MostrarUsuario(ByVal Codigo As Long)
    If Modo <> 0 Then Exit Sub
    CargarFicha Codigo
    If Seleccion > 0 Then MostrarDetalle
End Sub

Private Sub CargarFicha(ByVal Codigo As Long, Optional ByVal Guardado As Boolean = False)
    Dim Datos As ADODB.Recordset, Mensaje As String
    On Error GoTo ErrorCarga
    Set Datos = ListarUsuariosAdministracion(True, "", Codigo)
    If Datos.EOF Then Err.Raise vbObjectError + 2910, , "No se encontró la cuenta."
    LimpiarFicha
    Cargando = True
    Seleccion = CLng(Datos!Item)
    Version = Datos!VersionFila
    UsuarioActivo = CBool(Datos!Activo)
    TxtUsuario.Text = CStr(Datos!Usuari)
    TxtNombre.Text = CStr(Datos!NomUsu)
    TxtApellido.Text = CStr(Datos!ApeUsu)
    If Not IsNull(Datos!Email) Then TxtCorreo.Text = CStr(Datos!Email)
    If UCase$(Trim$(CStr(Datos!Access))) = "ADMINISTRADOR" Then CmbPerfil.ListIndex = 0 Else CmbPerfil.ListIndex = 1
    LblSeleccion.Caption = TxtUsuario.Text & " - " & TxtNombre.Text & " " & TxtApellido.Text
    LblEstadoUsuario.Caption = IIf(UsuarioActivo, "Activo", "Eliminado / inactivo")
    If IsNull(Datos!IdPerfilAcceso) Then CargarPerfilesUsuario Else CargarPerfilesUsuario CLng(Datos!IdPerfilAcceso)
    Datos.Close
    Cargando = False
    CargarArbol Seleccion
    ActualizarControles
    If Guardado Then Estado "Usuario guardado. Código: " & CStr(Codigo) Else Estado "Ficha de usuario."
    Exit Sub
ErrorCarga:
    Mensaje = Err.Description
    CerrarRecordset Datos
    Cargando = False
    LimpiarFicha
    ActualizarControles
    If Guardado Then
        MsgBox "Se guardó la cuenta con código " & CStr(Codigo) & ". Falló la recarga; seleccione la cuenta sin repetir el guardado.", vbExclamation, App.Title
    Else
        MsgBox "No se pudo cargar la cuenta: " & Mensaje, vbExclamation, App.Title
    End If
End Sub

Private Sub ActualizarControles()
    Dim Editando As Boolean, Claves As Boolean
    Editando = (Modo = 1 Or Modo = 2)
    Claves = (Modo = 1 Or Modo = 3)
    CmdDetalle.Enabled = (Modo = 0 And Seleccion > 0)
    CmdListado.Enabled = (Modo = 0 And Not MostrandoListado)
    GridUsuarios.Enabled = (Modo = 0)
    TxtBuscar.Enabled = (Modo = 0 And MostrandoListado)
    ChkEliminados.Enabled = TxtBuscar.Enabled
    CmdNuevo.Enabled = (Modo = 0)
    CmdModificar.Enabled = (Modo = 0 And Seleccion > 0 And UsuarioActivo)
    CmdClave.Enabled = CmdModificar.Enabled
    CmdEliminar.Enabled = (CmdModificar.Enabled And Seleccion <> IdUsuarioActual)
    CmdGuardar.Visible = (Modo <> 0)
    CmdCancelar.Visible = (Modo <> 0)
    CmdAccesos.Enabled = (Modo <> 3 And (Seleccion > 0 Or Modo = 1))
    TxtUsuario.Locked = (Modo <> 1)
    TxtNombre.Locked = Not Editando
    TxtApellido.Locked = Not Editando
    TxtCorreo.Locked = Not Editando
    CmbPerfil.Enabled = (Editando And Seleccion <> IdUsuarioActual)
    CmbPerfilAcceso.Enabled = (Editando And CmbPerfil.ListIndex = 1)
    CmbPerfilAcceso.Visible = (CmbPerfil.ListIndex = 1)
    LblPerfilAcceso.Visible = CmbPerfilAcceso.Visible
    TxtClave.Visible = Claves
    TxtConfirmacion.Visible = Claves
    LblClave.Visible = Claves
    LblConfirmacion.Visible = Claves
    ViewMenus.Enabled = True
    StatusBar1.Panels(1).Text = Choose(Modo + 1, "Consulta", "Crear", "Modificar", "Cambiar clave")
    If Seleccion = 0 Then StatusBar1.Panels(2).Text = "Sin selección" Else StatusBar1.Panels(2).Text = "Usuario: " & TxtUsuario.Text
End Sub

Private Sub Estado(ByVal Texto As String)
    StatusBar1.Panels(3).Text = Texto
End Sub

Private Sub CmdGeneral_Click()
    FraGeneral.Visible = True
    FraAccess.Visible = False
End Sub
Private Sub CmdAccesos_Click()
    If Modo = 3 Then Exit Sub
    FraGeneral.Visible = False
    FraAccess.Visible = True
End Sub
Private Sub CmdNuevo_Click()
    If Modo <> 0 Then Exit Sub
    On Error GoTo ErrorNuevo
    LimpiarFicha
    CargarPerfilesUsuario
    CargarArbol 0
    Modo = 1
    MostrarDetalle
    ActualizarControles
    Estado "Complete datos, clave y accesos. Guardar o Cancelar."
    TxtUsuario.SetFocus
    Exit Sub
ErrorNuevo:
    Cargando = False
    Modo = 0
    ActualizarControles
    MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdModificar_Click()
    If Modo <> 0 Or Seleccion = 0 Or Not UsuarioActivo Then Exit Sub
    Modo = 2
    MostrarDetalle
    ActualizarControles
    Estado "Edición de datos y permisos. Guardar o Cancelar."
    TxtNombre.SetFocus
End Sub
Private Sub CmdClave_Click()
    If Modo <> 0 Or Seleccion = 0 Or Not UsuarioActivo Then Exit Sub
    Modo = 3
    TxtClave.Text = ""
    TxtConfirmacion.Text = ""
    MostrarDetalle
    ActualizarControles
    Estado "Ingrese la clave nueva y su confirmación."
    TxtClave.SetFocus
End Sub

Private Function TextoPermisos() As String
    Dim I As Integer
    For I = 1 To 40
        If ModulosVigentes(I) Then
            If ViewMenus.Nodes("M" & CStr(I)).Checked Then TextoPermisos = TextoPermisos & "1" Else TextoPermisos = TextoPermisos & "0"
        Else
            TextoPermisos = TextoPermisos & "0"
        End If
    Next I
End Function

Private Sub CmdGuardar_Click()
    Dim Codigo As Long, Confirmado As Boolean, Mensaje As String
    On Error GoTo ErrorGuardado
    If Modo = 0 Then Exit Sub
    If Modo <> 3 Then
        If Len(Trim$(TxtNombre.Text)) = 0 Or Len(Trim$(TxtApellido.Text)) = 0 Or Len(Trim$(TxtUsuario.Text)) = 0 Then
            MsgBox "Complete usuario, nombres y apellidos.", vbExclamation, App.Title
            Exit Sub
        End If
    End If
    If Modo = 1 Or Modo = 3 Then
        If Len(TxtClave.Text) = 0 Or TxtClave.Text <> TxtConfirmacion.Text Then
            MsgBox "Ingrese la clave y su confirmación coincidente.", vbExclamation, App.Title
            Exit Sub
        End If
    End If
    If MsgBox("¿Guardar los cambios de la cuenta " & TxtUsuario.Text & "?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
    Codigo = GuardarDatosUsuario(Modo, Seleccion, Version, TxtUsuario.Text, TxtNombre.Text, TxtApellido.Text, TxtCorreo.Text, CmbPerfil.ListIndex, TxtClave.Text, TxtConfirmacion.Text, TextoPermisos(), PerfilElegido(), VersionPerfilElegido)
    Confirmado = True
    Modo = 0
    Cargando = True
    TxtBuscar.Text = ""
    Cargando = False
    CargarListado Codigo, True
    MostrarDetalle
    Exit Sub
ErrorGuardado:
    Mensaje = Err.Description
    If Confirmado Then
        Modo = 0
        LimpiarFicha
        ActualizarControles
        MsgBox "Se guardó el código " & CStr(Codigo) & ". Seleccione la cuenta sin repetir el guardado.", vbExclamation, App.Title
    Else
        MsgBox "No se pudo confirmar el guardado. Los datos se conservan; consulte antes de reintentar: " & Mensaje, vbExclamation, App.Title
    End If
End Sub

Private Sub CmdEliminar_Click()
    Dim Codigo As Long
    On Error GoTo ErrorBaja
    If Modo <> 0 Or Seleccion = 0 Or Not UsuarioActivo Then Exit Sub
    If MsgBox("¿Eliminar al usuario " & TxtUsuario.Text & " - " & TxtNombre.Text & " " & TxtApellido.Text & "?" & vbCrLf & "La baja será lógica; su historial y permisos se conservan. No podrá reactivarse.", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
    Codigo = GuardarDatosUsuario(4, Seleccion, Version, TxtUsuario.Text, "", "", "", CmbPerfil.ListIndex, "", "", "")
    CargarListado
    MostrarListado
    Estado "Usuario eliminado lógicamente. Historial conservado."
    Exit Sub
ErrorBaja:
    MsgBox "No se pudo confirmar la baja. Recargue antes de reintentar: " & Err.Description, vbExclamation, App.Title
End Sub
Private Function ConfirmarDescarte() As Boolean
    ConfirmarDescarte = True
    If Modo <> 0 Then ConfirmarDescarte = (MsgBox("¿Descartar los cambios de la cuenta?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) = vbYes)
End Function
Private Sub CmdCancelar_Click()
    If Not ConfirmarDescarte Then Exit Sub
    Modo = 0
    SeleccionarFila
    MostrarListado
    Estado "Edición cancelada."
End Sub
Private Sub CmbPerfil_Click()
    Dim I As Integer
    If Cargando Or (Modo <> 1 And Modo <> 2) Then Exit Sub
    Cargando = True
    If CmbPerfil.ListIndex = 0 Then
        CmbPerfilAcceso.ListIndex = 0
        VersionPerfilElegido = Null
        For I = 1 To 40
            If ModulosVigentes(I) Then ViewMenus.Nodes("M" & CStr(I)).Checked = True
        Next I
    Else
        ViewMenus.Nodes("M27").Checked = False
        ViewMenus.Nodes("M33").Checked = False
        ViewMenus.Nodes("M34").Checked = False
    End If
    Cargando = False
    ActualizarControles
End Sub
Private Sub ViewMenus_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim Codigo As Integer, Padre As Integer, I As Integer
    If Cargando Then Exit Sub
    Codigo = CInt(Mid$(Node.Key, 2))
    If (Modo <> 1 And Modo <> 2) Or CmbPerfil.ListIndex = 0 Then
        Cargando = True
        If CmbPerfil.ListIndex = 0 Then Node.Checked = True Else Node.Checked = PermisosConsulta(Codigo)
        Cargando = False
        Exit Sub
    End If
    If Codigo = 1 Or Codigo = 2 Or Codigo = 27 Or Codigo = 33 Or Codigo = 34 Then
        Cargando = True
        Node.Checked = (Codigo = 1 Or Codigo = 2)
        Cargando = False
        Exit Sub
    End If
    If PerfilElegido > 0 Then
        If MsgBox("¿Personalizar los accesos de esta cuenta?" & vbCrLf & "Conservará los permisos actuales como punto de partida. El perfil compartido no se modificará.", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then
            Cargando = True
            Node.Checked = PermisosConsulta(Codigo)
            Cargando = False
            Exit Sub
        End If
        Cargando = True
        CmbPerfilAcceso.ListIndex = 0
        VersionPerfilElegido = Null
        Cargando = False
    End If
    Cargando = True
    If Node.Checked Then
        Padre = Padres(Codigo)
        Do While Padre <> 0
            ViewMenus.Nodes("M" & CStr(Padre)).Checked = True
            Padre = Padres(Padre)
        Loop
    Else
        For I = 1 To 40
            If ModulosVigentes(I) Then
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

Private Sub ConfigurarGrilla()
    GridUsuarios.SelectionMode = flexSelectionByRow
    GridUsuarios.HighLight = flexHighlightAlways
    GridUsuarios.FocusRect = flexFocusNone
    GridUsuarios.AllowUserResizing = flexResizeNone
    GridUsuarios.Cols = 6
    GridUsuarios.FormatString = "Código|Usuario|Nombres y apellidos|Tipo de usuario|Perfil de acceso|Estado"
    GridUsuarios.ColWidth(0) = 720
    GridUsuarios.ColWidth(1) = 1200
    GridUsuarios.ColWidth(2) = 3000
    GridUsuarios.ColWidth(3) = 1500
    GridUsuarios.ColWidth(4) = 1800
    GridUsuarios.ColWidth(5) = 1000
End Sub

Private Sub CargarListado(Optional ByVal Codigo As Long = 0, Optional ByVal Guardado As Boolean = False)
    Dim Datos As ADODB.Recordset, Fila As Long, Elegida As Long, Mensaje As String
    On Error GoTo ErrorListado
    Cargando = True
    LimpiarFicha
    Cargando = True
    GridUsuarios.Rows = 2
    For Fila = 0 To 5
        GridUsuarios.TextMatrix(1, Fila) = ""
    Next Fila
    ActualizarControles
    Set Datos = ListarUsuariosAdministracion(ChkEliminados.Value = vbChecked, TxtBuscar.Text)
    CerrarRecordset ListaUsuarios
    Set ListaUsuarios = Datos
    If Datos.RecordCount > 0 Then GridUsuarios.Rows = Datos.RecordCount + 1
    Fila = 1
    Do While Not Datos.EOF
        GridUsuarios.TextMatrix(Fila, 0) = CStr(Datos!Item)
        GridUsuarios.TextMatrix(Fila, 1) = CStr(Datos!Usuari)
        GridUsuarios.TextMatrix(Fila, 2) = CStr(Datos!NomUsu) & " " & CStr(Datos!ApeUsu)
        GridUsuarios.TextMatrix(Fila, 3) = CStr(Datos!Access)
        GridUsuarios.TextMatrix(Fila, 4) = CStr(Datos!PerfilAcceso)
        GridUsuarios.TextMatrix(Fila, 5) = IIf(CBool(Datos!Activo), "Activo", "Eliminado")
        If CLng(Datos!Item) = Codigo Then Elegida = Fila
        Fila = Fila + 1
        Datos.MoveNext
    Loop
    If Elegida = 0 Then Elegida = 1
    GridUsuarios.Row = Elegida
    Cargando = False
    SeleccionarFila
    Estado CStr(Datos.RecordCount) & " usuarios. Enter: detalle. F5: actualizar."
    If Guardado Then Estado "Cuenta guardada. Código: " & CStr(Codigo)
    Exit Sub
ErrorListado:
    Mensaje = Err.Description
    CerrarRecordset Datos
    CerrarRecordset ListaUsuarios
    Cargando = False
    LimpiarFicha
    ActualizarControles
    If Guardado Then
        MsgBox "Se guardó la cuenta con código " & CStr(Codigo) & ". Falló la recarga; pulse F5 sin repetir el guardado.", vbExclamation, App.Title
    Else
        MsgBox "No se pudo consultar usuarios: " & Mensaje, vbExclamation, App.Title
    End If
End Sub

Private Sub SeleccionarFila()
    If Cargando Or Modo <> 0 Then Exit Sub
    If GridUsuarios.Row < 1 Then Exit Sub
    If Len(GridUsuarios.TextMatrix(GridUsuarios.Row, 0)) = 0 Then
        LimpiarFicha
        Cargando = True
        ViewMenus.Nodes.Clear
        Cargando = False
        ActualizarControles
        Exit Sub
    End If
    CargarFicha CLng(GridUsuarios.TextMatrix(GridUsuarios.Row, 0))
    Cargando = True
    GridUsuarios.Col = 0
    GridUsuarios.ColSel = GridUsuarios.Cols - 1
    GridUsuarios.RowSel = GridUsuarios.Row
    Cargando = False
End Sub

Private Sub MostrarListado()
    MostrandoListado = True
    FraFicha.Visible = False
    FraListado.Visible = True
    ActualizarControles
End Sub
Private Sub MostrarDetalle()
    MostrandoListado = False
    FraListado.Visible = False
    FraFicha.Visible = True
    CmdGeneral_Click
    ActualizarControles
End Sub
Private Sub CmdListado_Click()
    If Modo <> 0 Then Exit Sub
    MostrarListado
End Sub
Private Sub CmdDetalle_Click()
    If Modo = 0 And Seleccion > 0 Then MostrarDetalle
End Sub
Private Sub TxtBuscar_Change()
    If Not Cargando And Modo = 0 Then CargarListado
End Sub
Private Sub ChkEliminados_Click()
    If Not Cargando And Modo = 0 Then CargarListado
End Sub
Private Sub GridUsuarios_RowColChange()
    SeleccionarFila
End Sub
Private Sub GridUsuarios_Click()
    SeleccionarFila
End Sub
Private Sub GridUsuarios_DblClick()
    CmdDetalle_Click
End Sub
Private Sub GridUsuarios_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 0 Then
        KeyCode = 0
        SeleccionarFila
        CmdDetalle_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CerrarRecordset ListaUsuarios
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

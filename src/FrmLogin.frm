VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clever SoftPro 6.0 - Acceso"
   ClientHeight = 4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "FrmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight = 4185
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin Clever.Boton Cmdhelp 
      Height          =   375
      Left            =   4545
      TabIndex        =   5
      Top = 3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Ayuda"
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
      MICON           =   "FrmLogin.frx":08CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtPasword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3105
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top = 2790
      Width           =   2895
   End
   Begin VB.ComboBox CmbUsuario 
      Height          =   315
      Left            =   3090
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top = 2430
      Width           =   2925
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height = 2115
      Left            =   0
      Picture         =   "FrmLogin.frx":08E6
      ScaleHeight = 2085
      ScaleWidth      =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   6270
   End
   Begin Clever.Boton CmdAceptar 
      Height          =   375
      Left            =   1575
      TabIndex        =   3
      Top = 3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Aceptar"
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
      MICON           =   "FrmLogin.frx":6C6C
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
      Height          =   375
      Left            =   3060
      TabIndex        =   4
      Top = 3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Cancelar"
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
      MICON           =   "FrmLogin.frx":6C88
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   1935
      TabIndex        =   8
      Top = 2835
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   1980
      TabIndex        =   7
      Top = 2475
      Width           =   585
   End
   Begin VB.Label LblEstado
      Caption         =   "Conectando..."
      Height          =   345
      Left            =   405
      Top = 3150
      Width           =   5610
      TabIndex        =   6
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private md5Clave As MD5
Private Validando As Boolean
Private UsuariosCargados As Boolean

Private Sub Form_Load()
    Set md5Clave = New MD5
    CargarUsuarios
End Sub

Private Sub CargarUsuarios()
    Dim Usuarios As ADODB.Recordset
    Dim Mensaje As String
    Dim NombreWindows As String
    On Error GoTo ErrorCarga
    UsuariosCargados = False
    CmbUsuario.Clear
    TxtPasword.Text = ""
    NombreWindows = UsuarioWindows()
    If Not ConsultarUsuarios(Usuarios, Mensaje) Then GoTo Terminar
    Do While Not Usuarios.EOF
        If Not IsNull(Usuarios.Fields(0).Value) Then
            CmbUsuario.AddItem CStr(Usuarios.Fields(0).Value)
            If StrComp(CmbUsuario.List(CmbUsuario.NewIndex), NombreWindows, vbTextCompare) = 0 Then CmbUsuario.ListIndex = CmbUsuario.NewIndex
        End If
        Usuarios.MoveNext
    Loop
    UsuariosCargados = (CmbUsuario.ListCount > 0)
    If UsuariosCargados Then
        Mensaje = MENSAJE_ACCESO
    Else
        Mensaje = "No hay usuarios activos con clave MD5. F5 actualiza la lista."
    End If
Terminar:
    CerrarRecordset Usuarios
    LblEstado.Caption = Mensaje
    MdiPri.MostrarEstado Mensaje
    Exit Sub
ErrorCarga:
    Mensaje = "No se pudo cargar la lista. F5 o Aceptar permite reintentar."
    Resume Terminar
End Sub

Private Sub CmdAceptar_Click()
    AceptarAcceso
End Sub

Private Sub AceptarAcceso()
    Dim Accesos As ADODB.Recordset
    Dim Mensaje As String
    Dim Clave As String
    Dim Administrador As Boolean
    On Error GoTo ErrorAcceso
    If Validando Then Exit Sub
    If Not UsuariosCargados Then
        CargarUsuarios
        Exit Sub
    End If
    If CmbUsuario.ListIndex < 0 Or Len(TxtPasword.Text) = 0 Then
        LblEstado.Caption = "Seleccione su usuario e indique la contraseña."
        If CmbUsuario.ListIndex < 0 Then CmbUsuario.SetFocus Else TxtPasword.SetFocus
        Exit Sub
    End If
    Validando = True
    CmdAceptar.Enabled = False
    Clave = LCase$(md5Clave.DigestStrToHexStr(TxtPasword.Text))
    If ValidarUsuario(CmbUsuario.Text, Clave, Accesos, Administrador, Mensaje) Then
        If IniciarSesion(CmbUsuario.Text, Administrador, Accesos, Mensaje) Then
            CerrarRecordset Accesos
            TxtPasword.Text = ""
            Validando = False
            Unload Me
            Exit Sub
        End If
    End If
    CerrarRecordset Accesos
    LblEstado.Caption = Mensaje
    MdiPri.MostrarEstado Mensaje
    Validando = False
    CmdAceptar.Enabled = True
    TxtPasword.SetFocus
    TxtPasword.SelStart = 0
    TxtPasword.SelLength = Len(TxtPasword.Text)
    Exit Sub
ErrorAcceso:
    CerrarRecordset Accesos
    LimpiarSesion
    MdiPri.ActualizarAccesos
    Validando = False
    CmdAceptar.Enabled = True
    LblEstado.Caption = "No se pudo completar el acceso. Puede reintentar o cancelar."
End Sub

Private Sub CmdCancelar_Click()
    If Not Validando Then Unload Me
End Sub

Private Sub Cmdhelp_Click()
    MsgBox MENSAJE_ACCESO & vbCrLf & _
           "Enter: aceptar. Escape o Cancelar: salir del acceso." & vbCrLf & _
           "La cuenta del sistema es independiente del usuario de Windows utilizado para SQL Server.", vbInformation, "Ayuda de acceso"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Or Validando Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            KeyCode = 0
            AceptarAcceso
        Case vbKeyEscape
            KeyCode = 0
            CmdCancelar_Click
        Case vbKeyF5
            KeyCode = 0
            CargarUsuarios
    End Select
End Sub

Private Sub CmbUsuario_Click()
    TxtPasword.Text = ""
End Sub

Private Sub TxtPasword_GotFocus()
    TxtPasword.SelStart = 0
    TxtPasword.SelLength = Len(TxtPasword.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Validando Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TxtPasword.Text = ""
    Set md5Clave = Nothing
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

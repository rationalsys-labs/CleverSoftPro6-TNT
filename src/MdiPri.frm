VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.MDIForm MdiPri 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Clever SoftPro 6.0"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11070
   Icon            =   "MdiPri.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer TmrFondo
      Enabled = 0
      Interval = 40
      Left = 0
      Top = 0
   End
   Begin VB.Timer TmrCentrado
      Enabled = 0
      Interval = 30
      Left = 0
      Top = 0
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   1890
      Top             =   3510
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount = 24
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   16646297
      SelCheckBackColor=   13740436
      MenuBorderColor =   6956042
      SeparatorColor  =   -2147483632
      MenuBackColor   =   14609903
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   10027263
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   14215660
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   16646297
      ArrowNormalColor=   10027263
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Bmp:1 = "MdiPri.frx":402A
      Key:1 = "#A000005"
      Bmp:2 = "MdiPri.frx":4452
      Key:2 = "#A000006"
      Bmp:3 = "MdiPri.frx":487A
      Key:3 = "#A000007"
      Bmp:4 = "MdiPri.frx":4CA2
      Key:4 = "#A000009"
      Bmp:5 = "MdiPri.frx":50CA
      Key:5 = "#A000010"
      Bmp:6 = "MdiPri.frx":54F2
      Key:6 = "#A000011"
      Bmp:7 = "MdiPri.frx":591A
      Key:7 = "#A000012"
      Bmp:8 = "MdiPri.frx":5D42
      Key:8 = "#A000013"
      Bmp:9 = "MdiPri.frx":616A
      Key:9 = "#A000014"
      Bmp:10 = "MdiPri.frx":6592
      Key:10 = "#A000015"
      Bmp:11 = "MdiPri.frx":69BA
      Key:11 = "#A000016"
      Bmp:12 = "MdiPri.frx":720A
      Key:12 = "#A000019"
      Bmp:13 = "MdiPri.frx":7632
      Key:13 = "#A000020"
      Bmp:14 = "MdiPri.frx":7A5A
      Key:14 = "#A000021"
      Bmp:15 = "MdiPri.frx":8AFA
      Key:15 = "#MenuInfRP_Ver"
      Bmp:16 = "MdiPri.frx":8F22
      Key:16 = "#MenuInfRP_Act"
      Bmp:17 = "MdiPri.frx":934A
      Key:17 = "#MenuInfRP_desac"
      Bmp:18 = "MdiPri.frx":9772
      Key:18 = "#MenuInfRP_OnOff"
      Bmp:19 = "MdiPri.frx":9B9A
      Key:19 = "#MenuInfRP_Adicional"
      Bmp:20 = "MdiPri.frx":9FC2
      Key:20 = "#MenuInfRP_Ins"
      Bmp:21 = "MdiPri.frx":A3EA
      Key:21 = "#MenuInfRP_Event"
      Bmp:22 = "MdiPri.frx":A812
      Key:22 = "#MenuInfRP_Ret"
      Bmp:23 = "MdiPri.frx":AC3A
      Key:23 = "#MenuInfRP_con"
      Bmp:24 = "MdiPri.frx":B062
      Key:24 = "#MenuInfRP_des"
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
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Height          =   7380
      Left            =   9255
      ScaleHeight     =   7320
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      Begin VB.ListBox LstConf 
         Height          =   2400
         Left            =   135
         TabIndex        =   4
         Top             =   225
         Width           =   1500
      End
      Begin VB.PictureBox PicDibujo
         AutoRedraw = -1
         BorderStyle = 0
         ScaleMode = 3
         Visible = 0
         Left = 0
         Top = 0
         Width = 1500
         Height = 1500
      End
      Begin VB.PictureBox PicFondo 
         Height          =   1770
         Left            =   90
         Picture         =   "MdiPri.frx":B48A
         ScaleHeight     =   1710
         ScaleWidth      =   1470
         TabIndex        =   3
         Top             =   5490
         Width           =   1530
      End
   End
   Begin MSComctlLib.ImageList ICO 
      Left            =   1125
      Top             =   5580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":41DD0
            Key             =   "Cli"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":4236A
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":42D7C
            Key             =   "a3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":43656
            Key             =   "a4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":438EB
            Key             =   "a5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":43E85
            Key             =   "a6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":4441F
            Key             =   "a7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":449B9
            Key             =   "a8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":44F53
            Key             =   "a9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":454ED
            Key             =   "a10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":45A87
            Key             =   "a11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":46021
            Key             =   "a12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":465BB
            Key             =   "a13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":46B55
            Key             =   "a14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":470EF
            Key             =   "a15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":47689
            Key             =   "a16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":47C23
            Key             =   "a17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":481BD
            Key             =   "a18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":48757
            Key             =   "a19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":48CF1
            Key             =   "a20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":4928B
            Key             =   "a21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":49825
            Key             =   "a22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":49DBF
            Key             =   "a23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiPri.frx":4A359
            Key             =   "a24"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ICO"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Creacion de Queries Emitidas"
            ImageKey        =   "a5"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Control y Administracion de Queries Emitidas"
            ImageKey        =   "a6"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Finalizar Queries Emitidas"
            ImageKey        =   "a7"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Control del Diario RP"
            ImageKey        =   "a8"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Compilatorio RP"
            ImageKey        =   "a9"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Control de Facturas"
            ImageKey        =   "a10"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quejas y Reclamos"
            ImageKey        =   "a11"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cheques rechazados"
            ImageKey        =   "a12"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recojo de Cheques y Facturas"
            ImageKey        =   "a13"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Creacion de Cartas de Cobranza"
            ImageKey        =   "a14"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Gestion de Devoluciones"
            ImageKey        =   "a15"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Control y Manejo de Cargos"
            ImageKey        =   "a17"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantenimiento de Clientes"
            ImageKey        =   "Cli"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Importar Datos e Informacion al Sistema"
            ImageKey        =   "a18"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Informes personalizados"
            ImageKey        =   "a22"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
         EndProperty
      EndProperty
   End
   Begin VB.Menu A000001 
      Caption         =   "&Archivo"
      Begin VB.Menu A000002 
         Caption         =   "Salir del Sistema"
      End
      Begin VB.Menu CerSes 
         Caption         =   "Cerrar sesión"
      End
   End
   Begin VB.Menu A000003 
      Caption         =   "&Cobranzas"
      Begin VB.Menu A000004 
         Caption         =   "Solicitud de Queries"
         Begin VB.Menu A000005 
            Caption         =   "Correccion de Documentos - Anulaciones"
         End
         Begin VB.Menu A000006 
            Caption         =   "Regeneracion de Querie Emitidas"
         End
         Begin VB.Menu A000007 
            Caption         =   "Finalizar Queries Emitidas"
         End
      End
      Begin VB.Menu A000008 
         Caption         =   "Control de Rp's"
         Begin VB.Menu A000009 
            Caption         =   "Control Diario RP"
         End
         Begin VB.Menu A000010 
            Caption         =   "Compilatorio de Guias RP"
         End
      End
      Begin VB.Menu A000011 
         Caption         =   "Listado de Facturas Emitidas"
      End
      Begin VB.Menu A000012 
         Caption         =   "Quejas y Reclamos"
      End
      Begin VB.Menu A000013 
         Caption         =   "Control de Cheques rechazados"
      End
      Begin VB.Menu A000014 
         Caption         =   "Gestionar Recojo de Cheq o Fac"
      End
      Begin VB.Menu A000015 
         Caption         =   "Creacion de  Cartas de Cobranza"
      End
      Begin VB.Menu A000016 
         Caption         =   "Gestion de Devoluciones"
      End
   End
   Begin VB.Menu A000018 
      Caption         =   "&Operaciones"
      Begin VB.Menu A000019 
         Caption         =   "Control de Cargos"
      End
      Begin VB.Menu A000020 
         Caption         =   "Mantenimiento de Clientes"
      End
      Begin VB.Menu A000021 
         Caption         =   "Importación de información"
      End
   End
   Begin VB.Menu Informes 
      Caption         =   "&Informes"
      Begin VB.Menu A000040 
         Caption         =   "Informes personalizados"
      End
   End
   Begin VB.Menu A000026 
      Caption         =   "Configuracion"
      Begin VB.Menu A000039
         Caption = "Mantenimiento de Catálogos"
      End
      Begin VB.Menu A000027 
         Caption = "Perfiles de acceso"
         End
      Begin VB.Menu A000033 
         Caption         =   "Control de Usuarios"
         Begin VB.Menu A000034 
            Caption         =   "Administrar Usuarios"
         End
      End
   End
   Begin VB.Menu Ayuda 
      Caption         =   "A&yuda"
      Begin VB.Menu contenido 
         Caption         =   "Contenido"
      End
      Begin VB.Menu AcercClev 
         Caption         =   "Acerca de Clever SoftPro"
      End
   End
   Begin VB.Menu MenuInfRP 
      Caption         =   "MenuInfRP"
      Begin VB.Menu MenuInfRP_Ver 
         Caption         =   "Ver Información Completa"
      End
      Begin VB.Menu MenuInfRP_Act 
         Caption         =   "Activar Salida RP"
      End
      Begin VB.Menu MenuInfRP_desac 
         Caption         =   "Desactivar Salida RP"
      End
      Begin VB.Menu MenuInfRP_OnOff 
         Caption         =   "On / Off Guia RP"
      End
      Begin VB.Menu MenuInfRP_Adicional 
         Caption         =   "Asignar Servicio Adicional"
      End
      Begin VB.Menu MenuInfRP_Ins 
         Caption         =   "Insertar Guia Manual"
      End
      Begin VB.Menu MenuInfRP_Event 
         Caption         =   "Asignar Evento"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu MenuInfRP_Ret 
         Caption         =   "Activar Modalidad Retencion"
      End
      Begin VB.Menu MenuInfRP_con 
         Caption         =   "Activar Modalidad Contado"
      End
      Begin VB.Menu MenuInfRP_des 
         Caption         =   "Desactivar Gestiones"
      End
   End
End
Attribute VB_Name = "MdiPri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Iniciando As Boolean
Private Cerrando As Boolean
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal Objeto As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal Destino As Long, ByVal X As Long, ByVal Y As Long, ByVal Ancho As Long, ByVal Alto As Long, ByVal Origen As Long, ByVal OrigenX As Long, ByVal OrigenY As Long, ByVal OrigenAncho As Long, ByVal OrigenAlto As Long, ByVal Operacion As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, Rectangulo As Any, ByVal Region As Long, ByVal Opciones As Long) As Long
Private PintandoFondo As Boolean
Private AnchoFondo As Single, AltoFondo As Single
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal Modo As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, PuntoAnterior As Any) As Long

Private Sub MDIForm_Load()
    On Error GoTo ErrorCargaMdi
    Iniciando = True
    ActualizarFondo
    MenuInfRP.Visible = False
    ConfigurarBarra
    LimpiarSesion
    ActualizarAccesos
    Me.Show
    FrmLogin.Show vbModal, Me
    Iniciando = False
    If SesionActiva Or VistaPrevia Then
        Me.WindowState = vbMaximized
        ActualizarAccesos
    Else
        Unload Me
    End If
    Exit Sub
ErrorCargaMdi:
    MsgBox "No se pudo preparar el MDI: " & Err.Description, vbExclamation, App.Title
    LimpiarSesion
    Unload Me
End Sub

Public Sub ConfigurarBarra()
    Dim Codigos As Variant
    Dim I As Integer
    Codigos = Array(5, 6, 7, 0, 9, 10, 0, 11, 12, 13, 14, 15, 16, 0, 19, 20, 21, 0, 40)
    For I = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(I).Tag = CStr(Codigos(I - 1))
    Next I
    StatusBar1.Panels(1).AutoSize = sbrContents
    StatusBar1.Panels(2).AutoSize = sbrContents
    StatusBar1.Panels(3).AutoSize = sbrSpring
End Sub

Public Sub ActualizarAccesos()
    Dim I As Integer
    Dim Codigo As Integer
    Dim Menu As Object
    For I = 1 To 40
        If (I < 28 Or I > 32) And I <> 35 And (I < 36 Or I > 38) And (I < 22 Or I > 25) And I <> 17 Then
            Set Menu = Me.Controls("A" & Format$(I, "000000"))
            Menu.Enabled = AccesoPermitido(I) And FuncionDisponible(I)
        End If
    Next I
    ActualizarVisibilidad
    ' Archivo y Ayuda siempre permiten abandonar o consultar el sistema.
    A000001.Visible = True
    A000001.Enabled = True
    A000002.Visible = True
    A000002.Enabled = True
    CerSes.Enabled = (SesionActiva Or VistaPrevia) And Not Iniciando
    If VistaPrevia Then
        CerSes.Caption = "Volver al acceso"
    Else
        CerSes.Caption = "Cerrar sesión"
    End If
    Ayuda.Visible = True
    contenido.Enabled = True
    AcercClev.Enabled = True
    Informes.Enabled = AccesoPermitido(40) Or VistaPrevia
    A000040.Enabled = AccesoPermitido(40) And FuncionDisponible(40)
    Informes.Visible = AccesoPermitido(40) Or VistaPrevia
    MenuInfRP.Visible = False
    For I = 1 To Toolbar1.Buttons.Count
        Codigo = CInt(Toolbar1.Buttons(I).Tag)
        If Codigo > 0 Then
            Toolbar1.Buttons(I).Enabled = AccesoPermitido(Codigo) And FuncionDisponible(Codigo)
        Else
            Toolbar1.Buttons(I).Enabled = False
        End If
    Next I
    If SesionActiva Then
        Me.Caption = "Clever SoftPro 6.0 - " & VarUsuario
        StatusBar1.Panels(1).Text = "Usuario: " & VarUsuario
        If ModUsuario Then
            StatusBar1.Panels(2).Text = "Administrador"
        Else
            StatusBar1.Panels(2).Text = "Usuario del sistema"
        End If
        MostrarEstado "Sesión iniciada."
    Else
        Me.Caption = "Clever SoftPro 6.0"
        If VistaPrevia Then
            StatusBar1.Panels(1).Text = "Vista previa"
        Else
            StatusBar1.Panels(1).Text = "Sin sesión"
        End If
        If VistaPrevia Then
            StatusBar1.Panels(2).Text = "Sin autenticación"
        Else
            StatusBar1.Panels(2).Text = "Acceso pendiente"
        End If
        MostrarEstado MENSAJE_ACCESO
    End If
End Sub

Public Sub MostrarEstado(ByVal Mensaje As String)
    StatusBar1.Panels(3).Text = Mensaje
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Len(Button.Tag) > 0 Then AbrirModulo CInt(Button.Tag)
End Sub

Private Sub A000002_Click()
    Unload Me
End Sub

Private Sub CerSes_Click()
    If Iniciando Then Exit Sub
    If Not SesionActiva And Not VistaPrevia Then Exit Sub
    If SesionActiva Then
        If MsgBox("¿Desea cerrar la sesión de " & VarUsuario & "?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
    End If
    If Not CerrarFormulariosHijos() Then Exit Sub
    LimpiarSesion
    Iniciando = True
    ActualizarAccesos
    FrmLogin.Show vbModal, Me
    Iniciando = False
    If SesionActiva Or VistaPrevia Then
        Me.WindowState = vbMaximized
        ActualizarAccesos
    Else
        Unload Me
    End If
End Sub

Private Function CerrarFormulariosHijos() As Boolean
    Dim I As Integer
    Dim Cantidad As Integer
    For I = Forms.Count - 1 To 0 Step -1
        If Not Forms(I) Is Me Then
            Cantidad = Forms.Count
            Unload Forms(I)
            If Forms.Count = Cantidad Then Exit Function
        End If
    Next I
    CerrarFormulariosHijos = True
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Cerrando Then Exit Sub
    If SesionActiva Then
        If MsgBox("¿Desea salir de Clever SoftPro 6.0?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then
            Cancel = 1
            Exit Sub
        End If
    End If
    If Not CerrarFormulariosHijos() Then
        Cancel = 1
        Exit Sub
    End If
    Cerrando = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    CancelarCentradosPendientes
    LimpiarSesion
End Sub

Private Sub contenido_Click()
    Dim Informacion As String
    Informacion = "GUÍA RÁPIDA DE CLEVER SOFTPRO 6.0" & vbCrLf & vbCrLf
    Informacion = Informacion & "Importación: Operaciones > Importación de información. Elija el Excel, valide, revise incidencias y confirme la carga." & vbCrLf & vbCrLf
    Informacion = Informacion & "Clientes: use F9 donde se indique y confirme una fila. Solo los locales admiten edición y baja; los importados se conservan." & vbCrLf & vbCrLf
    Informacion = Informacion & "RP: registre eventos en Diario RP. Confirme la cartilla para pasar al Compilatorio y reimprimir." & vbCrLf & vbCrLf
    Informacion = Informacion & "Documentos y Queries: consulte comprobantes y guías, controle cargos y gestione solicitudes de corrección y su resolución." & vbCrLf & vbCrLf
    Informacion = Informacion & "Cobranzas: casos, cheques, recojos, cartas y devoluciones. Finalizar o anular conserva el historial y exige un motivo." & vbCrLf & vbCrLf
    Informacion = Informacion & "Correo: envío SMTP deshabilitado por compatibilidad." & vbCrLf & vbCrLf
    Informacion = Informacion & "Los módulos dependen de sus permisos. Archivo permite cerrar sesión o salir. Los importes y saldos vienen del sistema emisor."
    MsgBox Informacion, vbInformation, "Contenido - Guía rápida"
End Sub

Private Sub AcercClev_Click()
    Dim Informacion As String
    Informacion = "Clever SoftPro 6.0" & vbCrLf & "Versión interna " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision) & vbCrLf & vbCrLf
    Informacion = Informacion & "Sistema de gestión de cobranzas y seguimiento de servicios courier de importación. Trabaja con clientes, guías y comprobantes recibidos de un sistema externo." & vbCrLf & vbCrLf
    Informacion = Informacion & "Organiza la importación desde Excel, el Diario y Compilatorio RP, documentos emitidos, cargos e informes personalizados." & vbCrLf & vbCrLf
    Informacion = Informacion & "Incluye solicitudes y resolución de Queries, quejas y reclamos, cheques rechazados, recojos, cartas de cobranza y devoluciones, con historiales de gestión. Administra clientes locales, catálogos, usuarios, perfiles y permisos." & vbCrLf & vbCrLf
    Informacion = Informacion & "Tecnología: Visual Basic 6.0, SQL Server, ADO, controles OCX, Data Report y automatización de Microsoft Excel." & vbCrLf & vbCrLf
    Informacion = Informacion & "Diseñado y desarrollado por Juan Pablo Aguilar." & vbCrLf & "jaguilar.rationalsys@outlook.com"
    MsgBox Informacion, vbInformation, "Acerca de Clever SoftPro 6.0"
End Sub
Private Sub A000005_Click()
    AbrirModulo 5
End Sub

Private Sub A000006_Click()
    AbrirModulo 6
End Sub

Private Sub A000007_Click()
    AbrirModulo 7
End Sub

Private Sub A000009_Click()
    AbrirModulo 9
End Sub

Private Sub A000010_Click()
    AbrirModulo 10
End Sub

Private Sub A000011_Click()
    AbrirModulo 11
End Sub

Private Sub A000012_Click()
    AbrirModulo 12
End Sub

Private Sub A000013_Click()
    AbrirModulo 13
End Sub

Private Sub A000014_Click()
    AbrirModulo 14
End Sub

Private Sub A000015_Click()
    AbrirModulo 15
End Sub

Private Sub A000016_Click()
    AbrirModulo 16
End Sub


Private Sub A000019_Click()
    AbrirModulo 19
End Sub

Private Sub A000020_Click()
    AbrirModulo 20
End Sub

Private Sub A000021_Click()
    AbrirModulo 21
End Sub








Private Sub A000034_Click()
    AbrirModulo 34
End Sub




Private Sub ActualizarVisibilidad()
    A000003.Visible = AccesoPermitido(3) Or VistaPrevia
    A000018.Visible = AccesoPermitido(18) Or VistaPrevia
    A000026.Visible = AccesoPermitido(26) Or VistaPrevia
    MostrarGrupo Array(4, 8, 11, 12, 13, 14, 15, 16)
    MostrarGrupo Array(5, 6, 7)
    MostrarGrupo Array(9, 10)
    MostrarGrupo Array(19, 20, 21)
    MostrarGrupo Array(27, 33, 39)
    MostrarGrupo Array(34)
End Sub

Private Sub MostrarGrupo(ByVal Codigos As Variant)
    Dim I As Integer
    Dim AlgunoVisible As Boolean
    Dim Menu As Object
    ' Primero se muestran los autorizados: VB6 exige un hijo visible por submenú.
    For I = LBound(Codigos) To UBound(Codigos)
        If AccesoPermitido(CInt(Codigos(I))) Or VistaPrevia Then
            Set Menu = Me.Controls("A" & Format$(Codigos(I), "000000"))
            Menu.Visible = True
            AlgunoVisible = True
        End If
    Next I
    If Not AlgunoVisible Then Exit Sub
    For I = LBound(Codigos) To UBound(Codigos)
        Set Menu = Me.Controls("A" & Format$(Codigos(I), "000000"))
        Menu.Visible = AccesoPermitido(CInt(Codigos(I))) Or VistaPrevia
    Next I
    ' Sin permisos no se oculta el último hijo; permanece inhabilitado.
    ' El padre se oculta por su permiso y todos los eventos verifican el acceso.
End Sub

Private Sub A000039_Click()
    AbrirModulo 39
End Sub

Private Sub A000027_Click()
    AbrirModulo 27
End Sub

Private Sub TmrCentrado_Timer()
    AplicarCentradosPendientes
End Sub

Public Sub ActualizarFondo()
    Dim Ancho As Single, Alto As Single, ImagenAncho As Single, ImagenAlto As Single
    Dim Factor As Double, DibujoAncho As Single, DibujoAlto As Single
    Dim DCOrigen As Long, Anterior As Long, Resultado As Long
    If PintandoFondo Or Me.WindowState = vbMinimized Then Exit Sub
    Ancho = Me.ScaleWidth: Alto = Me.ScaleHeight
    If Ancho <= 0 Or Alto <= 0 Then Exit Sub
    If Ancho = AnchoFondo And Alto = AltoFondo Then Exit Sub
    On Error GoTo FondoOriginal
    PintandoFondo = True
    With PicDibujo
        .Width = Ancho: .Height = Alto
        .Cls
        ImagenAncho = .ScaleX(PicFondo.Picture.Width, vbHimetric, vbPixels)
        ImagenAlto = .ScaleY(PicFondo.Picture.Height, vbHimetric, vbPixels)
        If ImagenAncho <= 0 Or ImagenAlto <= 0 Then GoTo FondoOriginal
        Factor = .ScaleWidth / ImagenAncho
        If .ScaleHeight / ImagenAlto > Factor Then Factor = .ScaleHeight / ImagenAlto
        DibujoAncho = ImagenAncho * Factor: DibujoAlto = ImagenAlto * Factor
        SetStretchBltMode .hdc, 4
        SetBrushOrgEx .hdc, 0, 0, ByVal 0&
        DCOrigen = CreateCompatibleDC(.hdc)
        If DCOrigen = 0 Then GoTo FondoOriginal
        Anterior = SelectObject(DCOrigen, PicFondo.Picture.Handle)
        If Anterior = 0 Or Anterior = -1 Then GoTo FondoOriginal
        Resultado = StretchBlt(.hdc, (.ScaleWidth - DibujoAncho) / 2, (.ScaleHeight - DibujoAlto) / 2, DibujoAncho, DibujoAlto, DCOrigen, 0, 0, ImagenAncho, ImagenAlto, &HCC0020)
        SelectObject DCOrigen, Anterior
        DeleteDC DCOrigen
        DCOrigen = 0: Anterior = 0
        If Resultado = 0 Then GoTo FondoOriginal
        Set Me.Picture = .Image
    End With
    RedrawWindow Me.hWnd, ByVal 0&, 0, &H185
    AnchoFondo = Ancho: AltoFondo = Alto
    PintandoFondo = False
    Exit Sub
FondoOriginal:
    PintandoFondo = False
    ' Mantiene el recurso incorporado si no se puede crear el lienzo de ese tamaño.
    On Error Resume Next
    If DCOrigen <> 0 Then
        If Anterior <> 0 And Anterior <> -1 Then SelectObject DCOrigen, Anterior
        DeleteDC DCOrigen
    End If
    Set Me.Picture = PicFondo.Picture
End Sub

Private Sub MDIForm_Resize()
    ActualizarFondo
    ' VB6 termina de distribuir las barras y el cliente MDI despues de Resize.
    TmrFondo.Enabled = False
    If Me.WindowState <> vbMinimized Then TmrFondo.Enabled = True
End Sub

Private Sub TmrFondo_Timer()
    TmrFondo.Enabled = False
    ActualizarFondo
End Sub

Private Sub A000040_Click()
    AbrirModulo 40
End Sub

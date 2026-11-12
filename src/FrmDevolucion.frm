VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDevolucion 
   BackColor       =   &H00B4A79C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Devoluciones"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   12360
   Begin VB.Frame FraMotivo 
      BackColor       =   &H009E8D7E&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   3960
      TabIndex        =   15
      Top             =   360
      Visible = 0
      Width           =   4455
      Begin VB.ListBox LstMotivo 
         Appearance      =   0  'Flat
         Height          =   2760
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox TxtNewMot 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   4215
      End
   Begin VB.Label LabelDegrade2
      Left=0
      Top=120
      Width=4455
      Height=255
      Caption="MOTIVOS DE DEVOLUCIÓN"
      BackStyle=0
      BackStyle=1
      BackColor=&H0083443B&
      ForeColor=&H00FFFFFF&

   End
      Begin Clever.Boton CmdAcep 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         TX              =   "Aceptar"
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDevolucion.frx":0000
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
   Begin VB.Frame FraNew 
      BackColor       =   &H009E8D7E&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Visible = 0
      Width           =   10575
      Begin Clever.Boton CmdCrearCart 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   8640
         TabIndex        =   95
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         TX              =   "Crear Carta"
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDevolucion.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox Chkesp 
         Caption = "Sin documentos asociados"
         Height          =   315
         Left            =   8040
         TabIndex        =   97
         Top             =   240
         Width           =   2295
      End
   Begin VB.Label LabelDegrade1
      Left=0
      Top=0
      Width=10575
      Height=255
      Caption="NUEVA CARTA DE DEVOLUCIÓN"
      BackStyle=0
      BackStyle=1
      BackColor=&H0083443B&
      ForeColor=&H00FFFFFF&

   End
      Begin VB.CheckBox ChkCli 
         Caption = "Destinatario manual"
         Height          =   315
         Left            =   3240
         TabIndex        =   96
         Top             =   240
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         Height          =   375
         Left            =   6480
         TabIndex        =   92
         Top             =   510
         Width           =   1455
   Begin VB.Label LblCod
      Left=0
      Top=120
      Width=1455
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.ComboBox CmbTipFac 
         Height          =   315
         Left            =   8040
         Style = 2
         TabIndex        =   91
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   315
         Left            =   7080
         Style = 2
         TabIndex        =   90
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox CmdBancos 
         Height          =   315
         Left            =   5160
         Style = 2
         TabIndex        =   89
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox cmbTip 
         Height          =   315
         Left            =   5160
         Style = 2
         TabIndex        =   88
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox CmbMotivo 
         Height          =   315
         Left            =   1320
         Style = 2
         TabIndex        =   87
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox CmbModal 
         Enabled = 0
         Height          =   315
         Left            =   1320
         Style = 2
         TabIndex        =   86
         Top             =   960
         Width           =   2895
      End
   Begin VB.Label SkinLabel23
      Left=4320
      Top=960
      Width=615
      Height=255
      Caption="Forma:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel24
      Left=4320
      Top=1320
      Width=615
      Height=255
      Caption="Banco:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel25
      Left=4320
      Top=1680
      Width=615
      Height=255
      Caption="Cuenta:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel26
      Left=4320
      Top=2040
      Width=615
      Height=255
      Caption="Monto:"
      BackStyle=0

   End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   4995
         Picture         =   "FrmDevolucion.frx":0240
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   27
         Top             =   2040
         Width           =   135
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   4995
         Picture         =   "FrmDevolucion.frx":0449
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   26
         Top             =   1320
         Width           =   135
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   4995
         Picture         =   "FrmDevolucion.frx":0652
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   25
         Top             =   960
         Width           =   135
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1140
         Picture         =   "FrmDevolucion.frx":085B
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   24
         Top             =   2400
         Width           =   135
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1140
         Picture         =   "FrmDevolucion.frx":0A64
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   23
         Top             =   1680
         Width           =   135
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1140
         Picture         =   "FrmDevolucion.frx":0C6D
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   22
         Top             =   1320
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1140
         Picture         =   "FrmDevolucion.frx":0E76
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   21
         Top             =   960
         Width           =   135
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1140
         Picture         =   "FrmDevolucion.frx":107F
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   20
         Top             =   600
         Width           =   135
      End
      Begin VB.TextBox TxtCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAE6E5&
         Height          =   285
         Left            =   1320
         MaxLength = 100
         TabIndex        =   2
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox txtGui 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         MaxLength = 10
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ListBox LstGui 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "FrmDevolucion.frx":1288
         Left            =   8040
         List            =   "FrmDevolucion.frx":128A
         TabIndex        =   11
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox TxtContac 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength = 50
         TabIndex        =   9
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox txtMont 
         MaxLength = 18
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox TxtCta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength = 50
         TabIndex        =   7
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox TxtAbon 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength = 100
         TabIndex        =   6
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox TxtCan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength = 100
         TabIndex        =   5
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ListBox LstFac 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "FrmDevolucion.frx":128C
         Left            =   8040
         List            =   "FrmDevolucion.frx":128E
         TabIndex        =   4
         Top             =   915
         Width           =   2295
      End
      Begin VB.TextBox TxtFac 
         MaxLength = 14
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8880
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin Clever.Boton CmdConf 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   315
         Left            =   3840
         TabIndex        =   14
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         TX              =   ""
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDevolucion.frx":1290
         PICN            =   "FrmDevolucion.frx":12AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   Begin VB.Label SkinLabel16
      Left=240
      Top=240
      Width=855
      Height=255
      Caption="Registro:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel17
      Left=240
      Top=600
      Width=855
      Height=255
      Caption="Cliente:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel18
      Left=240
      Top=960
      Width=855
      Height=255
      Caption="Servicio:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel19
      Left=240
      Top=1320
      Width=855
      Height=255
      Caption="Motivo:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel20
      Left=240
      Top=1680
      Width=855
      Height=255
      Caption="Cancelado:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel21
      Left=240
      Top=2040
      Width=855
      Height=255
      Caption="Abonos:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel22
      Left=240
      Top=2400
      Width=855
      Height=255
      Caption="Contacto:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel27
      Left=6360
      Top=2040
      Width=615
      Height=255
      Caption="Moneda:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel28
      Left=8040
      Top=1680
      Width=615
      Height=255
      Caption="Guía:"
      BackStyle=0

   End
   Begin VB.Label LblCorrel
      Left=1320
      Top=240
      Width=1695
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin Clever.Boton CmdCancelar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=6600
      Top=2880
      Width=1935
      Height=375
      TX="Cancelar"
   End
   End
   Begin VB.ComboBox CmbFiltro 
      Style = 2
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Text            =   "Combo1"
      Top             =   120
      Width           =   3255
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   35
      Top             =   5100
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "20/04/2011"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "dd"
            TextSave        =   "dd"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   3360
      Width           =   10095
   Begin VB.Label SkinLabel6
      Left=120
      Top=120
      Width=975
      Height=255
      Caption="Registro:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel7
      Left=120
      Top=480
      Width=975
      Height=255
      Caption="Cliente:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel8
      Left=120
      Top=840
      Width=975
      Height=255
      Caption="Modalidad:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel9
      Left=120
      Top=1200
      Width=975
      Height=255
      Caption="Motivo:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel10
      Left=2700
      Top=120
      Width=615
      Height=255
      Caption="Fecha:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel11
      Left=3600
      Top=840
      Width=615
      Height=255
      Caption="Monto:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel12
      Left=3600
      Top=1200
      Width=615
      Height=255
      Caption="Moneda:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel13
      Left=5640
      Top=840
      Width=855
      Height=255
      Caption="Cuenta:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel14
      Left=5640
      Top=1200
      Width=855
      Height=255
      Caption="Contacto:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel15
      Left=8280
      Top=480
      Width=735
      Height=255
      Caption="Usuario:"
      BackStyle=0

   End
   Begin VB.Label VisCor
      Width=1350
      Left=1200
      Top=120
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label VisCli
      Left=1200
      Top=480
      Width=6975
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label VisMod
      Left=1200
      Top=840
      Width=2295
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label VisMot
      Left=1200
      Top=1200
      Width=2295
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label VisFec
      Left=3420
      Top=120
      Width=975
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label LblEstado
      Left=4860
      Top=120
      Width=1455
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label VisMon
      Left=4320
      Top=840
      Width=1215
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label VisDiv
      Left=4320
      Top=1200
      Width=1215
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label Viscta
      Left=6600
      Top=840
      Width=3255
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label VisCont
      Left=6600
      Top=1200
      Width=3255
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label VisUsu
      Left=9120
      Top=480
      Width=735
      Height=255
      Caption=""
      BackStyle=0

   End
   End
   Begin VB.Frame Frame2 
      Height = 2415
      Left            =   120
      TabIndex        =   39
      Top = 840
      Width           =   10095
      Begin MSFlexGridLib.MSFlexGrid GridDevoluciones 
         Height = 2265
         Left            =   10
         TabIndex        =   40
         Top             =   120
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   4842
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
   Begin VB.Frame FraPanel 
      BackColor       =   &H00B4A79C&
      Caption = ""
      Height          =   4455
      Left            =   10320
      TabIndex        =   12
      Top             =   480
      Width           =   1935
   Begin VB.Label SkinLabel1
      Left=120
      Top=2640
      Width=975
      Height=255
      Caption="Pendientes"
      BackStyle=0

   End
      Begin Clever.Boton CmdCrear 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Crear Carta"
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDevolucion.frx":2662
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdAnular 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Anular Devolucion"
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDevolucion.frx":267E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdExportar 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   0
         TabIndex        =   29
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX = "Exportar listado"
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDevolucion.frx":269A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdFinalizar 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   0
         TabIndex        =   30
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Finalizar Gestion"
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDevolucion.frx":26B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdRegen 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Regenerar Carta"
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDevolucion.frx":26D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   Begin VB.Label LabelDegrade3
      Left=0
      Top=2280
      Width=1920
      Height=255
      Caption="ESTADÍSTICAS"
      BackStyle=0
      BackStyle=1
      BackColor=&H0083443B&
      ForeColor=&H00FFFFFF&

   End
   Begin VB.Label SkinLabel2
      Left=120
      Top=3000
      Width=975
      Height=255
      Caption="Finalizadas"
      BackStyle=0

   End
   Begin VB.Label SkinLabel3
      Left=120
      Top=3360
      Width=975
      Height=255
      Caption="Activas"
      BackStyle=0

   End
   Begin VB.Label SkinLabel4
      Left=120
      Top=3720
      Width=975
      Height=255
      Caption="Anuladas"
      BackStyle=0

   End
   Begin VB.Label SkinLabel5
      Left=120
      Top=4080
      Width=975
      Height=255
      Caption="Visibles"
      BackStyle=0

   End
   Begin VB.Label CountPen
      Left=1320
      Top=2640
      Width=495
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label CountFin
      Left=1320
      Top=3000
      Width=495
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label CountAct
      Left=1320
      Top=3360
      Width=495
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label CountIna
      Left=1320
      Top=3720
      Width=495
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin VB.Label CountTot
      Left=1320
      Top=4080
      Width=495
      Height=255
      Caption=""
      BackStyle=0

   End
   Begin Clever.Boton CmdHistorial
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=0
      Top=1860
      Width=1935
      Height=375
      TX="Ver historial"
   End
   End
   Begin VB.TextBox TxtBuscar
      Left=3600
      Top=120
      Width=6400
      Height=315
      MaxLength=100

   End
   Begin Clever.Boton CmdActualizar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=10320
      Top=0
      Width=1935
      Height=375
      TX="Actualizar"
   End
End
Attribute VB_Name = "FrmDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Datos As ADODB.Recordset, Cliente As ADODB.Recordset, Docs As Object, Clave As String, IdGuardado As Long, Seleccionando As Boolean, Iniciando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(16) Then Unload Me: Exit Sub
 Iniciando = True: StatusBar.Style = sbrSimple
 CmbFiltro.Clear: CmbFiltro.AddItem "Todos": CmbFiltro.AddItem "Abierto": CmbFiltro.AddItem "Cerrado": CmbFiltro.AddItem "Anulado": CmbFiltro.ListIndex = 0
 CmbModal.Clear: CmbModal.AddItem "IMPORTACION": CmbModal.ListIndex = 0
 cmbTip.Clear: cmbTip.AddItem "TRANSFERENCIA": cmbTip.AddItem "CHEQUE": cmbTip.AddItem "PARA DEPOSITAR": cmbTip.ListIndex = 0
 CmbTipFac.Clear: CmbTipFac.AddItem "FT": CmbTipFac.AddItem "BV": CmbTipFac.AddItem "ND": CmbTipFac.ListIndex = 0
 CobCombo CmdBancos, 16, 1: CobCombo CmbMoneda, 16, 3: CargarMotivos
 Set Docs = CobNuevosDocumentos: FraNew.Visible = False: FraMotivo.Visible = False
 Iniciando = False: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title: Unload Me
End Sub
Public Sub Actualizar()
 Dim E As String, Abiertos As Long, Cerrados As Long, Anulados As Long
 On Error GoTo Fallo
 If CmbFiltro.ListIndex > 0 Then E = CmbFiltro.Text
 CerrarRecordset Datos: Set Datos = CobLeer(16, 0, 0, 0, Trim$(TxtBuscar.Text), E)
 Do While Not Datos.EOF
  Select Case CStr(Datos!EstadoGestion)
   Case "Abierto": Abiertos = Abiertos + 1
   Case "Cerrado": Cerrados = Cerrados + 1
   Case "Anulado": Anulados = Anulados + 1
  End Select
  Datos.MoveNext
 Loop
 If Datos.RecordCount > 0 Then Datos.MoveFirst
 CobGrilla GridDevoluciones, Datos, "Carta|Fecha|Cliente|Motivo|Moneda|Monto|Forma / banco|Estado", "NroCar|FecCar|Client|Motivo|Moneda|Monto|Transf|EstadoGestion"
 CountPen.Caption = CStr(Abiertos): CountFin.Caption = CStr(Cerrados): CountAct.Caption = CStr(Abiertos + Cerrados): CountIna.Caption = CStr(Anulados): CountTot.Caption = CStr(Datos.RecordCount)
 MostrarFicha
 StatusBar.SimpleText = CStr(Datos.RecordCount) & " devoluciones visibles (máximo 200). Los importes del sistema emisor se conservan."
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub MostrarFicha()
 Dim Id As Long
 On Error GoTo Vacio
 Id = CobFila(Datos, GridDevoluciones)
 VisCor.Caption = CStr(Datos!NroCar): VisCli.Caption = CStr(Datos!Client): VisMod.Caption = CStr(Datos!ImpExp): VisMot.Caption = CStr(Datos!Motivo)
 VisFec.Caption = FechaDocumento(Datos!FecCar): LblEstado.Caption = CStr(Datos!EstadoGestion): VisMon.Caption = Format$(Datos!Monto, "#,##0.00")
 VisDiv.Caption = CStr(Datos!Moneda): Viscta.Caption = CStr(Datos!CtaCte): VisCont.Caption = CStr(Datos!Contac): VisUsu.Caption = ValorTexto(Datos!Creador)
 Exit Sub
Vacio:
 VisCor.Caption = "": VisCli.Caption = "": VisMod.Caption = "": VisMot.Caption = "": VisFec.Caption = "": LblEstado.Caption = "": VisMon.Caption = "": VisDiv.Caption = "": Viscta.Caption = "": VisCont.Caption = "": VisUsu.Caption = ""
End Sub
Private Sub CmdCrear_Click()
 CerrarRecordset Cliente: Set Docs = CobNuevosDocumentos
 Clave = NuevaClaveQuery: IdGuardado = 0
 TxtCli.Text = "": LblCod.Caption = "": LblCorrel.Caption = "Nueva devolución": TxtCan.Text = "": TxtAbon.Text = "": TxtCta.Text = "": TxtContac.Text = "": txtMont.Text = "0.00": TxtFac.Text = "": txtGui.Text = "": LstFac.Clear: LstGui.Clear
 ChkCli.Value = 0: Chkesp.Value = 0
 FraNew.Visible = True: FraNew.ZOrder 0: FraPanel.Enabled = False: Frame1.Enabled = False: Frame2.Enabled = False: CmbFiltro.Enabled = False
End Sub
Private Sub CmdCancelar_Click()
 FraNew.Visible = False: FraMotivo.Visible = False: FraNew.Enabled = True: FraPanel.Enabled = True: Frame1.Enabled = True: Frame2.Enabled = True: CmbFiltro.Enabled = True
End Sub
Private Sub ChkCli_Click()
 CerrarRecordset Cliente: LblCod.Caption = "": TxtCli.Text = "": Set Docs = CobNuevosDocumentos: LstFac.Clear
 If ChkCli.Value = 1 Then Chkesp.Value = 1
End Sub
Private Sub Chkesp_Click()
 If Chkesp.Value = 1 Then Set Docs = CobNuevosDocumentos: LstFac.Clear
 CmbTipFac.Enabled = Chkesp.Value = 0: TxtFac.Enabled = Chkesp.Value = 0: LstFac.Enabled = Chkesp.Value = 0
End Sub
Private Sub TxtCli_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If ChkCli.Value = 1 Then Exit Sub
 If KeyCode <> vbKeyF9 And KeyCode <> vbKeyReturn Then Exit Sub
 KeyCode = 0: Set R = CobBuscarCliente(16, TxtCli.Text, Me)
 If R Is Nothing Then Exit Sub
 CerrarRecordset Cliente: Set Cliente = R: Set Docs = CobNuevosDocumentos: LstFac.Clear
 TxtCli.Text = CStr(R!CLIENT): LblCod.Caption = CStr(R!Asient): TxtContac.Text = Left$(ValorTexto(R!Contac), 50)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub TxtFac_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo Fallo
 If KeyCode <> vbKeyReturn Then Exit Sub
 KeyCode = 0: CobAgregarDocumento 16, Cliente, CmbMoneda.Text, CmbTipFac.Text, Trim$(TxtFac.Text), Docs, LstFac: TxtFac.Text = ""
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub LstFac_KeyDown(KeyCode As Integer, Shift As Integer)
 CobRetirarDocumento Docs, LstFac, KeyCode
End Sub
Private Sub txtGui_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim I As Integer, Numero As String
 On Error GoTo Fallo
 If KeyCode <> vbKeyReturn Then Exit Sub
 KeyCode = 0: Numero = Trim$(txtGui.Text)
 If Len(Numero) < 1 Or Len(Numero) > 10 Then Err.Raise 5, , "Ingrese una guía de hasta diez caracteres."
 If LstGui.ListCount >= 6 Then Err.Raise 5, , "Se permiten hasta seis guías."
 For I = 0 To LstGui.ListCount - 1
  If LstGui.List(I) = Numero Then Err.Raise 5, , "La guía ya está en la lista."
 Next I
 LstGui.AddItem Numero: txtGui.Text = ""
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub LstGui_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete And LstGui.ListIndex >= 0 Then LstGui.RemoveItem LstGui.ListIndex: KeyCode = 0
End Sub
Private Sub CmdCrearCart_Click()
 Dim S As String, R As ADODB.Recordset, I As Integer, Guardado As Boolean, Banco As Long
 On Error GoTo Fallo
 If ChkCli.Value = 0 Then
  If Cliente Is Nothing Then Err.Raise 5, , "Seleccione el cliente con F9."
  If TxtCli.Text <> CStr(Cliente!CLIENT) Then Err.Raise 5, , "Vuelva a seleccionar el cliente con F9."
 End If
 If CmdBancos.ListIndex >= 0 Then Banco = CobIdCombo(CmdBancos)
 If MsgBox("¿Registrar la devolución solicitada y generar su carta?", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 S = CobXMLBase(Cliente, CobIdCombo(CmbMoneda), txtMont.Text, TxtContac.Text)
 S = S & AtributoXML("Gastos", "0") & AtributoXML("Banco", CStr(Banco)) & AtributoXML("Motivo", "0") & AtributoXML("Cargo", "") & AtributoXML("Manual", CStr(ChkCli.Value)) & AtributoXML("SinDocumentos", CStr(Chkesp.Value)) & AtributoXML("NombreManual", Trim$(TxtCli.Text)) & AtributoXML("Modalidad", cmbTip.Text) & AtributoXML("Referencia", "") & AtributoXML("MotivoDevol", CmbMotivo.Text) & AtributoXML("Cuenta", Trim$(TxtCta.Text)) & AtributoXML("Cancelado", Trim$(TxtCan.Text)) & AtributoXML("Abonos", Trim$(TxtAbon.Text)) & ">" & CobXMLDocumentos(Docs)
 For I = 0 To LstGui.ListCount - 1: S = S & "<g" & AtributoXML("Numero", LstGui.List(I)) & "/>": Next I
 S = S & "</registro>"
 Set R = CobGuardar(16, Clave, S): IdGuardado = CLng(R!Item): Guardado = True: CerrarRecordset R
 CmdCancelar_Click
 Actualizar
 CobReporte 16, IdGuardado
 Exit Sub
Fallo:
 If Guardado Then MsgBox "La devolución quedó registrada. Use Regenerar carta.", vbInformation, App.Title Else MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdFinalizar_Click()
 On Error GoTo Fallo
 CobCerrarSeleccionado 16, Datos, GridDevoluciones, False: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdAnular_Click()
 On Error GoTo Fallo
 CobCerrarSeleccionado 16, Datos, GridDevoluciones, True: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdRegen_Click()
 On Error GoTo Fallo
 CobReporte 16, CobFila(Datos, GridDevoluciones)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdHistorial_Click()
 On Error GoTo Fallo
 CobHistorial 16, CobFila(Datos, GridDevoluciones)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdExportar_Click()
 On Error GoTo Fallo
 CobExportar Datos, "Devoluciones de dinero", "NroCar|FecCar|Client|Motivo|Moneda|Monto|Transf|Otros|ParDep|EstadoGestion"
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdConf_Click()
 FraMotivo.Visible = True: FraMotivo.ZOrder 0: FraNew.Enabled = False
End Sub
Private Sub CmdAcep_Click()
 FraMotivo.Visible = False: FraNew.Enabled = True
End Sub
Private Sub CargarMotivos()
 Dim R As ADODB.Recordset
 Set R = CobCatalogo(16, 7): CmbMotivo.Clear: LstMotivo.Clear
 Do While Not R.EOF
  CmbMotivo.AddItem CStr(R!Descripcion): LstMotivo.AddItem CStr(R!Descripcion): R.MoveNext
 Loop
 CerrarRecordset R
 If CmbMotivo.ListCount > 0 Then CmbMotivo.ListIndex = 0
End Sub
Private Sub TxtNewMot_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If KeyCode <> vbKeyReturn Then Exit Sub
 KeyCode = 0
 Set R = ConsultaOperacion("dbo.CobGuardarCatalogo", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Tipo", adInteger, 0, 7), Array("@Item", adInteger, 0, 0), Array("@Version", adBinary, 8, Null), Array("@Codigo", adVarWChar, 20, ""), Array("@Descripcion", adVarWChar, 150, Trim$(TxtNewMot.Text)), Array("@Baja", adBoolean, 0, False)), True)
 CerrarRecordset R: TxtNewMot.Text = "": CargarMotivos
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdActualizar_Click()
 Actualizar
End Sub
Private Sub CmbFiltro_Click()
 If Not Iniciando Then Actualizar
End Sub
Private Sub TxtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyF9 Then Actualizar: KeyCode = 0
End Sub
Private Sub GridDevoluciones_Click()
 SeleccionFilaQuery GridDevoluciones, Seleccionando: MostrarFicha
End Sub
Private Sub GridDevoluciones_RowColChange()
 If Seleccionando Then Exit Sub
 SeleccionFilaQuery GridDevoluciones, Seleccionando: MostrarFicha
End Sub
Private Sub GridDevoluciones_DblClick()
 CmdRegen_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos: CerrarRecordset Cliente
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

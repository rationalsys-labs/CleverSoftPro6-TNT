VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCheqRech 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Cheques Rechazados"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   13215
   Begin VB.Frame FraNew 
      Height          =   3375
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Visible = 0
      Width           =   10575
      Begin VB.Frame Frame13 
         Height          =   615
         Left            =   120
         TabIndex        =   84
         Top             =   2640
         Width           =   10335
         Begin Clever.Boton CmdCrearCart 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
            Height          =   375
            Left            =   6720
            TabIndex        =   85
            Top             =   165
            Width           =   1695
            _ExtentX        =   2990
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
            MICON           =   "FrmCheqRech.frx":0000
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Clever.Boton CmdCan 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
            Height          =   375
            Left            =   8520
            TabIndex        =   86
            Top             =   165
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            TX              =   "Cancelar"
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
            MICON           =   "FrmCheqRech.frx":001C
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
      Begin VB.Frame Frame12 
         Height          =   375
         Left            =   8760
         TabIndex        =   73
         Top             =   750
         Width           =   1575
   Begin VB.Label TxtRuc
      Left=30
      Top=120
      Width=1515
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.ComboBox TxtMot 
         Height          =   315
         Left            =   5520
         Style = 2
         TabIndex        =   72
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox CmbMon 
         Height          =   315
         Left            =   3360
         Style = 2
         TabIndex        =   71
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox TxtBank 
         Height          =   315
         Left            =   1320
         Style = 2
         TabIndex        =   70
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Height          =   375
         Left            =   1320
         TabIndex        =   68
         Top             =   390
         Width           =   1695
   Begin VB.Label LblCorrel
      Left=30
      Top=120
      Width=1635
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   3
         Left            =   8590
         Picture         =   "FrmCheqRech.frx":00F4
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   22
         Top             =   1920
         Width           =   135
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   2
         Left            =   8590
         Picture         =   "FrmCheqRech.frx":02FD
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   21
         Top             =   1560
         Width           =   135
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   5355
         Picture         =   "FrmCheqRech.frx":0506
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   20
         Top             =   2280
         Width           =   135
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   5355
         Picture         =   "FrmCheqRech.frx":070F
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   19
         Top             =   1920
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1140
         Picture         =   "FrmCheqRech.frx":0918
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   18
         Top             =   2280
         Width           =   135
      End
      Begin VB.TextBox TxtGas 
         MaxLength = 18
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8760
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox TxtMonCan 
         Locked = -1
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8760
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox TxtCarg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength = 50
         TabIndex        =   10
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtAten 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength = 50
         TabIndex        =   8
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox TxtnumCheq 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength = 50
         TabIndex        =   9
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox TxtMont 
         MaxLength = 18
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength = 100
         TabIndex        =   6
         Top             =   840
         Width           =   6495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1140
         Picture         =   "FrmCheqRech.frx":0B21
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   5
         Top             =   840
         Width           =   135
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1140
         Picture         =   "FrmCheqRech.frx":0D2A
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   4
         Top             =   1560
         Width           =   135
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   5355
         Picture         =   "FrmCheqRech.frx":0F33
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   3
         Top             =   1560
         Width           =   135
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   1140
         Picture         =   "FrmCheqRech.frx":113C
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   2
         Top             =   1920
         Width           =   135
      End
   Begin VB.Label LabelDegrade2
      Left=0
      Top=0
      Width=10575
      Height=255
      Caption="REGISTRO DE CHEQUE RECHAZADO"
      BackStyle=0
      BackStyle=1
      BackColor=&H0083443B&
      ForeColor=&H00FFFFFF&

   End
   Begin VB.Label SkinLabel16
      Left=240
      Top=480
      Width=855
      Height=255
      Caption="Registro:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel17
      Left=240
      Top=840
      Width=855
      Height=255
      Caption="Cliente:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel18
      Left=240
      Top=1200
      Width=855
      Height=255
      Caption="Dirección:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel19
      Left=240
      Top=1560
      Width=855
      Height=255
      Caption="Banco:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel20
      Left=240
      Top=1920
      Width=855
      Height=255
      Caption="Monto:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel21
      Left=240
      Top=2280
      Width=855
      Height=255
      Caption="Contacto:"
      BackStyle=0

   End
      Begin VB.Frame dfgdfg 
         Height          =   375
         Left            =   1320
         TabIndex        =   75
         Top             =   1110
         Width           =   9015
   Begin VB.Label TxtDir
      Left=30
      Top=120
      Width=8955
      Height=285
      Caption=""
      BackStyle=0

   End
      End
   Begin VB.Label SkinLabel11
      Left=7920
      Top=840
      Width=735
      Height=255
      Caption="RUC / DNI:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel12
      Left=4320
      Top=1560
      Width=855
      Height=255
      Caption="Motivo:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel13
      Left=4320
      Top=1920
      Width=855
      Height=255
      Caption="Cheque:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel14
      Left=4320
      Top=2280
      Width=855
      Height=255
      Caption="Cargo:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel15
      Left=7920
      Top=1560
      Width=615
      Height=255
      Caption="Giro:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel22
      Left=7920
      Top=1920
      Width=615
      Height=255
      Caption="Gastos:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel23
      Left=7920
      Top=2280
      Width=735
      Height=255
      Caption="Total:"
      BackStyle=0

   End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label LblCod 
         Appearance      =   0  'Flat
         BackColor       =   &H009E8D7E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9000
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
   Begin VB.TextBox TxtFechaGiro
      Left=8760
      Top=1560
      Width=1575
      Height=315
      MaxLength=10

   End
   End
   Begin VB.ComboBox CmbFiltro 
      Height          =   315
      Left            =   2400
      Style = 2
      TabIndex        =   26
      Top             =   120
      Width           =   3015
   End
   Begin VB.Frame FraInf 
      Height = 3135
      Left            =   2400
      TabIndex        =   27
      Top = 1080
      Width           =   10695
      Begin MSFlexGridLib.MSFlexGrid GridCheq 
         Height = 2985
         Left            =   15
         TabIndex        =   28
         Top             =   120
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   6535
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
   Begin VB.Frame FraPanel 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin Clever.Boton CmdCrear 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Crear Carta"
         ENAB            =   0   'False
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
         MICON           =   "FrmCheqRech.frx":1933
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
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Finalizar Carta"
         ENAB            =   0   'False
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
         MICON           =   "FrmCheqRech.frx":194F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdRegCart 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Re-Generar Carta"
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
         MICON           =   "FrmCheqRech.frx":196B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdReport 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   87
         Top             =   1395
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX = "Exportar a Excel"
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
         MICON           =   "FrmCheqRech.frx":1987
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   Begin Clever.Boton CmdHistorial
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=1890
      Width=1935
      Height=375
      TX="Ver historial"
   End
   Begin Clever.Boton CmdAnular
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=2310
      Width=1935
      Height=375
      TX="Anular expediente"
   End
   Begin Clever.Boton CmdConfig
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=2730
      Width=1935
      Height=375
      TX="Configuración"
   End
   Begin Clever.Boton CmdActualizar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=3150
      Width=1935
      Height=375
      TX="Actualizar"
   End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   2400
      TabIndex        =   29
      Top             =   4200
      Width           =   10695
      Begin VB.Frame Frame11 
         Height          =   375
         Left            =   5880
         TabIndex        =   55
         Top             =   750
         Width           =   4695
   Begin VB.Label lblmot
      Left=30
      Top=120
      Width=4635
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.Frame Frame10 
         Height          =   375
         Left            =   3840
         TabIndex        =   53
         Top             =   750
         Width           =   1335
   Begin VB.Label lbldiv
      Left=30
      Top=120
      Width=1275
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.Frame Frame9 
         Height          =   375
         Left            =   840
         TabIndex        =   51
         Top             =   750
         Width           =   2295
   Begin VB.Label lblcheq
      Left=30
      Top=120
      Width=2235
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.Frame Frame8 
         Height          =   375
         Left            =   9600
         TabIndex        =   49
         Top             =   430
         Width           =   975
   Begin VB.Label lblcan
      Left=30
      Top=120
      Width=915
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.Frame Frame7 
         Height          =   375
         Left            =   5880
         TabIndex        =   47
         Top             =   430
         Width           =   1935
   Begin VB.Label lblgas
      Left=30
      Top=120
      Width=1875
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.Frame Frame5 
         Height          =   375
         Left            =   3840
         TabIndex        =   45
         Top             =   430
         Width           =   1335
   Begin VB.Label lblmon
      Left=30
      Top=120
      Width=1275
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.Frame Frame4 
      Width=1215
      Left=9360
         Height          =   375
         TabIndex        =   43
         Top             =   110
   Begin VB.Label lblitem
      Top=120
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.Frame Frame3 
         Height          =   375
         Left            =   840
         TabIndex        =   42
         Top             =   430
         Width           =   2295
   Begin VB.Label lblbank
      Left=30
      Top=120
      Width=2235
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.Frame Frame6 
         Height          =   375
         Left            =   840
         TabIndex        =   40
         Top             =   110
         Width           =   6975
   Begin VB.Label lblCli
      Left=30
      Top=120
      Width=6915
      Height=285
      Caption=""
      BackStyle=0

   End
      End
   Begin VB.Label SkinLabel1
      Left=120
      Top=210
      Width=615
      Height=255
      Caption="Cliente"
      BackStyle=0

   End
   Begin VB.Label SkinLabel2
      Left=120
      Top=555
      Width=615
      Height=255
      Caption="Banco"
      BackStyle=0

   End
   Begin VB.Label SkinLabel3
      Left=120
      Top=870
      Width=615
      Height=255
      Caption="Cheque"
      BackStyle=0

   End
   Begin VB.Label SkinLabel4
      Left=3240
      Top=555
      Width=615
      Height=255
      Caption="Monto"
      BackStyle=0

   End
   Begin VB.Label SkinLabel5
      Left=3240
      Top=870
      Width=615
      Height=255
      Caption="Divisa"
      BackStyle=0

   End
   Begin VB.Label SkinLabel6
      Left=5280
      Top=555
      Width=615
      Height=255
      Caption="Gastos"
      BackStyle=0

   End
   Begin VB.Label SkinLabel7
      Left=5280
      Top=870
      Width=615
      Height=255
      Caption="Motivo"
      BackStyle=0

   End
   Begin VB.Label SkinLabel8
      Width=900
      Left=8400
      Top=210
      Height=255
      Caption="Registro"
      BackStyle=0

   End
   Begin VB.Label SkinLabel9
      Left=8160
      Top=555
      Width=1455
      Height=255
      Caption="Total"
      BackStyle=0

   End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   60
      Top             =   5460
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtBuscar
      Left=5520
      Top=120
      Width=7500
      Height=315
      MaxLength=100

   End
   Begin VB.Label LblDesde
      Left=2400
      Top=660
      Width=750
      Height=255
      Caption="Desde:"
      BackStyle=0
   End
   Begin VB.TextBox TxtDesde
      Left=3180
      Top=600
      Width=1300
      Height=315
      MaxLength=10

   End
   Begin VB.Label LblHasta
      Left=4800
      Top=660
      Width=750
      Height=255
      Caption="Hasta:"
      BackStyle=0
   End
   Begin VB.TextBox TxtHasta
      Left=5580
      Top=600
      Width=1300
      Height=315
      MaxLength=10

   End
End
Attribute VB_Name = "FrmCheqRech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Datos As ADODB.Recordset, Cliente As ADODB.Recordset, Clave As String, IdGuardado As Long, Seleccionando As Boolean, Iniciando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(13) Then Unload Me: Exit Sub
 Iniciando = True: StatusBar1.Style = sbrSimple
 CmbFiltro.Clear: CmbFiltro.AddItem "Todos": CmbFiltro.AddItem "Abierto": CmbFiltro.AddItem "Cerrado": CmbFiltro.AddItem "Anulado": CmbFiltro.ListIndex = 0
 CobCombo TxtBank, 13, 1: CobCombo TxtMot, 13, 2: CobCombo CmbMon, 13, 3
 FraNew.Visible = False: Iniciando = False: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title: Unload Me
End Sub
Public Sub Actualizar()
 Dim E As String, D As Variant, H As Variant, F As Date
 On Error GoTo Fallo
 D = Null: H = Null
 If Len(Trim$(TxtDesde.Text)) > 0 Then
  If Not FechaGestionValida(TxtDesde.Text, F) Then Err.Raise 5, , "Fecha Desde inválida; use dd/mm/aaaa."
  D = F
 End If
 If Len(Trim$(TxtHasta.Text)) > 0 Then
  If Not FechaGestionValida(TxtHasta.Text, F) Then Err.Raise 5, , "Fecha Hasta inválida; use dd/mm/aaaa."
  H = F
 End If
 If CmbFiltro.ListIndex > 0 Then E = CmbFiltro.Text
 CerrarRecordset Datos: Set Datos = CobLeer(13, 0, 0, 0, Trim$(TxtBuscar.Text), E, D, H)
 CobGrilla GridCheq, Datos, "Registro|Fecha|Cliente|Banco|Cheque|Moneda|Monto|Gastos|Total|Estado", "Asient|FecCar|Client|NomBank|NroChe|Divisa|Monto|Gastos|MonCan|EstadoGestion"
 MostrarFicha
 StatusBar1.SimpleText = CStr(Datos.RecordCount) & " registros visibles (máximo 200). F9 busca clientes en el registro nuevo."
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub MostrarFicha()
 Dim Id As Long
 On Error GoTo Vacio
 Id = CobFila(Datos, GridCheq)
 lblCli.Caption = CStr(Datos!Client): lblitem.Caption = CStr(Datos!Asient): lblbank.Caption = ValorTexto(Datos!NomBank)
 lblcheq.Caption = ValorTexto(Datos!NroChe): lblmon.Caption = Format$(Datos!Monto, "#,##0.00"): lbldiv.Caption = ValorTexto(Datos!Divisa)
 lblgas.Caption = Format$(Datos!Gastos, "#,##0.00"): lblcan.Caption = Format$(Datos!MonCan, "#,##0.00"): lblmot.Caption = ValorTexto(Datos!MotivoNombre)
 Exit Sub
Vacio:
 lblCli.Caption = "": lblitem.Caption = "": lblbank.Caption = "": lblcheq.Caption = "": lblmon.Caption = "": lbldiv.Caption = "": lblgas.Caption = "": lblcan.Caption = "": lblmot.Caption = ""
End Sub
Private Sub CmdCrear_Click()
 CerrarRecordset Cliente
 Clave = NuevaClaveQuery: IdGuardado = 0
 TxtCliente.Text = "": TxtRuc.Caption = "": TxtDir.Caption = "": LblCod.Caption = "": LblCorrel.Caption = "Nuevo registro"
 TxtnumCheq.Text = "": TxtMont.Text = "0.00": TxtGas.Text = "0.00": TxtMonCan.Text = "0.00": txtAten.Text = "": TxtCarg.Text = "": TxtFechaGiro.Text = Format$(Date, "dd/mm/yyyy")
 FraNew.Visible = True: FraNew.ZOrder 0: FraPanel.Enabled = False: FraInf.Enabled = False: CmbFiltro.Enabled = False
End Sub
Private Sub CmdCan_Click()
 FraNew.Visible = False: FraPanel.Enabled = True: FraInf.Enabled = True: CmbFiltro.Enabled = True
End Sub
Private Sub TxtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If KeyCode <> vbKeyF9 And KeyCode <> vbKeyReturn Then Exit Sub
 KeyCode = 0: Set R = CobBuscarCliente(13, TxtCliente.Text, Me)
 If R Is Nothing Then Exit Sub
 CerrarRecordset Cliente: Set Cliente = R
 TxtCliente.Text = CStr(R!CLIENT): TxtRuc.Caption = CStr(R!RUCDNI): TxtDir.Caption = CStr(R!DIRECC): LblCod.Caption = CStr(R!Asient): txtAten.Text = Left$(ValorTexto(R!Contac), 50)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Totalizar()
 If IsNumeric(TxtMont.Text) And IsNumeric(TxtGas.Text) Then TxtMonCan.Text = Format$(CCur(TxtMont.Text) + CCur(TxtGas.Text), "0.00") Else TxtMonCan.Text = ""
End Sub
Private Sub TxtMont_Change()
 Totalizar
End Sub
Private Sub TxtGas_Change()
 Totalizar
End Sub
Private Sub CmdCrearCart_Click()
 Dim S As String, R As ADODB.Recordset, Giro As Date, Guardado As Boolean
 On Error GoTo Fallo
 If Cliente Is Nothing Then Err.Raise 5, , "Seleccione el cliente con F9."
 If TxtCliente.Text <> CStr(Cliente!CLIENT) Then Err.Raise 5, , "Vuelva a seleccionar el cliente con F9."
 If Not CobFechaGiro(TxtFechaGiro.Text, Giro) Then Err.Raise 5, , "Fecha de giro inválida. Use dd/mm/aaaa."
 If MsgBox("¿Registrar el cheque rechazado y generar su carta?", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 S = CobXMLBase(Cliente, CobIdCombo(CmbMon), TxtMont.Text, txtAten.Text)
 S = S & AtributoXML("Gastos", NumeroXML(TxtGas.Text)) & AtributoXML("Banco", CStr(CobIdCombo(TxtBank))) & AtributoXML("Motivo", CStr(CobIdCombo(TxtMot))) & AtributoXML("FechaGiro", Format$(Giro, "yyyy-mm-dd")) & AtributoXML("Cheque", Trim$(TxtnumCheq.Text)) & AtributoXML("Cargo", Trim$(TxtCarg.Text)) & AtributoXML("Manual", "0") & AtributoXML("SinDocumentos", "1") & AtributoXML("Modalidad", "") & AtributoXML("Referencia", "") & "/>"
 Set R = CobGuardar(13, Clave, S): IdGuardado = CLng(R!Item): Guardado = True: CerrarRecordset R
 CmdCan_Click
 Actualizar
 CobReporte 13, IdGuardado
 Exit Sub
Fallo:
 If Guardado Then MsgBox "El cheque quedó registrado. Recargue para consultar la carta.", vbInformation, App.Title Else MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdRegCart_Click()
 On Error GoTo Fallo
 CobReporte 13, CobFila(Datos, GridCheq)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdFinalizar_Click()
 On Error GoTo Fallo
 CobCerrarSeleccionado 13, Datos, GridCheq, False
 Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdAnular_Click()
 On Error GoTo Fallo
 CobCerrarSeleccionado 13, Datos, GridCheq, True
 Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdHistorial_Click()
 On Error GoTo Fallo
 CobHistorial 13, CobFila(Datos, GridCheq)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdReport_Click()
 On Error GoTo Fallo
 CobExportar Datos, "Cheques rechazados", "Asient|FecCar|Client|NomBank|NroChe|Divisa|Monto|Gastos|MonCan|EstadoGestion"
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdConfig_Click()
 CobMostrarConfig 13
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
Private Sub GridCheq_Click()
 SeleccionFilaQuery GridCheq, Seleccionando: MostrarFicha
End Sub
Private Sub GridCheq_RowColChange()
 If Seleccionando Then Exit Sub
 SeleccionFilaQuery GridCheq, Seleccionando: MostrarFicha
End Sub
Private Sub GridCheq_DblClick()
 CmdRegCart_Click
End Sub
Private Sub GridCheq_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then CmdRegCart_Click: KeyCode = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Cliente: CerrarRecordset Datos
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDiarRp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modulo de Trabajo - Facturación RP"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   1965
   ClientWidth     =   14970
   KeyPreview = -1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14970
   Begin VB.Frame FraMod 
      Height          =   2895
      Left            =   4185
      TabIndex        =   62
      Top             =   3825
      Visible = 0
      Width           =   10710
   Begin Clever.Boton Cmd
      Left=6150
      Top=2340
      Width=1200
      Height=315
      Index=11
      TX="Cancelar"
      BTYPE=6
   End
      Begin VB.Frame FraCalcPes 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   2565
         TabIndex        =   82
         Top             =   1215
         Visible = 0
         Width           =   4980
         Begin VB.TextBox TxtPesVol 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3285
            TabIndex        =   77
            Top             =   240
            Width           =   1530
         End
         Begin VB.TextBox txtPesFis 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            TabIndex        =   76
            Top             =   225
            Width           =   1575
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Peso V:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2565
            TabIndex        =   84
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Peso F:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1680
         Left            =   2565
         TabIndex        =   79
         Top             =   1080
         Width           =   5010
         Begin VB.TextBox CmbIGV 
            Locked = -1
            Appearance      =   0  'Flat
            BackColor       =   &H00E3F5FF&
            Height          =   285
            Left            =   2385
            TabIndex        =   104
            Top             =   900
            Width           =   825
         End
         Begin VB.TextBox CmbRc 
            Locked = -1
            Appearance      =   0  'Flat
            BackColor       =   &H00E3F5FF&
            Height          =   285
            Left            =   2385
            TabIndex        =   103
            Top             =   540
            Width           =   825
         End
         Begin VB.TextBox CmbRcPub 
            Visible = 0
            Appearance      =   0  'Flat
            BackColor       =   &H00E3F5FF&
            Height          =   285
            Left            =   4095
            TabIndex        =   102
            Top             =   540
            Width           =   780
         End
         Begin VB.OptionButton OptPub 
            Visible = 0
            Caption         =   "Publica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3285
            TabIndex        =   101
            Top             =   585
            Width           =   825
         End
         Begin VB.OptionButton OptAsig 
            Visible = 0
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1845
            TabIndex        =   100
            Top             =   585
            Width           =   555
         End
         Begin VB.ComboBox CmbMoneda 
            BackColor       =   &H00E3F5FF&
            Height          =   315
            ItemData        =   "FrmDiarRp.frx":0000
            Left            =   2385
            List            =   "FrmDiarRp.frx":0002
            TabIndex        =   72
            Top             =   180
            Width = 1290
         End
         Begin VB.TextBox Txt 
            Index = 2
            Appearance      =   0  'Flat
            BackColor       =   &H00E3F5FF&
            Height          =   315
            Left            =   855
            TabIndex        =   69
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox Txt 
            Index = 5
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   855
            TabIndex        =   75
            Top             =   1260
            Width           =   915
         End
         Begin VB.TextBox Txt 
            Index = 4
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   855
            TabIndex        =   74
            Top             =   900
            Width           =   915
         End
         Begin VB.TextBox Txt 
            Index = 3
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   855
            TabIndex        =   73
            Top             =   540
            Width           =   915
         End
         Begin Clever.Boton Cmd 
            Index = 10
            Height          =   315
            Left            =   2385
            TabIndex        =   70
            Top             =   1260
            Width = 1170
            _ExtentX        =   4471
            _ExtentY        =   556
            BTYPE           =   14
            TX = "Guardar"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDiarRp.frx":0004
            PICN            =   "FrmDiarRp.frx":0020
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% IGV:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1845
            TabIndex        =   99
            Top             =   945
            Width           =   480
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Divisa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1845
            TabIndex        =   98
            Top             =   225
            Width           =   480
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   135
            TabIndex        =   97
            Top             =   225
            Width           =   735
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   135
            TabIndex        =   96
            Top             =   1305
            Width           =   735
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RC"
            Height          =   195
            Left            =   135
            TabIndex        =   81
            Top             =   585
            Width           =   225
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "IGV"
            Height          =   240
            Left            =   135
            TabIndex        =   80
            Top             =   945
            Width           =   645
         End
      End
      Begin VB.ComboBox CmbCliente 
         Style = 2
         Appearance      =   0  'Flat
         BackColor       =   &H00E3F5FF&
         Height          =   285
         Left            =   945
         TabIndex        =   64
         Top             =   765
         Width           =   6630
      End
      Begin VB.TextBox Txt 
         Index = 1
         Appearance      =   0  'Flat
         BackColor       =   &H00E3F5FF&
         Height          =   315
         Left            =   945
         TabIndex        =   66
         Top             =   1620
         Width           =   1230
      End
      Begin VB.TextBox Txt 
         Index = 0
         Appearance      =   0  'Flat
         BackColor       =   &H00E3F5FF&
         Height          =   315
         Left            =   945
         TabIndex        =   65
         Top             =   1260
         Width           =   1230
      End
      Begin VB.TextBox txtModTer 
         Locked = -1
         Appearance      =   0  'Flat
         BackColor       =   &H00E3F5FF&
         Height          =   315
         Left            =   945
         TabIndex        =   68
         Top             =   2340
         Width           =   1230
      End
      Begin VB.ComboBox CmbModPro 
         Enabled = 0
         BackColor       =   &H00E3F5FF&
         Height          =   315
         ItemData        =   "FrmDiarRp.frx":05BA
         Left            =   945
         List            =   "FrmDiarRp.frx":05C7
         TabIndex        =   67
         Top             =   1980
         Width           =   1230
      End
      Begin VB.TextBox Txt 
         Index = 6
         Appearance      =   0  'Flat
         BackColor       =   &H00E3F5FF&
         Height          =   2340
         Left            =   7695
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   71
         Top             =   360
         Width           =   2820
      End
      Begin VB.TextBox TxtCod 
         Height          =   330
         Left            =   0
         TabIndex        =   78
         Top             =   0
         Visible = 0
         Width           =   330
      End
      Begin VB.TextBox LblModGui 
         Locked = -1
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   63
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label LblDiv 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   95
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   135
         TabIndex        =   94
         Top             =   765
         Width           =   870
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso V:"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   135
         TabIndex        =   93
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso F:"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   135
         TabIndex        =   92
         Top             =   1260
         Width           =   870
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   135
         TabIndex        =   91
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Territorio:"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   135
         TabIndex        =   90
         Top             =   2340
         Width           =   735
      End
      Begin VB.Label LblModAsi 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   89
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Asiento"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   88
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Guia"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5265
         TabIndex        =   86
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   85
         Top             =   360
         Width           =   1695
      End
   End
   Begin Clever.Boton CmdExpand 
      Height          =   285
      Left            =   4185
      TabIndex        =   23
      Top             =   6750
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   503
      BTYPE           =   14
      TX = "Volver a la pizarra completa"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDiarRp.frx":05DA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid GridRP 
      SelectionMode = 1
      Height = 6240
      Left            =   4185
      TabIndex        =   2
      Top             =   810
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   12885
      _Version        =   393216
      Rows            =   4
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   14737632
      Appearance      =   0
   End
   Begin VB.TextBox TxtInforGui 
      Locked = -1
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9630
      TabIndex        =   19
      Top             =   8190
      Width           =   5235
   End
   Begin VB.TextBox TxtComent 
      Locked = -1
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4185
      TabIndex        =   18
      Top             =   8190
      Width           =   5325
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   17
      Top             =   8565
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "05/10/2010"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
         EndProperty
      EndProperty
   End
   Begin VB.TextBox InforGui 
      Locked = -1
      Appearance      =   0  'Flat
      BackColor       =   &H005E6068&
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   9630
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   7110
      Width           =   5220
   End
   Begin VB.TextBox Coment 
      Locked = -1
      Appearance      =   0  'Flat
      BackColor       =   &H005E6068&
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   4185
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   7110
      Width           =   5310
   End
   Begin VB.Frame FraFil 
      Height          =   690
      Left            =   1935
      TabIndex        =   8
      Top             =   45
      Width           =   12930
      Begin VB.CheckBox ChkSalida 
         Caption = "Con salida"
         Height          =   315
         Left            =   4005
         TabIndex        =   14
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label LblBusqueda 
         Caption = "Buscar guía / cliente:"
         Height          =   315
         Left            =   9540
         TabIndex        =   13
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox TxtBuscar 
         Height          =   315
         Left            =   11115
         TabIndex        =   12
         Top             =   225
         Width           =   1680
      End
      Begin VB.ComboBox CmbGrupo 
         Height          =   315
         Left            =   945
         Style = 2
         TabIndex        =   9
         Top             =   225
         Width           =   2445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar:"
         Height          =   195
         Left            =   3510
         TabIndex        =   11
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Totales RP:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.Frame FraEsta 
      Caption         =   "Estadisticas:"
      Height = 7695
      Left            =   1935
      TabIndex        =   3
      Top             =   765
      Width           =   2175
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   90
         TabIndex        =   24
         Top             =   4635
         Width           =   1995
      End
      Begin Clever.Boton CmdDetalle 
         Height          =   375
         Left            =   90
         TabIndex        =   47
         Top             =   3105
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         BTYPE           =   14
         TX = "Consultar ficha"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":106C
         PICN            =   "FrmDiarRp.frx":1088
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdActualizar 
         Height          =   375
         Left            =   90
         TabIndex        =   48
         Top             =   3510
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         BTYPE           =   14
         TX = "Actualizar datos"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":1A9A
         PICN            =   "FrmDiarRp.frx":1AB6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdMarcar 
         Height          =   375
         Left            =   90
         TabIndex        =   49
         Top             =   3915
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         BTYPE           =   14
         TX = "Marcar filas visibles"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":2050
         PICN            =   "FrmDiarRp.frx":206C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1440
         TabIndex        =   7
         Top             =   540
         Width           =   600
      End
      Begin VB.Label ConGui 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1440
         TabIndex        =   5
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "DOC RP Salida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   46
         Top             =   2565
         Width           =   1035
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   45
         Top             =   2565
         Width           =   555
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "ECO RP Salida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   44
         Top             =   2790
         Width           =   1020
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   43
         Top             =   2790
         Width           =   555
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Economys RP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   42
         Top             =   2115
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   41
         Top             =   2115
         Width           =   555
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "XPS RP Salida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   40
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   39
         Top             =   2340
         Width           =   555
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Muestras RP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   38
         Top             =   1665
         Width           =   900
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   37
         Top             =   1665
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Documentos RP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   36
         Top             =   1890
         Width           =   1125
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   35
         Top             =   1890
         Width           =   555
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Rp Status COT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   34
         Top             =   1215
         Width           =   1065
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   33
         Top             =   1215
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Rps Retenidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   32
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   31
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Rps Con Salida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   765
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   29
         Top             =   765
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Rp Status Ventas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   990
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   27
         Top             =   990
         Width           =   555
      End
      Begin VB.Label LblStaGui 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Status de Guias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   90
         TabIndex        =   25
         Top             =   4365
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº de Reporte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Guias Filtradas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Width           =   1050
      End
   End
   Begin VB.Frame FraHer 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   8835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1860
      Begin Clever.Boton Cmd 
         Index = 8
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   450
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Actualizar"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":2606
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton Cmd 
         Index = 2
         Height          =   330
         Left            =   135
         TabIndex        =   51
         Top             =   765
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Registrar evento"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":2622
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton Cmd 
         Index = 5
         Height          =   330
         Left            =   135
         TabIndex        =   52
         Top             =   1080
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Ver historial"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":263E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton Cmd 
         Index = 6
         Height          =   330
         Left            =   135
         TabIndex        =   53
         Top             =   1440
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Ver compilatorio"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":265A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdImportacion 
         Height          =   330
         Left            =   135
         TabIndex        =   54
         Top             =   2025
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Importación central"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":2676
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdGrupos 
         Height          =   330
         Left            =   135
         TabIndex        =   55
         Top             =   2340
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Administrar grupos"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":2692
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton Cmd 
         Index = 4
         Height          =   330
         Left            =   135
         TabIndex        =   56
         Top             =   2655
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Cerrar gestión"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":26AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton Cmd 
         Index = 3
         Height          =   330
         Left            =   135
         TabIndex        =   57
         Top             =   2970
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Confirmar cartilla"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":26CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton Cmd 
         Index = 1
         Height          =   330
         Left            =   135
         TabIndex        =   58
         Top             =   3285
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Activar / quitar salida"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":26E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton Cmd 
         Index = 0
         Height          =   330
         Left            =   135
         TabIndex        =   59
         Top             =   3600
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
         TX = "Asignar / modificar"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":2702
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton Cmd 
         Index = 9
         Height          =   330
         Left            =   135
         TabIndex        =   61
         Top             =   4230
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BTYPE           =   9
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   -2147483633
         FCOL            =   14737632
         FCOLO           =   8421504
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDiarRp.frx":273A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HERRAMIENTAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   90
         TabIndex        =   50
         Top             =   135
         Width           =   1665
      End
   End
   Begin VB.Label LblEstado
      Left=0
      Top=0
      Width=20
      Height=20
      Visible=0
   End
   Begin Clever.Boton Cmd
      Left=0
      Top=0
      Width=20
      Height=20
      Index=7
      Visible=0
   End
End
Attribute VB_Name = "FrmDiarRp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Const EsCompilatorio As Boolean = False
Private Datos As ADODB.Recordset
Private Editando As Boolean
Private Cargando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(IIf(EsCompilatorio, 10, 9)) Then Unload Me: Exit Sub
 PrepararVista
 CargarOpciones
 Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
 Unload Me
End Sub
Private Sub CargarOpciones()
 Dim R As ADODB.Recordset
 Cargando = True
 CmbGrupo.Clear
 CmbCliente.Clear
 Set R = ConsultaOperacion("SELECT Item,Asient,CLIENT FROM dbo.DbClientes WHERE Activo=1 ORDER BY CLIENT")
 Do While Not R.EOF
  CmbCliente.AddItem CStr(R!Asient) & " - " & CStr(R!CLIENT)
  CmbCliente.ItemData(CmbCliente.NewIndex) = CLng(R!Item)
  R.MoveNext
 Loop
 CerrarRecordset R
 CmbMoneda.Clear
 Set R = ConsultaOperacion("SELECT CodEst,Moneda FROM dbo.Int_Moneda WHERE ActEst=1 ORDER BY CodEst")
 Do While Not R.EOF
  CmbMoneda.AddItem CStr(R!Moneda): CmbMoneda.ItemData(CmbMoneda.NewIndex) = CLng(R!CodEst)
  R.MoveNext
 Loop
 CerrarRecordset R
 If Not EsCompilatorio Then
  CargarComboGruposRP CmbGrupo
 Else
  CmbGrupo.AddItem "General": CmbGrupo.ItemData(0) = -1
 End If
 CmbGrupo.ListIndex = 0
 CmbGrupo.Enabled = Not EsCompilatorio
 ChkSalida.Visible = Not EsCompilatorio
 Cargando = False
End Sub
Private Sub Actualizar()
 Dim F As Long, I As Long, Grupo As Long
 On Error GoTo Fallo
 Cargando = True
 CerrarRecordset Datos
 GridRP.Rows = 1
 Grupo = -1
 If CmbGrupo.ListIndex >= 0 Then Grupo = CmbGrupo.ItemData(CmbGrupo.ListIndex)
 Set Datos = ConsultaOperacion("dbo.ConsultarPizarraRP", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Compilatorio", adBoolean, 0, EsCompilatorio), Array("@Filtro", adVarWChar, 100, Left$(TxtBuscar.Text, 100)), Array("@Grupo", adInteger, 0, Grupo), Array("@SoloSalida", adBoolean, 0, ChkSalida.Value = 1)), True)
 GridRP.FormatString = "Marca|Guía|Cliente|Producto|Gestión|Salida|Total|Fecha"
 If EsCompilatorio Then GridRP.FormatString = "|Guía|Cliente|Producto|Situación|Cartilla|Importe|Fecha de envío"
 GridRP.ColWidth(0) = 600: GridRP.ColWidth(1) = 1400: GridRP.ColWidth(2) = 3200
 GridRP.ColWidth(3) = 850: GridRP.ColWidth(4) = 1300: GridRP.ColWidth(5) = 1300
 GridRP.ColWidth(6) = 1100: GridRP.ColWidth(7) = 1150
 Do While Not Datos.EOF
  GridRP.AddItem ""
  F = GridRP.Rows - 1
  GridRP.TextMatrix(F, 1) = CStr(Datos!NroGui)
  GridRP.TextMatrix(F, 2) = ValorTexto(Datos!Client)
  GridRP.TextMatrix(F, 3) = ValorTexto(Datos!TipGui)
  If EsCompilatorio Then
   GridRP.TextMatrix(F, 4) = ValorTexto(Datos!FINAL)
   GridRP.TextMatrix(F, 5) = ValorTexto(Datos!NroPla)
  Else
   GridRP.TextMatrix(F, 4) = ValorTexto(Datos!EstadoGestion)
   GridRP.TextMatrix(F, 5) = IIf(Val(ValorTexto(Datos!Salida)) = 1, "SI", "NO")
  End If
  GridRP.TextMatrix(F, 6) = Format$(CCur(IIf(IsNull(Datos!MonTot), 0, Datos!MonTot)), "0.00")
  If Not IsNull(Datos!PuGui) Then GridRP.TextMatrix(F, 7) = Format$(Datos!PuGui, "dd/mm/yyyy")
  Datos.MoveNext
 Loop
 Cargando = False
 If GridRP.Rows > 1 Then GridRP.Row = 1
 MostrarFicha
 Habilitar False
 ActualizarVista
 LblEstado.Caption = CStr(GridRP.Rows - 1) & " registros. Doble clic para marcar las guías del lote."
 Exit Sub
Fallo:
 Cargando = False
 CerrarRecordset Datos
 GridRP.Rows = 1
 MostrarFicha
 LblEstado.Caption = "No se pudo actualizar: " & Err.Description
End Sub
Private Function Seleccionar() As Boolean
 If Datos Is Nothing Then Exit Function
 If GridRP.Row < 1 Or GridRP.Row >= GridRP.Rows Or Datos.RecordCount = 0 Then Exit Function
 Datos.MoveFirst
 Datos.Move GridRP.Row - 1
 Seleccionar = True
End Function
Private Sub Elegir(ByVal Combo As ComboBox, ByVal Id As Long)
 Dim I As Long
 Combo.ListIndex = -1
 For I = 0 To Combo.ListCount - 1
  If Combo.ItemData(I) = Id Then Combo.ListIndex = I: Exit Sub
 Next I
End Sub
Private Sub MostrarFicha()
 Dim I As Long, Campos As Variant
 If Cargando Or Editando Then Exit Sub
 For I = 0 To 6: Txt(I).Text = "": Next I
 CmbCliente.ListIndex = -1: CmbMoneda.ListIndex = -1
 If Not Seleccionar Then Exit Sub
 Elegir CmbCliente, Val(ValorTexto(Datos!IdCliente))
 Elegir CmbMoneda, Val(ValorTexto(Datos!Moneda))
 Campos = Array("PesFis", "PesVol", "MonNet", "MonRc", "MonIgv", "MonTot", "Info")
 For I = 0 To 6
  Txt(I).Text = ValorTexto(Datos.Fields(CStr(Campos(I))).Value)
 Next I
End Sub
Private Sub Habilitar(ByVal Edicion As Boolean)
 Dim I As Long, Hay As Boolean
 Editando = Edicion
 Hay = (GridRP.Rows > 1)
 For I = 0 To 9: Cmd(I).Enabled = Not Edicion: Next I
 For I = 0 To 4: Cmd(I).Visible = Not EsCompilatorio: Next I
 Cmd(7).Visible = EsCompilatorio
 Cmd(10).Visible = Edicion: Cmd(11).Visible = Edicion
 GridRP.Enabled = Not Edicion: TxtBuscar.Enabled = Not Edicion
 CmbGrupo.Enabled = Not Edicion And Not EsCompilatorio
 ChkSalida.Enabled = Not Edicion
 CmbCliente.Enabled = Edicion: CmbMoneda.Enabled = Edicion
 For I = 0 To 6: Txt(I).Locked = Not Edicion: Next I
 If Not EsCompilatorio Then MostrarEditor Edicion
End Sub
Private Function FilasXML(ByVal Lote As Boolean) As String
 Dim I As Long, R As String
 If Datos Is Nothing Then Exit Function
 For I = 1 To GridRP.Rows - 1
  If (Lote And GridRP.TextMatrix(I, 0) = "X") Or (Not Lote And I = GridRP.Row) Then
   Datos.MoveFirst: Datos.Move I - 1
   R = R & "<r" & AtributoXML("Item", CStr(Datos!ITEM)) & AtributoXML("Version", VersionXML(Datos.Fields("VersionFila").Value)) & "/>"
  End If
 Next I
 FilasXML = R
End Function
Private Sub Ejecutar(ByVal Accion As String, ByVal Atributos As String, Optional ByVal Lote As Boolean = False)
 Dim XML As String, R As ADODB.Recordset, Confirmado As Boolean, Mensaje As String
 On Error GoTo Fallo
 XML = FilasXML(Lote)
 If Len(XML) = 0 Then MsgBox "Seleccione o marque las guías.", vbInformation, App.Title: Exit Sub
 If MsgBox("¿Confirmar " & Accion & " para las guías seleccionadas?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
 XML = "<operacion" & Atributos & ">" & XML & "</operacion>"
 Set R = ConsultaOperacion("dbo.GuardarPizarraRP", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Accion", adVarWChar, 20, Accion), Array("@Datos", adLongVarWChar, Len(XML), XML)), True)
 Confirmado = True
 Mensaje = "Operación confirmada."
 If Not IsNull(R!Planilla) Then Mensaje = "Cartilla " & CStr(R!Planilla) & " confirmada. Puede consultarla y reimprimirla desde Compilatorio."
 CerrarRecordset R
 Habilitar False
 Actualizar
 MsgBox Mensaje, vbInformation, App.Title
 Exit Sub
Fallo:
 Mensaje = Err.Description
 CerrarRecordset R
 If Confirmado Then Habilitar False: Mensaje = "Operación confirmada; recargue sin repetirla. " & Mensaje
 MsgBox Mensaje, vbExclamation, App.Title
End Sub
Private Sub Cmd_Click(Index As Integer)
 Dim A As String, I As Long, Estado As String, Comentario As String, R As ADODB.Recordset, Texto As String, Campos As Variant
 On Error GoTo Fallo
 Select Case Index
 Case 8
  CargarOpciones: Actualizar
 Case 9
  Unload Me
 Case 11
  If MsgBox("¿Descartar el borrador?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) = vbYes Then Habilitar False: MostrarFicha
 Case Else
  If Not Seleccionar Then Exit Sub
  Select Case Index
  Case 0
   Habilitar True
  Case 1
   Ejecutar "Salida", AtributoXML("Salida", IIf(Val(ValorTexto(Datos!Salida)) = 1, "0", "1"))
  Case 2
   FrmEventoRP.Preparar
   FrmEventoRP.Show vbModal, MdiPri
   If Not FrmEventoRP.Aceptado Then Unload FrmEventoRP: Exit Sub
   A = AtributoXML("Estado", CStr(FrmEventoRP.CodigoEstado)) & AtributoXML("Comentario", FrmEventoRP.Comentario)
   Unload FrmEventoRP
   Ejecutar "Evento", A
  Case 3
   Ejecutar "Confirmar", "", True
  Case 4
   Ejecutar "Cerrar", "", True
  Case 5
   Set R = ConsultaOperacion("SELECT h.FECHA,h.STATUS,CONVERT(nvarchar(1000),h.INFO) Info FROM dbo.DbsHisRP h WHERE h.HERENCIA=? ORDER BY h.FECHA,h.Item", Array(Array("Asient", adVarWChar, 20, CStr(Datos!Asient))))
   Do While Not R.EOF
    Texto = Texto & ValorTexto(R!FECHA) & " " & ValorTexto(R!STATUS) & " " & ValorTexto(R!Info) & vbCrLf
    R.MoveNext
   Loop
   CerrarRecordset R
   If Len(Texto) = 0 Then Texto = "Sin eventos."
   MostrarHistorialRP Texto
  Case 6
   If Not EsCompilatorio Then AbrirModulo 10: Exit Sub
   MostrarCartillaRP ValorTexto(Datos!NroPla)
  Case 7
   Ejecutar "Reactivar", ""
  Case 10
   If CmbCliente.ListIndex < 0 Or CmbMoneda.ListIndex < 0 Then Err.Raise 5, , "Seleccione cliente y moneda."
   A = AtributoXML("Cliente", CStr(CmbCliente.ItemData(CmbCliente.ListIndex))) & AtributoXML("Moneda", CStr(CmbMoneda.ItemData(CmbMoneda.ListIndex)))
   Campos = Array("PesoFisico", "PesoVolumen", "Neto", "Recargo", "IGV", "Total")
   For I = 0 To 5
    A = A & AtributoXML(CStr(Campos(I)), NumeroXML(Txt(I).Text))
   Next I
   A = A & AtributoXML("Info", Txt(6).Text)
   Ejecutar "Guardar", A
  End Select
 End Select
 Exit Sub
Fallo:
 CerrarRecordset R
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub GridRP_RowColChange()
 MostrarFicha
 RefrescarSeleccion
End Sub
Private Sub GridRP_DblClick()
 If GridRP.Row > 0 Then GridRP.TextMatrix(GridRP.Row, 0) = IIf(GridRP.TextMatrix(GridRP.Row, 0) = "X", "", "X")
End Sub
Private Sub GridRP_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then MostrarFicha: ConsultarFicha: KeyCode = 0
 If KeyCode = vbKeySpace Then GridRP_DblClick: KeyCode = 0
End Sub
Private Sub TxtBuscar_Change()
 If Not Cargando And Not Editando Then Actualizar
End Sub
Private Sub CmbGrupo_Click()
 If Not Cargando And Not Editando Then Actualizar
End Sub
Private Sub ChkSalida_Click()
 If Not Cargando And Not Editando Then Actualizar
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If Editando Then Cancel = (MsgBox("¿Descartar el borrador y cerrar?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes)
 If Not Cancel Then CerrarRecordset Datos
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF5 And Not Editando Then CargarOpciones: Actualizar: KeyCode = 0
 If KeyCode = vbKeyEscape Then
  If Editando Then Cmd_Click 11 Else Unload Me
 End If
End Sub

Private Sub PrepararVista()
 GridRP.Rows = 2
 GridRP.FixedRows = 1
 GridRP.FixedCols = 0
 FraMod.Visible = False
 CmdExpand.Visible = False
 InforGui.Locked = True
 Coment.Locked = True
 TxtComent.Text = "Historial de la guía seleccionada"
 TxtInforGui.Text = "Información de la guía"
 CmbModPro.Clear
 CmbModPro.AddItem "15D": CmbModPro.AddItem "15N": CmbModPro.AddItem "48N"
End Sub
Private Sub MostrarEditor(ByVal Edicion As Boolean)
 CmdGrupos.Enabled = Not Edicion
 FraMod.Visible = Edicion
 CmdExpand.Visible = Edicion
 If Edicion Then
  GridRP.Height = 2925
  FraMod.ZOrder 0
  RefrescarSeleccion
 Else
  GridRP.Height = 6240
 End If
End Sub
Private Sub ConsultarFicha()
 If Not Seleccionar Then Exit Sub
 FraMod.Visible = True
 CmdExpand.Visible = True
 GridRP.Height = 2925
 FraMod.ZOrder 0
 RefrescarSeleccion
End Sub
Private Sub RefrescarSeleccion()
 Dim R As ADODB.Recordset, Texto As String, I As Long
 On Error GoTo Salir
 If Cargando Or Editando Then Exit Sub
 If Not Seleccionar Then
  Coment.Text = "": InforGui.Text = "": List1.Clear
  Exit Sub
 End If
 LblModGui.Text = CStr(Datos!NroGui)
 LblModAsi.Caption = CStr(Datos!Asient)
 LblFecha.Caption = ValorTexto(Datos!PuGui)
 TxtCod.Text = ValorTexto(Datos!IdCliente)
 txtModTer.Text = ""
 CmbModPro.Text = CStr(Datos!TipGui)
 If Val(ValorTexto(Datos!MonNet)) > 0 Then
  CmbRc.Text = Format$(CCur(Datos!MonRc) * 100 / CCur(Datos!MonNet), "0.00")
  CmbIGV.Text = Format$(CCur(Datos!MonIgv) * 100 / (CCur(Datos!MonNet) + CCur(Datos!MonRc)), "0.00")
 Else
  CmbRc.Text = "0": CmbIGV.Text = "0"
 End If
 InforGui.Text = "Guía: " & CStr(Datos!NroGui) & vbCrLf & "Origen: " & CStr(Datos!Origen) & vbCrLf & "Producto: " & CStr(Datos!TipGui) & vbCrLf & "Cliente: " & ValorTexto(Datos!Client) & vbCrLf & "Salida: " & IIf(Val(ValorTexto(Datos!Salida)) = 1, "SI", "NO")
 List1.Clear
 Set R = ConsultaOperacion("SELECT STATUS,FECHA,CONVERT(nvarchar(1000),INFO) Info FROM dbo.DbsHisRP WHERE HERENCIA=? ORDER BY FECHA,Item", Array(Array("Asient", adVarWChar, 20, CStr(Datos!Asient))))
 Do While Not R.EOF
  Texto = Texto & ValorTexto(R!FECHA) & " - " & ValorTexto(R!STATUS) & ": " & ValorTexto(R!Info) & vbCrLf
  List1.AddItem ValorTexto(R!STATUS) & " - " & ValorTexto(R!Info)
  R.MoveNext
 Loop
 Coment.Text = Texto
Salir:
 CerrarRecordset R
End Sub
Private Sub ActualizarVista()
 Dim R As ADODB.Recordset, I As Long, C(0 To 10) As Long, S As String, Producto As String
 On Error GoTo Fallo
 Set R = ConsultaOperacion("SELECT TipGui,Salida,Reten,Estado FROM dbo.DbsDiarRp")
 Do While Not R.EOF
  C(0) = C(0) + 1
  If Val(ValorTexto(R!Salida)) = 1 Then C(1) = C(1) + 1
  If Val(ValorTexto(R!Estado)) = 4 Then C(2) = C(2) + 1
  If Val(ValorTexto(R!Estado)) = 3 Then C(3) = C(3) + 1
  If Val(ValorTexto(R!Reten)) = 1 Then C(4) = C(4) + 1
  Producto = ValorTexto(R!TipGui)
  Select Case Producto
   Case "15N": I = 5
   Case "15D": I = 6
   Case "48N": I = 7
   Case Else: I = 0
  End Select
  If I > 0 Then
   C(I) = C(I) + 1
   If Val(ValorTexto(R!Salida)) = 1 Then C(I + 3) = C(I + 3) + 1
  End If
  R.MoveNext
 Loop
 CerrarRecordset R
 ConGui.Caption = CStr(GridRP.Rows - 1)
 Label8.Caption = CStr(C(1)): Label6.Caption = CStr(C(2)): Label12.Caption = CStr(C(3)): Label10.Caption = CStr(C(4))
 Label16.Caption = CStr(C(5)): Label14.Caption = CStr(C(6)): Label20.Caption = CStr(C(7))
 Label18.Caption = CStr(C(8)): Label24.Caption = CStr(C(9)): Label22.Caption = CStr(C(10))
 Set R = ConsultaOperacion("SELECT ISNULL(MAX(Numero),0) Numero FROM dbo.CartillasRP")
 Label5.Caption = CStr(R!Numero)
 CerrarRecordset R
 StatusBar1.Panels(1).Text = "Consulta"
 StatusBar1.Panels(2).Text = CStr(GridRP.Rows - 1) & " guías en la pizarra seleccionada"
 RefrescarSeleccion
 Form_Resize
 Exit Sub
Fallo:
 CerrarRecordset R
 StatusBar1.Panels(1).Text = "Estadísticas: " & Err.Description
End Sub
Private Sub CmdDetalle_Click()
 ConsultarFicha
End Sub
Private Sub CmdExpand_Click()
 If Editando Then Cmd_Click 11 Else FraMod.Visible = False
 CmdExpand.Visible = False: GridRP.Height = 6240
End Sub
Private Sub CmdActualizar_Click()
 If Not Editando Then Cmd_Click 8
End Sub
Private Sub CmdMarcar_Click()
 Dim I As Long
 For I = 1 To GridRP.Rows - 1
  GridRP.TextMatrix(I, 0) = "X"
 Next I
End Sub
Private Sub CmdImportacion_Click()
 If Not Editando Then AbrirModulo 21
End Sub
Private Sub CmdGrupos_Click()
 Dim Anterior As Long, I As Long
 On Error GoTo Fallo
 If Editando Then Exit Sub
 Anterior = -1
 If CmbGrupo.ListIndex >= 0 Then Anterior = CmbGrupo.ItemData(CmbGrupo.ListIndex)
 FrmGruposRP.Show vbModal, MdiPri
 CargarOpciones
 Cargando = True
 For I = 0 To CmbGrupo.ListCount - 1
  If CmbGrupo.ItemData(I) = Anterior Then CmbGrupo.ListIndex = I: Exit For
 Next I
 Cargando = False
 Actualizar
 Exit Sub
Fallo:
 Cargando = False
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Form_Resize()
 Dim Ancho As Single, Alto As Single
 On Error GoTo Salir
 If Me.WindowState = vbMinimized Then Exit Sub
 Ancho = ScaleWidth: Alto = ScaleHeight - StatusBar1.Height
 If Ancho < 14000 Or Alto < 7400 Then Exit Sub
 FraHer.Height = Alto
 FraEsta.Height = Alto - FraEsta.Top - 150
 FraFil.Width = Ancho - FraFil.Left - 120
 TxtBuscar.Left = FraFil.Width - TxtBuscar.Width - 150
 LblBusqueda.Left = TxtBuscar.Left - LblBusqueda.Width - 150
 GridRP.Width = Ancho - GridRP.Left - 120
 If Not FraMod.Visible Then GridRP.Height = Alto - GridRP.Top - 1600
 Coment.Top = Alto - 1455
 Coment.Width = (GridRP.Width - 120) / 2
 InforGui.Top = Coment.Top
 InforGui.Left = Coment.Left + Coment.Width + 120
 InforGui.Width = Coment.Width
 TxtComent.Top = Alto - 375: TxtComent.Width = Coment.Width
 TxtInforGui.Top = TxtComent.Top: TxtInforGui.Left = InforGui.Left: TxtInforGui.Width = InforGui.Width
 FraMod.Top = Coment.Top - FraMod.Height - 150
 CmdExpand.Top = FraMod.Top - 375
 CmdExpand.Width = GridRP.Width
Salir:
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQueries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador de Solicitudes Queries"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   Icon            =   "FrmQueries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   12870
   Begin MSFlexGridLib.MSFlexGrid GridHis 
      Height          =   1455
      Left            =   2610
      TabIndex        =   20
      Top             =   495
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   2566
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
      Appearance      =   0
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
   Begin Clever.Boton CmdExpanHis 
      ToolTipText = "Alternar solicitudes anteriores / guías seleccionadas"
      Height          =   1455
      Left            =   5625
      TabIndex        =   22
      Top             =   495
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   2566
      BTYPE           =   14
      TX              =   "."
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmQueries.frx":1782
      PICN            =   "FrmQueries.frx":179E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   55
      Top             =   7335
      Width           =   12870
      _ExtentX        =   22701
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
   Begin VB.Frame Frame6 
      Caption         =   "Guias Ingresadas"
      Height = 1455
      Left = 2610
      TabIndex        =   62
      Top = 495
      Visible = 0
      Width = 2955
      Begin VB.ListBox LstGui 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height = 1095
         ItemData        =   "FrmQueries.frx":1D38
         Left            =   120
         List            =   "FrmQueries.frx":1D3A
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width = 2700
      End
   End
   Begin VB.ComboBox CmbFormato 
      Height          =   315
      ItemData        =   "FrmQueries.frx":1D3C
      Left            =   6030
      List            =   "FrmQueries.frx":1D3E
      Style = 2
      TabIndex        =   5
      Top             =   495
      Width           =   6720
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      ItemData        =   "FrmQueries.frx":1D40
      Left            =   8730
      List            =   "FrmQueries.frx":1D42
      Style = 2
      TabIndex        =   23
      Top             =   180
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   6030
      TabIndex        =   6
      Top             =   855
      Width           =   6720
      Begin Clever.Boton CmdBusCod 
         Visible = 0
         Height          =   285
         Left            =   6300
         TabIndex        =   19
         Top             =   225
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "."
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQueries.frx":1D44
         PICN            =   "FrmQueries.frx":1D60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtCod 
         Locked = -1
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5310
         TabIndex        =   17
         Top             =   225
         Width           =   930
      End
      Begin VB.TextBox RegMod 
         Locked = -1
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   990
         TabIndex        =   11
         Top             =   630
         Width           =   495
      End
      Begin VB.TextBox RegMont 
         Locked = -1
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2250
         TabIndex        =   10
         Top             =   630
         Width           =   915
      End
      Begin VB.TextBox RegDiv 
         Locked = -1
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3870
         TabIndex        =   9
         Top             =   630
         Width           =   615
      End
      Begin VB.TextBox RegFec 
         Locked = -1
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5310
         TabIndex        =   8
         Top             =   630
         Width           =   1335
      End
      Begin VB.TextBox RegCli 
         Locked = -1
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   990
         TabIndex        =   7
         Top             =   270
         Width           =   3495
      End
      Begin VB.Label LblDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   65
         Top             =   135
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Codigo:"
         Height          =   240
         Left            =   4590
         TabIndex        =   18
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Emision:"
         Height          =   240
         Left            =   4590
         TabIndex        =   16
         Top             =   675
         Width           =   600
      End
      Begin VB.Label Label6 
         Caption         =   "Divisa:"
         Height          =   240
         Left            =   3240
         TabIndex        =   15
         Top             =   675
         Width           =   600
      End
      Begin VB.Label Label5 
         Caption         =   "Monto:"
         Height          =   240
         Left            =   1575
         TabIndex        =   14
         Top             =   675
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Modal:"
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   675
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente:"
         Height          =   240
         Left            =   90
         TabIndex        =   12
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.TextBox RegDoc 
      Enabled = -1
      Appearance      =   0  'Flat
      BackColor = &H80000005&
      Height          =   285
      Left            =   10935
      MaxLength = 14
      TabIndex        =   3
      Top             =   180
      Width           =   1815
   End
   Begin VB.Frame FraHer 
      BackColor       =   &H0053555F&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   2310
      Begin Clever.Boton CmdInfoDoc 
         Height          =   375
         Left            =   105
         TabIndex        =   50
         Top             =   405
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Informacion de Documento"
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
         BCOL            =   8421504
         BCOLO           =   12632256
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQueries.frx":22FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdGen 
         Height          =   375
         Left            =   105
         TabIndex        =   1
         Top             =   1350
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   14
         TX = "Guardar y ver solicitud"
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
         BCOL            =   8421504
         BCOLO           =   12632256
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQueries.frx":2316
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdConfCas 
         Height          =   375
         Left            =   105
         TabIndex        =   67
         Top             =   810
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Configurar Casos Otros"
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
         BCOL            =   8421504
         BCOLO           =   12632256
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQueries.frx":2332
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblReg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Querie Nº "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   66
         Top             =   6750
         Width           =   2130
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Querie"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   61
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Visto de Aprobacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   60
         Top             =   3825
         Width           =   1815
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio del Problema:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   135
         TabIndex        =   59
         Top             =   3015
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Herramientas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   240
         Left            =   90
         TabIndex        =   2
         Top             =   135
         Width           =   2130
      End
   Begin VB.TextBox TxtFechaInicio
      Left=135
      Top=3240
      Width=2070
      Height=315
      MaxLength=10
   End
   Begin VB.TextBox TxtFechaVisto
      Left=135
      Top=4080
      Width=2070
      Height=315
      MaxLength=10
   End
   Begin VB.TextBox TxtFechaSolicitud
      Left=135
      Top=4920
      Width=2070
      Height=315
      MaxLength=10
   End
      Begin Clever.Boton CmdNueva 
         Height=375
         Left=105
         TabIndex        =   1
         Top=1830
         Width=2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   14
         TX="Nueva solicitud"
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
         BCOL            =   8421504
         BCOLO           =   12632256
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQueries.frx":2316
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
   Begin VB.Frame FraConfCas 
      Height          =   4560
      Left            =   6840
      TabIndex        =   91
      Top             =   2115
      Visible = 0
      Width           =   5775
      Begin VB.ListBox LstEliCas 
         Height          =   3180
         ItemData        =   "FrmQueries.frx":234E
         Left            =   1665
         List            =   "FrmQueries.frx":2350
         TabIndex        =   95
         Top             =   1125
         Width           =   3885
      End
      Begin VB.TextBox TxtNewCas 
         Height          =   285
         Left            =   1665
         TabIndex        =   94
         Top             =   765
         Width           =   2985
      End
      Begin VB.OptionButton OptAgr 
         Caption         =   "Agregar Caso"
         Height          =   285
         Left            =   225
         TabIndex        =   93
         Top             =   765
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton OptEli 
         Caption         =   "Eliminar Caso"
         Height          =   285
         Left            =   225
         TabIndex        =   92
         Top             =   1125
         Width           =   1275
      End
      Begin Clever.Boton CmdEliCas 
         Height          =   690
         Left            =   225
         TabIndex        =   96
         Top             =   1440
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1217
         BTYPE           =   14
         TX              =   "Eliminar Seleccionado"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQueries.frx":2352
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdNewCas 
         Height          =   285
         Left            =   4725
         TabIndex        =   97
         Top             =   765
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "Agregar"
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQueries.frx":236E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdCerCas 
         Height          =   375
         Left            =   225
         TabIndex        =   98
         Top             =   3915
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Cerrar"
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQueries.frx":238A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Administrar Casos Otros"
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
         Height          =   240
         Left            =   180
         TabIndex        =   99
         Top             =   270
         Width           =   5370
      End
   End
   Begin VB.Frame FraOpt 
      Enabled = -1
      Height          =   5235
      Left            =   2610
      TabIndex        =   25
      Top             =   1980
      Width           =   10140
      Begin VB.ComboBox CmbDet 
         Height          =   315
         Left            =   135
         Style = 2
         TabIndex        =   64
         Top             =   225
         Width           =   3975
      End
      Begin VB.TextBox TxtAdi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   135
         MaxLength = 255
         TabIndex        =   53
         Top             =   4815
         Width           =   9375
      End
      Begin VB.Frame Frame4 
         Height          =   2040
         Left            =   135
         TabIndex        =   27
         Top             =   2655
         Width           =   3975
         Begin VB.ComboBox CmbServ 
            Height          =   315
            Left            =   1215
            Style = 2
            TabIndex        =   46
            Top             =   630
            Width           =   2535
         End
         Begin VB.TextBox TipServ 
            Locked = -1
            Enabled         =   0   'False
            Height          =   315
            Left            =   2655
            TabIndex        =   45
            Text = "Importaciones"
            Top             =   1470
            Width           =   1095
         End
         Begin VB.ComboBox CmbResp 
            Height          =   315
            Left            =   1215
            Style = 2
            TabIndex        =   44
            Top             =   270
            Width           =   2535
         End
         Begin VB.TextBox TxtTerr 
            MaxLength = 150
            Height          =   315
            Left            =   1215
            TabIndex        =   43
            Top             =   990
            Width           =   1935
         End
         Begin VB.TextBox TxtRc 
            MaxLength = 8
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1215
            TabIndex        =   42
            Top             =   1470
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Serv:"
            Height          =   240
            Left            =   2025
            TabIndex        =   52
            Top             =   1485
            Width           =   600
         End
         Begin VB.Label Label18 
            Caption         =   "Recargo"
            Height          =   240
            Left            =   225
            TabIndex        =   51
            Top             =   1485
            Width           =   1005
         End
         Begin VB.Label Label17 
            Caption         =   "Territorio"
            Height          =   240
            Left            =   225
            TabIndex        =   49
            Top             =   1035
            Width           =   1005
         End
         Begin VB.Label Label16 
            Caption         =   "Servicio"
            Height          =   240
            Left            =   225
            TabIndex        =   48
            Top             =   675
            Width           =   1005
         End
         Begin VB.Label Label15 
            Caption         =   "Responsable"
            Height          =   240
            Left            =   225
            TabIndex        =   47
            Top             =   315
            Width           =   1005
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2130
         Left            =   135
         TabIndex        =   26
         Top             =   495
         Width           =   3975
         Begin VB.TextBox TxtContGui 
            Locked = -1
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2835
            TabIndex        =   40
            Top             =   1665
            Width           =   525
         End
         Begin VB.TextBox TxtNGuia 
            MaxLength = 10
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   38
            Top             =   1665
            Width           =   1695
         End
         Begin VB.OptionButton OptLibGuiSi 
            Appearance      =   0  'Flat
            Caption         =   "SI  Liberar la guia:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2145
            TabIndex        =   31
            Top             =   225
            Width           =   1695
         End
         Begin VB.OptionButton OptLibGuiNo 
            Appearance      =   0  'Flat
            Caption         =   "No Liberar la guia:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2145
            TabIndex        =   30
            Top             =   465
            Width           =   1695
         End
         Begin VB.CheckBox ChkAnlDir 
            Appearance      =   0  'Flat
            Caption         =   "Anular Directo"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   225
            TabIndex        =   29
            Top             =   465
            Width           =   1695
         End
         Begin VB.CheckBox ChkAnlNc 
            Appearance      =   0  'Flat
            Caption         =   "Anulación con NC"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   225
            TabIndex        =   28
            Top             =   225
            Width           =   1695
         End
         Begin Clever.Boton Boton3 
            Height          =   285
            Left            =   3375
            TabIndex        =   41
            Top             =   1665
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            BTYPE           =   14
            TX              =   "."
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
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmQueries.frx":23BD
            PICN            =   "FrmQueries.frx":23D9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Frame FraDes 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   180
            TabIndex        =   32
            Top             =   675
            Visible         =   0   'False
            Width           =   3615
            Begin VB.TextBox DesMont 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   810
               TabIndex        =   34
               Top             =   120
               Width           =   1695
            End
            Begin VB.ComboBox CmbDiv 
               Style = 2
               Height          =   315
               Left            =   2640
               TabIndex        =   33
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label11 
               Caption         =   "Monto:"
               Height          =   240
               Left            =   90
               TabIndex        =   35
               Top             =   135
               Width           =   645
            End
         End
         Begin VB.Label Label14 
            Caption         =   "Guia:"
            Height          =   240
            Left            =   270
            TabIndex        =   39
            Top             =   1680
            Width           =   645
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption = "La liberación es una instrucción al sistema origen."
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
            Height          =   240
            Left            =   270
            TabIndex        =   37
            Top             =   1395
            Width = 3690
         End
         Begin VB.Label Label12 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption = "Hasta 6 guías. Enter: agregar. Supr: retirar."
            Height          =   195
            Left            =   270
            TabIndex        =   36
            Top             =   1170
            Width = 3690
         End
      End
      Begin VB.Frame FraCasos 
         Height          =   4560
         Left            =   4230
         TabIndex        =   68
         Top             =   135
         Width           =   5775
         Begin VB.Frame FraOtr 
            BorderStyle     =   0  'None
            Height          =   3360
            Left            =   180
            TabIndex        =   70
            Top             =   1035
            Visible         =   0   'False
            Width           =   5415
            Begin VB.CheckBox ChkOtr 
               Appearance      =   0  'Flat
               Caption         =   "Descuento Especial"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   240
               Width           =   3255
            End
            Begin VB.CheckBox ChkCasOtr 
               Appearance      =   0  'Flat
               Caption         =   "Otros Varios"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   720
               Width           =   1575
            End
            Begin VB.ListBox LstOtr 
               BackColor       =   &H00C0C0C0&
               Enabled         =   0   'False
               Height          =   2010
               ItemData        =   "FrmQueries.frx":2973
               Left            =   135
               List            =   "FrmQueries.frx":297A
               TabIndex        =   71
               Top             =   1260
               Width           =   5115
            End
         End
         Begin VB.ComboBox CmbCas 
            Height          =   315
            ItemData        =   "FrmQueries.frx":2986
            Left            =   180
            List            =   "FrmQueries.frx":2988
            Style = 2
            TabIndex        =   69
            Top             =   585
            Width           =   5415
         End
         Begin MSComctlLib.TreeView ViewCas 
            Height          =   3330
            Left            =   180
            TabIndex        =   74
            Top             =   1035
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   5874
            _Version        =   393217
            Indentation     =   90
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Frame FraNone 
            Height          =   3435
            Left            =   180
            TabIndex        =   88
            Top             =   945
            Width           =   5415
            Begin VB.Label Label29 
               Alignment       =   2  'Center
               Caption         =   "Opciones no Disponibles sin especificar Documento"
               Height          =   375
               Left            =   540
               TabIndex        =   89
               Top             =   1530
               Width           =   4245
            End
         End
         Begin VB.Frame FraDatCli 
            BorderStyle     =   0  'None
            Height          =   3360
            Left            =   180
            TabIndex        =   75
            Top             =   1035
            Visible         =   0   'False
            Width           =   5415
            Begin VB.CheckBox chkDat 
               Appearance      =   0  'Flat
               Caption         =   "Mal ingreso de Datos"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Width           =   3975
            End
            Begin VB.CheckBox chkDat 
               Appearance      =   0  'Flat
               Caption         =   "Otra razón social o persona del mismo grupo"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   82
               Top             =   480
               Width           =   3495
            End
            Begin VB.CheckBox chkDat 
               Appearance      =   0  'Flat
               Caption         =   "Cambio de Razón Social"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   81
               Top             =   720
               Width           =   3375
            End
            Begin VB.CheckBox ChkRef 
               Appearance      =   0  'Flat
               Caption         =   "Cambio General de Cliente"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   80
               Top             =   1200
               Width           =   3255
            End
            Begin VB.TextBox TXTDATCLI 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               MaxLength = 255
               TabIndex        =   79
               Top             =   1560
               Width           =   4215
            End
            Begin VB.TextBox TXTDATRUC 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               MaxLength = 11
               TabIndex        =   78
               Top             =   1920
               Width           =   1695
            End
            Begin VB.TextBox TXTDATDIR 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               MaxLength = 255
               TabIndex        =   77
               Top             =   2280
               Width           =   4215
            End
            Begin VB.TextBox TXTDATTEL 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3600
               MaxLength = 20
               TabIndex        =   76
               Top             =   1920
               Width           =   1695
            End
            Begin VB.Label Label21 
               Caption         =   "Facturar a:"
               Height          =   240
               Left            =   135
               TabIndex        =   87
               Top             =   1620
               Width           =   915
            End
            Begin VB.Label Label22 
               Caption         =   "Ruc:"
               Height          =   240
               Left            =   135
               TabIndex        =   86
               Top             =   1980
               Width           =   915
            End
            Begin VB.Label Label23 
               Caption         =   "Dirección:"
               Height          =   240
               Left            =   135
               TabIndex        =   85
               Top             =   2295
               Width           =   915
            End
            Begin VB.Label Label24 
               Caption         =   "Telefono:"
               Height          =   240
               Left            =   2835
               TabIndex        =   84
               Top             =   1935
               Width           =   735
            End
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Razon Final de Errores en El Querie"
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
            Height          =   240
            Left            =   180
            TabIndex        =   90
            Top             =   270
            Width           =   5370
         End
      End
      Begin VB.Label LBLN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9540
         TabIndex        =   54
         Top             =   4815
         Width           =   465
      End
   End
   Begin VB.Frame FraVisDoc 
      Visible = 0
      Height          =   5235
      Left            =   2610
      TabIndex        =   100
      Top             =   1980
      Width           =   10140
      Begin VB.Frame Frame5 
         Height          =   4425
         Left            =   225
         TabIndex        =   101
         Top             =   585
         Width           =   9690
         Begin Clever.Boton CmdAceptar 
            Height          =   330
            Left            =   7650
            TabIndex        =   140
            Top             =   3870
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            BTYPE           =   14
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmQueries.frx":298A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Servicio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1935
            TabIndex        =   139
            Top             =   2610
            Width           =   795
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   270
            TabIndex        =   138
            Top             =   2610
            Width           =   435
         End
         Begin VB.Label LblSituacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bueno"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4500
            TabIndex        =   132
            Top             =   1890
            Width           =   3480
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Situacion:"
            Height          =   195
            Left            =   3240
            TabIndex        =   131
            Top             =   1890
            Width           =   705
         End
         Begin VB.Label LblDescripcion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bueno"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4500
            TabIndex        =   130
            Top             =   1620
            Width           =   3480
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   3240
            TabIndex        =   129
            Top             =   1620
            Width           =   885
         End
         Begin VB.Label LblDistrito 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "MIRAFLORES"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4500
            TabIndex        =   128
            Top             =   1350
            Width           =   3480
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Distrito:"
            Height          =   195
            Left            =   3240
            TabIndex        =   127
            Top             =   1350
            Width           =   525
         End
         Begin VB.Label LblDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SAGA FALABELLA S.A."
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4500
            TabIndex        =   126
            Top             =   1080
            Width           =   5055
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Direccion:"
            Height          =   195
            Left            =   3240
            TabIndex        =   125
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label LblRuc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "607152465"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4500
            TabIndex        =   124
            Top             =   810
            Width           =   1995
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Ruc / Dni:"
            Height          =   195
            Left            =   3240
            TabIndex        =   123
            Top             =   810
            Width           =   750
         End
         Begin VB.Label LblCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SAGA FALABELLA S.A."
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4500
            TabIndex        =   122
            Top             =   540
            Width           =   5055
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Cliente Asignado:"
            Height          =   195
            Left            =   3240
            TabIndex        =   121
            Top             =   540
            Width           =   1230
         End
         Begin VB.Label LblPorcentaje 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "12.00 %"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2385
            TabIndex        =   120
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label LblTotal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "S/.353.06"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1215
            TabIndex        =   119
            Top             =   1890
            Width           =   1140
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   180
            TabIndex        =   118
            Top             =   1890
            Width           =   405
         End
         Begin VB.Label LblGastos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "S/.0.00"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1215
            TabIndex        =   117
            Top             =   1620
            Width           =   1140
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Gastos:"
            Height          =   195
            Left            =   180
            TabIndex        =   116
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblIgv 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "S/.56.37"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1215
            TabIndex        =   115
            Top             =   1350
            Width           =   1140
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "IGV:"
            Height          =   195
            Left            =   180
            TabIndex        =   114
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label LblRecargo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "S/.28.64"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1215
            TabIndex        =   113
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Recargo:"
            Height          =   195
            Left            =   180
            TabIndex        =   112
            Top             =   1080
            Width           =   660
         End
         Begin VB.Label LblSubtotal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "S/.268.05"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1215
            TabIndex        =   111
            Top             =   810
            Width           =   1140
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal:"
            Height          =   195
            Left            =   180
            TabIndex        =   110
            Top             =   810
            Width           =   690
         End
         Begin VB.Label LblNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SAGA FALABELLA S.A."
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4500
            TabIndex        =   109
            Top             =   270
            Width           =   5055
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Cliente Factura:"
            Height          =   195
            Left            =   3240
            TabIndex        =   108
            Top             =   270
            Width           =   1110
         End
         Begin VB.Label LblVencimiento 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "05/05/1985"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1215
            TabIndex        =   107
            Top             =   540
            Width           =   1140
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento:"
            Height          =   195
            Left            =   180
            TabIndex        =   106
            Top             =   540
            Width           =   915
         End
         Begin VB.Label lblEmision 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1215
            TabIndex        =   105
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Emision:"
            Height          =   195
            Left            =   180
            TabIndex        =   104
            Top             =   270
            Width           =   585
         End
   Begin MSFlexGridLib.MSFlexGrid GridServicios
      Left=180
      Top=2925
      Width=7185
      Height=1215
      FixedCols=0
   End
      End
      Begin VB.Label LblDocumento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FT-0030181125"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1260
         TabIndex        =   103
         Top             =   315
         Width           =   1680
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   225
         TabIndex        =   102
         Top             =   315
         Width           =   870
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Tipo:"
      Height          =   240
      Left            =   8280
      TabIndex        =   24
      Top             =   225
      Width           =   420
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption = "Solicitudes anteriores / guías"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2610
      TabIndex        =   21
      Top             =   180
      Width           =   3300
   End
   Begin VB.Label Label2 
      Caption         =   "Documento"
      Height          =   240
      Left            =   9990
      TabIndex        =   4
      Top             =   225
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0053555F&
      BackStyle       =   1  'Opaque
      Height          =   8430
      Left            =   0
      Top             =   0
      Width           =   2490
   End
End
Attribute VB_Name = "FrmQueries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Documento As ADODB.Recordset, Servicios As ADODB.Recordset, Otros As ADODB.Recordset
Private Clave As String, IdGuardado As Long, TipoCargado As String, NumeroCargado As String
Private Iniciando As Boolean, Seleccionando As Boolean

Private Sub Form_Load()
 Dim C As Variant
 On Error GoTo Fallo
 If Not AccesoPermitido(5) Then Unload Me: Exit Sub
 StatusBar1.Style = sbrSimple
 Iniciando = True
 CmbTipo.Clear: CmbTipo.AddItem "FT": CmbTipo.AddItem "BV": CmbTipo.ListIndex = 0
 CmbFormato.Clear: CmbFormato.AddItem "Solicitud de corrección de documento": CmbFormato.AddItem "Solicitud de corrección - modelo alternativo": CmbFormato.ListIndex = 0
 CmbDet.Clear: CmbDet.AddItem "Anulación": CmbDet.AddItem "Descuento parcial": CmbDet.AddItem "Descuento al 100%": CmbDet.ListIndex = 0
 CmbCas.Clear
 For Each C In Array("Error en Datos del Cliente", "Tarifas Erradas", "Mal Servicio", "Condiciones de Facturacion", "Errores Otros - Especificar...")
  CmbCas.AddItem CStr(C)
 Next C
 CmbCas.ListIndex = 0
 ComboCatalogoQuery CmbResp, 1: ComboCatalogoQuery CmbServ, 2
 CargarOtros
 OptAgr.Value = True
 ConfigurarAccionCaso
 GrillaQuery GridHis, "Solicitud|Fecha|Estado": GridHis.ColWidth(0) = 1300: GridHis.ColWidth(1) = 1100: GridHis.ColWidth(2) = 1400
 GrillaQuery GridServicios, "Guía|Servicio": GridServicios.ColWidth(0) = 1700: GridServicios.ColWidth(1) = 5100
 ViewCas.Checkboxes = True
 RegDoc.Enabled = True: RegDoc.BackColor = vbWhite
 FraConfCas.Visible = False: FraVisDoc.Visible = False
 Iniciando = False
 NuevaSolicitud
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
 Unload Me
End Sub
Private Sub NuevaSolicitud()
 Dim I As Integer
 CerrarRecordset Documento: CerrarRecordset Servicios
 IdGuardado = 0: Clave = NuevaClaveQuery: NumeroCargado = "": TipoCargado = ""
 Iniciando = True
 RegDoc.Text = "": RegCli.Text = "": RegMod.Text = "": RegMont.Text = "": RegDiv.Text = "": RegFec.Text = "": txtCod.Text = ""
 TxtFechaInicio.Text = Format$(Date, "dd/mm/yyyy"): TxtFechaVisto.Text = TxtFechaInicio.Text: TxtFechaSolicitud.Text = TxtFechaInicio.Text
 TxtFechaInicio.Enabled = True: TxtFechaVisto.Enabled = True: TxtFechaSolicitud.Enabled = True
 TxtAdi.Text = "": TxtRc.Text = "0": TxtTerr.Text = "": TipServ.Text = "Importaciones"
 TxtNGuia.Text = "": LstGui.Clear: TxtContGui.Text = "0": DesMont.Text = "0": CmbDiv.Clear
 CmbDet.ListIndex = 0: CmbCas.ListIndex = 0: ChkAnlNc.Value = 1: ChkAnlDir.Value = 0: OptLibGuiNo.Value = True
 For I = 0 To 2: chkDat(I).Value = 0: Next I
 ChkRef.Value = 0: ChkOtr.Value = 0: ChkCasOtr.Value = 0: LstOtr.ListIndex = -1
 TXTDATCLI.Text = "": TXTDATRUC.Text = "": TXTDATDIR.Text = "": TXTDATTEL.Text = ""
 LblReg.Caption = "Nueva solicitud": CmdGen.Enabled = False
 GridHis.Rows = 1: Frame6.Visible = False: FraVisDoc.Visible = False
 FraOpt.Enabled = False: FraNone.Visible = True: FraNone.ZOrder 0
 Iniciando = False
 ActualizarTipo
 MostrarCaso
 StatusBar1.SimpleText = "Ingrese tipo y número completo del documento y presione Enter."
End Sub
Private Sub CmdNueva_Click()
 On Error GoTo Fallo
 If IdGuardado = 0 And Not Documento Is Nothing Then
  If MsgBox("¿Descartar los datos de la solicitud actual?", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 End If
 NuevaSolicitud
 RegDoc.SetFocus
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub RegDoc_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyF9 Then CargarDocumento: KeyCode = 0
End Sub
Private Sub RegDoc_Change()
 If Iniciando Then Exit Sub
 CmdGen.Enabled = False
End Sub
Private Sub CmbTipo_Click()
 If Not Iniciando Then CmdGen.Enabled = False
End Sub
Private Sub CargarDocumento()
 Dim D As ADODB.Recordset
 On Error GoTo Fallo
 If Len(Trim$(RegDoc.Text)) = 0 Then Exit Sub
 If IdGuardado > 0 Then Err.Raise 5, , "Use Nueva solicitud para registrar otro caso."
 Set D = DocumentoQuery(CmbTipo.Text, Trim$(RegDoc.Text))
 If D.EOF Then Err.Raise 5, , "No se encontró el documento importado. Revise tipo y número completo."
 If CBool(D!Anulado) Then Err.Raise 5, , "El documento está anulado en el sistema origen."
 If CCur(D!Importe) <= 0 Then Err.Raise 5, , "El documento no tiene importe válido."
 CerrarRecordset Documento: Set Documento = D
 CerrarRecordset Servicios: Set Servicios = ServiciosQuery(CLng(D!IdDocumento))
 TipoCargado = CmbTipo.Text: NumeroCargado = CStr(D!Documento)
 RegDoc.Text = NumeroCargado
 RegCli.Text = CStr(D!Cliente): txtCod.Text = CStr(D!CodigoCliente)
 RegMont.Text = Format$(D!Importe, "#,##0.00"): RegDiv.Text = CStr(D!Divisa): RegFec.Text = FechaDocumento(D!Emision)
 RegMod.Text = "IMP": TxtTerr.Text = ValorTexto(D!Territorio): TxtRc.Text = "0"
 CmbDiv.Clear: CmbDiv.AddItem CStr(D!Divisa): CmbDiv.ListIndex = 0: CmbDiv.Enabled = False
 DesMont.Text = Format$(D!Importe, "0.00")
 TXTDATCLI.Text = CStr(D!Cliente): TXTDATRUC.Text = ValorTexto(D!RUCDNI): TXTDATDIR.Text = ValorTexto(D!DIRECC): TXTDATTEL.Text = ""
 LstGui.Clear: TxtContGui.Text = "0": Label13.Visible = Servicios.EOF
 If Servicios.EOF Then Label13.Caption = "Documento sin guías vinculadas."
 FraOpt.Enabled = True: CmdGen.Enabled = True: FraNone.Visible = False
 CargarHistoria
 MostrarCaso
 StatusBar1.SimpleText = "Documento cargado. Los cambios solicitados se registran como instrucciones."
 Exit Sub
Fallo:
 CerrarRecordset D
 CmdGen.Enabled = False: FraOpt.Enabled = False
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CargarHistoria()
 Dim R As ADODB.Recordset
 Set R = ListaQueries(5, NumeroCargado)
 GridHis.Rows = 1
 Do While Not R.EOF
  If CStr(R!NumDoc) = NumeroCargado And CStr(R!TipoDocumento) = TipoCargado Then
   GridHis.AddItem CStr(R!Asient) & vbTab & FechaDocumento(R!FecSol) & vbTab & IIf(CStr(R!SitDoc) = "ANL", "Anulada", CStr(R!Estado))
   GridHis.RowData(GridHis.Rows - 1) = CLng(R!Item)
  End If
  R.MoveNext
 Loop
 CerrarRecordset R
 If GridHis.Rows > 1 Then GridHis.Row = 1: SeleccionFilaQuery GridHis, Seleccionando
End Sub
Private Sub CmbDet_Click()
 If Not Iniciando Then ActualizarTipo
End Sub
Private Sub ActualizarTipo()
 ChkAnlNc.Enabled = (CmbDet.ListIndex = 0): ChkAnlDir.Enabled = ChkAnlNc.Enabled
 FraDes.Visible = (CmbDet.ListIndex > 0)
 DesMont.Locked = (CmbDet.ListIndex <> 1)
 If Not Documento Is Nothing Then
  If CmbDet.ListIndex <> 1 Then DesMont.Text = Format$(Documento!Importe, "0.00")
 End If
End Sub
Private Sub ChkAnlNc_Click()
 If ChkAnlNc.Value = 1 Then ChkAnlDir.Value = 0
End Sub
Private Sub ChkAnlDir_Click()
 If ChkAnlDir.Value = 1 Then ChkAnlNc.Value = 0
End Sub
Private Sub CmbCas_Click()
 If Not Iniciando Then MostrarCaso
End Sub
Private Sub MostrarCaso()
 Dim A As Variant, I As Integer, Prefijo As String
 FraDatCli.Visible = False: FraOtr.Visible = False: ViewCas.Visible = False
 ViewCas.Nodes.Clear
 If Documento Is Nothing Then FraNone.Visible = True: FraNone.ZOrder 0: Exit Sub
 FraNone.Visible = False
 Select Case CmbCas.ListIndex
  Case 0
   FraDatCli.Visible = True: FraDatCli.ZOrder 0
  Case 1
   Prefijo = "ChkTar": A = Array("Descuentos no incluidos", "Error tarifario", "Zona", "Peso físico / volumen", "Producto")
  Case 2
   Prefijo = "ChkServ": A = Array("No salió", "No llegó", "Mal estado", "Fuera de tiempo", "Collect")
  Case 3
   Prefijo = "ChkCon": A = Array("Por guía", "Por área", "Por producto")
  Case 4
   FraOtr.Visible = True: FraOtr.ZOrder 0
 End Select
 If Len(Prefijo) > 0 Then
  For I = 0 To UBound(A)
   ViewCas.Nodes.Add , , Prefijo & CStr(I + 1), CStr(A(I))
  Next I
  ViewCas.Visible = True: ViewCas.ZOrder 0
 End If
End Sub
Private Sub TxtNGuia_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then AgregarGuia: KeyCode = 0
End Sub
Private Sub Boton3_Click()
 AgregarGuia
End Sub
Private Sub AgregarGuia()
 Dim I As Integer, Encontrada As Boolean, Numero As String
 On Error GoTo Fallo
 If Documento Is Nothing Or IdGuardado > 0 Then Exit Sub
 Numero = Trim$(TxtNGuia.Text)
 If Len(Numero) = 0 Then Frame6.Visible = Not Frame6.Visible: Frame6.ZOrder 0: Exit Sub
 If LstGui.ListCount >= 6 Then Err.Raise 5, , "Se permiten hasta seis guías por solicitud."
 For I = 0 To LstGui.ListCount - 1
  If LstGui.List(I) = Numero Then Err.Raise 5, , "La guía ya fue agregada."
 Next I
 If Servicios.RecordCount > 0 Then Servicios.MoveFirst
 Do While Not Servicios.EOF
  If CStr(Servicios!Guia) = Numero Then Encontrada = True: Exit Do
  Servicios.MoveNext
 Loop
 If Not Encontrada Then Err.Raise 5, , "La guía no pertenece al documento cargado."
 LstGui.AddItem Numero: TxtContGui.Text = CStr(LstGui.ListCount): TxtNGuia.Text = ""
 Frame6.Visible = True: Frame6.ZOrder 0
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub LstGui_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete And IdGuardado = 0 And LstGui.ListIndex >= 0 Then
  LstGui.RemoveItem LstGui.ListIndex: TxtContGui.Text = CStr(LstGui.ListCount): KeyCode = 0
 End If
End Sub
Private Sub CmdExpanHis_Click()
 Frame6.Visible = Not Frame6.Visible
 If Frame6.Visible Then Frame6.ZOrder 0
End Sub
Private Sub GridHis_Click()
 SeleccionFilaQuery GridHis, Seleccionando
End Sub
Private Sub GridHis_RowColChange()
 SeleccionFilaQuery GridHis, Seleccionando
End Sub
Private Sub GridHis_DblClick()
 If GridHis.Row > 0 Then VerQuery GridHis.RowData(GridHis.Row), 5
End Sub
Private Sub GridServicios_Click()
 SeleccionFilaQuery GridServicios, Seleccionando
End Sub
Private Sub GridServicios_RowColChange()
 SeleccionFilaQuery GridServicios, Seleccionando
End Sub
Private Sub TxtAdi_Change()
 LBLN.Caption = CStr(Len(TxtAdi.Text))
End Sub
Private Function SolicitudXML() As String
 Dim S As String, Clase As String, F As Date, Inicio As Date, Visto As Date, N As Variant, I As Integer, Activo As Boolean, Otro As Long
 If Documento Is Nothing Then Err.Raise 5, , "Cargue primero un documento."
 If TipoCargado <> CmbTipo.Text Or NumeroCargado <> Trim$(RegDoc.Text) Then Err.Raise 5, , "El número cambió. Presione Enter para cargar el documento."
 If Not FechaGestionValida(TxtFechaSolicitud.Text, F) Or Not FechaGestionValida(TxtFechaInicio.Text, Inicio) Or Not FechaGestionValida(TxtFechaVisto.Text, Visto) Then Err.Raise 5, , "Ingrese las tres fechas con formato dd/mm/aaaa y hasta hoy."
 If CmbResp.ListIndex < 0 Or CmbServ.ListIndex < 0 Then Err.Raise 5, , "Seleccione responsable y servicio."
 Select Case CmbDet.ListIndex
  Case 0
   If ChkAnlNc.Value = 0 And ChkAnlDir.Value = 0 Then Err.Raise 5, , "Seleccione anulación directa o con NC."
   If ChkAnlDir.Value = 1 Then Clase = "ANL ANL" Else Clase = "ANL NC"
  Case 1: Clase = "DES NC"
  Case 2: Clase = "DES 100"
 End Select
 S = "<solicitud" & AtributoXML("Documento", CStr(Documento!IdDocumento)) & AtributoXML("Version", VersionXML(Documento.Fields("VersionFila").Value)) & AtributoXML("Clase", Clase) & AtributoXML("Caso", CmbCas.Text)
 S = S & AtributoXML("Fecha", Format$(F, "yyyy-mm-dd")) & AtributoXML("Inicio", Format$(Inicio, "yyyy-mm-dd")) & AtributoXML("Visto", Format$(Visto, "yyyy-mm-dd"))
 S = S & AtributoXML("Monto", NumeroXML(DesMont.Text)) & AtributoXML("Recargo", NumeroXML(TxtRc.Text)) & AtributoXML("Territorio", Trim$(TxtTerr.Text))
 S = S & AtributoXML("Responsable", CStr(CmbResp.ItemData(CmbResp.ListIndex))) & AtributoXML("Servicio", CStr(CmbServ.ItemData(CmbServ.ListIndex))) & AtributoXML("Formato", CStr(CmbFormato.ListIndex + 1))
 S = S & AtributoXML("Adicional", Trim$(TxtAdi.Text)) & AtributoXML("SiLibGuia", IIf(OptLibGuiSi.Value, "1", "0")) & AtributoXML("NoLibGuia", IIf(OptLibGuiNo.Value, "1", "0"))
 For Each N In Array("ClientePropuesto", "RucPropuesto", "DireccionPropuesta", "TelefonoPropuesto")
  Select Case CStr(N)
   Case "ClientePropuesto": Clase = TXTDATCLI.Text
   Case "RucPropuesto": Clase = TXTDATRUC.Text
   Case "DireccionPropuesta": Clase = TXTDATDIR.Text
   Case "TelefonoPropuesto": Clase = TXTDATTEL.Text
  End Select
  If CmbCas.ListIndex <> 0 Then Clase = ""
  S = S & AtributoXML(CStr(N), Trim$(Clase))
 Next N
 For Each N In Split("ChkMal1 ChkMal2 ChkMal3 CambioCliente ChkTar1 ChkTar2 ChkTar3 ChkTar4 ChkTar5 ChkServ1 ChkServ2 ChkServ3 ChkServ4 ChkServ5 ChkCon1 ChkCon2 ChkCon3 ChkDescuento", " ")
  Activo = False
  If CmbCas.ListIndex = 0 Then
   If Left$(CStr(N), 6) = "ChkMal" Then Activo = (chkDat(CInt(Right$(CStr(N), 1)) - 1).Value = 1)
   If CStr(N) = "CambioCliente" Then Activo = (ChkRef.Value = 1)
  ElseIf CmbCas.ListIndex = 4 Then
   If CStr(N) = "ChkDescuento" Then Activo = (ChkOtr.Value = 1)
  Else
   For I = 1 To ViewCas.Nodes.Count
    If ViewCas.Nodes(I).Key = CStr(N) Then Activo = ViewCas.Nodes(I).Checked
   Next I
  End If
  S = S & AtributoXML(CStr(N), IIf(Activo, "1", "0"))
 Next N
 If CmbCas.ListIndex = 4 And ChkCasOtr.Value = 1 Then
  If LstOtr.ListIndex < 0 Then Err.Raise 5, , "Seleccione el caso Otros."
  Otro = LstOtr.ItemData(LstOtr.ListIndex)
 End If
 S = S & AtributoXML("Otro", CStr(Otro)) & ">"
 For I = 0 To LstGui.ListCount - 1
  S = S & "<g" & AtributoXML("Numero", LstGui.List(I)) & "/>"
 Next I
 SolicitudXML = S & "</solicitud>"
End Function
Private Sub CmdGen_Click()
 Dim R As ADODB.Recordset, S As String
 On Error GoTo Fallo
 If IdGuardado > 0 Then VerQuery IdGuardado, 5: Exit Sub
 S = SolicitudXML
 If MsgBox("¿Guardar la solicitud de corrección? Se conservará como expediente para su gestión.", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 Set R = GuardarQuery(Clave, S)
 If R.EOF Then Err.Raise 5, , "El servidor no confirmó la solicitud."
 IdGuardado = CLng(R!Item): LblReg.Caption = CStr(R!Asient)
 CerrarRecordset R
 FraOpt.Enabled = False: TxtFechaInicio.Enabled = False: TxtFechaVisto.Enabled = False: TxtFechaSolicitud.Enabled = False
 StatusBar1.SimpleText = "Solicitud guardada. Consulte o finalice el expediente en Queries."
 VerQuery IdGuardado, 5
 CargarHistoria
 Exit Sub
Fallo:
 S = Err.Description: CerrarRecordset R
 If IdGuardado > 0 Then S = "La solicitud quedó guardada. No se pudo actualizar la vista: " & S
 MsgBox S, vbExclamation, App.Title
End Sub
Private Sub CmdInfoDoc_Click()
 On Error GoTo Fallo
 If Documento Is Nothing Then Err.Raise 5, , "Cargue un documento primero."
 LblDocumento.Caption = TipoCargado & " " & NumeroCargado
 lblEmision.Caption = FechaDocumento(Documento!Emision): LblVencimiento.Caption = FechaDocumento(Documento!Vencimiento)
 LblSubtotal.Caption = Format$(Documento!Subtotal, "#,##0.00"): LblRecargo.Caption = Format$(Documento!Recargo, "#,##0.00")
 LblIgv.Caption = Format$(Documento!IGV, "#,##0.00"): LblTotal.Caption = Format$(Documento!Importe, "#,##0.00")
 LblGastos.Caption = "": LblPorcentaje.Caption = "": LblNombre.Caption = Documento!Cliente: LblCliente.Caption = Documento!Cliente
 LblRuc.Caption = ValorTexto(Documento!RUCDNI): LblDireccion.Caption = ValorTexto(Documento!DIRECC): LblDistrito.Caption = ValorTexto(Documento!Distrito)
 LblDescripcion.Caption = "Documento de importación": LblSituacion.Caption = ValorTexto(Documento!Situacion)
 GridServicios.Rows = 1
 If Servicios.RecordCount > 0 Then Servicios.MoveFirst
 Do While Not Servicios.EOF
  GridServicios.AddItem CeldaDocumento(Servicios!Guia) & vbTab & CeldaDocumento(Servicios!Servicio): Servicios.MoveNext
 Loop
 If GridServicios.Rows > 1 Then GridServicios.Row = 1: SeleccionFilaQuery GridServicios, Seleccionando
 FraVisDoc.Visible = True: FraVisDoc.ZOrder 0
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdAceptar_Click()
 FraVisDoc.Visible = False
End Sub
Private Sub CargarOtros()
 CerrarRecordset Otros: Set Otros = CatalogoQuery(3)
 LstOtr.Clear: LstEliCas.Clear
 Do While Not Otros.EOF
  LstOtr.AddItem CStr(Otros!Descripcion): LstOtr.ItemData(LstOtr.NewIndex) = CLng(Otros!Item)
  LstEliCas.AddItem CStr(Otros!Descripcion): LstEliCas.ItemData(LstEliCas.NewIndex) = CLng(Otros!Item)
  Otros.MoveNext
 Loop
End Sub
Private Sub CmdConfCas_Click()
 FraConfCas.Visible = True: FraConfCas.ZOrder 0: TxtNewCas.MaxLength = 150
End Sub
Private Sub ConfigurarAccionCaso()
 TxtNewCas.Enabled = OptAgr.Value: CmdNewCas.Enabled = OptAgr.Value
 LstEliCas.Enabled = OptEli.Value: CmdEliCas.Enabled = OptEli.Value
End Sub
Private Sub OptAgr_Click()
 ConfigurarAccionCaso
End Sub
Private Sub OptEli_Click()
 ConfigurarAccionCaso
End Sub
Private Sub CmdCerCas_Click()
 FraConfCas.Visible = False
End Sub
Private Sub CmdNewCas_Click()
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If Len(Trim$(TxtNewCas.Text)) = 0 Then Err.Raise 5, , "Ingrese la descripción del caso."
 Set R = GuardarCasoQuery(0, Null, Trim$(TxtNewCas.Text), False): CerrarRecordset R
 TxtNewCas.Text = "": CargarOtros
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdEliCas_Click()
 Dim R As ADODB.Recordset, Id As Long
 On Error GoTo Fallo
 If LstEliCas.ListIndex < 0 Then Err.Raise 5, , "Seleccione el caso que desea dar de baja."
 If MsgBox("¿Dar de baja el caso seleccionado? Las solicitudes existentes conservarán su descripción.", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 Id = LstEliCas.ItemData(LstEliCas.ListIndex)
 Otros.MoveFirst: Otros.Find "Item=" & CStr(Id)
 If Otros.EOF Then Err.Raise 5, , "Recargue los casos."
 Set R = GuardarCasoQuery(Id, Otros.Fields("VersionFila").Value, "", True): CerrarRecordset R
 CargarOtros
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Documento: CerrarRecordset Servicios: CerrarRecordset Otros
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

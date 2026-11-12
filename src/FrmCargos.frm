VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCargos 
   BorderStyle     =   1  'Fixed Single
   Caption = "Control de cargos y registro de estados"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14775
   Icon            =   "FrmCargos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   14775
   Begin VB.Frame Frame2 
      Height          =   1140
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Width           =   10050
      Begin VB.Label LblServ 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7695
         TabIndex        =   26
         Top             =   675
         Width           =   2220
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Serv:"
         Height          =   240
         Left            =   7245
         TabIndex        =   25
         Top             =   675
         Width           =   375
      End
      Begin VB.Label LblGui 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7695
         TabIndex        =   24
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
         Height          =   195
         Left            =   7245
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.Label LblMon 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5715
         TabIndex        =   22
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         Height          =   195
         Left            =   5175
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label LblDiv 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4005
         TabIndex        =   13
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Divisa:"
         Height          =   195
         Left            =   3510
         TabIndex        =   12
         Top             =   360
         Width           =   480
      End
      Begin VB.Label LblCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1125
         TabIndex        =   11
         Top             =   675
         Width           =   5910
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   675
         Width           =   525
      End
      Begin VB.Label LblDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1125
         TabIndex        =   9
         Top             =   315
         Width           =   1950
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   315
         Width           =   870
      End
   End
   Begin VB.Frame FraHer 
      BackColor       =   &H0053555F&
      BorderStyle     =   0  'None
      Height          =   7440
      Left            =   60
      TabIndex        =   62
      Top             =   45
      Width           =   2310
      Begin MSComctlLib.TreeView ViewBanda 
         Height          =   2580
         Left            =   90
         TabIndex        =   63
         Top             =   405
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   4551
         _Version        =   393217
         Indentation     =   178
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin Clever.Boton CmdExport 
         Height          =   330
         Left            =   90
         TabIndex        =   64
         Top             =   3195
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         BTYPE           =   14
         TX              =   "Exportar a Excel"
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
         MICON           =   "FrmCargos.frx":1782
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdReg 
         Height          =   330
         Left            =   90
         TabIndex        =   65
         Top             =   3735
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         BTYPE           =   14
         TX = "Ver / reimprimir"
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
         MICON           =   "FrmCargos.frx":179E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdAnl 
         Height          =   330
         Left            =   90
         TabIndex        =   66
         Top             =   4140
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         BTYPE           =   14
         TX = "Vigencia de registros"
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
         MICON           =   "FrmCargos.frx":17BA
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
         Height          =   330
         Left            =   90
         TabIndex        =   67
         Top             =   6975
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         BTYPE           =   14
         TX = "Generar cargos"
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
         MICON           =   "FrmCargos.frx":17D6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdLim 
         Height          =   330
         Left            =   90
         TabIndex        =   68
         Top             =   6570
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         BTYPE           =   14
         TX = "Vaciar banda"
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
         MICON           =   "FrmCargos.frx":17F2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblRow 
         Caption         =   "0"
         Height          =   285
         Left            =   180
         TabIndex        =   85
         Top             =   4590
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Niveles de Bandas"
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
         TabIndex        =   69
         Top             =   45
         Width           =   1860
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1140
      Left            =   12645
      TabIndex        =   14
      Top             =   360
      Width           =   2040
      Begin VB.Label VisCas 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   810
         TabIndex        =   20
         Top             =   765
         Width           =   1140
      End
      Begin VB.Label VisSit 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   810
         TabIndex        =   18
         Top             =   450
         Width           =   1140
      End
      Begin VB.Label VisEst 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   810
         TabIndex        =   16
         Top             =   135
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   45
         TabIndex        =   15
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Situacion:"
         Height          =   195
         Left            =   45
         TabIndex        =   17
         Top             =   495
         Width           =   705
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Con Caso:"
         Height          =   195
         Left            =   45
         TabIndex        =   19
         Top             =   810
         Width           =   735
      End
   End
   Begin VB.ComboBox CmbTip 
      BackColor = &H80000005&
      Enabled = -1
      Height          =   315
      ItemData        =   "FrmCargos.frx":180E
      Left = 7050
      List            =   "FrmCargos.frx":1810
      Style = 2
      TabIndex        =   5
      Top             =   45
      Width = 900
   End
   Begin VB.TextBox TxtDoc 
      MaxLength = 50
      Appearance      =   0  'Flat
      BackColor = &H80000005&
      Enabled = -1
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   45
      Width = 3000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2745
      Top             =   7110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargos.frx":1812
            Key             =   "a1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargos.frx":1DAC
            Key             =   "a2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargos.frx":2346
            Key             =   "A3"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   7650
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4059
            MinWidth        =   4059
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15876
            MinWidth        =   15876
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   6090
      Left            =   12645
      TabIndex        =   6
      Top             =   1485
      Width           =   2040
      Begin VB.Label ConOtr 
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
         Left            =   1350
         TabIndex        =   51
         Top             =   3465
         Width           =   555
      End
      Begin VB.Label ConNc 
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
         Left            =   1350
         TabIndex        =   50
         Top             =   3195
         Width           =   555
      End
      Begin VB.Label ConPer 
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
         Left            =   1350
         TabIndex        =   49
         Top             =   2925
         Width           =   555
      End
      Begin VB.Label ConRba 
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
         Index           =   4
         Left            =   1350
         TabIndex        =   48
         Top             =   2655
         Width           =   555
      End
      Begin VB.Label ConRba 
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
         Index           =   3
         Left            =   1350
         TabIndex        =   47
         Top             =   2385
         Width           =   555
      End
      Begin VB.Label ConRba 
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
         Index           =   2
         Left            =   1350
         TabIndex        =   46
         Top             =   2115
         Width           =   555
      End
      Begin VB.Label ConRba 
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
         Index           =   1
         Left            =   1350
         TabIndex        =   45
         Top             =   1845
         Width           =   555
      End
      Begin VB.Label ConRba 
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
         Index           =   0
         Left            =   1350
         TabIndex        =   44
         Top             =   1575
         Width           =   555
      End
      Begin VB.Label ConAnl 
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
         Left            =   1350
         TabIndex        =   43
         Top             =   855
         Width           =   555
      End
      Begin VB.Label ConAct 
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
         Left            =   1350
         TabIndex        =   42
         Top             =   585
         Width           =   555
      End
      Begin VB.Label LblInf 
         Caption = ""
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1860
         Left            =   90
         TabIndex        =   41
         Top             =   4140
         Width           =   1860
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "Informaciones"
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
         Left            =   90
         TabIndex        =   40
         Top             =   3825
         Width           =   1860
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Reparto Otros"
         Height          =   195
         Left            =   135
         TabIndex        =   39
         Top             =   3465
         Width           =   990
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Nota de Credito"
         Height          =   195
         Left            =   135
         TabIndex        =   38
         Top             =   3195
         Width           =   1110
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Personalizado"
         Height          =   195
         Left            =   135
         TabIndex        =   37
         Top             =   2925
         Width           =   990
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Registros B5"
         Height          =   195
         Left            =   135
         TabIndex        =   36
         Top             =   2655
         Width           =   900
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Registros B4"
         Height          =   195
         Left            =   135
         TabIndex        =   35
         Top             =   2385
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Registros B3"
         Height          =   195
         Left            =   135
         TabIndex        =   34
         Top             =   2115
         Width           =   900
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Registros B2"
         Height          =   195
         Left            =   135
         TabIndex        =   33
         Top             =   1845
         Width           =   900
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Registros B1"
         Height          =   195
         Left            =   135
         TabIndex        =   32
         Top             =   1575
         Width           =   900
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "Cantidades"
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
         Left            =   90
         TabIndex        =   31
         Top             =   1215
         Width           =   1860
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00404080&
         Caption         =   "Situación"
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
         Left            =   90
         TabIndex        =   30
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Anulados"
         Height          =   195
         Left            =   135
         TabIndex        =   29
         Top             =   855
         Width           =   660
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Activos"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   585
         Width           =   525
      End
   End
   Begin VB.Frame FraGenCarg 
      Caption         =   "Generar Cargos"
      Height          =   1500
      Left            =   2520
      TabIndex        =   53
      Top             =   6075
      Visible         =   0   'False
      Width           =   10050
      Begin VB.ComboBox CmbRep 
         Height          =   315
         ItemData        =   "FrmCargos.frx":28E0
         Left            =   135
         List            =   "FrmCargos.frx":28E2
         Style = 2
         TabIndex        =   54
         Top             =   630
         Width           =   6975
      End
      Begin Clever.Boton CmdRepReg 
         Height          =   330
         Left            =   7380
         TabIndex        =   56
         Top             =   270
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   582
         BTYPE           =   14
         TX = "Registrar y ver reporte"
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
         MICON           =   "FrmCargos.frx":28E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdSolRep 
         Height          =   330
         Left            =   7380
         TabIndex        =   57
         Top             =   630
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   582
         BTYPE           =   14
         TX = "Vista previa"
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
         MICON           =   "FrmCargos.frx":2900
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdSolReg 
         Height          =   330
         Left            =   7380
         TabIndex        =   58
         Top             =   990
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   582
         BTYPE           =   14
         TX              =   "Solo Registrar"
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
         MICON           =   "FrmCargos.frx":291C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdCanReg 
         Height          =   330
         Left            =   135
         TabIndex        =   70
         Top             =   990
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         BTYPE           =   14
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   12632256
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCargos.frx":2938
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblCorrel 
         Alignment       =   1  'Right Justify
         Caption         =   "000000"
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
         Left            =   6300
         TabIndex        =   61
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Correlativo:"
         Height          =   240
         Left            =   5400
         TabIndex        =   60
         Top             =   315
         Width           =   825
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha:"
         Height          =   240
         Left            =   135
         TabIndex        =   59
         Top             =   315
         Width           =   735
      End
   Begin VB.TextBox TxtFecha
      Left=855
      Top=270
      Width=1695
      Height=315
      MaxLength=10
   End
   End
   Begin MSFlexGridLib.MSFlexGrid GridCargos 
      Height          =   6000
      Left            =   2520
      TabIndex        =   1
      Top             =   1575
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   10583
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.Frame FraRegen 
      Height          =   6090
      Left            =   2520
      TabIndex        =   71
      Top             =   1485
      Visible         =   0   'False
      Width           =   10050
      Begin VB.OptionButton OptList 
         Caption         =   "Listado 50 Ultimos:"
         Height          =   240
         Left            =   225
         TabIndex        =   82
         Top             =   1305
         Width           =   1995
      End
      Begin VB.TextBox TxtRegist 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   225
         TabIndex        =   81
         Top             =   855
         Width           =   1410
      End
      Begin VB.OptionButton OptRegist 
         Caption         =   "Registro:"
         Height          =   240
         Left            =   225
         TabIndex        =   80
         Top             =   585
         Value           =   -1  'True
         Width           =   960
      End
      Begin MSFlexGridLib.MSFlexGrid GridRegenerar 
         Height          =   4425
         Left            =   2340
         TabIndex        =   72
         Top             =   585
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   7805
         _Version        =   393216
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid GridRegen 
         Height          =   3390
         Left            =   180
         TabIndex        =   73
         Top             =   1620
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   5980
         _Version        =   393216
         FixedCols       =   0
         Enabled         =   0   'False
         SelectionMode   =   1
      End
      Begin Clever.Boton CmdRegenerar 
         Height          =   375
         Left            =   5220
         TabIndex        =   74
         Top             =   5130
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Regenerar Reporte"
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
         MICON           =   "FrmCargos.frx":2954
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdRegenerarExcel 
         Height          =   375
         Left            =   5220
         TabIndex        =   75
         Top             =   5535
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Regenerar en Excel"
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
         MICON           =   "FrmCargos.frx":2970
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
         Height          =   375
         Left            =   7560
         TabIndex        =   76
         Top             =   5130
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   661
         BTYPE           =   5
         TX = "Anular registro"
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
         MICON           =   "FrmCargos.frx":298C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdReactivar 
         Height          =   375
         Left            =   7560
         TabIndex        =   77
         Top             =   5535
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   661
         BTYPE           =   5
         TX = "Reactivar registro"
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
         MICON           =   "FrmCargos.frx":29A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton cmdAcepRegen 
         Height          =   375
         Left            =   180
         TabIndex        =   78
         Top             =   5535
         Width           =   2040
         _ExtentX        =   3598
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
         MICON           =   "FrmCargos.frx":29C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdOkRegist 
         Height          =   285
         Left            =   1755
         TabIndex        =   83
         Top             =   855
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "OK"
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
         MICON           =   "FrmCargos.frx":29E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblTitRegist 
         Alignment       =   1  'Right Justify
         Caption = ""
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5580
         TabIndex        =   84
         Top             =   270
         Width           =   4335
      End
      Begin VB.Label Label11 
         Caption         =   "Regeneracion de Reportes de Registro de Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   79
         Top             =   270
         Width           =   5415
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption = "Enter: agregar"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   12555
      TabIndex        =   52
      Top             =   45
      Width = 2100
   End
   Begin VB.Label LblReg 
      Caption         =   "BANDA EXPORT  - BANDA Nº 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left = 8100
      TabIndex        =   27
      Top             =   90
      Width = 4200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left = 6600
      TabIndex        =   4
      Top             =   90
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Documento:"
      Height          =   195
      Left            =   2610
      TabIndex        =   3
      Top             =   90
      Width           =   870
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
Attribute VB_Name = "FrmCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Bandas(1 To 8) As ADODB.Recordset
Private Estados As ADODB.Recordset, Registros As ADODB.Recordset, DetalleRegistro As ADODB.Recordset
Private Banda As Integer, Cargando As Boolean, PosicionInicialAplicada As Boolean
Private Sub Form_Load()
 Dim I As Integer, Nodo As MSComctlLib.Node
 On Error GoTo Fallo
 If Not AccesoPermitido(19) Then Unload Me: Exit Sub
 Cargando = True
 For I = 1 To 8: Set Bandas(I) = NuevaBandaCargos(): Next I
 ViewBanda.Nodes.Clear
 For I = 1 To 5
  ViewBanda.Nodes.Add , , "B" & CStr(I), "Banda Nº " & CStr(I)
 Next I
 ViewBanda.Nodes.Add , , "B6", "Personalizada"
 ViewBanda.Nodes.Add , , "B7", "Notas de crédito"
 ViewBanda.Nodes.Add , , "B8", "Reparto otros"
 CmbTip.Clear
 For Each Nodo In ViewBanda.Nodes: Nodo.Expanded = True: Next Nodo
 CmbTip.AddItem "FT": CmbTip.AddItem "BV": CmbTip.AddItem "NC": CmbTip.AddItem "ND"
 CmbTip.ListIndex = 0
 TxtDoc.BackColor = vbWindowBackground
 StatusBar.Panels.Clear: StatusBar.Panels.Add
 StatusBar.Panels(1).AutoSize = sbrSpring
 TxtFecha.Text = Format$(Date, "dd/mm/yyyy")
 LblInf.Caption = "Escriba el número completo y pulse Enter. Supr retira una fila; F5 actualiza la banda. Doble clic muestra el historial."
 Label14.Caption = "Registro:"
 Label5.Caption = "Enter: agregar"
 LblTitRegist.Caption = ""
 PrepararGrillaDocumento GridRegen, "Registro|Estado"
 GridRegen.ColWidth(0) = 850: GridRegen.ColWidth(1) = 1100
 PrepararGrillaDocumento GridRegenerar, "Tipo|Documento|Emisión|Cliente|Moneda|Importe"
 GridRegenerar.ColWidth(0) = 500: GridRegenerar.ColWidth(1) = 1800
 GridRegenerar.ColWidth(2) = 1100: GridRegenerar.ColWidth(3) = 2200: GridRegenerar.ColWidth(4) = 800: GridRegenerar.ColWidth(5) = 1000
 CargarEstados
 Banda = 1
 ViewBanda.Nodes("B1").Selected = True
 Cargando = False
 CmdCanReg_Click
 PintarBanda
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
 Unload Me
End Sub
Private Sub CargarEstados()
 CerrarRecordset Estados
 Set Estados = EstadosCargos()
 CmbRep.Clear
 Do While Not Estados.EOF
  CmbRep.AddItem CStr(Estados!TitEst)
  CmbRep.ItemData(CmbRep.NewIndex) = CLng(Estados!ITEM)
  Estados.MoveNext
 Loop
 If CmbRep.ListCount > 0 Then CmbRep.ListIndex = 0
End Sub
Private Sub ViewBanda_Click()
 If Cargando Then Exit Sub
 If ViewBanda.SelectedItem Is Nothing Then Exit Sub
 Banda = CInt(Mid$(ViewBanda.SelectedItem.Key, 2))
 PintarBanda
End Sub
Private Sub PintarBanda()
 Dim R As ADODB.Recordset, I As Integer, Activos As Long, Anulados As Long, Texto As String
 If Banda < 1 Or Banda > 8 Then Exit Sub
 Set R = Bandas(Banda)
 Cargando = True
 PrepararGrillaDocumento GridCargos, "Tipo|Documento|Emisión|Cliente|Moneda|Importe|Situación|Gestión"
 GridCargos.ColWidth(0) = 450: GridCargos.ColWidth(1) = 1700: GridCargos.ColWidth(2) = 1000
 GridCargos.ColWidth(3) = 3000: GridCargos.ColWidth(4) = 650: GridCargos.ColWidth(5) = 1100
 GridCargos.ColWidth(6) = 900: GridCargos.ColWidth(7) = 1800
 If Banda = 8 Then
  GridCargos.Cols = 12
  GridCargos.TextMatrix(0, 8) = "Contacto": GridCargos.ColWidth(8) = 2000
  GridCargos.TextMatrix(0, 9) = "Dirección": GridCargos.ColWidth(9) = 4000
  GridCargos.TextMatrix(0, 10) = "Distrito": GridCargos.ColWidth(10) = 1800
  GridCargos.TextMatrix(0, 11) = "Teléfono": GridCargos.ColWidth(11) = 1800
 End If
 If R.RecordCount > 0 Then R.MoveFirst
 Do While Not R.EOF
  Texto = CStr(R!Tipo) & vbTab & CeldaDocumento(R!Documento) & vbTab & FechaDocumento(R!Emision) & vbTab & CeldaDocumento(R!Cliente) & vbTab & CeldaDocumento(R!Divisa) & vbTab & Format$(R!Importe, "#,##0.00") & vbTab & CeldaDocumento(R!Situacion) & vbTab & CeldaDocumento(R!Estado)
  If Banda = 8 Then Texto = Texto & vbTab & CeldaDocumento(R!ContactoEntrega) & vbTab & CeldaDocumento(R!DireccionEntrega) & vbTab & CeldaDocumento(R!DistritoEntrega) & vbTab & CeldaDocumento(R!TelefonoEntrega)
  GridCargos.AddItem Texto
  If CBool(R!Anulado) Then Anulados = Anulados + 1 Else Activos = Activos + 1
  R.MoveNext
 Loop
 For I = 1 To 5: ConRba(I - 1).Caption = CStr(Bandas(I).RecordCount): Next I
 ConPer.Caption = CStr(Bandas(6).RecordCount): ConNc.Caption = CStr(Bandas(7).RecordCount): ConOtr.Caption = CStr(Bandas(8).RecordCount)
 ConAct.Caption = CStr(Activos): ConAnl.Caption = CStr(Anulados)
 LblReg.Caption = UCase$(ViewBanda.Nodes("B" & CStr(Banda)).Text)
 If GridCargos.Rows > 1 Then GridCargos.Row = 1
 Cargando = False
 MostrarSeleccion
 StatusBar.Panels(1).Text = CStr(R.RecordCount) & " documentos en esta banda. Los documentos emitidos se conservan."
End Sub
Private Function SeleccionBanda() As Boolean
 If Banda < 1 Or Banda > 8 Then Exit Function
 If GridCargos.Row < 1 Or GridCargos.Row >= GridCargos.Rows Then Exit Function
 If Bandas(Banda).RecordCount = 0 Then Exit Function
 Bandas(Banda).MoveFirst: Bandas(Banda).Move GridCargos.Row - 1
 SeleccionBanda = True
End Function
Private Sub MostrarSeleccion()
 Dim R As ADODB.Recordset
 LblDoc.Caption = "": LblCli.Caption = "": LblDiv.Caption = "": LblMon.Caption = ""
 LblGui.Caption = "": LblServ.Caption = "": VisEst.Caption = "": VisSit.Caption = "": VisCas.Caption = ""
 If Cargando Then Exit Sub
 If Not SeleccionBanda Then Exit Sub
 Set R = Bandas(Banda)
 LblDoc.Caption = CStr(R!Documento): LblCli.Caption = CStr(R!Cliente): LblCli.ToolTipText = Left$(LblCli.Caption, 255)
 LblDiv.Caption = CStr(R!Divisa): LblMon.Caption = Format$(R!Importe, "#,##0.00")
 LblGui.Caption = ValorTexto(R!Guias): LblGui.ToolTipText = Left$(LblGui.Caption, 255)
 LblServ.Caption = ValorTexto(R!Servicios): LblServ.ToolTipText = Left$(LblServ.Caption, 255)
 VisEst.Caption = ValorTexto(R!Situacion): VisSit.Caption = ValorTexto(R!Estado)
 VisSit.ToolTipText = VisSit.Caption
End Sub
Private Sub GridCargos_RowColChange()
 If Not Cargando Then MostrarSeleccion
End Sub
Private Sub TxtDoc_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then AgregarDocumento: KeyAscii = 0
End Sub
Private Sub AgregarDocumento()
 Dim R As ADODB.Recordset, B As ADODB.Recordset, Agregando As Boolean
 On Error GoTo Fallo
 If Banda < 1 Or Len(Trim$(TxtDoc.Text)) = 0 Then Exit Sub
 If Banda = 7 And CmbTip.Text <> "NC" Then Err.Raise 5, , "Seleccione NC para la banda de notas de crédito."
 Set R = ConsultarDocumentos(CmbTip.Text, Trim$(TxtDoc.Text), True, 19)
 If R.EOF Then Err.Raise 5, , "No se encontró el documento de ese tipo. Ingrese el número completo."
 Set B = Bandas(Banda)
 If B.RecordCount >= 200 Then Err.Raise 5, , "La banda admite como máximo 200 documentos."
 If B.RecordCount > 0 Then B.MoveFirst
 Do While Not B.EOF
  If B!Tipo = R!Tipo And B!IdDocumento = R!IdDocumento Then Err.Raise 5, , "El documento ya está en esta banda."
  B.MoveNext
 Loop
 Agregando = True
 B.AddNew
 CopiarDocumentoBanda B, R
 Agregando = False
 CerrarRecordset R
 TxtDoc.Text = ""
 PintarBanda
 Exit Sub
Fallo:
 If Agregando And Not B Is Nothing Then
  If B.EditMode <> adEditNone Then B.CancelUpdate
 End If
 CerrarRecordset R
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub GridCargos_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete Then
  If SeleccionBanda Then
   Bandas(Banda).Delete
   Bandas(Banda).MoveNext
   PintarBanda
  End If
  KeyCode = 0
 End If
 If KeyCode = vbKeyReturn Then GridCargos_DblClick: KeyCode = 0
End Sub
Private Sub GridCargos_DblClick()
 On Error GoTo Fallo
 If Not SeleccionBanda Then Exit Sub
 FrmCargosHistorial.Preparar CStr(Bandas(Banda)!Tipo), CLng(Bandas(Banda)!IdDocumento)
 FrmCargosHistorial.Show vbModal, MdiPri
 Unload FrmCargosHistorial
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdLim_Click()
 If Bandas(Banda).RecordCount = 0 Then Exit Sub
 If MsgBox("¿Vaciar los documentos de esta banda? Solo se retira la lista de trabajo; no se borran documentos ni registros guardados.", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
 CerrarRecordset Bandas(Banda)
 Set Bandas(Banda) = NuevaBandaCargos()
 PintarBanda
End Sub
Private Sub RefrescarBanda()
 Dim B As ADODB.Recordset, R As ADODB.Recordset
 On Error GoTo Fallo
 Set B = Bandas(Banda)
 If B.RecordCount > 0 Then B.MoveFirst
 Do While Not B.EOF
  Set R = ConsultarDocumentos(CStr(B!Tipo), CStr(B!Documento), True, 19)
  If R.EOF Then Err.Raise 5, , "Un documento ya no está disponible. Retírelo de la banda."
  CopiarDocumentoBanda B, R
  CerrarRecordset R
  B.MoveNext
 Loop
 PintarBanda
 Exit Sub
Fallo:
 CerrarRecordset R
 Err.Raise vbObjectError + 3510, , Err.Description
End Sub
Private Sub CmdGen_Click()
 On Error GoTo Fallo
 If Bandas(Banda).RecordCount = 0 Then Err.Raise 5, , "Agregue documentos a la banda."
 CargarEstados
 FraGenCarg.Visible = True
 GridCargos.Height = 4470
 FraHer.Enabled = False: TxtDoc.Enabled = False: CmbTip.Enabled = False
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmbRep_Click()
 LblCorrel.Caption = "—"
 If Estados Is Nothing Then Exit Sub
 If CmbRep.ListIndex < 0 Then Exit Sub
 Estados.MoveFirst
 Do While Not Estados.EOF
  If CLng(Estados!ITEM) = CmbRep.ItemData(CmbRep.ListIndex) Then
   If ValorTexto(Estados!ActCor) = "1" Then LblCorrel.Caption = Format$(Val(ValorTexto(Estados!Correl)) + 1, "000000")
   Exit Do
  End If
  Estados.MoveNext
 Loop
End Sub
Private Sub CmdCanReg_Click()
 FraGenCarg.Visible = False: FraRegen.Visible = False: GridCargos.Visible = True
 GridCargos.Height = 6000
 FraHer.Enabled = True: TxtDoc.Enabled = True: CmbTip.Enabled = True
End Sub
Private Sub ProcesarBanda(ByVal Registrar As Boolean, ByVal Reporte As Boolean)
 Dim Fecha As Date, R As ADODB.Recordset, Numero As String, Correlativo As String, Titulo As String, Confirmado As Boolean
 On Error GoTo Fallo
 If CmbRep.ListIndex < 0 Then Err.Raise 5, , "Seleccione el estado de gestión."
 If Not FechaGestionValida(TxtFecha.Text, Fecha) Then Err.Raise 5, , "Ingrese una fecha válida en formato dd/mm/aaaa, hasta hoy."
 If Bandas(Banda).RecordCount = 0 Then Err.Raise 5, , "La banda está vacía."
 If CLng(ConAnl.Caption) > 0 Then Err.Raise 5, , "La banda contiene documentos anulados por el sistema externo. Retírelos para generar cargos."
 Titulo = CmbRep.Text
 If Registrar Then
  If MsgBox("¿Registrar " & CStr(Bandas(Banda).RecordCount) & " documentos con el estado «" & Titulo & "»?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
  Set R = GuardarCargo(CmbRep.ItemData(CmbRep.ListIndex), Fecha, Bandas(Banda))
  Confirmado = True
  Numero = Format$(R!Numero, "000000")
  Correlativo = ValorTexto(R!Correlativo)
  If Len(Correlativo) > 0 Then Correlativo = Format$(CLng(Correlativo), "000000")
  CerrarRecordset R
  If Reporte Then
   Set R = RegistrosCargos(CLng(Numero))
   Set DetalleRegistro = DetalleCargo(CLng(R!Item))
   MostrarReporteCargos DetalleRegistro, Titulo, Numero, Correlativo, Fecha
   CerrarRecordset R
  End If
  RefrescarBanda
  MsgBox "Registro " & Numero & " guardado. Se conservaron los datos del documento externo.", vbInformation, App.Title
 ElseIf Reporte Then
  MostrarReporteCargos Bandas(Banda), "VISTA PREVIA - " & Titulo, "Sin registrar", "", Fecha
 End If
 CmdCanReg_Click
 Exit Sub
Fallo:
 CerrarRecordset R
 If Confirmado Then
  MsgBox "El registro " & Numero & " quedó guardado. No repita el guardado. No se pudo completar la vista o actualización: " & Err.Description & ". Consulte Registros para reimprimir.", vbExclamation, App.Title
 Else
  MsgBox Err.Description, vbExclamation, App.Title
 End If
End Sub
Private Sub CmdRepReg_Click()
 ProcesarBanda True, True
End Sub
Private Sub CmdSolReg_Click()
 ProcesarBanda True, False
End Sub
Private Sub CmdSolRep_Click()
 ProcesarBanda False, True
End Sub
Private Sub CmdExport_Click()
 On Error GoTo Fallo
 ExportarDocumentosExcel Bandas(Banda), LblReg.Caption
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub AbrirRegistros()
 On Error GoTo Fallo
 FraRegen.Visible = True: GridCargos.Visible = False: FraGenCarg.Visible = False
 FraHer.Enabled = False: TxtDoc.Enabled = False: CmbTip.Enabled = False
 Cargando = True: OptList.Value = True: Cargando = False
 CargarRegistros 0
 Exit Sub
Fallo:
 Cargando = False
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CargarRegistros(ByVal Numero As Long)
 CerrarRecordset Registros: CerrarRecordset DetalleRegistro
 Set Registros = RegistrosCargos(Numero)
 Cargando = True
 GridRegen.Rows = 1: GridRegenerar.Rows = 1
 LblTitRegist.Caption = ""
 Do While Not Registros.EOF
  GridRegen.AddItem Format$(Registros!Numero, "000000") & vbTab & IIf(CBool(Registros!Activo), "Activo", "Retirado")
  Registros.MoveNext
 Loop
 If GridRegen.Rows > 1 Then GridRegen.Row = 1
 Cargando = False
 CargarSeleccionRegistro
End Sub
Private Sub CargarSeleccionRegistro()
 On Error GoTo Fallo
 CerrarRecordset DetalleRegistro
 GridRegenerar.Rows = 1
 CmdRegenerar.Enabled = False: CmdRegenerarExcel.Enabled = False: CmdAnular.Enabled = False: CmdReactivar.Enabled = False
 LblTitRegist.Caption = ""
 If Registros Is Nothing Then Exit Sub
 If GridRegen.Row < 1 Or GridRegen.Row >= GridRegen.Rows Then Exit Sub
 Registros.MoveFirst: Registros.Move GridRegen.Row - 1
 Set DetalleRegistro = DetalleCargo(CLng(Registros!Item))
 LblTitRegist.Caption = CStr(Registros!Titulo) & " - " & FechaDocumento(Registros!FechaEstado)
 LblTitRegist.ToolTipText = LblTitRegist.Caption & " / " & CStr(Registros!Usuario)
 Do While Not DetalleRegistro.EOF
  GridRegenerar.AddItem CStr(DetalleRegistro!Tipo) & vbTab & CeldaDocumento(DetalleRegistro!Documento) & vbTab & FechaDocumento(DetalleRegistro!Emision) & vbTab & CeldaDocumento(DetalleRegistro!Cliente) & vbTab & CeldaDocumento(DetalleRegistro!Divisa) & vbTab & Format$(DetalleRegistro!Importe, "#,##0.00")
  DetalleRegistro.MoveNext
 Loop
 CmdRegenerar.Enabled = (GridRegenerar.Rows > 1): CmdRegenerarExcel.Enabled = CmdRegenerar.Enabled
 CmdAnular.Enabled = CBool(Registros!Activo): CmdReactivar.Enabled = Not CmdAnular.Enabled
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdReg_Click()
 AbrirRegistros
End Sub
Private Sub CmdAnl_Click()
 AbrirRegistros
End Sub
Private Sub GridRegen_RowColChange()
 If Not Cargando Then CargarSeleccionRegistro
End Sub
Private Sub OptList_Click()
 If Not Cargando Then CmdOkRegist_Click
End Sub
Private Sub OptRegist_Click()
 If Cargando Then Exit Sub
 TxtRegist.Enabled = True
 TxtRegist.SetFocus
End Sub
Private Sub CmdOkRegist_Click()
 Dim Numero As Long
 On Error GoTo Fallo
 If OptRegist.Value Then
  If Len(Trim$(TxtRegist.Text)) = 0 Or Trim$(TxtRegist.Text) Like "*[!0-9]*" Then Err.Raise 5, , "Ingrese el número de registro."
  Numero = CLng(TxtRegist.Text)
  If Numero <= 0 Then Err.Raise 5, , "Registro inválido."
 End If
 CargarRegistros Numero
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub TxtRegist_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then CmdOkRegist_Click: KeyAscii = 0
End Sub
Private Sub cmdAcepRegen_Click()
 CmdCanReg_Click
End Sub
Private Sub CambiarVigencia(ByVal Activo As Boolean)
 Dim Numero As Long, Confirmado As Boolean, Mensaje As String
 On Error GoTo Fallo
 If Registros Is Nothing Then Exit Sub
 If Registros.RecordCount = 0 Then Exit Sub
 If Registros.EOF Then Exit Sub
 If CBool(Registros!Activo) = Activo Then Exit Sub
 Numero = CLng(Registros!Numero)
 Mensaje = IIf(Activo, "¿Reactivar", "¿Anular") & " el registro de cargos " & Format$(Numero, "000000") & "?" & vbCrLf & "Se recalculará la gestión de sus documentos. Las facturas, importes, guías y sus situaciones externas se conservan; el historial permanece visible."
 If MsgBox(Mensaje, vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
 CambiarCargo CLng(Registros!Item), Registros.Fields("VersionFila").Value, Activo
 Confirmado = True
 CargarRegistros Numero
 Exit Sub
Fallo:
 If Confirmado Then Mensaje = "El cambio quedó guardado; recargue el registro. " Else Mensaje = ""
 MsgBox Mensaje & Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdAnular_Click()
 CambiarVigencia False
End Sub
Private Sub CmdReactivar_Click()
 CambiarVigencia True
End Sub
Private Sub CmdRegenerar_Click()
 Dim Titulo As String
 On Error GoTo Fallo
 If DetalleRegistro Is Nothing Then Exit Sub
 If DetalleRegistro.RecordCount = 0 Then Exit Sub
 Titulo = CStr(Registros!Titulo)
 If Not CBool(Registros!Activo) Then Titulo = "[ANULADO] " & Titulo
 MostrarReporteCargos DetalleRegistro, Titulo, Format$(Registros!Numero, "000000"), ValorTexto(Registros!Correlativo), CDate(Registros!FechaEstado)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdRegenerarExcel_Click()
 On Error GoTo Fallo
 If DetalleRegistro Is Nothing Then Exit Sub
 ExportarDocumentosExcel DetalleRegistro, "Registro de cargos " & CStr(Registros!Numero)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo Fallo
 If KeyCode = vbKeyF5 Then
  If FraRegen.Visible Then CmdOkRegist_Click Else RefrescarBanda
  KeyCode = 0
 End If
 If KeyCode = vbKeyEscape Then
  If FraGenCarg.Visible Or FraRegen.Visible Then CmdCanReg_Click Else Unload Me
 End If
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Dim I As Integer, Cantidad As Long
 For I = 1 To 8
  If Not Bandas(I) Is Nothing Then Cantidad = Cantidad + Bandas(I).RecordCount
 Next I
 If Cantidad > 0 Then
  If MsgBox("¿Cerrar y descartar las listas de trabajo de las bandas? Los registros de cargos guardados se conservan.", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Cancel = True
 End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Dim I As Integer
 For I = 1 To 8: CerrarRecordset Bandas(I): Next I
 CerrarRecordset Estados: CerrarRecordset Registros: CerrarRecordset DetalleRegistro
End Sub
Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

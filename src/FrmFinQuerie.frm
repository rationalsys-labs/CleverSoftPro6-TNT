VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFinQuerie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finalizacion de Queries Emitidas"
   ClientHeight = 7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight = 7950
   ScaleWidth      =   10455
   Begin VB.Frame FraPend 
      Height = 7530
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin MSFlexGridLib.MSFlexGrid GridPend 
         Height = 2595
         Left            =   120
         TabIndex        =   1
         Top = 660
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5318
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
      Begin VB.Frame FraNc 
         Height          =   2295
         Left            =   0
         TabIndex        =   18
         Top = 5100
         Visible = 0
         Width           =   10215
         Begin Clever.Boton CmdBus 
            Height          =   285
            Left            =   4440
            TabIndex        =   19
            Top             =   600
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            BTYPE           =   3
            TX = "..."
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            MICON           =   "FrmFinQuerie.frx":0000
            PICN            =   "FrmFinQuerie.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox TxtNc 
            MaxLength = 14
            Height          =   285
            Left            =   2520
            TabIndex        =   20
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox TxtDocNc 
            MaxLength = 14
            Height          =   315
            Left            =   6000
            TabIndex        =   23
            Top             =   365
            Width           =   1815
         End
         Begin VB.ListBox LstDocNc 
            BackColor       =   &H00FFFFFF&
            Height          =   1230
            ItemData        =   "FrmFinQuerie.frx":05B6
            Left            =   5040
            List            =   "FrmFinQuerie.frx":05BD
            TabIndex        =   22
            Top             =   810
            Width           =   2775
         End
         Begin VB.ComboBox CmbDivNc 
            Height          =   315
            ItemData        =   "FrmFinQuerie.frx":05CC
            Left            =   5040
            List            =   "FrmFinQuerie.frx":05CE
            Style = 2
            TabIndex        =   21
            Top             =   360
            Width           =   855
         End
   Begin VB.Label LabelDegrade2
      Left=120
      Top=240
      Width=3975
      Height=255
      Caption="RESOLUCIÓN CON NOTA DE CRÉDITO"
      BackStyle=0
      BackStyle=1
      BackColor=&H0083443B&
      ForeColor=&H00FFFFFF&

   End
         Begin Clever.Boton CmdFinQueNc 
            Height          =   375
            Left            =   7920
            TabIndex        =   25
            Top             =   1680
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            BTYPE           =   6
            TX              =   "Finalizar Querie"
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
            BCOLO           =   6185064
            FCOL            =   16777215
            FCOLO           =   8454143
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmFinQuerie.frx":05D0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Clever.Boton CmdLimNc 
            Height          =   375
            Left            =   7920
            TabIndex        =   26
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            BTYPE           =   6
            TX              =   "Limpiar Datos"
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
            BCOLO           =   6185064
            FCOL            =   16777215
            FCOLO           =   8454143
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmFinQuerie.frx":05EC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   1935
            Left            =   4920
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label LblTipNc 
            BackColor       =   &H00004080&
            Caption         =   "Anulacion"
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
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label LblNcFac 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   38
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label21 
            Caption         =   "Factura"
            Height          =   255
            Left            =   2520
            TabIndex        =   37
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label LblNcQue 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   36
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Querie"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label LblMonNc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   34
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label24 
            Caption         =   "Monto"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "Aplicado"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label LblExcNc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   30
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label30 
            Caption         =   "Exceso"
            Height          =   255
            Left            =   2520
            TabIndex        =   29
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label18 
            WordWrap = -1
            BackStyle = 0
            Caption = "Enter: agregar reemplazo. Supr: retirar de la lista."
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor = 0
            Height          =   855
            Left            =   7920
            TabIndex        =   28
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label LblDivNc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   1440
            Width           =   495
         End
   Begin VB.TextBox TxtAplicado
      Left=840
      Top=1785
      Width=1455
      Height=285
      MaxLength=18
   End
      End
      Begin VB.Frame FraDir 
         Height          =   2295
         Left            =   0
         TabIndex        =   2
         Top = 5100
         Visible = 0
         Width           =   10215
         Begin VB.ComboBox CmbDivDir 
            Height          =   315
            ItemData        =   "FrmFinQuerie.frx":0608
            Left            =   5040
            List            =   "FrmFinQuerie.frx":060A
            Style = 2
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
         Begin VB.ListBox LstDocDir 
            Height          =   1230
            ItemData        =   "FrmFinQuerie.frx":060C
            Left            =   5040
            List            =   "FrmFinQuerie.frx":0613
            TabIndex        =   4
            Top             =   810
            Width           =   2775
         End
         Begin VB.TextBox TxtDocDir 
            MaxLength = 14
            Height          =   315
            Left            =   6000
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
   Begin VB.Label LabelDegrade1
      Left=120
      Top=240
      Width=3975
      Height=255
      Caption="ANULACIÓN DIRECTA"
      BackStyle=0
      BackStyle=1
      BackColor=&H0083443B&
      ForeColor=&H00FFFFFF&

   End
         Begin Clever.Boton CmdIngdir 
            Height          =   375
            Left            =   7920
            TabIndex        =   7
            Top             =   1680
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            BTYPE           =   6
            TX              =   "Finalizar Querie"
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
            BCOLO           =   6185064
            FCOL            =   16777215
            FCOLO           =   8454143
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmFinQuerie.frx":0622
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
            Height          =   375
            Left            =   7920
            TabIndex        =   8
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            BTYPE           =   6
            TX              =   "Limpiar Datos"
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
            BCOLO           =   6185064
            FCOL            =   16777215
            FCOLO           =   8454143
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmFinQuerie.frx":063E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   1935
            Left            =   4920
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label5 
            Caption         =   "Querie"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   615
         End
         Begin VB.Label LblDirQue 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   16
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Factura"
            Height          =   255
            Left            =   2520
            TabIndex        =   15
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   615
         End
         Begin VB.Label LblDirClient 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   13
            Top             =   960
            Width           =   3975
         End
         Begin VB.Label Label11 
            Caption         =   "Caso"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label LblDirCaso 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   11
            Top             =   1320
            Width           =   3975
         End
         Begin VB.Label LblDirFac 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label31 
            WordWrap = -1
            BackStyle = 0
            Caption = "Enter: agregar reemplazo. Supr: retirar de la lista."
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor = 0
            Height          =   855
            Left            =   7920
            TabIndex        =   9
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Shape Shape3 
         Height          =   495
         Left            =   120
         Top             =   3360
         Width           =   9975
      End
      Begin VB.Label Label8 
         Caption         =   "Anulado Directo"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   46
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Anulado Con NC"
         Height          =   255
         Left            =   2640
         TabIndex        =   45
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   44
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Descuento Parcial"
         Height          =   255
         Left            =   5160
         TabIndex        =   43
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6720
         TabIndex        =   42
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Descuento al 100%"
         Height          =   255
         Left            =   7920
         TabIndex        =   41
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9360
         TabIndex        =   40
         Top             =   3480
         Width           =   615
      End
   Begin VB.Label LblSolucion
      Left=120
      Top=4005
      Width=8200
      Height=285
      Caption="Solución confirmada por el sistema origen (obligatoria):"
   End
   Begin VB.TextBox TxtResolucion
      Left=120
      Top=4320
      Width=9975
      Height=675
      MaxLength=1000
      MultiLine=-1
      ScrollBars=2
   End
         Begin Clever.Boton CmdActualizar 
            Height=330
            Left=8640
            TabIndex        =   25
            Top=3855
            Width=1455
            _ExtentX        =   3625
            _ExtentY        =   661
            BTYPE           =   6
            TX="Actualizar"
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
            BCOLO           =   6185064
            FCOL            =   16777215
            FCOLO           =   8454143
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmFinQuerie.frx":05D0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
   Begin VB.Label LblBuscar
      Left=120
      Top=240
      Width=1800
      Height=255
      Caption="Solicitud / doc. / cliente:"
   End
   Begin VB.TextBox TxtBuscar
      Left=2040
      Top=195
      Width=8055
      Height=315
      MaxLength=100
   End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   48
      Top             =   6405
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/05/2010"
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
End
Attribute VB_Name = "FrmFinQuerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Datos As ADODB.Recordset, Actual As ADODB.Recordset, Nota As ADODB.Recordset
Private Reemplazos As Object
Private IdActual As Long, Seleccionando As Boolean, Iniciando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(7) Then Unload Me: Exit Sub
 StatusBar1.Style = sbrSimple
 Iniciando = True
 GrillaQuery GridPend, "Solicitud|Fecha|Tipo|Documento|Cliente|Solicitud de|Caso|Moneda|Solicitado"
 GridPend.ColWidth(0) = 1300: GridPend.ColWidth(1) = 1200: GridPend.ColWidth(2) = 600: GridPend.ColWidth(3) = 1800
 GridPend.ColWidth(4) = 3400: GridPend.ColWidth(5) = 1300: GridPend.ColWidth(6) = 3000: GridPend.ColWidth(7) = 1100: GridPend.ColWidth(8) = 1300
 CmbDivNc.Clear: CmbDivNc.AddItem "FT": CmbDivNc.AddItem "BV": CmbDivNc.ListIndex = 0
 CmbDivDir.Clear: CmbDivDir.AddItem "FT": CmbDivDir.AddItem "BV": CmbDivDir.ListIndex = 0
 Set Reemplazos = CreateObject("Scripting.Dictionary")
 FraNc.Visible = False: FraDir.Visible = False
 Iniciando = False: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
 Unload Me
End Sub
Private Sub Actualizar()
 Dim R As ADODB.Recordset, I As Integer, N(0 To 3) As Long
 On Error GoTo Fallo
 Set R = ListaQueries(7, Trim$(TxtBuscar.Text), "Abierto")
 Iniciando = True: IdActual = 0: GridPend.Rows = 1
 CerrarRecordset Actual: CerrarRecordset Nota: Reemplazos.RemoveAll
 FraNc.Visible = False: FraDir.Visible = False: TxtResolucion.Text = "": TxtResolucion.Enabled = False
 CerrarRecordset Datos: Set Datos = R
 Do While Not R.EOF
  GridPend.AddItem CStr(R!Asient) & vbTab & FechaDocumento(R!FecSol) & vbTab & CStr(R!TipoDocumento) & vbTab & CStr(R!NumDoc) & vbTab & CeldaDocumento(R!CliDoc) & vbTab & CStr(R!TipQue) & vbTab & CStr(R!TipCas) & vbTab & CStr(R!DivSol) & vbTab & Format$(R!SolDes, "#,##0.00")
  GridPend.RowData(GridPend.Rows - 1) = CLng(R!Item)
  Select Case CStr(R!TipQue)
   Case "ANL ANL": N(0) = N(0) + 1
   Case "ANL NC": N(1) = N(1) + 1
   Case "DES NC": N(2) = N(2) + 1
   Case "DES 100": N(3) = N(3) + 1
  End Select
  R.MoveNext
 Loop
 Label10.Caption = CStr(N(0)): Label13.Caption = CStr(N(1)): Label15.Caption = CStr(N(2)): Label17.Caption = CStr(N(3))
 Iniciando = False
 StatusBar1.SimpleText = CStr(R.RecordCount) & " pendientes visibles (máximo 200). Doble clic: ver solicitud."
 If GridPend.Rows > 1 Then GridPend.Row = 1: Seleccionar
 Exit Sub
Fallo:
 Iniciando = False
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Seleccionar()
 Dim Id As Long
 On Error GoTo Fallo
 If Iniciando Or GridPend.Row < 1 Or GridPend.Rows < 2 Then Exit Sub
 SeleccionFilaQuery GridPend, Seleccionando
 Id = GridPend.RowData(GridPend.Row)
 If Id = IdActual Then Exit Sub
 CerrarRecordset Actual: Set Actual = DatosQuery(Id, 0, 7)
 If Actual.EOF Then Err.Raise 5, , "La solicitud ya no está disponible. Actualice la lista."
 IdActual = Id: LimpiarResolucion
 FraDir.Visible = (CStr(Actual!TipQue) = "ANL ANL"): FraNc.Visible = Not FraDir.Visible
 TxtResolucion.Enabled = True
 LblDirQue.Caption = CStr(Actual!Asient): LblDirFac.Caption = CStr(Actual!NumDoc): LblDirClient.Caption = CStr(Actual!CliDoc): LblDirCaso.Caption = CStr(Actual!TipCas)
 LblNcQue.Caption = CStr(Actual!Asient): LblNcFac.Caption = CStr(Actual!NumDoc): LblTipNc.Caption = CStr(Actual!TipQue)
 LblMonNc.Caption = Format$(Actual!SolDes, "#,##0.00"): LblDivNc.Caption = CStr(Actual!DivSol)
 Exit Sub
Fallo:
 IdActual = 0: FraNc.Visible = False: FraDir.Visible = False: TxtResolucion.Enabled = False
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub LimpiarResolucion()
 CerrarRecordset Nota
 Reemplazos.RemoveAll: LstDocNc.Clear: LstDocDir.Clear
 TxtNc.Text = "": TxtDocNc.Text = "": TxtDocDir.Text = "": TxtAplicado.Text = "0": TxtResolucion.Text = "": LblExcNc.Caption = ""
End Sub
Private Sub CmdActualizar_Click()
 If IdActual > 0 And (Len(TxtResolucion.Text) > 0 Or Reemplazos.Count > 0 Or Not Nota Is Nothing) Then
  If MsgBox("¿Descartar los datos de cierre y actualizar los pendientes?", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 End If
 Actualizar
End Sub
Private Sub GridPend_Click()
 Seleccionar
End Sub
Private Sub TxtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyF9 Then CmdActualizar_Click: KeyCode = 0
End Sub
Private Sub GridPend_RowColChange()
 If Not Seleccionando Then Seleccionar
End Sub
Private Sub GridPend_DblClick()
 If IdActual > 0 Then VerQuery IdActual, 7
End Sub
Private Sub GridPend_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Seleccionar: KeyCode = 0
End Sub
Private Sub TxtNc_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF9 Then CmdBus_Click: KeyCode = 0
 If KeyCode = vbKeyReturn Then CargarNota: KeyCode = 0
End Sub
Private Sub TxtNc_Change()
 If Not Nota Is Nothing Then
  If Trim$(TxtNc.Text) <> CStr(Nota!Documento) Then CerrarRecordset Nota
 End If
End Sub
Private Sub CmdBus_Click()
 On Error GoTo Fallo
 If IdActual = 0 Then Err.Raise 5, , "Seleccione una solicitud pendiente."
 Load FrmBuscarNCQuery
 FrmBuscarNCQuery.Preparar IdActual, Trim$(TxtNc.Text)
 FrmBuscarNCQuery.Show vbModal, MdiPri
 If FrmBuscarNCQuery.Aceptado Then
  TxtNc.Text = FrmBuscarNCQuery.NumeroSeleccionado
  Unload FrmBuscarNCQuery: CargarNota
 Else
  Unload FrmBuscarNCQuery
 End If
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CargarNota()
 Dim R As ADODB.Recordset, Encontrada As Boolean
 On Error GoTo Fallo
 If IdActual = 0 Then Err.Raise 5, , "Seleccione una solicitud."
 CerrarRecordset Nota
 Set R = NotasQuery(IdActual, Trim$(TxtNc.Text))
 Do While Not R.EOF
  If CStr(R!Documento) = Trim$(TxtNc.Text) Then Encontrada = True: Exit Do
  R.MoveNext
 Loop
 If Not Encontrada Then Err.Raise 5, , "No existe una NC disponible con ese número, cliente, moneda y referencia. Use F9 para buscar."
 Set Nota = R
 TxtAplicado.Text = Format$(Actual!SolDes, "0.00")
 If CCur(Nota!Importe) < CCur(Actual!SolDes) Then TxtAplicado.Text = Format$(Nota!Importe, "0.00")
 CalcularAplicado
 Exit Sub
Fallo:
 CerrarRecordset R
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub TxtAplicado_Change()
 CalcularAplicado
End Sub
Private Sub CalcularAplicado()
 If Nota Is Nothing Then LblExcNc.Caption = "": Exit Sub
 If IsNumeric(TxtAplicado.Text) Then LblExcNc.Caption = Format$(CCur(Nota!Importe) - CCur(TxtAplicado.Text), "#,##0.00") Else LblExcNc.Caption = ""
End Sub
Private Sub TxtDocNc_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then AgregarReemplazo False: KeyCode = 0
End Sub
Private Sub TxtDocDir_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then AgregarReemplazo True: KeyCode = 0
End Sub
Private Sub AgregarReemplazo(ByVal Directo As Boolean)
 Dim R As ADODB.Recordset, Tipo As String, Numero As String, L As ListBox, Id As Long, Mensaje As String
 On Error GoTo Fallo
 If IdActual = 0 Then Err.Raise 5, , "Seleccione una solicitud."
 If Directo Then Tipo = CmbDivDir.Text: Numero = Trim$(TxtDocDir.Text): Set L = LstDocDir Else Tipo = CmbDivNc.Text: Numero = Trim$(TxtDocNc.Text): Set L = LstDocNc
 If Len(Numero) = 0 Then Exit Sub
 Set R = DocumentoQuery(Tipo, Numero, 7)
 If R.EOF Then Err.Raise 5, , "No se encontró el documento de reemplazo importado."
 Id = CLng(R!IdDocumento)
 If Id = CLng(Actual!IdFactura) Or CBool(R!Anulado) Or CLng(R!IdCliente) <> CLng(Actual!IdCliente) Or CStr(R!Divisa) <> CStr(Actual!DivSol) Then Err.Raise 5, , "El reemplazo debe ser otro documento vigente del mismo cliente y moneda."
 If Reemplazos.Exists(CStr(Id)) Then Err.Raise 5, , "El documento ya está en la lista."
 If Reemplazos.Count >= 200 Then Err.Raise 5, , "Máximo 200 documentos de reemplazo."
 Reemplazos.Add CStr(Id), VersionXML(R.Fields("VersionFila").Value)
 L.AddItem Tipo & " " & Numero: L.ItemData(L.NewIndex) = Id
 If Directo Then TxtDocDir.Text = "" Else TxtDocNc.Text = ""
 CerrarRecordset R
 Exit Sub
Fallo:
 Mensaje = Err.Description: CerrarRecordset R
 MsgBox Mensaje, vbExclamation, App.Title
End Sub
Private Sub Retirar(ByVal L As ListBox, ByRef KeyCode As Integer)
 If KeyCode <> vbKeyDelete Or L.ListIndex < 0 Then Exit Sub
 Reemplazos.Remove CStr(L.ItemData(L.ListIndex)): L.RemoveItem L.ListIndex: KeyCode = 0
End Sub
Private Sub LstDocNc_KeyDown(KeyCode As Integer, Shift As Integer)
 Retirar LstDocNc, KeyCode
End Sub
Private Sub LstDocDir_KeyDown(KeyCode As Integer, Shift As Integer)
 Retirar LstDocDir, KeyCode
End Sub
Private Sub CmdLimNc_Click()
 LimpiarResolucion
End Sub
Private Sub CmdLim_Click()
 LimpiarResolucion
End Sub
Private Sub CmdFinQueNc_Click()
 Finalizar False
End Sub
Private Sub CmdIngdir_Click()
 Finalizar True
End Sub
Private Sub Finalizar(ByVal Directo As Boolean)
 Dim XML As String, Id As Variant, IdNota As Long, V As Variant, Monto As Currency, Guardado As Boolean, Mensaje As String
 On Error GoTo Fallo
 If IdActual = 0 Then Err.Raise 5, , "Seleccione una solicitud pendiente."
 If Len(Trim$(TxtResolucion.Text)) = 0 Then Err.Raise 5, , "Describa la solución confirmada por el sistema origen."
 V = Null
 If Not Directo Then
  If Nota Is Nothing Then Err.Raise 5, , "Seleccione y cargue la nota de crédito (Enter o F9)."
  If Not IsNumeric(TxtAplicado.Text) Then Err.Raise 5, , "Ingrese un importe aplicado válido."
  IdNota = CLng(Nota!IdDocumento): V = Nota.Fields("VersionFila").Value: Monto = CCur(TxtAplicado.Text)
 End If
 XML = "<documentos>"
 For Each Id In Reemplazos.Keys
  XML = XML & "<d" & AtributoXML("Id", CStr(Id)) & AtributoXML("Version", CStr(Reemplazos(Id))) & "/>"
 Next Id
 XML = XML & "</documentos>"
 If MsgBox("¿Finalizar " & CStr(Actual!Asient) & " con esta solución? El expediente quedará cerrado y conservará sus documentos asociados.", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 CerrarSolicitudQuery IdActual, Actual.Fields("VersionFila").Value, IdNota, V, Monto, Trim$(TxtResolucion.Text), XML
 Guardado = True: Actualizar
 MsgBox "Solución registrada. La solicitud pasó al historial de cerradas.", vbInformation, App.Title
 Exit Sub
Fallo:
 Mensaje = Err.Description
 If Guardado Then Mensaje = "El cierre quedó registrado. Recargue los pendientes. " & Mensaje
 MsgBox Mensaje, vbExclamation, App.Title
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos: CerrarRecordset Actual: CerrarRecordset Nota
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

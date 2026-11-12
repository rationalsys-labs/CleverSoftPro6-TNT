VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{DA729E34-689F-49EA-A856-B57046630B73}#1.0#0"; "progressbar-xp.ocx"
Begin VB.Form FrmDatQuerie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de Solicitudes Querie"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   12720
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2175
      Begin VB.Frame Frame7 
         Height          =   375
         Left            =   1320
         TabIndex        =   43
         Top             =   4950
         Width           =   615
   Begin VB.Label LBLNULL
      Left=10
      Top=120
      Width=580
      Height=285
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame6 
         Height          =   375
         Left            =   1320
         TabIndex        =   41
         Top             =   4590
         Width           =   615
   Begin VB.Label LBLACT
      Left=10
      Top=120
      Width=580
      Height=285
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame5 
         Height          =   375
         Left            =   1320
         TabIndex        =   39
         Top             =   4230
         Width           =   615
   Begin VB.Label LBLDES
      Left=10
      Top=120
      Width=580
      Height=285
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame4 
         Height          =   375
         Left            =   1320
         TabIndex        =   37
         Top             =   3870
         Width           =   615
   Begin VB.Label LBLNC
      Left=10
      Top=120
      Width=580
      Height=285
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame1 
         Height          =   375
         Left            =   1320
         TabIndex        =   35
         Top             =   3510
         Width           =   615
   Begin VB.Label LblAnl
      Left=10
      Top=120
      Width=580
      Height=285
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame3 
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Top             =   3150
         Width           =   615
   Begin VB.Label LblQuerie
      Left=10
      Top=120
      Width=580
      Height=285
      Caption="0"
      BackStyle=0

   End
      End
      Begin Clever.Boton CmdInfoDoc 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   9
         TX = "Exportar lista a Excel"
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":0234
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Anular Querie"
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":0250
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdGenerar 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   5760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   9
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":026C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdFil 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Activar Filtros"
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":0288
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdMas 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Regenerar en Masivo"
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":02A4
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
      Top=3240
      Width=855
      Height=255
      Caption="Visibles"
      BackStyle=0

   End
   Begin VB.Label SkinLabel17
      Left=240
      Top=3600
      Width=855
      Height=255
      Caption="Anul. directa"
      BackStyle=0

   End
   Begin VB.Label SkinLabel18
      Left=240
      Top=3960
      Width=855
      Height=255
      Caption="Anul. NC"
      BackStyle=0

   End
   Begin VB.Label SkinLabel19
      Left=240
      Top=4320
      Width=855
      Height=255
      Caption="Descuentos"
      BackStyle=0

   End
   Begin VB.Label SkinLabel20
      Left=240
      Top=4680
      Width=855
      Height=255
      Caption="Abiertas"
      BackStyle=0

   End
   Begin VB.Label SkinLabel21
      Left=240
      Top=5040
      Width=855
      Height=255
      Caption="Anuladas"
      BackStyle=0

   End
      Begin Clever.Boton CmdConf 
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   1920
         Visible = -1
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   9
         TX = "Solución e historial"
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":0554
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdActualizar 
         Height=375
         Left=120
         TabIndex        =   3
         Top=2430
         Width=1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   9
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":0250
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6330
      Width           =   12720
      _ExtentX        =   22437
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
   Begin VB.Frame FraFil 
      Height = 1170
      Left            =   2400
      TabIndex        =   7
      Top             =   0
      Visible = 0
      Width           =   10215
   Begin VB.Label SkinLabel2
      Left=4800
      Top=240
      Width=615
      Height=255
      Caption="Desde:"
      BackStyle=0

   End
      Begin VB.ComboBox CmbQuer 
         Style = 2
         Height          =   315
         ItemData        =   "FrmDatQuerie.frx":05DA
         Left            =   3000
         List            =   "FrmDatQuerie.frx":05EA
         TabIndex        =   9
         Text            =   "Registros Totales"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox CmbOrigenes 
         Style = 2
         Height          =   315
         ItemData        =   "FrmDatQuerie.frx":062F
         Left            =   120
         List            =   "FrmDatQuerie.frx":0631
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
      Begin Clever.Boton CmdCargReg 
         Height          =   315
         Left            =   8640
         TabIndex        =   12
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   6
         TX = "Consultar"
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":0633
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdXFil 
         Height          =   315
         Left            =   9720
         TabIndex        =   13
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   3
         TX = "X"
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
         MICON           =   "FrmDatQuerie.frx":064F
         PICN            =   "FrmDatQuerie.frx":066B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   Begin VB.Label SkinLabel3
      Left=6720
      Top=240
      Width=615
      Height=255
      Caption="Hasta:"
      BackStyle=0

   End
   Begin VB.TextBox TxtDesde
      Left=5400
      Top=240
      Width=1215
      Height=315
      MaxLength=10
   End
   Begin VB.TextBox TxtHasta
      Left=7320
      Top=240
      Width=1215
      Height=315
      MaxLength=10
   End
   Begin VB.Label LblBuscar
      Left=120
      Top=735
      Width=1800
      Height=255
      Caption="Solicitud / doc. / cliente:"
   End
   Begin VB.TextBox TxtBuscar
      Left=2040
      Top=690
      Width=6495
      Height=315
      MaxLength=100
   End
   End
   Begin VB.Frame FraMas 
      Height          =   2175
      Left            =   2400
      TabIndex        =   14
      Top             =   4080
      Visible = 0
      Width           =   10215
      Begin Clever.Boton CmdGenRep 
         Height          =   315
         Left            =   8760
         TabIndex        =   15
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BTYPE           =   6
         TX              =   "Generar"
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":0D2F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdLimMas 
         Height          =   315
         Left            =   8760
         TabIndex        =   16
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BTYPE           =   6
         TX              =   "Limpiar"
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
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDatQuerie.frx":0D4B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame8 
         Height          =   1935
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   8535
         Begin MSFlexGridLib.MSFlexGrid GridGenMas 
            Height          =   1785
            Left            =   15
            TabIndex        =   46
            Top             =   120
            Width           =   8495
            _ExtentX        =   14975
            _ExtentY        =   3149
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
   End
   Begin VB.Frame Frame9 
      Height          =   6255
      Left            =   2400
      TabIndex        =   47
      Top             =   0
      Width           =   10215
      Begin MSFlexGridLib.MSFlexGrid GridQuerie 
         Height          =   6105
         Left            =   15
         TabIndex        =   48
         Top             =   120
         Width           =   10170
         _ExtentX        =   17939
         _ExtentY        =   10769
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
End
Attribute VB_Name = "FrmDatQuerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Datos As ADODB.Recordset
Private Banda As Collection
Private Seleccionando As Boolean, Iniciando As Boolean
Private Sub Form_Load()
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If Not AccesoPermitido(6) Then Unload Me: Exit Sub
 StatusBar1.Style = sbrSimple
 Iniciando = True
 Set Banda = New Collection
 GrillaQuery GridQuerie, "Solicitud|Fecha|Tipo|Documento|Cliente|Solicitud de|Caso|Moneda|Importe|Estado|Creado por"
 GridQuerie.ColWidth(0) = 1300: GridQuerie.ColWidth(1) = 1200: GridQuerie.ColWidth(2) = 600
 GridQuerie.ColWidth(3) = 1800: GridQuerie.ColWidth(4) = 3500: GridQuerie.ColWidth(5) = 1300: GridQuerie.ColWidth(6) = 3100
 GrillaQuery GridGenMas, "Solicitud|Documento|Cliente": GridGenMas.ColWidth(0) = 1500: GridGenMas.ColWidth(1) = 1800: GridGenMas.ColWidth(2) = 4800
 CmbQuer.Clear: CmbQuer.AddItem "Todas": CmbQuer.AddItem "Abierto": CmbQuer.AddItem "Cerrado": CmbQuer.AddItem "Anuladas": CmbQuer.ListIndex = 0
 CmbOrigenes.Clear: CmbOrigenes.AddItem "Todos los usuarios": CmbOrigenes.ItemData(0) = 0
 Set R = CatalogoQuery(4, 6)
 Do While Not R.EOF
  CmbOrigenes.AddItem CStr(R!Descripcion): CmbOrigenes.ItemData(CmbOrigenes.NewIndex) = CLng(R!Item): R.MoveNext
 Loop
 CerrarRecordset R: CmbOrigenes.ListIndex = 0
 FraFil.Visible = False: FraMas.Visible = False: AjustarLista
 Iniciando = False: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
 Unload Me
End Sub
Private Sub AjustarLista()
 Frame9.Top = IIf(FraFil.Visible, 1200, 0)
 Frame9.Height = 6255 - Frame9.Top
 If FraMas.Visible Then Frame9.Height = 4080 - Frame9.Top
 GridQuerie.Height = Frame9.Height - 150
 If FraFil.Visible Then FraFil.ZOrder 0
 If FraMas.Visible Then FraMas.ZOrder 0
End Sub
Private Sub Actualizar()
 Dim D As Variant, H As Variant, F As Date, Estado As String, R As ADODB.Recordset
 Dim Anl As Long, NC As Long, Des As Long, Abiertas As Long, Nulas As Long
 On Error GoTo Fallo
 D = Null: H = Null
 If FraFil.Visible Then
  If Len(Trim$(TxtDesde.Text)) > 0 Then
   If Not FechaGestionValida(TxtDesde.Text, F) Then Err.Raise 5, , "Fecha Desde inválida. Use dd/mm/aaaa."
   D = F
  End If
  If Len(Trim$(TxtHasta.Text)) > 0 Then
   If Not FechaGestionValida(TxtHasta.Text, F) Then Err.Raise 5, , "Fecha Hasta inválida. Use dd/mm/aaaa."
   H = F
  End If
  If CmbQuer.ListIndex > 0 Then Estado = CmbQuer.Text
  Set R = ListaQueries(6, Trim$(TxtBuscar.Text), Estado, D, H, CmbOrigenes.ItemData(CmbOrigenes.ListIndex))
 Else
  Set R = ListaQueries(6)
 End If
 CerrarRecordset Datos: Set Datos = R
 GridQuerie.Rows = 1
 Do While Not R.EOF
  GridQuerie.AddItem CeldaDocumento(R!Asient) & vbTab & FechaDocumento(R!FecSol) & vbTab & CeldaDocumento(R!TipoDocumento) & vbTab & CeldaDocumento(R!NumDoc) & vbTab & CeldaDocumento(R!CliDoc) & vbTab & CeldaDocumento(R!TipQue) & vbTab & CeldaDocumento(R!TipCas) & vbTab & CeldaDocumento(R!DivSol) & vbTab & Format$(R!MonSol, "#,##0.00") & vbTab & IIf(CStr(R!SitDoc) = "ANL", "Anulada", CStr(R!Estado)) & vbTab & CeldaDocumento(R!Creador)
  GridQuerie.RowData(GridQuerie.Rows - 1) = CLng(R!Item)
  If CStr(R!SitDoc) = "ANL" Then
   Nulas = Nulas + 1
  Else
   If CStr(R!Estado) = "Abierto" Then Abiertas = Abiertas + 1
   Select Case CStr(R!TipQue)
    Case "ANL ANL": Anl = Anl + 1
    Case "ANL NC": NC = NC + 1
    Case Else: Des = Des + 1
   End Select
  End If
  R.MoveNext
 Loop
 LblQuerie.Caption = CStr(R.RecordCount): LblAnl.Caption = CStr(Anl): LBLNC.Caption = CStr(NC): LBLDES.Caption = CStr(Des): LBLACT.Caption = CStr(Abiertas): LBLNULL.Caption = CStr(Nulas)
 StatusBar1.SimpleText = CStr(R.RecordCount) & " solicitudes visibles (máximo 200). Use filtros para buscar en el historial."
 If GridQuerie.Rows > 1 Then GridQuerie.Row = 1: SeleccionFilaQuery GridQuerie, Seleccionando
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Function Seleccionada() As Long
 If GridQuerie.Row < 1 Or GridQuerie.Rows < 2 Then Err.Raise 5, , "Seleccione una solicitud."
 Seleccionada = GridQuerie.RowData(GridQuerie.Row)
 Datos.MoveFirst: Datos.Find "Item=" & CStr(Seleccionada)
 If Datos.EOF Then Err.Raise 5, , "Recargue la lista de solicitudes."
End Function
Private Sub CmdActualizar_Click()
 Actualizar
End Sub
Private Sub CmdCargReg_Click()
 Actualizar
End Sub
Private Sub CmdFil_Click()
 FraFil.Visible = Not FraFil.Visible: AjustarLista
 If Not FraFil.Visible Then Actualizar
End Sub
Private Sub CmdXFil_Click()
 FraFil.Visible = False: AjustarLista: Actualizar
End Sub
Private Sub TxtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyF9 Then Actualizar: KeyCode = 0
End Sub
Private Sub CmdGenerar_Click()
 On Error GoTo Fallo
 VerQuery Seleccionada
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdConf_Click()
 On Error GoTo Fallo
 VerHistorialQuery Seleccionada
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdAnl_Click()
 Dim Id As Long, Motivo As String, Guardado As Boolean
 On Error GoTo Fallo
 Id = Seleccionada
 If CStr(Datos!SitDoc) <> "ACT" Or CStr(Datos!Estado) <> "Abierto" Then Err.Raise 5, , "Solo se pueden anular solicitudes abiertas y vigentes."
 Load FrmMotivoQuery
 FrmMotivoQuery.Aceptado = False: FrmMotivoQuery.TxtMotivo.Text = ""
 FrmMotivoQuery.Show vbModal, MdiPri
 If Not FrmMotivoQuery.Aceptado Then Unload FrmMotivoQuery: Exit Sub
 Motivo = Trim$(FrmMotivoQuery.TxtMotivo.Text): Unload FrmMotivoQuery
 If MsgBox("¿Anular la solicitud " & CStr(Datos!Asient) & "? La factura y sus importes pertenecen al sistema origen.", vbExclamation + vbYesNo, App.Title) <> vbYes Then Exit Sub
 AnularSolicitudQuery Id, Datos.Fields("VersionFila").Value, Motivo
 Guardado = True: Actualizar
 Exit Sub
Fallo:
 If Guardado Then MsgBox "La solicitud quedó anulada. Recargue la lista.", vbInformation, App.Title Else MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdInfoDoc_Click()
 On Error GoTo Fallo
 ExportarQueries Datos
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdMas_Click()
 FraMas.Visible = True: AjustarLista: AgregarBanda
End Sub
Private Sub AgregarBanda()
 Dim Id As Long, V As Variant
 On Error GoTo Fallo
 Id = Seleccionada
 For Each V In Banda
  If CLng(V) = Id Then Exit Sub
 Next V
 If Banda.Count >= 200 Then Err.Raise 5, , "Máximo 200 solicitudes por reporte."
 Banda.Add Id
 GridGenMas.AddItem CStr(Datos!Asient) & vbTab & CStr(Datos!NumDoc) & vbTab & CeldaDocumento(Datos!CliDoc)
 GridGenMas.RowData(GridGenMas.Rows - 1) = Id
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdLimMas_Click()
 Set Banda = New Collection: GridGenMas.Rows = 1: FraMas.Visible = False: AjustarLista
End Sub
Private Sub CmdGenRep_Click()
 MostrarQueries Banda
End Sub
Private Sub GridQuerie_DblClick()
 If FraMas.Visible Then AgregarBanda Else CmdGenerar_Click
End Sub
Private Sub GridQuerie_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then GridQuerie_DblClick: KeyCode = 0
End Sub
Private Sub GridQuerie_Click()
 SeleccionFilaQuery GridQuerie, Seleccionando
End Sub
Private Sub GridQuerie_RowColChange()
 SeleccionFilaQuery GridQuerie, Seleccionando
End Sub
Private Sub GridGenMas_Click()
 SeleccionFilaQuery GridGenMas, Seleccionando
End Sub
Private Sub GridGenMas_RowColChange()
 SeleccionFilaQuery GridGenMas, Seleccionando
End Sub
Private Sub GridGenMas_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim I As Long
 If KeyCode <> vbKeyDelete Or GridGenMas.Row < 1 Then Exit Sub
 I = GridGenMas.Row: Banda.Remove I: GridGenMas.RemoveItem I: KeyCode = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

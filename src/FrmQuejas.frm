VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmQuejas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modulo de Quejas y Reclamos"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   13215
   Begin VB.Frame Frame1 
      Caption         =   "Quejas y Reclamos"
      Height          =   6135
      Left            =   2400
      TabIndex        =   12
      Top             =   0
      Width           =   10695
      Begin MSFlexGridLib.MSFlexGrid GridQuejas 
         Height          =   5415
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9551
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
      Begin VB.TextBox TxtBus 
         Appearance      =   0  'Flat
         Height          =   285
         Left = 3600
         TabIndex        =   14
         Top             =   240
         Width = 6750
      End
      Begin VB.ComboBox CmbBus 
         Style = 2
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Text            =   "Nombre de Cliente"
         Top             =   240
         Width = 3300
      End
   End
   Begin VB.Frame Frame2 
      BackColor = &H8000000F&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin Clever.Boton CmdCheque 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX = "Cliente y nuevo caso"
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
         MICON           =   "FrmQuejas.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdFactura 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Nuevo Caso"
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
         MICON           =   "FrmQuejas.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdConf 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Configuracion General"
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
         MICON           =   "FrmQuejas.frx":0038
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
         Left            =   120
         TabIndex        =   5
         Top             =   1680
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
         MICON           =   "FrmQuejas.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label PaseCas 
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PaseCod 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblCli 
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   5400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblCod 
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   5640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PaseNew 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
      ForeColor=&H0083693C&
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1935
      End
   Begin Clever.Boton CmdActualizar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=2160
      Width=1935
      Height=375
      TX="Actualizar"
   End
   Begin Clever.Boton CmdCasos
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=2580
      Width=1935
      Height=375
      TX="Ver casos"
   End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6285
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "23/06/2010"
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
Attribute VB_Name = "FrmQuejas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Datos As ADODB.Recordset, Seleccionando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(12) Then Unload Me: Exit Sub
 CmbBus.Clear: CmbBus.AddItem "Cliente / código / RUC": CmbBus.ListIndex = 0
 StatusBar1.Style = sbrSimple
 PaseNew.Visible = False: PaseCod.Visible = False: PaseCas.Visible = False: LblCli.Visible = False: LblCod.Visible = False
 Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title: Unload Me
End Sub
Public Sub Actualizar()
 On Error GoTo Fallo
 CerrarRecordset Datos: Set Datos = CobLeer(12, 0, 0, 0, Trim$(TxtBus.Text))
 CobGrilla GridQuejas, Datos, "Código|Cliente|RUC / DNI|Apertura|Usuario|Casos|Abiertos", "Asient|CLIENT|RUCDNI|FecApe|Usuario|Casos|Abiertos", "IdCliente"
 GridQuejas.ColWidth(0) = 1500: GridQuejas.ColWidth(1) = 4500: GridQuejas.ColWidth(2) = 1600
 StatusBar1.SimpleText = CStr(Datos.RecordCount) & " clientes en seguimiento (máximo 200). Enter: casos del cliente."
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdCheque_Click()
 Dim C As ADODB.Recordset
 On Error GoTo Fallo
 Set C = CobBuscarCliente(12, TxtBus.Text, Me)
 If C Is Nothing Then Exit Sub
 Load FrmQuejaCaso: FrmQuejaCaso.Preparar CLng(C!Item), CStr(C!CLIENT)
 FrmQuejaCaso.Show vbModal, MdiPri: Unload FrmQuejaCaso
 CerrarRecordset C: Actualizar
 Exit Sub
Fallo:
 CerrarRecordset C: MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdFactura_Click()
 Dim Id As Long
 On Error GoTo Fallo
 Id = CobFila(Datos, GridQuejas, "IdCliente")
 Load FrmQuejaCaso: FrmQuejaCaso.Preparar Id, CStr(Datos!CLIENT)
 FrmQuejaCaso.Show vbModal, MdiPri: Unload FrmQuejaCaso: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdCasos_Click()
 Dim Id As Long
 On Error GoTo Fallo
 Id = CobFila(Datos, GridQuejas, "IdCliente")
 Load FrmQuejasCasos: FrmQuejasCasos.Preparar Id, CStr(Datos!CLIENT)
 FrmQuejasCasos.Show vbModal, MdiPri: Unload FrmQuejasCasos: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdConf_Click()
 FrmQuejasConfig.Show vbModal, MdiPri
End Sub
Private Sub CmdExportar_Click()
 On Error GoTo Fallo
 CobExportar Datos, "Seguimiento de quejas por cliente", "Asient|CLIENT|RUCDNI|FecApe|Usuario|Casos|Abiertos"
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdActualizar_Click()
 Actualizar
End Sub
Private Sub TxtBus_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyF9 Then Actualizar: KeyCode = 0
End Sub
Private Sub GridQuejas_Click()
 SeleccionFilaQuery GridQuejas, Seleccionando
End Sub
Private Sub GridQuejas_RowColChange()
 SeleccionFilaQuery GridQuejas, Seleccionando
End Sub
Private Sub GridQuejas_DblClick()
 CmdCasos_Click
End Sub
Private Sub GridQuejas_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then CmdCasos_Click: KeyCode = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

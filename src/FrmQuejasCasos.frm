VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmQuejasCasos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quejas por Cliente"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   13455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Quejas por Cliente"
      Height          =   6135
      Left            =   2400
      TabIndex        =   7
      Top             =   0
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid GridCasos 
         Height          =   5775
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   10186
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
   End
   Begin Clever.Boton CmdAceptar 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   5640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      FOCUSR          =   -1  'True
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmQuejasCasos.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   120
      TabIndex        =   2
      Top             =   120
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
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Crear Caso"
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
         MICON           =   "FrmQuejasCasos.frx":001C
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
         TabIndex        =   4
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Finalizar Caso"
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
         MICON           =   "FrmQuejasCasos.frx":0038
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
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX = "Ver historial"
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
         MICON           =   "FrmQuejasCasos.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6270
      Width           =   13455
      _ExtentX        =   23733
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
Attribute VB_Name = "FrmQuejasCasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private IdCliente As Long, Datos As ADODB.Recordset, NombreCliente As String, Seleccionando As Boolean
Public Sub Preparar(ByVal Id As Long, ByVal Nombre As String)
 IdCliente = Id: NombreCliente = Nombre: Me.Caption = "Casos de " & Nombre: Actualizar
End Sub
Private Sub Actualizar()
 On Error GoTo Fallo
 CerrarRecordset Datos: Set Datos = CobLeer(12, 1, 0, IdCliente)
 CobGrilla GridCasos, Datos, "Caso|Apertura|Concepto|Estado|Usuario|Cierre|Cerrado por", "Asient|FecCas|DescripcionCaso|Estado|Usuari|FecFin|UsuFin"
 GridCasos.ColWidth(2) = 3100: StatusBar1.Style = sbrSimple
 StatusBar1.SimpleText = CStr(Datos.RecordCount) & " casos (máximo 200). Enter: historial de gestiones."
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdCrear_Click()
 Load FrmQuejaCaso: FrmQuejaCaso.Preparar IdCliente, NombreCliente
 FrmQuejaCaso.Show vbModal, Me: Unload FrmQuejaCaso: Actualizar
End Sub
Private Sub AbrirGestiones(ByVal Cerrar As Boolean)
 Dim Id As Long
 On Error GoTo Fallo
 Id = CobFila(Datos, GridCasos)
 Load FrmQuejasGestiones: FrmQuejasGestiones.Preparar Id, Cerrar
 FrmQuejasGestiones.Show vbModal, Me: Unload FrmQuejasGestiones: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdFinalizar_Click()
 AbrirGestiones True
End Sub
Private Sub CmdHistorial_Click()
 AbrirGestiones False
End Sub
Private Sub CmdAceptar_Click()
 Unload Me
End Sub
Private Sub GridCasos_Click()
 SeleccionFilaQuery GridCasos, Seleccionando
End Sub
Private Sub GridCasos_RowColChange()
 SeleccionFilaQuery GridCasos, Seleccionando
End Sub
Private Sub GridCasos_DblClick()
 AbrirGestiones False
End Sub
Private Sub GridCasos_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then AbrirGestiones False: KeyCode = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

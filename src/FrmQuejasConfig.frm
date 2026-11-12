VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQuejasConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuraciones"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Administrar Status para Casos"
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   5655
      Begin MSFlexGridLib.MSFlexGrid GridSta 
         Height          =   1815
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3201
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
      Begin VB.TextBox txtdes 
         MaxLength = 50
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtSta 
         MaxLength = 20
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin Clever.Boton mdAgreSta 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQuejasConfig.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Descrip:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Administrar Concepto de Casos"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin MSFlexGridLib.MSFlexGrid GridCas 
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3201
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
      Begin VB.TextBox TxtCas 
         MaxLength = 50
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin Clever.Boton CmdAgreCas 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
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
         FOCUSR          =   -1  'True
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmQuejasConfig.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
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
      Left            =   4080
      TabIndex        =   10
      Top             =   6360
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
      MICON           =   "FrmQuejasConfig.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   6240
      Width           =   6135
   End
   Begin VB.Label LblBaja
      Left=120
      Top=6180
      Width=5200
      Height=255
      Caption="Supr sobre una fila: baja lógica; conserva el historial."
      BackStyle=0
   End
End
Attribute VB_Name = "FrmQuejasConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Casos As ADODB.Recordset, Estados As ADODB.Recordset, Seleccionando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(12) Then Unload Me: Exit Sub
 Cargar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title: Unload Me
End Sub
Private Sub Cargar()
 CerrarRecordset Casos: CerrarRecordset Estados
 Set Casos = CobCatalogo(12, 4): Set Estados = CobCatalogo(12, 5)
 CobGrilla GridCas, Casos, "Código|Descripción|Activo", "Item|Descripcion|Activo"
 GridCas.ColWidth(1) = 3600: GridCas.ColWidth(2) = 900
 CobGrilla GridSta, Estados, "Código|Estado|Descripción|Activo", "Item|Codigo|Descripcion|Activo"
 GridSta.ColWidth(0) = 0: GridSta.ColWidth(1) = 900: GridSta.ColWidth(2) = 3300: GridSta.ColWidth(3) = 900
End Sub
Private Sub Guardar(ByVal Tipo As Integer, ByVal Id As Long, ByVal V As Variant, ByVal Codigo As String, ByVal Descripcion As String, ByVal Baja As Boolean)
 Dim R As ADODB.Recordset
 Set R = ConsultaOperacion("dbo.CobGuardarCatalogo", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Tipo", adInteger, 0, Tipo), Array("@Item", adInteger, 0, Id), Array("@Version", adBinary, 8, V), Array("@Codigo", adVarWChar, 20, Codigo), Array("@Descripcion", adVarWChar, 150, Descripcion), Array("@Baja", adBoolean, 0, Baja)), True)
 CerrarRecordset R: Cargar
End Sub
Private Sub CmdAgreCas_Click()
 On Error GoTo Fallo
 Guardar 4, 0, Null, "", Trim$(TxtCas.Text), False: TxtCas.Text = ""
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub mdAgreSta_Click()
 On Error GoTo Fallo
 Guardar 5, 0, Null, UCase$(Trim$(txtSta.Text)), Trim$(txtdes.Text), False
 txtSta.Text = "": txtdes.Text = ""
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Baja(ByVal Tipo As Integer)
 Dim Id As Long, R As ADODB.Recordset, G As MSFlexGrid
 On Error GoTo Fallo
 If Tipo = 4 Then Set R = Casos: Set G = GridCas Else Set R = Estados: Set G = GridSta
 Id = CobFila(R, G)
 If MsgBox("¿Dar de baja el registro seleccionado? Las gestiones existentes conservarán su referencia.", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 Guardar Tipo, Id, R.Fields("VersionFila").Value, "", CStr(R!Descripcion), True
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub GridCas_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete Then Baja 4: KeyCode = 0
End Sub
Private Sub GridSta_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete Then Baja 5: KeyCode = 0
End Sub
Private Sub GridCas_Click()
 SeleccionFilaQuery GridCas, Seleccionando
End Sub
Private Sub GridCas_RowColChange()
 SeleccionFilaQuery GridCas, Seleccionando
End Sub
Private Sub GridSta_Click()
 SeleccionFilaQuery GridSta, Seleccionando
End Sub
Private Sub GridSta_RowColChange()
 SeleccionFilaQuery GridSta, Seleccionando
End Sub
Private Sub CmdAceptar_Click()
 Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Casos: CerrarRecordset Estados
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

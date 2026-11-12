VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDocsCobranzas 
   BorderStyle     =   1  'Fixed Single
   Caption = "Documentos asociados"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin MSFlexGridLib.MSFlexGrid GridDocs 
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3413
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.TextBox TxtNumero 
         MaxLength = 50
         Appearance      =   0  'Flat
         Height          =   285
         Left=1020
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   Begin VB.ComboBox CmbTipo
      Left=120
      Top=240
      Width=810
      Height=315
      Style=2
   End
   Begin Clever.Boton CmdAgregar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=3600
      Top=210
      Width=2775
      Height=375
      TX="Agregar (Enter)"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   2880
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
      MICON           =   "FrmBuscarNCQuery.frx":0000
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
      Top             =   2760
      Width           =   6735
   End
   Begin VB.ListBox LstAux
      Left=0
      Top=0
      Width=100
      Height=100
      Visible=0
   End
End
Attribute VB_Name = "FrmDocsCobranzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Public Docs As Object, Aceptado As Boolean
Private Modulo As Integer, Cliente As ADODB.Recordset, Moneda As String, Seleccionando As Boolean
Public Sub Preparar(ByVal Codigo As Integer, ByVal DatosCliente As ADODB.Recordset, ByVal Divisa As String, ByVal Original As Object)
 Dim Id As Variant
 Modulo = Codigo: Set Cliente = DatosCliente: Moneda = Divisa: Aceptado = False
 Set Docs = CobNuevosDocumentos
 For Each Id In Original.Keys: Docs.Add CStr(Id), Original(Id): Next Id
 CmbTipo.Clear: CmbTipo.AddItem "FT": CmbTipo.AddItem "BV": CmbTipo.AddItem "ND": CmbTipo.ListIndex = 0
 Pintar
End Sub
Private Sub Pintar()
 Dim Id As Variant
 GrillaQuery GridDocs, "Documento|Importe"
 GridDocs.ColWidth(0) = 4200: GridDocs.ColWidth(1) = 1600
 For Each Id In Docs.Keys
  GridDocs.AddItem CStr(Docs(Id)(1)) & vbTab & Format$(Docs(Id)(2), "#,##0.00"): GridDocs.RowData(GridDocs.Rows - 1) = CLng(Id)
 Next Id
 If GridDocs.Rows > 1 Then GridDocs.Row = 1: SeleccionFilaQuery GridDocs, Seleccionando
End Sub
Private Sub CmdAgregar_Click()
 Dim L As ListBox
 On Error GoTo Fallo
 ' La lista oculta usa el mismo método de selección/validación que Recojos y Devoluciones.
 CobAgregarDocumento Modulo, Cliente, Moneda, CmbTipo.Text, Trim$(TxtNumero.Text), Docs, LstAux
 TxtNumero.Text = "": Pintar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub TxtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then CmdAgregar_Click: KeyCode = 0
End Sub
Private Sub GridDocs_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete And GridDocs.Row > 0 And GridDocs.Rows > 1 Then
  Docs.Remove CStr(GridDocs.RowData(GridDocs.Row)): Pintar: KeyCode = 0
 End If
End Sub
Private Sub GridDocs_Click()
 SeleccionFilaQuery GridDocs, Seleccionando
End Sub
Private Sub GridDocs_RowColChange()
 SeleccionFilaQuery GridDocs, Seleccionando
End Sub
Private Sub CmdAceptar_Click()
 Aceptado = True: Me.Hide
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = vbFormControlMenu Then Cancel = 1: Aceptado = False: Me.Hide
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

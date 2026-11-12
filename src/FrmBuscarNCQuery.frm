VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBuscarNCQuery 
   BorderStyle     =   1  'Fixed Single
   Caption = "Notas de crédito compatibles"
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
      Begin MSFlexGridLib.MSFlexGrid GridNC 
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
      Begin VB.TextBox TxtBusNc 
         MaxLength = 50
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin Clever.Boton CmdAceptar 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   6
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   6185064
      FCOL            =   16777215
      FCOLO           =   8454143
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
End
Attribute VB_Name = "FrmBuscarNCQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Public IdQuery As Long, Aceptado As Boolean, NumeroSeleccionado As String
Private Datos As ADODB.Recordset, Seleccionando As Boolean
Private Sub Form_Load()
 If Not AccesoPermitido(7) Then Unload Me: Exit Sub
 GrillaQuery GridNC, "Nota de crédito|Emisión|Moneda|Importe|Referencia"
 GridNC.ColWidth(0) = 1700: GridNC.ColWidth(1) = 1200: GridNC.ColWidth(2) = 1000: GridNC.ColWidth(3) = 1100: GridNC.ColWidth(4) = 1900
End Sub
Public Sub Preparar(ByVal Item As Long, ByVal Buscar As String)
 IdQuery = Item: Aceptado = False: NumeroSeleccionado = "": TxtBusNc.Text = Buscar
 BuscarNotas
End Sub
Private Sub BuscarNotas()
 On Error GoTo Fallo
 CerrarRecordset Datos: Set Datos = NotasQuery(IdQuery, Trim$(TxtBusNc.Text))
 GridNC.Rows = 1
 Do While Not Datos.EOF
  GridNC.AddItem CStr(Datos!Documento) & vbTab & FechaDocumento(Datos!Emision) & vbTab & CStr(Datos!Divisa) & vbTab & Format$(Datos!Importe, "#,##0.00") & vbTab & CStr(Datos!Referencia)
  Datos.MoveNext
 Loop
 If GridNC.Rows > 1 Then GridNC.Row = 1: SeleccionFilaQuery GridNC, Seleccionando
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub TxtBusNc_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyF9 Then BuscarNotas: KeyCode = 0
End Sub
Private Sub CmdAceptar_Click()
 If GridNC.Rows < 2 Or GridNC.Row < 1 Then Exit Sub
 NumeroSeleccionado = GridNC.TextMatrix(GridNC.Row, 0): Aceptado = True: Me.Hide
End Sub
Private Sub GridNC_DblClick()
 CmdAceptar_Click
End Sub
Private Sub GridNC_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then CmdAceptar_Click: KeyCode = 0
End Sub
Private Sub GridNC_Click()
 SeleccionFilaQuery GridNC, Seleccionando
End Sub
Private Sub GridNC_RowColChange()
 SeleccionFilaQuery GridNC, Seleccionando
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode = vbFormControlMenu Then Cancel = 1: Aceptado = False: Me.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

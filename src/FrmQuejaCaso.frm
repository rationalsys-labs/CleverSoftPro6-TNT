VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQuejaCaso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creacion de Nuevo Caso y Nueva Gestion"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo Caso"
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   9615
      Begin MSFlexGridLib.MSFlexGrid GridSta 
         Height          =   3135
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5530
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
      Begin VB.TextBox TxtCas 
         MaxLength = 4000
         Height          =   3135
         Left            =   4200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1800
         Width           =   5175
      End
      Begin VB.TextBox TxtContac 
         MaxLength = 50
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox TxtTelef 
         MaxLength = 20
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox CmbTipo 
         Style = 2
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   2820
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Caso:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Status:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Contacto:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0083693C&
         BackStyle       =   1  'Opaque
         Height          =   1335
         Left            =   5160
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label LblMen 
         Alignment       =   2  'Center
         BackColor       =   &H0083693C&
         ForeColor       =   &H0080FFFF&
         Height          =   1095
         Left            =   5280
         TabIndex        =   13
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label LblCodCas 
         BackColor       =   &H00808080&
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   255
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
      Left            =   6240
      TabIndex        =   4
      Top             =   6120
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
      MICON           =   "FrmQuejaCaso.frx":0A2E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton CmdCan 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   6120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      FOCUSR          =   -1  'True
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmQuejaCaso.frx":0A4A
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
      Caption         =   "Codigo De Caso - 000001"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0083693C&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   2895
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   9720
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label LblCod 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8280
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LblCli 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmQuejaCaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private IdCliente As Long, Estados As ADODB.Recordset, Clave As String, Guardado As Boolean, Seleccionando As Boolean
Public Sub Preparar(ByVal Id As Long, ByVal Nombre As String)
 IdCliente = Id: Clave = NuevaClaveQuery: Guardado = False
 LblCli.Caption = Nombre: LblCod.Caption = CStr(Id): LblCorrel.Caption = "Nuevo caso"
End Sub
Private Sub Form_Load()
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If Not AccesoPermitido(12) Then Unload Me: Exit Sub
 Set R = CobCatalogo(12, 4): CmbTipo.Clear
 Do While Not R.EOF
  If CBool(R!Activo) Then CmbTipo.AddItem CStr(R!Descripcion): CmbTipo.ItemData(CmbTipo.NewIndex) = CLng(R!Item)
  R.MoveNext
 Loop
 CerrarRecordset R
 If CmbTipo.ListCount > 0 Then CmbTipo.ListIndex = 0
 Set Estados = CobCatalogo(12, 5): Estados.Filter = "Activo=True"
 CobGrilla GridSta, Estados, "Código|Estado|Descripción", "Item|Codigo|Descripcion"
 GridSta.ColWidth(0) = 0: GridSta.ColWidth(1) = 900: GridSta.ColWidth(2) = 2800
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title: Unload Me
End Sub
Private Sub CmdAceptar_Click()
 Dim S As String, R As ADODB.Recordset, Id As Long
 On Error GoTo Fallo
 If Guardado Then Me.Hide: Exit Sub
 Id = CobFila(Estados, GridSta)
 S = "<gestion" & AtributoXML("Caso", "0") & AtributoXML("Cliente", CStr(IdCliente)) & AtributoXML("Tipo", CStr(CobIdCombo(CmbTipo))) & AtributoXML("Estado", CStr(Id)) & AtributoXML("Contacto", Trim$(TxtContac.Text)) & AtributoXML("Telefono", Trim$(TxtTelef.Text)) & AtributoXML("Texto", Trim$(TxtCas.Text)) & "/>"
 If MsgBox("¿Crear el caso con su primera gestión?", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 Set R = CobGuardar(12, Clave, S): Guardado = True: CerrarRecordset R
 Me.Hide
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdCan_Click()
 Me.Hide
End Sub
Private Sub GridSta_Click()
 SeleccionFilaQuery GridSta, Seleccionando
End Sub
Private Sub GridSta_RowColChange()
 SeleccionFilaQuery GridSta, Seleccionando
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Estados
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

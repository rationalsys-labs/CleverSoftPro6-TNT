VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQuejaDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informacion Historica"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   7935
      Begin VB.TextBox TxtGes 
         Locked = -1
         Appearance      =   0  'Flat
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   2760
         Width           =   7695
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label LblUsu 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   7800
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Caso:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblCas 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label LblSta 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label LblDes 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblFec 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6360
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Contacto:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label LblCont 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label15 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Lbltelef 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.Label LblCli 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label20 
         Caption         =   "Ruc:"
         Height          =   255
         Left            =   5640
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LblRuc 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6360
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Direccion:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label LblDir 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   1200
         Width           =   6735
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
      Left            =   3120
      TabIndex        =   1
      Top             =   5640
      Width           =   2055
      _ExtentX        =   3625
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
      MICON           =   "FrmQuejaDetalle.frx":0000
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
Attribute VB_Name = "FrmQuejaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Public Sub Preparar(ByVal Caso As ADODB.Recordset, ByVal Gestion As ADODB.Recordset)
 LblCas.Caption = ValorTexto(Caso!DescripcionCaso): LblCli.Caption = ValorTexto(Caso!Cliente)
 LblSta.Caption = ValorTexto(Gestion!CodigoEstado): LblDes.Caption = ValorTexto(Gestion!DescripcionEstado)
 LblFec.Caption = Format$(Gestion!FecSta, "dd/mm/yyyy hh:nn:ss"): LblCont.Caption = ValorTexto(Gestion!Contac)
 Lbltelef.Caption = ValorTexto(Gestion!Telef1): LblUsu.Caption = ValorTexto(Gestion!Usuario)
 LblRuc.Caption = ValorTexto(Caso!RUCCliente): LblDir.Caption = ValorTexto(Caso!DireccionCliente): TxtGes.Text = ValorTexto(Gestion!Gestion)
End Sub
Private Sub CmdAceptar_Click()
 Unload Me
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Clever SoftPro 6.0 - Presentación"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   5475
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LblVersion
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 6.0.1"
      ForeColor       =   &H00503016&
      Height          =   225
      Left            =   6600
      Top             =   180
      Width           =   3240
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblAutor
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "© Juan Pablo Aguilar"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5400
      Top             =   4545
      Width           =   4500
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblCorreo
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "jaguilar.rationalsys@outlook.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5400
      Top             =   4770
      Width           =   4500
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblDerechos
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Todos los derechos reservados."
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5400
      Top             =   4995
      Width           =   4500
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrInicio
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Public Listo As Boolean

Private Sub Form_Load()
    Listo = False
    LblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    ' Se conserva la presentación de cinco segundos mediante un Timer de VB6.
    tmrInicio.Enabled = True
End Sub

Private Sub tmrInicio_Timer()
    tmrInicio.Enabled = False
    Listo = True
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrInicio.Enabled = False
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

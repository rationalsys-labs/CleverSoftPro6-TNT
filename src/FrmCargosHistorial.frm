VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCargosHistorial 
   BorderStyle     =   1  'Fixed Single
   Caption = "Historial de gestión del documento"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   Icon            =   "FrmCargosHistorial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid GridStatus 
      Height          =   3525
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   6218
      _Version        =   393216
      FixedCols       =   0
   End
   Begin Clever.Boton cmdAceptar 
      Height          =   375
      Left            =   8190
      TabIndex        =   1
      Top             =   3690
      Width           =   2040
      _extentx        =   3598
      _extenty        =   661
      btype           =   5
      tx              =   "Aceptar"
      enab            =   -1  'True
      font            =   "FrmCargosHistorial.frx":1782
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   13160660
      bcolo           =   13160660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "FrmCargosHistorial.frx":17AA
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
End
Attribute VB_Name = "FrmCargosHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Public Sub Preparar(ByVal Tipo As String, ByVal Item As Long)
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 PrepararGrillaDocumento GridStatus, "Fecha|Estado|Usuario|Situación del registro|Registro"
 GridStatus.ColWidth(0) = 1300: GridStatus.ColWidth(1) = 3000: GridStatus.ColWidth(2) = 1700
 GridStatus.ColWidth(3) = 2300: GridStatus.ColWidth(4) = 1200
 Set R = DetalleDocumento(Tipo, Item, 2, 19)
 Do While Not R.EOF
  GridStatus.AddItem FechaDocumento(R!Fecha) & vbTab & CeldaDocumento(R!Estado) & vbTab & CeldaDocumento(R!Usuario) & vbTab & IIf(CBool(R!Activo), "Activo", "Retirado") & vbTab & CeldaDocumento(R!Registro)
  R.MoveNext
 Loop
 CerrarRecordset R
 Exit Sub
Fallo:
 CerrarRecordset R
 Err.Raise vbObjectError + 3511, , Err.Description
End Sub
Private Sub cmdAceptar_Click()
 Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

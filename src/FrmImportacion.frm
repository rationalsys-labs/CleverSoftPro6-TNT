VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImportacion 
   KeyPreview = -1
   BorderStyle     =   1  'Fixed Single
   Caption = "Importación de información"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   13800
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   240
      Left            =   10440
      TabIndex        =   28
      Top             =   1665
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ListBox LstErr 
      Height          =   840
      Left            =   90
      TabIndex        =   26
      Top             =   7065
      Visible         =   0   'False
      Width           =   2580
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   15
      Top             =   7995
      Width           =   13800
      _ExtentX        =   24342
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4059
            MinWidth        =   4059
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15876
            MinWidth        =   15876
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportacion.frx":0000
            Key             =   "Dat"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportacion.frx":059A
            Key             =   "a2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportacion.frx":0B34
            Key             =   "a3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImportacion.frx":10CE
            Key             =   "a4"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GridCarg 
      Height          =   3660
      Left            =   2880
      TabIndex        =   13
      Top             =   4230
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   6456
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estadisticas"
      Height          =   2175
      Left            =   8730
      TabIndex        =   12
      Top             =   1980
      Width           =   4920
      Begin VB.Label Det_5 
         Caption = "0"
         Height          =   285
         Left            =   1620
         TabIndex        =   25
         Top             =   1665
         Width           =   3165
      End
      Begin VB.Label Det_4 
         Caption = "0"
         Height          =   285
         Left            =   1620
         TabIndex        =   24
         Top             =   1350
         Width           =   3165
      End
      Begin VB.Label Det_3 
         Caption = "0"
         Height          =   285
         Left            =   1620
         TabIndex        =   23
         Top             =   1035
         Width           =   3165
      End
      Begin VB.Label Det_2 
         Caption = "0"
         Height          =   285
         Left            =   1620
         TabIndex        =   22
         Top             =   720
         Width           =   3165
      End
      Begin VB.Label Det_1 
         Caption = "0"
         Height          =   285
         Left            =   1620
         TabIndex        =   21
         Top             =   405
         Width           =   3165
      End
      Begin VB.Label Est_5 
         AutoSize        =   -1  'True
         Caption = "Total de filas"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   1665
         Width = 1455
      End
      Begin VB.Label Est_4 
         AutoSize        =   -1  'True
         Caption = "Servicios"
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   1350
         Width = 1455
      End
      Begin VB.Label Est_3 
         AutoSize        =   -1  'True
         Caption = "Documentos"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   1035
         Width = 1455
      End
      Begin VB.Label Est_2 
         AutoSize        =   -1  'True
         Caption = "Guías"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   720
         Width = 1455
      End
      Begin VB.Label Est_1 
         AutoSize        =   -1  'True
         Caption = "Clientes"
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   405
         Width = 1455
      End
   End
   Begin VB.TextBox TxtArchivo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4635
      TabIndex        =   10
      Top             =   1350
      Width           =   9015
   End
   Begin Clever.Boton CmdGuardar 
      Enabled = 0
      Height          =   375
      Left            =   11970
      TabIndex        =   8
      Top             =   495
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      BTYPE           =   14
      TX = "Confirmar carga"
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmImportacion.frx":1668
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   7470
      TabIndex        =   5
      Top             =   405
      Width           =   4380
      Begin Clever.Boton CmdValidar 
         Height          =   510
         Left            =   3465
         TabIndex        =   6
         Top             =   180
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   900
         BTYPE           =   8
         TX = "Validar"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "FrmImportacion.frx":1684
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblEstado 
         Caption = "Seleccione un libro y valide su contenido."
         Height          =   420
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   3255
      End
   End
   Begin VB.ComboBox CmbArchivo 
      Height          =   315
      Left            =   4455
      Style = 2
      TabIndex        =   3
      Top             =   855
      Width           =   2925
   End
   Begin VB.ComboBox CmbDBase 
      Height          =   315
      Left            =   4455
      Style = 2
      TabIndex        =   1
      Top             =   495
      Width           =   2925
   End
   Begin Clever.Boton CmdDatImpRech 
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   540
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   582
      BTYPE           =   8
      TX = "Detalle de incidencias"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   128
      BCOLO           =   14737632
      FCOL            =   16777215
      FCOLO           =   128
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmImportacion.frx":16A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton CmdExaminar 
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   1350
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   503
      BTYPE           =   14
      TX              =   "Examinar"
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmImportacion.frx":16BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Visible = 0
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13800
      _ExtentX        =   24342
      _ExtentY        =   635
      ButtonWidth     =   2249
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rechazados"
            ImageKey        =   "Dat"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.TreeView ViewErr 
      Height          =   2070
      Left            =   2880
      TabIndex        =   27
      Top             =   2070
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   3651
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Detalle de Proceso"
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Archivo a Importar"
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   900
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data a Actualizar:"
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   540
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BackColor = &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   8430
      Left            =   0
      Top             =   -45
      Width           =   2760
   End
   Begin VB.TextBox TxtResultado
      Left=0
      Top=0
      Width=100
      Height=100
      Visible=0
      MultiLine=-1
   End
   Begin Clever.Boton CmdSalir
      Left=135
      Top=7450
      Width=2490
      Height=375
      TX="Salir"
      BTYPE=6
   End
End
Attribute VB_Name = "FrmImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean

Private XMLValidado As String
Private ArchivoValidado As String
Private Ocupado As Boolean
Private Sub TxtArchivo_Change()
 XMLValidado = ""
 CmdGuardar.Enabled = False
End Sub
Private Sub CmdExaminar_Click()
 On Error GoTo Fallo
 TxtArchivo.Text = ElegirLibroImportacion()
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Habilitar(ByVal Valor As Boolean)
 Ocupado = Not Valor
 CmdExaminar.Enabled = Valor
 CmdValidar.Enabled = Valor
 TxtArchivo.Enabled = Valor
 CmdSalir.Enabled = Valor
End Sub
Private Sub CmdValidar_Click()
 Dim Resumen As String, Datos As ADODB.Recordset
 On Error GoTo Fallo
 XMLValidado = ""
 CmdGuardar.Enabled = False
 Habilitar False
 Progress.Value = 0
 ViewErr.Nodes.Clear
 ViewErr.Nodes.Add , , , "Leyendo libro Excel..."
 LblEstado.Caption = "Leyendo y validando. No cierre la aplicación."
 Me.Refresh
 XMLValidado = LeerLibroImportacion(Trim$(TxtArchivo.Text), Resumen)
 Progress.Value = 50
 ViewErr.Nodes.Add , , , "Comprobando datos y relaciones..."
 Set Datos = ProcesarImportacion(Trim$(TxtArchivo.Text), XMLValidado, False)
 Progress.Value = 100
 ArchivoValidado = Trim$(TxtArchivo.Text)
 MostrarDatosLibro
 TxtResultado.Text = Resumen & vbCrLf & "VALIDACIÓN CORRECTA: " & CStr(Datos!Filas) & " filas. No se confirmó ningún cambio." & vbCrLf & "Revise el archivo y pulse Confirmar carga."
 CerrarRecordset Datos
 Habilitar True
 CmdGuardar.Enabled = True
 LblEstado.Caption = "Archivo validado. Pendiente de confirmar."
 Exit Sub
Fallo:
 Progress.Value = 0
 TxtResultado.Text = Err.Description
 XMLValidado = ""
 CerrarRecordset Datos
 Habilitar True
 LblEstado.Caption = "Validación rechazada. Corrija el archivo y vuelva a validar."
End Sub
Private Sub CmdGuardar_Click()
 Dim Datos As ADODB.Recordset, Confirmado As Boolean
 On Error GoTo Fallo
 If Len(XMLValidado) = 0 Then Exit Sub
 If MsgBox("¿Confirmar la carga validada de " & ArchivoValidado & "?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
 Habilitar False
 CmdGuardar.Enabled = False
 Set Datos = ProcesarImportacion(ArchivoValidado, XMLValidado, True)
 Confirmado = True
 XMLValidado = ""
 TxtResultado.Text = "CARGA CONFIRMADA: " & CStr(Datos!Filas) & " filas procesadas. Los registros idénticos no se duplican."
 CerrarRecordset Datos
 Habilitar True
 LblEstado.Caption = "Carga terminada."
 Exit Sub
Fallo:
 CerrarRecordset Datos
 Habilitar True
 If Confirmado Then
  XMLValidado = ""
  TxtResultado.Text = "La carga fue confirmada. No repita la escritura por este error de presentación: " & Err.Description
 Else
  TxtResultado.Text = "La carga no se confirmó: " & Err.Description
  CmdGuardar.Enabled = (Len(XMLValidado) > 0)
 End If
End Sub
Private Sub Form_Load()
 If Not AccesoPermitido(21) Then Unload Me: Exit Sub
 InicializarVista
End Sub
Private Sub CmdSalir_Click()
 Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Cancel = Ocupado
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyEscape And Not Ocupado Then Unload Me
End Sub

Private Sub InicializarVista()
 CmdValidar.Enabled = True
 CmdGuardar.Enabled = False
 CmdDatImpRech.Enabled = False
 CmbArchivo.Clear
 CmbArchivo.AddItem "Libro Excel (.xls / .xlsx)"
 CmbArchivo.ListIndex = 0
 CmbArchivo.Enabled = False
 CmbDBase.Clear
 CmbDBase.AddItem "Clientes"
 CmbDBase.AddItem "Guías"
 CmbDBase.AddItem "Documentos"
 CmbDBase.AddItem "Servicios facturados"
 CmbDBase.ListIndex = 0
 GridCarg.Rows = 1
 GridCarg.Cols = 2
 GridCarg.FormatString = "Fila|Vista previa del libro validado"
 ViewErr.Nodes.Clear
 ViewErr.Nodes.Add , , , "Seleccione el libro de información externa."
 Progress.Value = 0
 StatusBar.Panels(1).Text = "Consulta"
End Sub
Private Sub TxtResultado_Change()
 CmdDatImpRech.Enabled = (Len(TxtResultado.Text) > 0)
 Dim Lineas As Variant, I As Long
 ViewErr.Nodes.Clear
 Lineas = Split(TxtResultado.Text, vbCrLf)
 For I = 0 To UBound(Lineas)
  If Len(Lineas(I)) > 0 Then ViewErr.Nodes.Add , , , CStr(Lineas(I))
 Next I
 StatusBar.Panels(1).Text = LblEstado.Caption
End Sub
Private Sub CmbDBase_Click()
 If Len(XMLValidado) > 0 Then MostrarDatosLibro
End Sub
Private Sub MostrarDatosLibro()
 Dim Documento As Object, Filas As Object, Fila As Object, Nombre As String, Campos As Variant, Nombres As Variant
 Dim I As Long, F As Long, Cuenta(1 To 4) As Long, Todas As Object
 On Error GoTo FalloVista
 Set Documento = CreateObject("MSXML2.DOMDocument.6.0")
 Documento.async = False
 If Not Documento.LoadXML(XMLValidado) Then Err.Raise 5, , "No se pudo presentar la vista previa."
 Set Todas = Documento.selectNodes("/carga/r")
 For Each Fila In Todas
  Select Case Fila.getAttribute("hoja")
   Case "Clientes": Cuenta(1) = Cuenta(1) + 1
   Case "Guias": Cuenta(2) = Cuenta(2) + 1
   Case "Documentos": Cuenta(3) = Cuenta(3) + 1
   Case "Servicios": Cuenta(4) = Cuenta(4) + 1
  End Select
 Next Fila
 Det_1.Caption = CStr(Cuenta(1)): Det_2.Caption = CStr(Cuenta(2))
 Det_3.Caption = CStr(Cuenta(3)): Det_4.Caption = CStr(Cuenta(4))
 Det_5.Caption = CStr(Todas.Length)
 Nombres = Array("Clientes", "Guias", "Documentos", "Servicios")
 Nombre = CStr(Nombres(CmbDBase.ListIndex))
 Campos = Split(CamposImportacion(Nombre), ",")
 GridCarg.Rows = 1
 GridCarg.Cols = UBound(Campos) + 2
 GridCarg.TextMatrix(0, 0) = "Fila"
 GridCarg.ColWidth(0) = 600
 For I = 0 To UBound(Campos)
  GridCarg.TextMatrix(0, I + 1) = CStr(Campos(I))
  GridCarg.ColWidth(I + 1) = 1600
 Next I
 If Nombre = "Guias" Then
  GridCarg.ColWidth(1) = 1650
  GridCarg.ColWidth(2) = 1000
  GridCarg.ColWidth(3) = 1800
  GridCarg.ColWidth(4) = GridCarg.Width - 7950
  GridCarg.ColWidth(5) = 1400
  GridCarg.ColWidth(6) = 1300
 End If
 Set Filas = Documento.selectNodes("/carga/r[@hoja='" & Nombre & "']")
 For Each Fila In Filas
  If GridCarg.Rows >= 501 Then Exit For
  GridCarg.AddItem ""
  F = GridCarg.Rows - 1
  GridCarg.TextMatrix(F, 0) = CStr(Fila.getAttribute("fila"))
  For I = 0 To UBound(Campos)
   GridCarg.TextMatrix(F, I + 1) = CStr(Fila.getAttribute(CStr(Campos(I))))
  Next I
 Next Fila
 StatusBar.Panels(1).Width = 6500
 StatusBar.Panels(1).Text = Nombre & ": " & CStr(Filas.Length) & " filas; vista previa hasta 500."
 Exit Sub
FalloVista:
 StatusBar.Panels(1).Text = "No se pudo actualizar la vista previa: " & Err.Description
End Sub
Private Sub CmdDatImpRech_Click()
 MostrarHistorialRP TxtResultado.Text
End Sub
Private Sub Form_Resize()
 On Error GoTo Salir
 If Me.WindowState = vbMinimized Then Exit Sub
 If ScaleWidth < 12000 Or ScaleHeight < 7000 Then Exit Sub
 Shape1.Height = ScaleHeight
 CmdSalir.Top = ScaleHeight - StatusBar.Height - 600
 TxtArchivo.Width = ScaleWidth - TxtArchivo.Left - 150
 CmdGuardar.Left = ScaleWidth - CmdGuardar.Width - 150
 Frame2.Left = ScaleWidth - Frame2.Width - 150
 ViewErr.Width = Frame2.Left - ViewErr.Left - 90
 Progress.Left = Frame2.Left + 1710
 Frame1.Left = CmdGuardar.Left - Frame1.Width - 120
 GridCarg.Width = ScaleWidth - GridCarg.Left - 150
 GridCarg.Height = ScaleHeight - StatusBar.Height - GridCarg.Top - 100
Salir:
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub

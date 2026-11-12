VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFacturas 
   BackColor       =   &H8000000F&
   BorderStyle     =   1  'Fixed Single
   Caption = "Listado de documentos emitidos"
   ClientHeight = 5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight = 5400
   ScaleWidth      =   12390
   Begin MSComctlLib.TabStrip TabDocumento
      Left = 0
      Top = 120
      Width = 12390
      Height = 4830
      TabIndex = 8
      _Version = 393216
   End
   Begin VB.ComboBox CmbFiltro 
      Style = 2
      Height          =   315
      Left            =   9000
      TabIndex        =   16
      Text            =   "Combo1"
      Top = 645
      Width           =   3255
   End
   Begin MSComctlLib.ImageList ICO 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFacturas.frx":2A87
            Key             =   "a1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFacturas.frx":3499
            Key             =   "A2"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Fra_Inf 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top = 1080
      Width           =   12135
      Begin MSFlexGridLib.MSFlexGrid GridFac 
         Height          =   3705
         Left            =   15
         TabIndex        =   12
         Top             =   120
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   6535
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   -2147483633
         BackColorBkg    =   16777215
         GridColor       =   15790320
         GridColorFixed  =   15790320
         GridLinesFixed  =   1
         MergeCells      =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Fra_Can 
      BackColor       =   &H8000000F&
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top = 1080
      Visible         =   0   'False
      Width           =   12135
      Begin MSFlexGridLib.MSFlexGrid GridGuias 
         Height          =   3015
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width = 11655
         _ExtentX        =   8070
         _ExtentY        =   5318
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         Appearance      =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption = "Guías y servicios del documento"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width = 4500
      End
   End
   Begin VB.Frame Fra_Det 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top = 1080
      Visible         =   0   'False
      Width           =   12135
      Begin MSFlexGridLib.MSFlexGrid GridDet 
         Height          =   3735
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   6588
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   -2147483633
         BackColorBkg    =   16777215
         GridColorFixed  =   14737632
         GridLinesFixed  =   1
         MergeCells      =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Fra_dat 
      BackColor       =   &H8000000F&
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top = 1080
      Visible         =   0   'False
      Width           =   12135
      Begin VB.Frame Frame1
         BackColor = &H8000000F&
         BorderStyle = 0
         Height = 3855
         Left = 240
         Top = 240
         Width = 4575
         Begin VB.Label LblDato0
            Left = 120
            Top = 180
            Width = 1080
            Height = 255
            Caption = "Documento:"
            BackStyle = 0
         End
         Begin VB.TextBox TxtDocumento
            Left = 1260
            Top = 150
            Width = 3195
            Height = 315
            Locked = -1
            BackColor = &H80000005&
            ForeColor = &H80000008&
            TabIndex = 34
         End
         Begin VB.Label LblDato1
            Left = 120
            Top = 600
            Width = 1080
            Height = 255
            Caption = "Cliente:"
            BackStyle = 0
         End
         Begin VB.TextBox TxtCliente
            Left = 1260
            Top = 570
            Width = 3195
            Height = 315
            Locked = -1
            BackColor = &H80000005&
            ForeColor = &H80000008&
            TabIndex = 35
         End
         Begin VB.Label LblDato2
            Left = 120
            Top = 1020
            Width = 1080
            Height = 255
            Caption = "Código:"
            BackStyle = 0
         End
         Begin VB.TextBox TxtCodigo
            Left = 1260
            Top = 990
            Width = 3195
            Height = 315
            Locked = -1
            BackColor = &H80000005&
            ForeColor = &H80000008&
            TabIndex = 36
         End
         Begin VB.Label LblDato3
            Left = 120
            Top = 1440
            Width = 1080
            Height = 255
            Caption = "Situación:"
            BackStyle = 0
         End
         Begin VB.TextBox TxtSituacion
            Left = 1260
            Top = 1410
            Width = 3195
            Height = 315
            Locked = -1
            BackColor = &H80000005&
            ForeColor = &H80000008&
            TabIndex = 37
         End
         Begin VB.Label LblDato4
            Left = 120
            Top = 1860
            Width = 1080
            Height = 255
            Caption = "Gestión:"
            BackStyle = 0
         End
         Begin VB.TextBox TxtGestion
            Left = 1260
            Top = 1830
            Width = 3195
            Height = 315
            Locked = -1
            BackColor = &H80000005&
            ForeColor = &H80000008&
            TabIndex = 38
         End
         Begin VB.Label LblDato5
            Left = 120
            Top = 2280
            Width = 1080
            Height = 255
            Caption = "Fecha:"
            BackStyle = 0
         End
         Begin VB.TextBox TxtFecha
            Left = 1260
            Top = 2250
            Width = 3195
            Height = 315
            Locked = -1
            BackColor = &H80000005&
            ForeColor = &H80000008&
            TabIndex = 39
         End
         Begin VB.Label LblDato6
            Left = 120
            Top = 2700
            Width = 1080
            Height = 255
            Caption = "Guías:"
            BackStyle = 0
         End
         Begin VB.TextBox TxtGuias
            Left = 1260
            Top = 2670
            Width = 3195
            Height = 315
            Locked = -1
            BackColor = &H80000005&
            ForeColor = &H80000008&
            TabIndex = 40
         End
         Begin VB.Label LblDato7
            Left = 120
            Top = 3120
            Width = 1080
            Height = 255
            Caption = "Courier:"
            BackStyle = 0
         End
         Begin VB.TextBox TxtCourier
            Left = 1260
            Top = 3090
            Width = 3195
            Height = 315
            Locked = -1
            BackColor = &H80000005&
            ForeColor = &H80000008&
            TabIndex = 41
         End
      End
      Begin MSComctlLib.TreeView LstViewHis 
         Height          =   2895
         Left            =   5040
         TabIndex        =   2
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5106
         _Version        =   393217
         Indentation     =   178
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
      End
   Begin VB.Label LabelDegrade2
      Left=8520
      Top=360
      Width=3375
      Height=255
      Caption="HISTORIAL DE COURIER"
      BackStyle=0
      BackStyle=1
      BackColor=&H8000000F&
      ForeColor=&H80000012&

   End
   Begin VB.Label LabelDegrade1
      Left=5040
      Top=360
      Width=3375
      Height=255
      Caption="HISTORIAL DE GESTIÓN"
      BackStyle=0
      BackStyle=1
      BackColor=&H8000000F&
      ForeColor=&H80000012&

   End
      Begin MSComctlLib.TreeView ViewCourier 
         Height          =   2895
         Left            =   8520
         TabIndex        =   5
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5106
         _Version        =   393217
         Indentation     =   178
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   33
      Top             =   4440
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Label LblBuscar
      Left=120
      Top=690
      Width=1200
      Height=255
      Caption="Buscar:"
      BackStyle = 0
   End
   Begin VB.TextBox TxtBuscar
      Left=1440
      Top=645
      Width=7200
      Height=315
      MaxLength=100
   End
End
Attribute VB_Name = "FrmFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Datos As ADODB.Recordset
Private Cargando As Boolean
Private SeleccionandoNotas As Boolean
Private PosicionInicialAplicada As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(11) Then Unload Me: Exit Sub
 Cargando = True
 StatusBar.Panels.Clear: StatusBar.Panels.Add
 StatusBar.Panels(1).AutoSize = sbrSpring
 TabDocumento.Tabs.Clear
 TabDocumento.Tabs.Add , , "Listado de documentos"
 TabDocumento.Tabs.Add , , "Datos e históricos"
 TabDocumento.Tabs.Add , , "Notas relacionadas"
 TabDocumento.Tabs.Add , , "Guías y servicios"
 TabDocumento.ZOrder 1
 Set LstViewHis.ImageList = ICO: Set ViewCourier.ImageList = ICO
 CmbFiltro.Clear
 CmbFiltro.AddItem "Todos los documentos"
 CmbFiltro.AddItem "FT - Facturas"
 CmbFiltro.AddItem "BV - Boletas"
 CmbFiltro.AddItem "NC - Notas de crédito"
 CmbFiltro.AddItem "ND - Notas de débito"
 CmbFiltro.ListIndex = 0
 PrepararGrillaDocumento GridFac, "Tipo|Documento|Emisión|Vencimiento|Moneda|Importe|Saldo|Situación|Gestión|Cliente"
 GridFac.MergeCells = flexMergeNever
 GridFac.BackColorSel = vbHighlight
 GridFac.ForeColorSel = vbHighlightText
 GridFac.ColWidth(0) = 500: GridFac.ColWidth(1) = 1700
 GridFac.ColWidth(2) = 1000: GridFac.ColWidth(3) = 1000
 GridFac.ColWidth(4) = 700: GridFac.ColWidth(5) = 1100: GridFac.ColWidth(6) = 1100
 GridFac.ColWidth(7) = 850: GridFac.ColWidth(8) = 1700: GridFac.ColWidth(9) = 3400
 PrepararGrillaDocumento GridGuias, "Guía|Servicio|Origen|Envío|Producto|Gestión"
 GridGuias.ColWidth(0) = 1500: GridGuias.ColWidth(1) = 4200
 GridGuias.ColWidth(2) = 900: GridGuias.ColWidth(3) = 1300: GridGuias.ColWidth(4) = 1100: GridGuias.ColWidth(5) = 1100
 PrepararGrillaDocumento GridDet, "Tipo|Documento|Emisión|Moneda|Importe|Situación|Referencia"
 GridDet.ColWidth(0) = 600: GridDet.ColWidth(1) = 2200: GridDet.ColWidth(2) = 1300
 GridDet.ColWidth(3) = 900: GridDet.ColWidth(4) = 1400: GridDet.ColWidth(5) = 1400: GridDet.ColWidth(6) = 2200
 GridDet.MergeCells = flexMergeNever
 GridDet.BackColorSel = vbHighlight: GridDet.ForeColorSel = vbHighlightText
 Cargando = False
 Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
 Unload Me
End Sub
Private Sub Actualizar(Optional ByVal Exacto As Boolean = False)
 Dim Tipo As String, Saldo As String
 On Error GoTo Fallo
 If Cargando Then Exit Sub
 If CmbFiltro.ListIndex > 0 Then Tipo = Left$(CmbFiltro.Text, 2)
 CerrarRecordset Datos
 Set Datos = ConsultarDocumentos(Tipo, Trim$(TxtBuscar.Text), Exacto)
 Cargando = True
 GridFac.Rows = 1
 Do While Not Datos.EOF
  Saldo = ""
  If Not IsNull(Datos!Saldo) Then Saldo = Format$(Datos!Saldo, "#,##0.00")
  GridFac.AddItem CStr(Datos!Tipo) & vbTab & CeldaDocumento(Datos!Documento) & vbTab & FechaDocumento(Datos!Emision) & vbTab & FechaDocumento(Datos!Vencimiento) & vbTab & CeldaDocumento(Datos!Divisa) & vbTab & Format$(Datos!Importe, "#,##0.00") & vbTab & Saldo & vbTab & CeldaDocumento(Datos!Situacion) & vbTab & CeldaDocumento(Datos!Estado) & vbTab & CeldaDocumento(Datos!Cliente)
  Datos.MoveNext
 Loop
 If GridFac.Rows > 1 Then GridFac.Row = 1
 Cargando = False
 SeleccionarFila
 MostrarPanel 0
 StatusBar.Panels(1).Text = CStr(Datos.RecordCount) & " documentos. Máximo 200; escriba para filtrar."
 Exit Sub
Fallo:
 Cargando = False
 GridFac.Rows = 1
 StatusBar.Panels(1).Text = "No se pudo consultar. F5 permite reintentar."
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Function SeleccionValida() As Boolean
 If Datos Is Nothing Then Exit Function
 If GridFac.Row < 1 Or GridFac.Row >= GridFac.Rows Then Exit Function
 If Datos.RecordCount = 0 Then Exit Function
 Datos.MoveFirst: Datos.Move GridFac.Row - 1
 SeleccionValida = True
End Function
Private Sub MostrarPanel(ByVal Indice As Integer)
 Dim TopPanel As Single, EstadoCarga As Boolean
 EstadoCarga = Cargando: Cargando = True
 Set TabDocumento.SelectedItem = TabDocumento.Tabs(Indice + 1)
 Cargando = EstadoCarga
 Fra_Inf.Visible = (Indice = 0): Fra_dat.Visible = (Indice = 1)
 Fra_Det.Visible = (Indice = 2): Fra_Can.Visible = (Indice = 3)
 LblBuscar.Visible = (Indice = 0): TxtBuscar.Visible = (Indice = 0): CmbFiltro.Visible = (Indice = 0)
 TopPanel = IIf(Indice = 0, 1080, 600)
 AjustarMarco Fra_Inf, TopPanel
 AjustarMarco Fra_dat, TopPanel
 AjustarMarco Fra_Det, TopPanel
 AjustarMarco Fra_Can, TopPanel
 GridFac.Height = Fra_Inf.Height - GridFac.Top - 120
 GridDet.Height = Fra_Det.Height - GridDet.Top - 120
 GridGuias.Height = Fra_Can.Height - GridGuias.Top - 120
 Frame1.Height = Fra_dat.Height - Frame1.Top - 240
 LstViewHis.Height = Fra_dat.Height - LstViewHis.Top - 240
 ViewCourier.Height = Fra_dat.Height - ViewCourier.Top - 240
End Sub
Private Sub AjustarMarco(ByVal Marco As VB.Frame, ByVal Arriba As Single)
 Marco.Move Marco.Left, Arriba, Marco.Width, ScaleHeight - StatusBar.Height - Arriba - 120
End Sub
Private Sub TabDocumento_Click()
 If Cargando Then Exit Sub
 If TabDocumento.SelectedItem.Index > 1 Then
  If Not SeleccionValida Then MostrarPanel 0: Exit Sub
 End If
 Select Case TabDocumento.SelectedItem.Index
  Case 1: MostrarPanel 0
  Case 2: Cmd_Dat_Click
  Case 3: Cmd_Det_Click
  Case 4: Cmd_Can_Click
 End Select
End Sub
Private Sub Cmd_Inf_Click()
 MostrarPanel 0
End Sub
Private Sub Cmd_Dat_Click()
 Dim R As ADODB.Recordset, Texto As String
 On Error GoTo Fallo
 If Not SeleccionValida Then Exit Sub
 TxtDocumento.Text = ValorTexto(Datos!Documento)
 TxtCliente.Text = ValorTexto(Datos!Cliente): TxtCliente.ToolTipText = Left$(TxtCliente.Text, 255)
 TxtCodigo.Text = ValorTexto(Datos!CodigoCliente)
 TxtSituacion.Text = ValorTexto(Datos!Situacion): TxtGestion.Text = ValorTexto(Datos!Estado)
 TxtFecha.Text = FechaDocumento(Datos!FechaEstado): TxtGuias.Text = ValorTexto(Datos!Guias)
 TxtGuias.ToolTipText = Left$(TxtGuias.Text, 255): TxtCourier.Text = ValorTexto(Datos!Courier)
 LstViewHis.Nodes.Clear: ViewCourier.Nodes.Clear
 Set R = DetalleDocumento(CStr(Datos!Tipo), CLng(Datos!IdDocumento), 2)
 Do While Not R.EOF
  Texto = FechaDocumento(R!Fecha) & " - " & ValorTexto(R!Estado) & " / " & ValorTexto(R!Usuario)
  If Not CBool(R!Activo) Then Texto = Texto & " [retirado]"
  LstViewHis.Nodes.Add , , , Texto, 1, 1
  R.MoveNext
 Loop
 CerrarRecordset R
 Set R = DetalleDocumento(CStr(Datos!Tipo), CLng(Datos!IdDocumento), 3)
 Do While Not R.EOF
  ViewCourier.Nodes.Add , , , FechaDocumento(R!Fecha) & " - " & ValorTexto(R!Responsable), 2, 2
  R.MoveNext
 Loop
 CerrarRecordset R
 MostrarPanel 1
 Exit Sub
Fallo:
 CerrarRecordset R
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Cmd_Can_Click()
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If Not SeleccionValida Then Exit Sub
 GridGuias.Rows = 1
 Set R = DetalleDocumento(CStr(Datos!Tipo), CLng(Datos!IdDocumento), 1)
 Do While Not R.EOF
  GridGuias.AddItem CeldaDocumento(R!Guia) & vbTab & CeldaDocumento(R!Servicio) & vbTab & CeldaDocumento(R!Origen) & vbTab & FechaDocumento(R!FechaEnvio) & vbTab & CeldaDocumento(R!Producto) & vbTab & CeldaDocumento(R!Gestion)
  R.MoveNext
 Loop
 CerrarRecordset R
 MostrarPanel 3
 Exit Sub
Fallo:
 CerrarRecordset R
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Cmd_Det_Click()
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If Not SeleccionValida Then Exit Sub
 GridDet.Rows = 1
 Set R = DetalleDocumento(CStr(Datos!Tipo), CLng(Datos!IdDocumento), 4)
 Do While Not R.EOF
  GridDet.AddItem CeldaDocumento(R!Tipo) & vbTab & CeldaDocumento(R!Documento) & vbTab & FechaDocumento(R!Emision) & vbTab & CeldaDocumento(R!Divisa) & vbTab & Format$(R!Importe, "#,##0.00") & vbTab & CeldaDocumento(R!Situacion) & vbTab & CeldaDocumento(R!Referencia)
  R.MoveNext
 Loop
 CerrarRecordset R
 SeleccionarNota
 MostrarPanel 2
 Exit Sub
Fallo:
 CerrarRecordset R
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub TxtBuscar_Change()
 If Not Cargando Then Actualizar
End Sub
Private Sub CmbFiltro_Click()
 If Not Cargando Then Actualizar
End Sub
Private Sub GridFac_DblClick()
 SeleccionarFila
 Cmd_Dat_Click
End Sub
Private Sub SeleccionarFila()
 Dim PrimeraColumna As Long
 If Cargando Then Exit Sub
 If GridFac.Row < 1 Or GridFac.Row >= GridFac.Rows Then Exit Sub
 Cargando = True
 PrimeraColumna = GridFac.LeftCol
 GridFac.RowSel = GridFac.Row
 GridFac.Col = 0
 GridFac.ColSel = GridFac.Cols - 1
 GridFac.LeftCol = PrimeraColumna
 Cargando = False
End Sub
Private Sub GridFac_RowColChange()
 SeleccionarFila
End Sub
Private Sub GridFac_Click()
 SeleccionarFila
End Sub
Private Sub GridFac_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Cmd_Dat_Click: KeyCode = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF5 Then Actualizar: KeyCode = 0
 If KeyCode = vbKeyEscape Then
  If Fra_Inf.Visible Then Unload Me Else MostrarPanel 0
 End If
End Sub
Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub
Private Sub SeleccionarNota()
 Dim PrimeraColumna As Long
 If Cargando Or SeleccionandoNotas Then Exit Sub
 If GridDet.Row < 1 Or GridDet.Row >= GridDet.Rows Then Exit Sub
 SeleccionandoNotas = True
 PrimeraColumna = GridDet.LeftCol
 GridDet.RowSel = GridDet.Row: GridDet.Col = 0: GridDet.ColSel = GridDet.Cols - 1
 GridDet.LeftCol = PrimeraColumna
 SeleccionandoNotas = False
End Sub
Private Sub GridDet_Click()
 SeleccionarNota
End Sub
Private Sub GridDet_RowColChange()
 SeleccionarNota
End Sub
Private Sub GridDet_DblClick()
 AbrirNota
End Sub
Private Sub GridDet_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then AbrirNota: KeyCode = 0
End Sub
Private Sub AbrirNota()
 Dim Tipo As String, Numero As String, I As Integer
 If GridDet.Row < 1 Or GridDet.Row >= GridDet.Rows Then Exit Sub
 Tipo = GridDet.TextMatrix(GridDet.Row, 0): Numero = GridDet.TextMatrix(GridDet.Row, 1)
 Cargando = True
 TxtBuscar.Text = Numero: CmbFiltro.ListIndex = 0
 For I = 1 To CmbFiltro.ListCount - 1
  If Left$(CmbFiltro.List(I), 2) = Tipo Then CmbFiltro.ListIndex = I: Exit For
 Next I
 Cargando = False
 Actualizar True
 If GridFac.Rows > 1 Then GridFac.SetFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos
End Sub

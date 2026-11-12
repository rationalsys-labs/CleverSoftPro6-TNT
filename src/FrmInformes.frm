VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInformes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes personalizados"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   12255
   Begin VB.CommandButton CmdgenExcel 
      Caption         =   "Generar Excel"
      Height          =   420
      Left            =   10395
      TabIndex        =   25
      Top             =   5130
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros: columna, condición y valor"
      Height          =   2760
      Left            =   5445
      TabIndex        =   6
      Top             =   495
      Width           =   6765
      Begin VB.ComboBox CmbCol 
         Height          =   315
         Index           =   0
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   270
         Width           =   1680
      End
      Begin VB.Frame Frame5 
         Height          =   375
         Left            =   90
         TabIndex        =   22
         Top             =   2205
         Width           =   6405
         Begin VB.ComboBox CmbCol 
            Height          =   315
            Index           =   3
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   0
            Width           =   1680
         End
         Begin VB.TextBox TxtFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   3645
            TabIndex        =   24
            Top             =   0
            Width           =   2625
         End
         Begin VB.ComboBox CmbFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.Frame Frame4 
         Height          =   645
         Left            =   90
         TabIndex        =   17
         Top             =   1530
         Width           =   6405
         Begin VB.ComboBox CmbCol 
            Height          =   315
            Index           =   2
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   0
            Width           =   1680
         End
         Begin VB.OptionButton OptO 
            Caption         =   "O"
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   1035
            TabIndex        =   21
            Top             =   360
            Width           =   510
         End
         Begin VB.OptionButton OptY 
            Caption         =   "Y"
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   450
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   510
         End
         Begin VB.TextBox TxtFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   3645
            TabIndex        =   19
            Top             =   0
            Width           =   2625
         End
         Begin VB.ComboBox CmbFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.Frame Frame3 
         Height          =   645
         Left            =   90
         TabIndex        =   12
         Top             =   900
         Width           =   6405
         Begin VB.ComboBox CmbCol 
            Height          =   315
            Index           =   1
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   0
            Width           =   1680
         End
         Begin VB.OptionButton OptO 
            Caption         =   "O"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   1035
            TabIndex        =   16
            Top             =   360
            Width           =   510
         End
         Begin VB.OptionButton OptY 
            Caption         =   "Y"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   450
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   510
         End
         Begin VB.TextBox TxtFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   3645
            TabIndex        =   14
            Top             =   0
            Width           =   2625
         End
         Begin VB.ComboBox CmbFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   6405
         Begin VB.OptionButton OptO 
            Caption         =   "O"
            Height          =   240
            Index           =   0
            Left            =   1035
            TabIndex        =   11
            Top             =   360
            Width           =   510
         End
         Begin VB.OptionButton OptY 
            Caption         =   "Y"
            Height          =   240
            Index           =   0
            Left            =   450
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   510
         End
         Begin VB.TextBox TxtFil 
            Height          =   315
            Index           =   0
            Left            =   3645
            TabIndex        =   9
            Top             =   0
            Width           =   2625
         End
         Begin VB.ComboBox CmbFil 
            Height          =   315
            Index           =   0
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   0
            Width           =   1680
         End
      End
   End
   Begin VB.ComboBox CmbOrigen 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   90
      Width           =   3795
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Informe"
      Height          =   420
      Left            =   10395
      TabIndex        =   3
      Top             =   4680
      Width           =   1770
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar Informe"
      Height          =   420
      Left            =   5445
      TabIndex        =   2
      Top             =   5130
      Width           =   1770
   End
   Begin MSComctlLib.TreeView ViewMenus 
      Height          =   4950
      Left            =   90
      TabIndex        =   0
      Top             =   585
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   8731
      _Version        =   393217
      Indentation     =   90
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   5625
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Origen de Reporte"
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   135
      Width           =   1365
   End
   Begin VB.CommandButton CmdVista
      Caption="Ver resultado"
      Left=5445
      Top=3510
      Width=3195
      Height=420
   End
   Begin VB.CommandButton CmdLimpiar
      Caption="Limpiar filtros"
      Left=8910
      Top=3510
      Width=3255
      Height=420
   End
   Begin VB.Label LblAyuda
      Caption="Marque las columnas. Los filtros Y/O se aplican de izquierda a derecha. Fechas: dd/mm/aaaa."
      Left=5445
      Top=4050
      Width=6630
      Height=495
   End
End
Attribute VB_Name = "FrmInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Preparando As Boolean, PosicionAplicada As Boolean

Private Sub Form_Load()
    On Error GoTo FalloCarga
    Dim Posicion As Integer, Operadores As Variant, Operador As Variant
    Dim Origenes As ADODB.Recordset
    If Not AccesoPermitido(40) Then Unload Me: Exit Sub
    Me.Icon = MDIPri.ICO.ListImages("a22").Picture
    ViewMenus.ImageList = MDIPri.ICO
    Set Origenes = OrigenesInforme()
    Do Until Origenes.EOF
        CmbOrigen.AddItem CStr(Origenes!Nombre): CmbOrigen.ItemData(CmbOrigen.NewIndex) = CLng(Origenes!Origen): Origenes.MoveNext
    Loop
    CerrarRecordset Origenes
    Operadores = Array("Es igual a", "No es igual a", "Es mayor que", "Es mayor o igual a", "Es menor que", "Es menor o igual a", "Comienza por", "No comienza por", "Termina Con", "No termina con", "Contiene", "No Contiene")
    For Posicion = 0 To 3
        For Each Operador In Operadores: CmbFil(Posicion).AddItem CStr(Operador): Next
    Next
    Dim Panel As MSComctlLib.Panel
    For Each Panel In StatusBar1.Panels: Panel.MinWidth = 15: Next
    StatusBar1.Panels(1).Width = 2700: StatusBar1.Panels(2).AutoSize = sbrSpring: StatusBar1.Panels(3).Width = 2400
    HabilitarFiltros
    Exit Sub
FalloCarga:
    MsgBox Err.Description, vbExclamation, App.Title: Unload Me
End Sub

Private Sub CmbOrigen_Click()
    Dim Datos As ADODB.Recordset, Posicion As Integer
    On Error GoTo ErrInt
    Preparando = True: ViewMenus.Nodes.Clear
    For Posicion = 0 To 3
        CmbCol(Posicion).Clear: CmbFil(Posicion).ListIndex = -1: TxtFil(Posicion).Text = ""
        If Posicion < 3 Then OptY(Posicion).Value = True
    Next
    If CmbOrigen.ListIndex < 0 Then Preparando = False: HabilitarFiltros: Exit Sub
    Set Datos = CamposInforme(CmbOrigen.ItemData(CmbOrigen.ListIndex))
    Do Until Datos.EOF
        ViewMenus.Nodes.Add , , CStr(Datos!Campo), TituloCampoInforme(CStr(Datos!Campo)), "a22"
        Datos.MoveNext
    Loop
    CerrarRecordset Datos: Preparando = False: HabilitarFiltros
    Exit Sub
ErrInt:
    CerrarRecordset Datos: ViewMenus.Nodes.Clear: Preparando = False: HabilitarFiltros
    MsgBox Err.Description, vbExclamation, App.Title
End Sub

Public Sub ActualizarColumnas()
    Dim Posicion As Integer, Indice As Integer, Anterior As String, Nodo As MSComctlLib.Node
    Preparando = True
    For Posicion = 0 To 3
        Anterior = CmbCol(Posicion).Text: CmbCol(Posicion).Clear
        For Each Nodo In ViewMenus.Nodes
            If Nodo.Checked Then CmbCol(Posicion).AddItem Nodo.Key
        Next
        For Indice = 0 To CmbCol(Posicion).ListCount - 1
            If CmbCol(Posicion).List(Indice) = Anterior Then CmbCol(Posicion).ListIndex = Indice: Exit For
        Next
        If CmbCol(Posicion).ListIndex < 0 Then CmbFil(Posicion).ListIndex = -1: TxtFil(Posicion).Text = ""
    Next
    Preparando = False: HabilitarFiltros
End Sub

Private Sub HabilitarFiltros()
    Dim Posicion As Integer, Habilitar As Boolean
    Preparando = True
    Habilitar = (CmbOrigen.ListIndex >= 0 And CmbCol(0).ListCount > 0)
    For Posicion = 0 To 3
        CmbCol(Posicion).Enabled = Habilitar: CmbFil(Posicion).Enabled = Habilitar: TxtFil(Posicion).Enabled = Habilitar
        If Habilitar Then CmbFil(Posicion).BackColor = vbWindowBackground: TxtFil(Posicion).BackColor = vbWindowBackground Else CmbFil(Posicion).BackColor = &H80000003: TxtFil(Posicion).BackColor = &H80000003
        If Posicion < 3 Then OptY(Posicion).Enabled = Habilitar: OptO(Posicion).Enabled = Habilitar
        Habilitar = Habilitar And CmbCol(Posicion).ListIndex >= 0 And CmbFil(Posicion).ListIndex >= 0
        If Not Habilitar And Posicion < 3 Then
            CmbCol(Posicion + 1).ListIndex = -1: CmbFil(Posicion + 1).ListIndex = -1: TxtFil(Posicion + 1).Text = ""
        End If
    Next
    Preparando = False
    ActualizarAcciones
End Sub

Private Sub ViewMenus_NodeCheck(ByVal Node As MSComctlLib.Node)
    If Not Preparando Then ActualizarColumnas
End Sub

Private Sub CmbCol_Click(Index As Integer)
    If Not Preparando Then HabilitarFiltros
End Sub

Private Sub CmbFil_Click(Index As Integer)
    If Not Preparando Then HabilitarFiltros
End Sub

Public Function GenerarInforme() As ADODB.Recordset
    Dim Columnas As New Collection, Filtros(0 To 3) As FiltroInforme, Nodo As MSComctlLib.Node, Posicion As Integer, Cantidad As Integer
    For Each Nodo In ViewMenus.Nodes
        If Nodo.Checked Then Columnas.Add Nodo.Key
    Next
    For Posicion = 0 To 3
        If CmbCol(Posicion).Enabled Then
            If CmbCol(Posicion).ListIndex >= 0 Or CmbFil(Posicion).ListIndex >= 0 Or TxtFil(Posicion).Text <> "" Then
                If CmbCol(Posicion).ListIndex < 0 Or CmbFil(Posicion).ListIndex < 0 Then Err.Raise 5, , "Complete el campo y la condicion de cada filtro."
                Filtros(Cantidad).Campo = CmbCol(Posicion).Text
                Filtros(Cantidad).Operador = CmbFil(Posicion).ListIndex + 1: Filtros(Cantidad).Valor = TxtFil(Posicion).Text
                If Posicion < 3 Then Filtros(Cantidad).OpcionO = OptO(Posicion).Value
                Cantidad = Cantidad + 1
            End If
        End If
    Next
    If CmbOrigen.ListIndex < 0 Then Err.Raise 5, , "Seleccione el origen."
    Set GenerarInforme = ConsultarInforme(CmbOrigen.ItemData(CmbOrigen.ListIndex), Columnas, Filtros, Cantidad)
End Function

Public Sub ExportarInforme(Optional ByVal Ruta As String = "")
    Dim Datos As ADODB.Recordset, Numero As Long, Descripcion As String
    On Error GoTo ErrExportar
    Set Datos = GenerarInforme()
    StatusBar1.Panels(3).Text = CStr(Datos.RecordCount) & " registros"
    ExportarInformeExcel Datos, "Clever SoftPro - " & CmbOrigen.Text, Ruta, ContextoInforme
    Datos.Close
    Exit Sub
ErrExportar:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    On Error GoTo 0
    Err.Raise Numero, , Descripcion
End Sub

Private Sub CmdgenExcel_Click()
    On Error GoTo ErrInt
    ExportarInforme
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Command1_Click()
    Dim Datos As ADODB.Recordset, Ruta As String, Mensaje As String
    On Error GoTo Fallo
    Set Datos = GenerarInforme()
    If Datos.RecordCount = 0 Then Err.Raise 5, , "El informe no contiene registros."
    Ruta = ElegirArchivoInforme(Me.hWnd)
    If Len(Ruta) > 0 Then GuardarInformeTexto Datos, Ruta, True
    CerrarRecordset Datos
    Exit Sub
Fallo:
    Mensaje = Err.Description: CerrarRecordset Datos: MsgBox Mensaje, vbExclamation, App.Title
End Sub

Private Sub CmdImprimir_Click()
    Dim Datos As ADODB.Recordset
    On Error GoTo ErrInt
    Set Datos = GenerarInforme()
    If Datos.RecordCount = 0 Then Err.Raise 5, , "El informe no contiene registros."
    If MsgBox("Se imprimiran " & CStr(Datos.RecordCount) & " registros. Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes Then ImprimirInforme Datos, "Clever SoftPro - " & CmbOrigen.Text, ContextoInforme
    Datos.Close
    Exit Sub
ErrInt:
    Dim Descripcion As String
    Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    On Error GoTo 0
    MsgBox Descripcion, vbExclamation, "Clever SoftPro"
End Sub
Private Sub ActualizarAcciones()
    Dim Nodo As MSComctlLib.Node, Cantidad As Long
    For Each Nodo In ViewMenus.Nodes
        If Nodo.Checked Then Cantidad = Cantidad + 1
    Next
    CmdgenExcel.Enabled = (Cantidad > 0): Command1.Enabled = (Cantidad > 0): CmdImprimir.Enabled = (Cantidad > 0): CmdVista.Enabled = (Cantidad > 0)
    If CmbOrigen.ListIndex >= 0 Then StatusBar1.Panels(1).Text = CmbOrigen.Text Else StatusBar1.Panels(1).Text = "Seleccione el origen del informe"
    StatusBar1.Panels(2).Text = CStr(Cantidad) & " columnas | Filtros Y/O de izquierda a derecha"
    StatusBar1.Panels(3).Text = "Consulta pendiente"
End Sub
Private Sub TxtFil_Change(Index As Integer)
    If Not Preparando Then ActualizarAcciones
End Sub
Private Sub OptY_Click(Index As Integer)
    If Not Preparando Then ActualizarAcciones
End Sub
Private Sub OptO_Click(Index As Integer)
    If Not Preparando Then ActualizarAcciones
End Sub
Private Sub Form_Activate()
    If Not PosicionAplicada Then PrepararPosicionVentana Me: PosicionAplicada = True
End Sub

Public Function ContextoInforme() As String
 Dim I As Integer
 ContextoInforme = CmbOrigen.Text
 For I = 0 To 3
  If CmbCol(I).Enabled And CmbCol(I).ListIndex >= 0 And CmbFil(I).ListIndex >= 0 Then
   If I > 0 Then ContextoInforme = ContextoInforme & IIf(OptO(I - 1).Value, " O ", " Y ") Else ContextoInforme = ContextoInforme & ": "
   ContextoInforme = ContextoInforme & CmbCol(I).Text & " " & CmbFil(I).Text & " " & TxtFil(I).Text
  End If
 Next I
End Function
Private Sub CmdVista_Click()
 Dim R As ADODB.Recordset, Mensaje As String
 On Error GoTo Fallo
 Set R = GenerarInforme(): StatusBar1.Panels(3).Text = CStr(R.RecordCount) & " registros"
 Load FrmVistaInforme: FrmVistaInforme.Preparar R, CmbOrigen.Text, ContextoInforme
 CerrarRecordset R: FrmVistaInforme.Show vbModal, MdiPri: Unload FrmVistaInforme
 Exit Sub
Fallo:
 Mensaje = Err.Description: CerrarRecordset R: MsgBox Mensaje, vbExclamation, App.Title
End Sub
Private Sub CmdLimpiar_Click()
 Dim I As Integer
 Preparando = True
 For I = 0 To 3
  CmbCol(I).ListIndex = -1: CmbFil(I).ListIndex = -1: TxtFil(I).Text = ""
  If I < 3 Then OptY(I).Value = True
 Next I
 Preparando = False: HabilitarFiltros
End Sub

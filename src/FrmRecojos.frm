VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmRecojos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion de Recojo de Cheques y Facturas"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   12240
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Height          =   6255
      Left            =   180
      TabIndex        =   2
      Top             =   90
      Width           =   2175
      Begin Clever.Boton CmdConf 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Configuracion General"
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
         MICON           =   "FrmRecojos.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   Begin Clever.Boton CmdNueva
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=720
      Width=1935
      Height=375
      TX="Nueva orden"
   End
   Begin Clever.Boton CmdHistorial
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=1200
      Width=1935
      Height=375
      TX="Historial y envíos"
   End
   Begin Clever.Boton CmdFinalizar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=1620
      Width=1935
      Height=375
      TX="Finalizar gestión"
   End
   Begin Clever.Boton CmdAnular
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=2040
      Width=1935
      Height=375
      TX="Anular orden"
   End
   Begin Clever.Boton CmdActualizar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=2460
      Width=1935
      Height=375
      TX="Actualizar"
   End
   Begin VB.Label LblBuscar
      Left=120
      Top=3000
      Width=1935
      Height=255
      Caption="Solicitud / cliente:"
      BackStyle=0
   End
   Begin VB.TextBox TxtBuscar
      Left=120
      Top=3300
      Width=1935
      Height=315
      MaxLength=100

   End
   Begin VB.ComboBox CmbFiltro
      Left=120
      Top=3750
      Width=1935
      Height=315
      Style=2
   End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   6465
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "22/03/2011"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin Clever.Boton Cmd_Cheq 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      TX              =   "Recojo de Cheques"
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
      MICON           =   "FrmRecojos.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton Cmd_Fac 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      TX              =   "Recojo de Facturas"
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
      MICON           =   "FrmRecojos.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame FraRec 
      Height          =   5775
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   9615
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   7575
         Begin MSFlexGridLib.MSFlexGrid GridOrd 
            Height          =   2505
            Left            =   15
            TabIndex        =   23
            Top             =   120
            Width           =   7500
            _ExtentX        =   13229
            _ExtentY        =   4419
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   16777215
            BackColorFixed  =   16185077
            BackColorBkg    =   16777215
            GridColor       =   15790320
            GridColorFixed  =   15790320
            GridLinesFixed  =   1
            MergeCells      =   1
            AllowUserResizing=   1
            BorderStyle     =   0
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
   Begin VB.Label LblTitle
      Left=120
      Top=2820
      Width=7575
      Height=255
      Caption="NUEVA ORDEN DE RECOJO"
      BackStyle=0
      BackStyle=1
      BackColor=&H0083443B&
      ForeColor=&H00FFFFFF&

   End
      Begin VB.ListBox lstStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDEDE&
         ForeColor       =   &H00000000&
         Height          =   1980
         Left            =   7800
         TabIndex        =   13
         Top             =   3120
         Width           =   1680
      End
      Begin VB.ComboBox CmbMoneda 
         Style = 2
         Height          =   315
         ItemData        =   "FrmRecojos.frx":0054
         Left            =   6240
         List            =   "FrmRecojos.frx":005E
         TabIndex        =   9
         Top             =   5280
         Width           =   1455
      End
      Begin VB.TextBox TxtDestino 
         MaxLength = 1000
         BackColor       =   &H00DEDEDE&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   5280
         Width           =   6015
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   7575
   Begin VB.Label SkinLabel1
      Left=120
      Top=240
      Width=615
      Height=255
      Caption="Cliente:"
      BackStyle=0

   End
         Begin VB.ComboBox CmbDiv 
            Height          =   315
            ItemData        =   "FrmRecojos.frx":00ED
            Left            =   4920
            List            =   "FrmRecojos.frx":00EF
            Style = 2
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.ListBox LstFac 
            Appearance      =   0  'Flat
            Height          =   1395
            Left            =   4920
            TabIndex        =   16
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox TxtFac 
            MaxLength = 14
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5880
            TabIndex        =   15
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtCliente 
            MaxLength = 100
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   840
            TabIndex        =   12
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox txtinfo 
            MaxLength = 4000
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   1395
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   600
            Width           =   4695
         End
      End
      Begin Clever.Boton CmdReenviar 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   315
         Left            =   7800
         TabIndex        =   10
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         TX              =   "Re-Enviar Orden"
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
         MICON           =   "FrmRecojos.frx":00F1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdEnviar 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   315
         Left            =   7800
         TabIndex        =   14
         Top             =   5280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         TX = "Registrar y enviar"
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
         MICON           =   "FrmRecojos.frx":010D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   Begin VB.Label LblMoneda
      Left=6240
      Top=5100
      Width=1455
      Height=255
      Caption="Moneda:"
      BackStyle=0
   End
   End
   Begin VB.Label LblModal 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   9720
      TabIndex        =   18
      Top             =   120
      Visible = 0
      Width           =   375
   End
   Begin VB.Label LblCod 
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   7680
      TabIndex        =   0
      Top             =   6480
      Visible = 0
      Width           =   1095
   End
End
Attribute VB_Name = "FrmRecojos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Datos As ADODB.Recordset, Cliente As ADODB.Recordset, Docs As Object, Clave As String, Modalidad As String, IdGuardado As Long, Seleccionando As Boolean, Iniciando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(14) Then Unload Me: Exit Sub
 Iniciando = True: StatusBar1.Style = sbrSimple
 CmbFiltro.Clear: CmbFiltro.AddItem "Todos": CmbFiltro.AddItem "Abierto": CmbFiltro.AddItem "Cerrado": CmbFiltro.AddItem "Anulado": CmbFiltro.ListIndex = 0
 CmbDiv.Clear: CmbDiv.AddItem "FT": CmbDiv.AddItem "BV": CmbDiv.ListIndex = 0
 CobCombo CmbMoneda, 14, 3
 Set Docs = CobNuevosDocumentos: Iniciando = False
 Nueva "CHEQUE": Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title: Unload Me
End Sub
Private Sub Nueva(ByVal Tipo As String)
 Dim Cfg As ADODB.Recordset
 Modalidad = Tipo: CerrarRecordset Cliente: Set Docs = CobNuevosDocumentos
 Clave = NuevaClaveQuery: IdGuardado = 0
 TxtCliente.Text = "": txtinfo.Text = "": TxtFac.Text = "": LstFac.Clear: LblCod.Caption = ""
 Frame3.Enabled = True: TxtDestino.Enabled = True: CmdEnviar.Enabled = True
 CmbDiv.Enabled = Tipo = "FACTURA": TxtFac.Enabled = Tipo = "FACTURA": LstFac.Enabled = Tipo = "FACTURA"
 LblTitle.Caption = "NUEVA ORDEN - RECOJO DE " & IIf(Tipo = "FACTURA", "FACTURAS", "CHEQUES")
 Set Cfg = CobConfig(14)
 If Tipo = "FACTURA" Then TxtDestino.Text = CStr(Cfg!DestinoFacturas) Else TxtDestino.Text = CStr(Cfg!DestinoCheques)
 CerrarRecordset Cfg: lstStatus.Clear
End Sub
Public Sub Actualizar()
 Dim E As String
 On Error GoTo Fallo
 If CmbFiltro.ListIndex > 0 Then E = CmbFiltro.Text
 CerrarRecordset Datos: Set Datos = CobLeer(14, 0, 0, 0, Trim$(TxtBuscar.Text), E)
 CobGrilla GridOrd, Datos, "Orden|Fecha|Cliente|Modalidad|Estado|Último envío", "Asient|FecEnv|Cliente|ModGes|EstadoGestion|EstadoEnvio"
 StatusBar1.SimpleText = CStr(Datos.RecordCount) & " órdenes visibles (máximo 200)."
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Cmd_Cheq_Click()
 On Error GoTo Fallo
 Nueva "CHEQUE"
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Cmd_Fac_Click()
 On Error GoTo Fallo
 Nueva "FACTURA"
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdNueva_Click()
 If Modalidad = "FACTURA" Then Cmd_Fac_Click Else Cmd_Cheq_Click
End Sub
Private Sub TxtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If KeyCode <> vbKeyF9 And KeyCode <> vbKeyReturn Then Exit Sub
 KeyCode = 0: Set R = CobBuscarCliente(14, TxtCliente.Text, Me)
 If R Is Nothing Then Exit Sub
 CerrarRecordset Cliente: Set Cliente = R: Set Docs = CobNuevosDocumentos: LstFac.Clear
 TxtCliente.Text = CStr(R!CLIENT): LblCod.Caption = CStr(R!Asient)
 txtinfo.Text = "Solicitamos gestionar el recojo de " & IIf(Modalidad = "FACTURA", "facturas", "cheques") & "." & vbCrLf & "Cliente: " & CStr(R!CLIENT) & vbCrLf & "Dirección: " & CStr(R!DIRECC) & vbCrLf & "Contacto: " & ValorTexto(R!Contac) & vbCrLf & "Teléfono: " & ValorTexto(R!TELEF1) & vbCrLf & "RUC / DNI: " & CStr(R!RUCDNI) & vbCrLf & "Horario y observaciones: "
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub TxtFac_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo Fallo
 If KeyCode <> vbKeyReturn Then Exit Sub
 KeyCode = 0
 CobAgregarDocumento 14, Cliente, CmbMoneda.Text, CmbDiv.Text, Trim$(TxtFac.Text), Docs, LstFac
 TxtFac.Text = ""
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub LstFac_KeyDown(KeyCode As Integer, Shift As Integer)
 CobRetirarDocumento Docs, LstFac, KeyCode
End Sub
Private Sub CmdEnviar_Click()
 Dim R As ADODB.Recordset, S As String, C As String, I As Long, Guardado As Boolean, Contacto As String
 On Error GoTo Fallo
 If IdGuardado > 0 Then Err.Raise 5, , "La orden ya está guardada. Use Reenviar orden o Nueva orden."
 If Cliente Is Nothing Then Err.Raise 5, , "Seleccione un cliente con F9."
 If TxtCliente.Text <> CStr(Cliente!CLIENT) Then Err.Raise 5, , "Vuelva a seleccionar el cliente con F9."
 If MsgBox("¿Registrar la orden y solicitar el envío del correo a " & Trim$(TxtDestino.Text) & "?", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 Contacto = Left$(ValorTexto(Cliente!Contac), 50)
 If Len(Contacto) = 0 Then Contacto = "Responsable de recojo"
 C = Trim$(txtinfo.Text)
 If Docs.Count > 0 Then
  C = C & vbCrLf & vbCrLf & "DOCUMENTOS:"
  For I = 0 To LstFac.ListCount - 1: C = C & vbCrLf & LstFac.List(I): Next I
 End If
 If Len(C) > 4000 Then Err.Raise 5, , "El mensaje con sus documentos supera 4000 caracteres. Reduzca el texto."
 S = CobXMLBase(Cliente, CobIdCombo(CmbMoneda), "0", Contacto)
 S = S & AtributoXML("Banco", "0") & AtributoXML("Motivo", "0") & AtributoXML("Gastos", "0") & AtributoXML("Cargo", "") & AtributoXML("Manual", "0") & AtributoXML("SinDocumentos", "0") & AtributoXML("Modalidad", Modalidad) & AtributoXML("Referencia", "") & AtributoXML("Destino", Trim$(TxtDestino.Text)) & AtributoXML("Mensaje", C) & ">" & CobXMLDocumentos(Docs) & "</registro>"
 CmdEnviar.Enabled = False
 Set R = CobGuardar(14, Clave, S): IdGuardado = CLng(R!Item): Guardado = True: CerrarRecordset R
 Frame3.Enabled = False: TxtDestino.Enabled = False
 StatusBar1.SimpleText = "Orden guardada. Procesando envío...": DoEvents
 S = CobEnviar(14, IdGuardado): lstStatus.AddItem S
 Actualizar
 MsgBox S, vbInformation, App.Title
 Exit Sub
Fallo:
 If IdGuardado = 0 Then CmdEnviar.Enabled = True
 If Guardado Then MsgBox "La orden quedó guardada. Consulte el historial antes de reenviar.", vbInformation, App.Title Else MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdReenviar_Click()
 Dim Id As Long, S As String
 On Error GoTo Fallo
 Id = CobFila(Datos, GridOrd)
 If MsgBox("¿Solicitar nuevamente el envío de esta orden?", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 S = CobEnviar(14, Id): lstStatus.AddItem S: Actualizar
 MsgBox S, vbInformation, App.Title
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdFinalizar_Click()
 On Error GoTo Fallo
 CobCerrarSeleccionado 14, Datos, GridOrd, False: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdAnular_Click()
 On Error GoTo Fallo
 CobCerrarSeleccionado 14, Datos, GridOrd, True: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdHistorial_Click()
 On Error GoTo Fallo
 CobHistorial 14, CobFila(Datos, GridOrd)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdConf_Click()
 CobMostrarConfig 14
End Sub
Private Sub CmdActualizar_Click()
 Actualizar
End Sub
Private Sub CmbFiltro_Click()
 If Not Iniciando Then Actualizar
End Sub
Private Sub TxtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyF9 Then Actualizar: KeyCode = 0
End Sub
Private Sub GridOrd_Click()
 SeleccionFilaQuery GridOrd, Seleccionando
End Sub
Private Sub GridOrd_RowColChange()
 SeleccionFilaQuery GridOrd, Seleccionando
End Sub
Private Sub GridOrd_DblClick()
 CmdHistorial_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos: CerrarRecordset Cliente
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

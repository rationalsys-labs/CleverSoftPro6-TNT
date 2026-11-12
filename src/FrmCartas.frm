VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCartas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control y Creacion de Cartas de Cobranza"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   14145
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin Clever.Boton CmdNewCart 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Crear Primera Carta"
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
         MICON           =   "FrmCartas.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdFactura 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Generar Segunda Carta"
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
         MICON           =   "FrmCartas.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdReport 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX = "Exportar listado"
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
         MICON           =   "FrmCartas.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdConf 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX = "Configuración"
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
         MICON           =   "FrmCartas.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdReenviar 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX = "Reenviar aviso SMTP"
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
         MICON           =   "FrmCartas.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   Begin Clever.Boton CmdRegenerar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=1380
      Width=1935
      Height=375
      TX="Ver / reimprimir"
   End
   Begin Clever.Boton CmdDocumentos
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=2760
      Width=1935
      Height=375
      TX="Documentos asociados"
   End
   Begin Clever.Boton CmdHistorial
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=3180
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
      Top=3600
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
      Top=4020
      Width=1935
      Height=375
      TX="Anular expediente"
   End
   Begin Clever.Boton CmdActualizar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=120
      Top=4440
      Width=1935
      Height=375
      TX="Actualizar"
   End
   End
   Begin VB.Frame FraNew 
      Caption         =   "Cartas de Cobranza"
      Height          =   2295
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
      Visible = 0
      Width           =   11655
      Begin VB.Frame Frame5 
         Height          =   375
         Left            =   1800
         TabIndex        =   31
         Top             =   1710
         Width           =   5175
   Begin VB.Label SkinLabel7
      Left=30
      Top=120
      Width=5115
      Height=285
      Caption=""
      BackStyle=0

   End
      End
      Begin VB.Frame Frame4 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1710
         Width           =   1575
   Begin VB.Label LblCorrel
      Left=30
      Top=120
      Width=1515
      Height=285
      Caption=""
      BackStyle=0

   End
      End
   Begin VB.Label SkinLabel1
      Left=120
      Top=240
      Width=735
      Height=255
      Caption="Cliente:"
      BackStyle=0

   End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   7080
         TabIndex        =   16
         Top             =   120
         Width           =   4455
         Begin VB.ComboBox TxtRef 
            Style = 2
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox TxtCarg 
            MaxLength = 50
            Height          =   315
            Left            =   1200
            TabIndex        =   10
            Top             =   360
            Width           =   3135
         End
         Begin VB.ComboBox CmbDiv 
            Style = 2
            Height          =   315
            ItemData        =   "FrmCartas.frx":022A
            Left            =   3480
            List            =   "FrmCartas.frx":022C
            TabIndex        =   13
            Text            =   "S/."
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox TxtMont 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength = 18
            TabIndex        =   12
            Top             =   1080
            Width           =   2175
         End
   Begin VB.Label SkinLabel4
      Left=240
      Top=720
      Width=855
      Height=255
      Caption="Referencia:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel5
      Left=240
      Top=360
      Width=735
      Height=255
      Caption="Cargo:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel6
      Left=240
      Top=1080
      Width=855
      Height=255
      Caption="Monto:"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   6855
         Begin VB.CheckBox ChkMail 
            Value = 1
            Caption = "Enviar aviso SMTP al guardar"
            Height          =   375
            Left            =   3720
            TabIndex        =   20
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox TxtAtenc 
            MaxLength = 50
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   9
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox TxtDir 
            MaxLength = 255
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   8
            Top             =   240
            Width           =   5535
         End
   Begin VB.Label SkinLabel2
      Left=240
      Top=240
      Width=735
      Height=255
      Caption="Dirección:"
      BackStyle=0

   End
   Begin VB.Label SkinLabel3
      Left=240
      Top=600
      Width=735
      Height=255
      Caption="Atención:"
      BackStyle=0

   End
      End
      Begin VB.TextBox TxtCli 
         MaxLength = 100
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   6015
      End
      Begin Clever.Boton CmdCrear 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   7560
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Crear Carta"
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
         MICON           =   "FrmCartas.frx":045A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdCancelar 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   9600
         TabIndex        =   17
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
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
         MICON           =   "FrmCartas.frx":0476
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblCodCli 
         BackColor       =   &H00808080&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame FraInf 
      Height = 4335
      Left            =   2400
      TabIndex        =   27
      Top = 720
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid GridCheq 
         Height = 4185
         Left            =   15
         TabIndex        =   28
         Top             =   120
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8652
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   5205
      Width           =   14145
      _ExtentX        =   24950
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
            Object.Width           =   16757
            MinWidth        =   16757
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox CmbFiltro
      Left=2400
      Top=120
      Width=2600
      Height=315
      Style=2
   End
   Begin VB.TextBox TxtBuscar
      Left=5160
      Top=120
      Width=8895
      Height=315
      MaxLength=100

   End
End
Attribute VB_Name = "FrmCartas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Datos As ADODB.Recordset, Cliente As ADODB.Recordset, Docs As Object, Clave As String, IdGuardado As Long, Seleccionando As Boolean, Iniciando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(15) Then Unload Me: Exit Sub
 Iniciando = True: StatusBar1.Style = sbrSimple
 CmbFiltro.Clear: CmbFiltro.AddItem "Todos": CmbFiltro.AddItem "Abierto": CmbFiltro.AddItem "Cerrado": CmbFiltro.AddItem "Anulado": CmbFiltro.ListIndex = 0
 TxtRef.Clear: TxtRef.AddItem "Cheque Rechazado": TxtRef.AddItem "Facturas pendientes de Pago": TxtRef.ListIndex = 1
 CobCombo CmbDiv, 15, 3: Set Docs = CobNuevosDocumentos
 FraNew.Visible = False: Iniciando = False: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title: Unload Me
End Sub
Public Sub Actualizar()
 Dim E As String
 On Error GoTo Fallo
 If CmbFiltro.ListIndex > 0 Then E = CmbFiltro.Text
 CerrarRecordset Datos: Set Datos = CobLeer(15, 0, 0, 0, Trim$(TxtBuscar.Text), E)
 CobGrilla GridCheq, Datos, "Carta|Fecha|Cliente|Referencia|Moneda|Monto|Emisiones|Estado", "Correl|FecCar|CliCar|RefeCar|DivCar|MontCar|Emisiones|EstadoGestion"
 StatusBar1.SimpleText = CStr(Datos.RecordCount) & " cartas visibles (máximo 200). Doble clic: última emisión guardada."
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdNewCart_Click()
 CerrarRecordset Cliente: Set Docs = CobNuevosDocumentos
 Clave = NuevaClaveQuery: IdGuardado = 0
 TxtCli.Text = "": TxtDir.Text = "": TxtAtenc.Text = "": TxtCarg.Text = "": TxtMont.Text = "0.00": LblCodCli.Caption = "": LblCorrel.Caption = "Nueva carta": ChkMail.Value = 1
 FraNew.Visible = True: FraNew.ZOrder 0: GridCheq.Height = 1920
End Sub
Private Sub CmdCancelar_Click()
 FraNew.Visible = False: GridCheq.Height = 4185
End Sub
Private Sub TxtCli_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 If KeyCode <> vbKeyF9 And KeyCode <> vbKeyReturn Then Exit Sub
 KeyCode = 0: Set R = CobBuscarCliente(15, TxtCli.Text, Me)
 If R Is Nothing Then Exit Sub
 CerrarRecordset Cliente: Set Cliente = R: Set Docs = CobNuevosDocumentos
 TxtCli.Text = CStr(R!CLIENT): LblCodCli.Caption = CStr(R!Asient): TxtDir.Text = CStr(R!DIRECC): TxtAtenc.Text = Left$(ValorTexto(R!Contac), 50)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdDocumentos_Click()
 On Error GoTo Fallo
 If Cliente Is Nothing Or Not FraNew.Visible Then Err.Raise 5, , "Abra una nueva carta y seleccione su cliente primero."
 Load FrmDocsCobranzas: FrmDocsCobranzas.Preparar 15, Cliente, CmbDiv.Text, Docs
 FrmDocsCobranzas.Show vbModal, MdiPri
 If FrmDocsCobranzas.Aceptado Then Set Docs = FrmDocsCobranzas.Docs
 Unload FrmDocsCobranzas
 StatusBar1.SimpleText = CStr(Docs.Count) & " documentos asociados a la carta."
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdCrear_Click()
 Dim R As ADODB.Recordset, S As String, Guardado As Boolean
 On Error GoTo Fallo
 If Cliente Is Nothing Then Err.Raise 5, , "Seleccione el cliente con F9."
 If TxtCli.Text <> CStr(Cliente!CLIENT) Then Err.Raise 5, , "Seleccione nuevamente el cliente con F9."
 If MsgBox("¿Guardar la primera carta de cobranza?", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 S = CobXMLBase(Cliente, CobIdCombo(CmbDiv), TxtMont.Text, TxtAtenc.Text)
 S = S & AtributoXML("Direccion", Trim$(TxtDir.Text)) & AtributoXML("Cargo", Trim$(TxtCarg.Text)) & AtributoXML("Referencia", TxtRef.Text) & AtributoXML("Banco", "0") & AtributoXML("Motivo", "0") & AtributoXML("Gastos", "0") & AtributoXML("Manual", "0") & AtributoXML("SinDocumentos", "0") & AtributoXML("Modalidad", "") & ">" & CobXMLDocumentos(Docs) & "</registro>"
 Set R = CobGuardar(15, Clave, S): IdGuardado = CLng(R!Item): Guardado = True: CerrarRecordset R
 CmdCancelar_Click
 Actualizar
 CobReporte 15, IdGuardado, 1
 If ChkMail.Value = 1 Then MsgBox CobEnviar(15, IdGuardado, 1), vbInformation, App.Title
 Exit Sub
Fallo:
 If Guardado Then MsgBox "La carta quedó guardada. Consulte sus emisiones para continuar.", vbInformation, App.Title Else MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdFactura_Click()
 Dim Id As Long, Guardado As Boolean
 On Error GoTo Fallo
 Id = CobFila(Datos, GridCheq)
 If MsgBox("¿Registrar la segunda emisión de esta carta?", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 CobSegunda Id, Datos.Fields("VersionFila").Value: Guardado = True
 Actualizar
 CobReporte 15, Id, 2
 If MsgBox("¿Enviar el aviso SMTP de esta segunda carta?", vbQuestion + vbYesNo, App.Title) = vbYes Then MsgBox CobEnviar(15, Id, 2), vbInformation, App.Title
 Exit Sub
Fallo:
 If Guardado Then MsgBox "La segunda carta quedó emitida. Use Ver / reimprimir.", vbInformation, App.Title Else MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdRegenerar_Click()
 On Error GoTo Fallo
 CobReporte 15, CobFila(Datos, GridCheq)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdReenviar_Click()
 Dim Id As Long, E As Integer
 On Error GoTo Fallo
 Id = CobFila(Datos, GridCheq): E = IIf(CLng(Datos!Emisiones) > 1, 2, 1)
 If MsgBox("¿Solicitar el envío del aviso de esta carta?", vbQuestion + vbYesNo, App.Title) = vbYes Then MsgBox CobEnviar(15, Id, E), vbInformation, App.Title
 Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdConf_Click()
 CobMostrarConfig 15
End Sub
Private Sub CmdHistorial_Click()
 On Error GoTo Fallo
 CobHistorial 15, CobFila(Datos, GridCheq)
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdFinalizar_Click()
 On Error GoTo Fallo
 CobCerrarSeleccionado 15, Datos, GridCheq, False: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdAnular_Click()
 On Error GoTo Fallo
 CobCerrarSeleccionado 15, Datos, GridCheq, True: Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdReport_Click()
 On Error GoTo Fallo
 CobExportar Datos, "Cartas de cobranza", "Correl|FecCar|CliCar|RefeCar|DivCar|MontCar|Emisiones|EstadoGestion"
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
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
Private Sub GridCheq_Click()
 SeleccionFilaQuery GridCheq, Seleccionando
End Sub
Private Sub GridCheq_RowColChange()
 SeleccionFilaQuery GridCheq, Seleccionando
End Sub
Private Sub GridCheq_DblClick()
 CmdRegenerar_Click
End Sub
Private Sub GridCheq_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then CmdRegenerar_Click: KeyCode = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos: CerrarRecordset Cliente
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

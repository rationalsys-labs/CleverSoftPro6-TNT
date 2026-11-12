VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmQuejasGestiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Gestion de Casos"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin Clever.Boton CmdCrear 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX              =   "Crear Nueva Gestion"
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
         MICON           =   "FrmQuejasGestiones.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdDetalle 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         TX = "Ver detalle"
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
         MICON           =   "FrmQuejasGestiones.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
      ForeColor=&H0083693C&
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HERRAMIENTAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   6255
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "23/06/2010"
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
   Begin VB.Frame FraGenGes 
      Height          =   6075
      Left            =   2400
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   8175
      Begin MSFlexGridLib.MSFlexGrid GridSta 
         Height          =   2535
         Left            =   4200
         TabIndex        =   18
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4471
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
      Begin VB.Frame Frame3 
         Caption         =   "Informacion"
         Height          =   1335
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3975
         Begin VB.TextBox TxtContac 
            MaxLength = 50
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox TxtTelef 
            MaxLength = 20
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Contacto:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Telefono:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.TextBox TxtCas 
         MaxLength = 4000
         Appearance      =   0  'Flat
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3120
         Width           =   7935
      End
      Begin Clever.Boton CmdGuardar 
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         TX              =   "Guardar"
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
         MICON           =   "FrmQuejasGestiones.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "Gestion"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label LblCorrel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Caso:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Generar Nueva Gestion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -15
         TabIndex        =   6
         Top             =   60
         Width           =   8175
      End
   Begin Clever.Boton CmdCancelar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=4200
      Top=2640
      Width=1575
      Height=375
      TX="Cancelar"
   End
   End
   Begin VB.Frame TabQuejas 
      Caption         =   "Listado de Quejas por Cliente"
      Height          =   6015
      Left            =   2400
      TabIndex        =   17
      Top             =   120
      Width           =   8175
      Begin MSFlexGridLib.MSFlexGrid GridGestion 
         Height          =   5655
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   9975
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "FrmQuejasGestiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Caso As ADODB.Recordset, Datos As ADODB.Recordset, Estados As ADODB.Recordset
Private IdCaso As Long, Clave As String, Seleccionando As Boolean
Public Sub Preparar(ByVal Id As Long, Optional ByVal Cerrar As Boolean = False)
 Dim I As Long
 IdCaso = Id: Actualizar
 If Cerrar Then
  CmdCrear_Click
  For I = 1 To GridSta.Rows - 1
   If GridSta.TextMatrix(I, 1) = "CCR" Then GridSta.Row = I: SeleccionFilaQuery GridSta, Seleccionando: Exit For
  Next I
 End If
End Sub
Private Sub Actualizar()
 CerrarRecordset Caso: Set Caso = CobLeer(12, 1, IdCaso)
 If Caso.EOF Then Err.Raise 5, , "El caso no existe."
 Me.Caption = CStr(Caso!Asient) & " - " & CStr(Caso!Cliente): LblCorrel.Caption = CStr(Caso!Asient)
 CerrarRecordset Datos: Set Datos = CobLeer(12, 2, IdCaso)
 CobGrilla GridGestion, Datos, "Estado|Fecha|Descripción|Contacto|Usuario", "CodigoEstado|FecSta|DescripcionEstado|Contac|Usuario"
 CerrarRecordset Estados: Set Estados = CobCatalogo(12, 5): Estados.Filter = "Activo=True"
 CobGrilla GridSta, Estados, "Código|Estado|Descripción", "Item|Codigo|Descripcion"
 GridSta.ColWidth(0) = 0: GridSta.ColWidth(1) = 900: GridSta.ColWidth(2) = 2800
 FraGenGes.Visible = False: TabQuejas.Visible = True
 CmdCrear.Enabled = IsNull(Caso!FecFin): StatusBar1.Style = sbrSimple
 StatusBar1.SimpleText = IIf(IsNull(Caso!FecFin), "Caso abierto. Cada gestión conserva su autor y fecha.", "Caso cerrado. Historial de solo lectura.")
End Sub
Private Sub CmdCrear_Click()
 On Error GoTo Fallo
 If Not IsNull(Caso!FecFin) Then Err.Raise 5, , "El caso está cerrado."
 Clave = NuevaClaveQuery: TxtCas.Text = "": TxtContac.Text = "": TxtTelef.Text = ""
 FraGenGes.Visible = True: FraGenGes.ZOrder 0: TabQuejas.Visible = False
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdGuardar_Click()
 Dim S As String, Id As Long, R As ADODB.Recordset, Guardado As Boolean
 On Error GoTo Fallo
 Id = CobFila(Estados, GridSta)
 If MsgBox("¿Registrar la gestión seleccionada? El estado CCR finaliza el caso.", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 S = "<gestion" & AtributoXML("Caso", CStr(IdCaso)) & AtributoXML("Cliente", CStr(Caso!IdCliente)) & AtributoXML("Tipo", CStr(Caso!TipCas)) & AtributoXML("Version", VersionXML(Caso.Fields("VersionFila").Value)) & AtributoXML("Estado", CStr(Id)) & AtributoXML("Contacto", Trim$(TxtContac.Text)) & AtributoXML("Telefono", Trim$(TxtTelef.Text)) & AtributoXML("Texto", Trim$(TxtCas.Text)) & "/>"
 Set R = CobGuardar(12, Clave, S): Guardado = True: CerrarRecordset R: Actualizar
 Exit Sub
Fallo:
 If Guardado Then MsgBox "La gestión quedó guardada. Recargue el historial.", vbInformation, App.Title Else MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub CmdCancelar_Click()
 FraGenGes.Visible = False: TabQuejas.Visible = True
End Sub
Private Sub CmdDetalle_Click()
 Dim Id As Long
 On Error GoTo Fallo
 Id = CobFila(Datos, GridGestion)
 Load FrmQuejaDetalle: FrmQuejaDetalle.Preparar Caso, Datos
 FrmQuejaDetalle.Show vbModal, Me: Unload FrmQuejaDetalle
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub GridGestion_Click()
 SeleccionFilaQuery GridGestion, Seleccionando
End Sub
Private Sub GridGestion_RowColChange()
 SeleccionFilaQuery GridGestion, Seleccionando
End Sub
Private Sub GridGestion_DblClick()
 CmdDetalle_Click
End Sub
Private Sub GridGestion_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then CmdDetalle_Click: KeyCode = 0
End Sub
Private Sub GridSta_Click()
 SeleccionFilaQuery GridSta, Seleccionando
End Sub
Private Sub GridSta_RowColChange()
 SeleccionFilaQuery GridSta, Seleccionando
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Caso: CerrarRecordset Datos: CerrarRecordset Estados
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

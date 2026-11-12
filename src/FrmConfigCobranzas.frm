VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmConfigCobranzas
 Caption="Configuración de cobranzas y correo automático"
 BorderStyle=1
 ClientWidth=13200
 ClientHeight=8400
 MaxButton=0
 MinButton=0
 ShowInTaskbar=0
 StartUpPosition=1
 BeginProperty Font
  Name="Tahoma"
  Size=8.25
 EndProperty
   Begin Clever.Boton CmdTab
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=240
      Top=180
      Width=4080
      Height=375
      Index=0
      TX="Institución y destinos"
   End
   Begin Clever.Boton CmdTab
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=4440
      Top=180
      Width=4080
      Height=375
      Index=1
      TX="Servidor SMTP"
   End
   Begin Clever.Boton CmdTab
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=8640
      Top=180
      Width=4080
      Height=375
      Index=2
      TX="Destinos por territorio"
   End
 Begin VB.Frame FraPanel
  Index=0
  Left=240
  Top=750
  Width=12720
  Height=6750
  Caption=""
   Begin VB.Label LblCampo0
      Left=180
      Top=450
      Width=2400
      Height=255
      Caption="Empresa:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=390
      Width=9780
      Height=315
      MaxLength=150
      Index=0
   End
   Begin VB.Label LblCampo1
      Left=180
      Top=1110
      Width=2400
      Height=255
      Caption="Cuenta en soles:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=1050
      Width=9780
      Height=315
      MaxLength=150
      Index=1
   End
   Begin VB.Label LblCampo2
      Left=180
      Top=1770
      Width=2400
      Height=255
      Caption="Cuenta en dólares:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=1710
      Width=9780
      Height=315
      MaxLength=150
      Index=2
   End
   Begin VB.Label LblCampo3
      Left=180
      Top=2430
      Width=2400
      Height=255
      Caption="Responsable:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=2370
      Width=9780
      Height=315
      MaxLength=150
      Index=3
   End
   Begin VB.Label LblCampo4
      Left=180
      Top=3090
      Width=2400
      Height=255
      Caption="Correo de contacto:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=3030
      Width=9780
      Height=315
      MaxLength=150
      Index=4
   End
   Begin VB.Label LblCampo5
      Left=180
      Top=3750
      Width=2400
      Height=255
      Caption="Teléfono / anexo:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=3690
      Width=9780
      Height=315
      MaxLength=150
      Index=5
   End
   Begin VB.Label LblCampo6
      Left=180
      Top=4410
      Width=2400
      Height=255
      Caption="Destino recojo cheques:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=4350
      Width=9780
      Height=315
      MaxLength=1000
      Index=6
   End
   Begin VB.Label LblCampo7
      Left=180
      Top=5070
      Width=2400
      Height=255
      Caption="Destino recojo facturas:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=5010
      Width=9780
      Height=315
      MaxLength=1000
      Index=7
   End
   Begin VB.Label LblCampo8
      Left=180
      Top=5730
      Width=2400
      Height=255
      Caption="Destino avisos de cartas:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=5670
      Width=9780
      Height=315
      MaxLength=1000
      Index=8
   End
 End
 Begin VB.Frame FraPanel
  Index=1
  Left=240
  Top=750
  Width=12720
  Height=6750
  Caption=""
   Begin VB.Label LblCampo9
      Left=180
      Top=450
      Width=2400
      Height=255
      Caption="Servidor SMTP:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=390
      Width=9780
      Height=315
      MaxLength=255
      Index=9
   End
   Begin VB.Label LblCampo10
      Left=180
      Top=1110
      Width=2400
      Height=255
      Caption="Puerto STARTTLS:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=1050
      Width=9780
      Height=315
      MaxLength=255
      Index=10
   End
   Begin VB.Label LblCampo11
      Left=180
      Top=1770
      Width=2400
      Height=255
      Caption="Usuario SMTP:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=1710
      Width=9780
      Height=315
      MaxLength=255
      Index=11
   End
   Begin VB.Label LblCampo12
      Left=180
      Top=2430
      Width=2400
      Height=255
      Caption="Correo remitente:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=2370
      Width=9780
      Height=315
      MaxLength=255
      Index=12
   End
   Begin VB.Label LblCampo13
      Left=180
      Top=3090
      Width=2400
      Height=255
      Caption="Clave SMTP:"
      BackStyle=0
   End
   Begin VB.TextBox Txt
      Left=2700
      Top=3030
      Width=9780
      Height=315
      MaxLength=255
      Index=13
      PasswordChar="*"
   End
   Begin VB.CheckBox ChkTLS
      Left=2700
      Top=3600
      Width=6600
      Height=315
      Caption="Usar STARTTLS (recomendado; puerto habitual 587)"
      Value=1
   End
   Begin VB.Label LblClave
      Left=2700
      Top=4140
      Width=9400
      Height=255
      Caption="Deje la clave vacía para conservar la que ya está guardada."
      BackStyle=0
   End
   Begin VB.Label LblEquipo
      Left=2700
      Top=4590
      Width=9400
      Height=255
      Caption="La clave se guarda protegida en este equipo."
      BackStyle=0
   End
 End
 Begin VB.Frame FraPanel
  Index=2
  Left=240
  Top=750
  Width=12720
  Height=6750
  Caption=""
   Begin VB.Label LblTer
      Left=180
      Top=300
      Width=11500
      Height=255
      Caption="Seleccione un territorio existente y configure sus destinatarios separados por punto y coma."
      BackStyle=0
   End
   Begin MSFlexGridLib.MSFlexGrid GridTerritorios
      Left=180
      Top=780
      Width=12300
      Height=3600
      FixedCols=0
   End
   Begin VB.Label LblDestTer
      Left=180
      Top=4650
      Width=11500
      Height=255
      Caption="Destinatarios para el territorio seleccionado (vacío: usa el destino general de cartas):"
      BackStyle=0
   End
   Begin VB.TextBox TxtTerritorio
      Left=180
      Top=5070
      Width=12300
      Height=315
      MaxLength=1000

   End
   Begin Clever.Boton CmdGuardarTerritorio
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=9480
      Top=5730
      Width=3000
      Height=375
      TX="Guardar destinatarios"
   End
 End
   Begin VB.Label LblEstado
      Left=240
      Top=7650
      Width=12500
      Height=255
      Caption=""
      BackStyle=0
   End
   Begin Clever.Boton CmdGuardar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=9360
      Top=7980
      Width=1680
      Height=375
      TX="Guardar"
   End
   Begin Clever.Boton CmdCerrar
      FCOLO=16777215
      FCOL=16777215
      BCOLO=10390910
      BCOL=8612156
      COLTYPE=2
      BTYPE=9
      Left=11280
      Top=7980
      Width=1680
      Height=375
      TX="Cerrar"
   End
End
Attribute VB_Name = "FrmConfigCobranzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Modulo As Integer, Datos As ADODB.Recordset, Territorios As ADODB.Recordset, Seleccionando As Boolean
Public Sub Preparar(ByVal Codigo As Integer)
 Modulo = Codigo: Cargar
End Sub
Private Sub Cargar()
 Dim Campos As Variant, I As Integer
 CerrarRecordset Datos: Set Datos = CobConfig(Modulo)
 Campos = Split("Empresa|CuentaSoles|CuentaDolares|Contacto|CorreoConfirmacion|Anexo|DestinoCheques|DestinoFacturas|DestinoCartas|HostSMTP|PuertoSMTP|UsuarioSMTP|RemitenteSMTP", "|")
 For I = 0 To 12: Txt(I).Text = ValorTexto(Datos.Fields(CStr(Campos(I))).Value): Txt(I).Locked = Not ModUsuario: Next I
 Txt(13).Text = "": Txt(13).Locked = Not ModUsuario
 ChkTLS.Value = IIf(CBool(Datos!UsarTLS), 1, 0): ChkTLS.Enabled = ModUsuario
 LblClave.Caption = IIf(CBool(Datos!TieneClave), "Hay una clave guardada. Déjela vacía para conservarla.", "No hay clave SMTP guardada.")
 LblEquipo.Caption = "Clave protegida en este equipo. Equipo configurado: " & ValorTexto(Datos!EquipoSMTP)
 CmdGuardar.Enabled = ModUsuario: CmdGuardarTerritorio.Enabled = ModUsuario: TxtTerritorio.Locked = Not ModUsuario
 CargarTerritorios
 CmdTab_Click 0
End Sub
Private Sub CargarTerritorios()
 CerrarRecordset Territorios: Set Territorios = CobTerritoriosCorreo(Modulo)
 CobGrilla GridTerritorios, Territorios, "Territorio|Destinatarios", "Territorio|Destinatarios"
 GridTerritorios.ColWidth(0) = 3000: GridTerritorios.ColWidth(1) = 9000
 MostrarTerritorio
End Sub
Private Sub CmdTab_Click(Index As Integer)
 Dim I As Integer
 For I = 0 To 2: FraPanel(I).Visible = I = Index: Next I
 FraPanel(Index).ZOrder 0
End Sub
Private Sub CmdGuardar_Click()
 Dim XML As String, I As Integer, Nombres As Variant
 On Error GoTo Fallo
 If Not IsNumeric(Txt(10).Text) Then Err.Raise 5, , "El puerto SMTP debe ser un número."
 Nombres = Split("Empresa|Soles|Dolares|Contacto|Correo|Anexo|Cheques|Facturas|Cartas|Host|Puerto|Usuario|Remitente", "|")
 XML = "<config"
 For I = 0 To 12: XML = XML & AtributoXML(CStr(Nombres(I)), Trim$(Txt(I).Text)): Next I
 XML = XML & AtributoXML("TLS", CStr(ChkTLS.Value)) & AtributoXML("Equipo", Environ$("COMPUTERNAME")) & "/>"
 CobGuardarConfig Modulo, Datos.Fields("VersionFila").Value, XML, Txt(13).Text
 Cargar
 LblEstado.Caption = "Configuración guardada. Envío SMTP deshabilitado por compatibilidad."
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub MostrarTerritorio()
 Dim Id As Long
 On Error GoTo Vacio
 Id = CobFila(Territorios, GridTerritorios): TxtTerritorio.Text = ValorTexto(Territorios!Destinatarios)
 Exit Sub
Vacio:
 TxtTerritorio.Text = ""
End Sub
Private Sub CmdGuardarTerritorio_Click()
 Dim Id As Long, R As ADODB.Recordset
 On Error GoTo Fallo
 Id = CobFila(Territorios, GridTerritorios)
 Set R = CobTerritoriosCorreo(Modulo, True, Id, Territorios.Fields("VersionFila").Value, Trim$(TxtTerritorio.Text))
 CerrarRecordset R: CargarTerritorios: LblEstado.Caption = "Destinatarios del territorio guardados."
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub GridTerritorios_Click()
 SeleccionFilaQuery GridTerritorios, Seleccionando: MostrarTerritorio
End Sub
Private Sub GridTerritorios_RowColChange()
 If Seleccionando Then Exit Sub
 SeleccionFilaQuery GridTerritorios, Seleccionando: MostrarTerritorio
End Sub
Private Sub CmdCerrar_Click()
 Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos: CerrarRecordset Territorios
End Sub

Private Sub Form_Activate()
 If PosicionInicialAplicada Then Exit Sub
 PrepararPosicionVentana Me
 PosicionInicialAplicada = True
End Sub

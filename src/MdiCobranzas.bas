Attribute VB_Name = "MdiCobranzas"
Option Explicit
Private Type BLOB_COB
 Cantidad As Long
 Puntero As Long
End Type
Private Declare Function CryptProtectData Lib "crypt32.dll" (ByRef Entrada As BLOB_COB, ByVal Descripcion As Long, ByVal Entropia As Long, ByVal Reservado As Long, ByVal Aviso As Long, ByVal Opciones As Long, ByRef Salida As BLOB_COB) As Long
Private Declare Sub CopiarCob Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destino As Any, ByVal Origen As Long, ByVal Cantidad As Long)
Private Declare Function LocalFreeCob Lib "kernel32" Alias "LocalFree" (ByVal Memoria As Long) As Long
Private RCheque As ADODB.Recordset, RCarta1 As ADODB.Recordset, RCarta2 As ADODB.Recordset, RCarta3 As ADODB.Recordset, RDevol As ADODB.Recordset

Public Function CobLeer(ByVal Modulo As Integer, Optional ByVal Seccion As Integer = 0, Optional ByVal Item As Long = 0, Optional ByVal Cliente As Long = 0, Optional ByVal Buscar As String = "", Optional ByVal Estado As String = "", Optional ByVal Desde As Variant, Optional ByVal Hasta As Variant) As ADODB.Recordset
 If IsMissing(Desde) Then Desde = Null
 If IsMissing(Hasta) Then Hasta = Null
 Set CobLeer = ConsultaOperacion("dbo.CobConsultar", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Seccion", adInteger, 0, Seccion), Array("@Item", adInteger, 0, Item), Array("@Cliente", adInteger, 0, Cliente), Array("@Buscar", adVarWChar, 100, Buscar), Array("@Estado", adVarWChar, 10, Estado), Array("@Desde", adDBTimeStamp, 0, Desde), Array("@Hasta", adDBTimeStamp, 0, Hasta)), True)
End Function
Public Function CobCatalogo(ByVal Modulo As Integer, ByVal Tipo As Integer) As ADODB.Recordset
 Set CobCatalogo = ConsultaOperacion("dbo.CobCatalogos", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Catalogo", adInteger, 0, Tipo)), True)
End Function
Public Function CobClientes(ByVal Modulo As Integer, Optional ByVal Buscar As String = "", Optional ByVal Item As Long = 0) As ADODB.Recordset
 Set CobClientes = ConsultaOperacion("dbo.CobBuscarClientes", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Buscar", adVarWChar, 100, Buscar), Array("@Item", adInteger, 0, Item)), True)
End Function
Public Function CobBuscarCliente(ByVal Modulo As Integer, ByVal Texto As String, ByVal Duenio As Form) As ADODB.Recordset
 Dim Id As Long
 Load FrmBuscarClienteCobranza
 FrmBuscarClienteCobranza.Preparar Modulo, Texto
 If Duenio.MDIChild Then FrmBuscarClienteCobranza.Show vbModal, MdiPri Else FrmBuscarClienteCobranza.Show vbModal, Duenio
 Id = FrmBuscarClienteCobranza.IdElegido
 Unload FrmBuscarClienteCobranza
 If Id > 0 Then Set CobBuscarCliente = CobClientes(Modulo, "", Id)
End Function
Public Function CobDoc(ByVal Modulo As Integer, ByVal Tipo As String, ByVal Numero As String) As ADODB.Recordset
 Set CobDoc = ConsultaOperacion("dbo.CobDocumento", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Tipo", adVarWChar, 2, Tipo), Array("@Numero", adVarWChar, 50, Trim$(Numero))), True)
End Function
Public Function CobGuardar(ByVal Modulo As Integer, ByVal Clave As String, ByVal XML As String) As ADODB.Recordset
 Set CobGuardar = ConsultaOperacion(IIf(Modulo = 12, "dbo.CobGuardarQueja", "dbo.CobGuardarRegistro"), CobParametrosGuardar(Modulo, Clave, XML), True)
End Function
Private Function CobParametrosGuardar(ByVal Modulo As Integer, ByVal Clave As String, ByVal XML As String) As Variant
 If Modulo = 12 Then
  CobParametrosGuardar = Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Clave", adVarChar, 36, Clave), Array("@Solicitud", adLongVarWChar, Len(XML), XML))
 Else
  CobParametrosGuardar = Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Clave", adVarChar, 36, Clave), Array("@Solicitud", adLongVarWChar, Len(XML), XML))
 End If
End Function
Public Sub CobCerrarSeleccionado(ByVal Modulo As Integer, ByVal Datos As ADODB.Recordset, ByVal Grilla As MSFlexGrid, ByVal Anular As Boolean)
 Dim Id As Long, Version As Variant
 Id = CobFila(Datos, Grilla)
 Version = Datos.Fields("VersionFila").Value
 CobCerrar Modulo, Id, Version, Anular
End Sub
Public Sub CobCerrar(ByVal Modulo As Integer, ByVal Id As Long, ByVal Version As Variant, ByVal Anular As Boolean)
 Dim R As ADODB.Recordset, Motivo As String
 Load FrmMotivoQuery
 FrmMotivoQuery.Caption = IIf(Anular, "Anular expediente", "Finalizar gestión")
 FrmMotivoQuery.LblMotivo.Caption = IIf(Anular, "Motivo de anulación:", "Resultado confirmado de la gestión:")
 FrmMotivoQuery.Aceptado = False: FrmMotivoQuery.TxtMotivo.Text = ""
 FrmMotivoQuery.Show vbModal, MdiPri
 If Not FrmMotivoQuery.Aceptado Then Unload FrmMotivoQuery: Exit Sub
 Motivo = Trim$(FrmMotivoQuery.TxtMotivo.Text): Unload FrmMotivoQuery
 If MsgBox(IIf(Anular, "¿Anular este expediente?", "¿Finalizar esta gestión?") & " Se conservarán sus datos e historial.", vbQuestion + vbYesNo, App.Title) <> vbYes Then Exit Sub
 Set R = ConsultaOperacion("dbo.CobCerrarRegistro", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Item", adInteger, 0, Id), Array("@Version", adBinary, 8, Version), Array("@Anular", adBoolean, 0, Anular), Array("@Motivo", adVarWChar, 1000, Motivo)), True)
 CerrarRecordset R
End Sub
Public Sub CobSegunda(ByVal Id As Long, ByVal Version As Variant)
 Dim R As ADODB.Recordset
 Set R = ConsultaOperacion("dbo.CobSegundaCarta", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Item", adInteger, 0, Id), Array("@Version", adBinary, 8, Version)), True)
 CerrarRecordset R
End Sub
Public Function CobIdCombo(ByVal C As ComboBox) As Long
 If C.ListIndex < 0 Then Err.Raise 5, , "Seleccione un valor de " & C.Name & "."
 CobIdCombo = C.ItemData(C.ListIndex)
End Function
Public Function CobFechaGiro(ByVal Texto As String, ByRef Fecha As Date) As Boolean
 Dim P As Variant
 On Error GoTo Fallo
 If Not Trim$(Texto) Like "##/##/####" Then Exit Function
 P = Split(Trim$(Texto), "/"): Fecha = DateSerial(CInt(P(2)), CInt(P(1)), CInt(P(0)))
 CobFechaGiro = Year(Fecha) = CInt(P(2)) And Month(Fecha) = CInt(P(1)) And Day(Fecha) = CInt(P(0)) And Year(Fecha) >= 1900
Fallo:
End Function
Public Sub CobCombo(ByVal C As ComboBox, ByVal Modulo As Integer, ByVal Tipo As Integer)
 Dim R As ADODB.Recordset
 Set R = CobCatalogo(Modulo, Tipo): C.Clear
 Do While Not R.EOF
  C.AddItem CStr(R!Descripcion): C.ItemData(C.NewIndex) = CLng(R!Item): R.MoveNext
 Loop
 CerrarRecordset R
 If C.ListCount > 0 Then C.ListIndex = 0
End Sub
Public Function CobFila(ByVal Datos As ADODB.Recordset, ByVal Grilla As MSFlexGrid, Optional ByVal Campo As String = "Item") As Long
 If Datos Is Nothing Then Err.Raise 5, , "No hay datos disponibles."
 If Datos.RecordCount = 0 Or Grilla.Rows < 2 Or Grilla.Row < 1 Then Err.Raise 5, , "Seleccione una fila."
 CobFila = Grilla.RowData(Grilla.Row)
 Datos.MoveFirst: Datos.Find Campo & "=" & CStr(CobFila)
 If Datos.EOF Then Err.Raise 5, , "Recargue la lista antes de continuar."
End Function
Public Sub CobGrilla(ByVal G As MSFlexGrid, ByVal R As ADODB.Recordset, ByVal Cabecera As String, ByVal Campos As String, Optional ByVal Id As String = "Item")
 Dim N As Variant, S As String, V As Variant, Ocupado As Boolean, I As Integer
 GrillaQuery G, Cabecera
 G.BackColor = vbWhite: G.BackColorBkg = vbWhite
 Do While Not R.EOF
  S = ""
  For Each N In Split(Campos, "|")
   V = R.Fields(CStr(N)).Value
   If Len(S) > 0 Then S = S & vbTab
   If IsNull(V) Then
    S = S & " "
   ElseIf VarType(V) = vbDate Then
    S = S & FechaDocumento(V)
   ElseIf VarType(V) = vbCurrency Then
    S = S & Format$(V, "#,##0.00")
   Else
    S = S & CeldaDocumento(V)
   End If
  Next N
  G.AddItem S: G.RowData(G.Rows - 1) = CLng(R.Fields(Id).Value): R.MoveNext
 Loop
 For I = 0 To G.Cols - 1: G.ColWidth(I) = 1500: Next I
 If G.Cols > 2 Then G.ColWidth(2) = 3400
 If G.Rows > 1 Then G.Row = 1: SeleccionFilaQuery G, Ocupado
End Sub
Public Function CobXMLBase(ByVal Cliente As ADODB.Recordset, ByVal Moneda As Long, ByVal Monto As String, ByVal Contacto As String) As String
 ' No usar IIf con campos: VB6 evalúa ambas ramas y un destinatario manual no tiene recordset.
 CobXMLBase = "<registro"
 If Cliente Is Nothing Then
  CobXMLBase = CobXMLBase & AtributoXML("Cliente", "0") & AtributoXML("VersionCliente", "0x0000000000000000")
 Else
  If Cliente.EOF Then Err.Raise 5, , "Seleccione el cliente con F9."
  CobXMLBase = CobXMLBase & AtributoXML("Cliente", CStr(Cliente!Item)) & AtributoXML("VersionCliente", VersionXML(Cliente.Fields("VersionFila").Value))
 End If
 CobXMLBase = CobXMLBase & AtributoXML("Moneda", CStr(Moneda)) & AtributoXML("Monto", NumeroXML(Monto)) & AtributoXML("Contacto", Trim$(Contacto)) & AtributoXML("Fecha", Format$(Date, "yyyy-mm-dd"))
End Function
Public Function CobNuevosDocumentos() As Object
 Set CobNuevosDocumentos = CreateObject("Scripting.Dictionary")
End Function
Public Sub CobAgregarDocumento(ByVal Modulo As Integer, ByVal DatosCliente As ADODB.Recordset, ByVal Moneda As String, ByVal Tipo As String, ByVal Numero As String, ByVal Docs As Object, ByVal Lista As ListBox)
 Dim R As ADODB.Recordset, Id As Long, M As String
 On Error GoTo Fallo
 If DatosCliente Is Nothing Then Err.Raise 5, , "Seleccione primero el cliente con F9."
 If Docs.Count >= IIf(Modulo = 16, 6, 200) Then Err.Raise 5, , "Se alcanzó el máximo de documentos."
 Set R = CobDoc(Modulo, Tipo, Numero)
 If R.EOF Then Err.Raise 5, , "No se encontró el documento importado."
 If CLng(R!IdCliente) <> CLng(DatosCliente!Item) Or CStr(R!Divisa) <> Moneda Or CBool(R!Anulado) Then Err.Raise 5, , "El documento debe estar vigente y pertenecer al cliente y moneda seleccionados."
 Id = CLng(R!IdDocumento)
 If Docs.Exists(CStr(Id)) Then Err.Raise 5, , "El documento ya está en la lista."
 Docs.Add CStr(Id), Array(VersionXML(R.Fields("VersionFila").Value), Tipo & " " & CStr(R!Documento), CCur(R!Importe))
 Lista.AddItem Tipo & " " & CStr(R!Documento): Lista.ItemData(Lista.NewIndex) = Id
 CerrarRecordset R
 Exit Sub
Fallo:
 M = Err.Description: CerrarRecordset R: Err.Raise 5, , M
End Sub
Public Function CobXMLDocumentos(ByVal Docs As Object) As String
 Dim Id As Variant
 For Each Id In Docs.Keys
  CobXMLDocumentos = CobXMLDocumentos & "<d" & AtributoXML("Id", CStr(Id)) & AtributoXML("Version", CStr(Docs(Id)(0))) & "/>"
 Next Id
End Function
Public Sub CobRetirarDocumento(ByVal Docs As Object, ByVal L As ListBox, ByRef Tecla As Integer)
 If Tecla <> vbKeyDelete Or L.ListIndex < 0 Then Exit Sub
 Docs.Remove CStr(L.ItemData(L.ListIndex)): L.RemoveItem L.ListIndex: Tecla = 0
End Sub
Public Sub CobExportar(ByVal R As ADODB.Recordset, ByVal Titulo As String, ByVal Campos As String)
 Dim E As Object, L As Object, H As Object, N As Variant, I As Integer, F As Long, V As Variant, M As String
 On Error GoTo Fallo
 If R Is Nothing Then Err.Raise 5, , "No hay datos para exportar."
 If R.RecordCount = 0 Then Err.Raise 5, , "No hay datos para exportar."
 Set E = CreateObject("Excel.Application"): Set L = E.Workbooks.Add: Set H = L.Worksheets(1)
 H.Cells.NumberFormat = "@": H.Cells(1, 1).Value2 = Titulo: I = 1
 For Each N In Split(Campos, "|")
  H.Cells(3, I).Value2 = CStr(N)
  If R.Fields(CStr(N)).Type = adCurrency Then H.Columns(I).NumberFormat = "#,##0.00"
  I = I + 1
 Next N
 R.MoveFirst: F = 4
 Do While Not R.EOF
  I = 1
  For Each N In Split(Campos, "|")
   V = R.Fields(CStr(N)).Value
   If Not IsNull(V) Then
    If VarType(V) = vbDate Then V = FechaDocumento(V)
    H.Cells(F, I).Value2 = V
   End If
   I = I + 1
  Next N
  R.MoveNext: F = F + 1
 Loop
 H.Rows(3).Font.Bold = True: H.Columns.ColumnWidth = 22
 E.Visible = True: E.UserControl = True
 Set H = Nothing: Set L = Nothing: Set E = Nothing
 Exit Sub
Fallo:
 M = Err.Description
 On Error Resume Next
 If Not L Is Nothing Then L.Close False
 If Not E Is Nothing Then E.Quit
 On Error GoTo 0
 Err.Raise 5, , M
End Sub
Public Sub CobHistorial(ByVal Modulo As Integer, ByVal Item As Long)
 Dim R As ADODB.Recordset, T As String
 Set R = CobLeer(Modulo, 4, Item)
 Do While Not R.EOF
  T = T & Format$(R!Fecha, "dd/mm/yyyy hh:nn:ss") & " / " & R!Usuario & " / " & R!Accion & vbCrLf & R!Detalle & vbCrLf & vbCrLf: R.MoveNext
 Loop
 CerrarRecordset R
 If Modulo = 14 Or Modulo = 15 Then
  T = T & "INTENTOS DE CORREO" & vbCrLf
  Set R = CobLeer(Modulo, 7, Item)
  Do While Not R.EOF
   T = T & FechaDocumento(R!Fecha) & " / " & R!Estado & " / " & R!Destinatario & vbCrLf & ValorTexto(R!Detalle) & vbCrLf: R.MoveNext
  Loop
  CerrarRecordset R
 End If
 Load FrmHistorialRP: FrmHistorialRP.Caption = "Historial de cobranza": FrmHistorialRP.TxtHistorial.Text = T
 FrmHistorialRP.Show vbModal, MdiPri
End Sub
Public Function CobConfig(ByVal Modulo As Integer) As ADODB.Recordset
 Set CobConfig = ConsultaOperacion("dbo.CobLeerConfiguracion", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo)), True)
End Function
Public Sub CobMostrarConfig(ByVal Modulo As Integer)
 Load FrmConfigCobranzas: FrmConfigCobranzas.Preparar Modulo: FrmConfigCobranzas.Show vbModal, MdiPri
 Unload FrmConfigCobranzas
End Sub
Public Function CobProtegerClave(ByVal Texto As String) As Variant
 Dim B() As Byte, C() As Byte, E As BLOB_COB, S As BLOB_COB
 If Len(Texto) = 0 Then Err.Raise 5, , "Ingrese la clave SMTP."
 B = Texto: E.Cantidad = UBound(B) + 1: E.Puntero = VarPtr(B(0))
 If CryptProtectData(E, 0, 0, 0, 0, 5, S) = 0 Then Err.Raise 5, , "No se pudo proteger la clave en este equipo."
 ReDim C(0 To S.Cantidad - 1)
 CopiarCob C(0), S.Puntero, S.Cantidad: LocalFreeCob S.Puntero
 Erase B: CobProtegerClave = C
End Function
Public Sub CobGuardarConfig(ByVal Modulo As Integer, ByVal Version As Variant, ByVal XML As String, ByVal NuevaClave As String)
 Dim B As Variant, R As ADODB.Recordset, Tamano As Long
 B = Null: Tamano = 1
 If Len(NuevaClave) > 0 Then B = CobProtegerClave(NuevaClave): Tamano = UBound(B) + 1
 Set R = ConsultaOperacion("dbo.CobGuardarConfiguracion", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Version", adBinary, 8, Version), Array("@Datos", adLongVarWChar, Len(XML), XML), Array("@CambiarClave", adBoolean, 0, Len(NuevaClave) > 0), Array("@ClaveProtegida", adLongVarBinary, Tamano, B)), True)
 CerrarRecordset R
End Sub
Public Function CobTerritoriosCorreo(ByVal Modulo As Integer, Optional ByVal Guardar As Boolean = False, Optional ByVal Id As Long = 0, Optional ByVal Version As Variant, Optional ByVal Correos As String = "") As ADODB.Recordset
 If IsMissing(Version) Then Version = Null
 Set CobTerritoriosCorreo = ConsultaOperacion("dbo.CobCorreosTerritorio", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Guardar", adBoolean, 0, Guardar), Array("@IdTerritorio", adInteger, 0, Id), Array("@Version", adBinary, 8, Version), Array("@Correos", adVarWChar, 1000, Correos)), True)
End Function
Public Function CobEnviar(ByVal Modulo As Integer, ByVal Item As Long, Optional ByVal Emision As Integer = 1) As String
 CobEnviar = "Envío SMTP deshabilitado por compatibilidad."
End Function

Public Function CobAtributo(ByVal X As Object, ByVal Nombre As String) As String
 Dim N As Object
 Set N = X.selectSingleNode("/registro/@" & Nombre)
 If Not N Is Nothing Then CobAtributo = N.Text
End Function
Private Function CobEmpresa(ByVal X As Object, ByVal Campo As String) As String
 Dim N As Object
 Set N = X.selectSingleNode("/registro/empresa/" & Campo)
 If Not N Is Nothing Then CobEmpresa = N.Text
End Function
Private Function CobCuenta(ByVal X As Object) As String
 If CobAtributo(X, "Moneda") = "1" Then CobCuenta = CobEmpresa(X, "CuentaSoles")
 If CobAtributo(X, "Moneda") = "2" Then CobCuenta = CobEmpresa(X, "CuentaDolares")
End Function
Private Function CobCamposReporte(ByVal Cantidad As Integer, Optional ByVal Espacio As Boolean = True) As ADODB.Recordset
 Dim R As New ADODB.Recordset, I As Integer
 If Espacio Then
  For I = 1 To Cantidad: R.Fields.Append "REGISTRO " & CStr(I), adVarWChar, 1000: Next I
 Else
  R.Fields.Append "REGISTRO", adVarWChar, 1000
 End If
 R.Open: Set CobCamposReporte = R
End Function
Public Sub CobReporte(ByVal Modulo As Integer, ByVal Item As Long, Optional ByVal Emision As Integer = 0)
 Dim Q As ADODB.Recordset, R As ADODB.Recordset, Docs As ADODB.Recordset, X As Object, Informe As Object, C As Object
 Dim V As Variant, I As Integer, Texto As String, Cuenta As String, T As String, M As String, Fecha As Date, Numero As String
 On Error GoTo Fallo
 Set Q = CobLeer(Modulo, 0, Item)
 If Q.EOF Then Err.Raise 5, , "No se encontró el expediente."
 If IsNull(Q!Solicitud) Then Err.Raise 5, , "El registro histórico no tiene fotografía del formato. Requiere revisión de sus datos."
 Set X = XmlQuery(CStr(Q!Solicitud))
 If Modulo = 13 Then
  Set R = CobCamposReporte(9)
  R.AddNew Array("REGISTRO 1", "REGISTRO 2", "REGISTRO 3", "REGISTRO 4", "REGISTRO 5", "REGISTRO 6", "REGISTRO 7", "REGISTRO 8", "REGISTRO 9"), Array(Q!Contac, Q!Cargo, CobAtributo(X, "BancoNombre"), FechaDocumento(Q!FecGir), Q!NroChe, CStr(Q!Divisa) & " " & Format$(Q!Monto, "#,##0.00"), CobAtributo(X, "MotivoNombre"), CStr(Q!Divisa) & " " & Format$(Q!Gastos, "#,##0.00"), CStr(Q!Divisa) & " " & Format$(Q!MonCan, "#,##0.00"))
  Unload RepChequeCobranza: CerrarRecordset RCheque: Set RCheque = R: Set Informe = RepChequeCobranza
  Set Informe.DataSource = RCheque
  With Informe.Sections("Section1").Controls
   .Item("Lblfec").Caption = Format$(Q!FecCar, "Long Date")
   .Item("LblCont").Caption = CStr(Year(Q!FecCar)) & " - " & Q!Asient
   .Item("LBLCLI").Caption = Q!Client: .Item("LBLRUC").Caption = Q!NroRUC: .Item("LBLDIR").Caption = Q!Direcc
   Cuenta = CobCuenta(X)
   Texto = "Sírvase coordinar la regularización del monto indicado con el área de cobranzas."
   If Len(Cuenta) > 0 Then Texto = Texto & " Cuenta: " & Cuenta & "."
   Texto = Texto & " " & CobEmpresa(X, "CorreoConfirmacion") & " " & CobEmpresa(X, "Anexo")
   .Item("LBLCUERPO").Caption = Texto
  End With
 ElseIf Modulo = 15 Then
  Set Docs = CobLeer(15, 5, Item)
  If Docs.EOF Then Err.Raise 5, , "La carta no tiene emisiones guardadas."
  If Emision = 0 Then Docs.MoveLast Else Docs.Find "Numero=" & CStr(Emision)
  If Docs.EOF Then Err.Raise 5, , "No se encontró esa emisión de carta."
  Fecha = CDate(Docs!Fecha): Set X = XmlQuery(CStr(Docs!Datos))
  Set R = CobCamposReporte(5)
  R.AddNew Array("REGISTRO 1", "REGISTRO 2", "REGISTRO 3", "REGISTRO 4", "REGISTRO 5"), Array(Q!CliCar, Q!DirCar, Q!ContCar, Q!CargCar, Q!RefeCar)
  Select Case CInt(Docs!Modelo)
   Case 1: Unload RepCartaCheque: CerrarRecordset RCarta1: Set RCarta1 = R: Set Informe = RepCartaCheque: Set Informe.DataSource = RCarta1
   Case 2: Unload RepCartaFactura: CerrarRecordset RCarta2: Set RCarta2 = R: Set Informe = RepCartaFactura: Set Informe.DataSource = RCarta2
   Case 3: Unload RepCartaSegunda: CerrarRecordset RCarta3: Set RCarta3 = R: Set Informe = RepCartaSegunda: Set Informe.DataSource = RCarta3
  End Select
  With Informe.Sections("Cuerpo").Controls
   .Item("LblFec").Caption = Format$(Fecha, "Long Date") & " / " & Q!Correl & " / Emisión " & CStr(Docs!Numero)
   .Item("LbMont").Caption = "por servicios prestados, el importe de " & CStr(Q!DivCar) & " " & Format$(Q!MontCar, "#,##0.00") & ", conforme se"
   .Item("LBLANEXO").Caption = CobEmpresa(X, "Contacto") & " / " & CobEmpresa(X, "CorreoConfirmacion") & " / " & CobEmpresa(X, "Anexo")
   Cuenta = CobCuenta(X)
   .Item("LblCuentas").Caption = IIf(Len(Cuenta) > 0, "Cuenta para regularización: " & Cuenta, "Solicite los datos de pago al responsable de cobranzas.")
  End With
 ElseIf Modulo = 16 Then
  Set Docs = CobLeer(16, 3, Item): Set R = CobCamposReporte(1, False)
  T = "Sin documentos asociados"
  If Not Docs.EOF Then
   T = CStr(Docs!Documento): Docs.MoveNext
  End If
  Do While Not Docs.EOF
   R.AddNew Array("REGISTRO"), Array(CStr(Docs!Documento)): Docs.MoveNext
  Loop
  If R.RecordCount = 0 Then R.AddNew Array("REGISTRO"), Array("")
  Unload RepDevolucionCobranza: CerrarRecordset RDevol: Set RDevol = R: Set Informe = RepDevolucionCobranza: Set Informe.DataSource = RDevol
  With Informe.Sections("PARTE1").Controls
   .Item("LBLMODAL").Caption = "originado en servicios de IMPORTACIÓN"
   .Item("LBLFEC").Caption = FechaDocumento(Q!FecCar): .Item("LBLCLI").Caption = Q!Client
   .Item("LBLGUI").Caption = Q!NroGui: .Item("LBLFAC").Caption = T: .Item("LBLCORREL").Caption = Q!NroCar
   .Item("LBLTRAN").Caption = IIf(CobAtributo(X, "Modalidad") = "TRANSFERENCIA", "X", "")
   .Item("LBLPAGO").Caption = IIf(CobAtributo(X, "Modalidad") <> "TRANSFERENCIA", "X", "")
  End With
  With Informe.Sections("PARTE2").Controls
   .Item("LBLCAN").Caption = Q!Cancel: .Item("LBLABO").Caption = Q!Abonos: .Item("LBLMOT").Caption = Q!Motivo
   .Item("LBLTRA").Caption = Q!Transf: .Item("LBLOTR").Caption = Q!Otros: .Item("LBLDEP").Caption = Q!ParDep
   .Item("LBLCTA").Caption = Q!CtaCte: .Item("LBLMONT").Caption = Format$(Q!Monto, "#,##0.00"): .Item("LBLDIV").Caption = Q!Moneda: .Item("LBLCONT").Caption = Q!Contac
  End With
 Else
  Err.Raise 5, , "Este módulo consulta su historial en pantalla."
 End If
 If CStr(Q!EstadoGestion) = "Anulado" Then
  Informe.Caption = "[ANULADO] " & Informe.Caption
  Select Case Modulo
   Case 13: Set C = Informe.Sections("Section1").Controls("LblCont")
   Case 15: Set C = Informe.Sections("Cuerpo").Controls("LblFec")
   Case 16: Set C = Informe.Sections("PARTE1").Controls("LBLCORREL")
  End Select
  C.Caption = "[ANULADO] " & C.Caption
 End If
 R.MoveFirst: Informe.Show: CentrarReporteMDI Informe
 CerrarRecordset Q: CerrarRecordset Docs
 Exit Sub
Fallo:
 M = Err.Description: CerrarRecordset Q: CerrarRecordset Docs
 MsgBox "No se pudo abrir el reporte: " & M, vbExclamation, App.Title
End Sub

Attribute VB_Name = "MdiQueries"
Option Explicit
Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef Id As Any) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByRef Id As Any, ByVal Texto As Long, ByVal Maximo As Long) As Long
Private ReporteQuery1 As ADODB.Recordset, ReporteQuery2 As ADODB.Recordset

Public Function NuevaClaveQuery() As String
 Dim Id(0 To 15) As Byte, S As String
 If CoCreateGuid(Id(0)) <> 0 Then Err.Raise 5, , "No se pudo preparar la identificación de la solicitud."
 S = String$(39, vbNullChar)
 If StringFromGUID2(Id(0), StrPtr(S), 39) = 0 Then Err.Raise 5, , "No se pudo identificar la solicitud."
 NuevaClaveQuery = Mid$(S, 2, 36)
End Function
Public Function DocumentoQuery(ByVal Tipo As String, ByVal Numero As String, Optional ByVal Modulo As Integer = 5) As ADODB.Recordset
 Tipo = UCase$(Trim$(Tipo)): Numero = UCase$(Trim$(Numero))
 If Len(Numero) = 0 Then Err.Raise 5, , "Ingrese el número del documento."
 If Mid$(Numero, 3, 1) = "-" Then
  If Left$(Numero, 2) <> Tipo Then Err.Raise 5, , "El prefijo del documento no coincide con el tipo seleccionado."
 Else
  Numero = Tipo & "-" & Numero
 End If
 If Len(Numero) > 14 Then Err.Raise 5, , "El documento completo admite hasta 14 caracteres."
 Set DocumentoQuery = ConsultaOperacion("dbo.BuscarDocumentoQuery", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Tipo", adVarWChar, 2, Tipo), Array("@Numero", adVarWChar, 50, Trim$(Numero))), True)
End Function
Public Function CatalogoQuery(ByVal Catalogo As Integer, Optional ByVal Modulo As Integer = 5) As ADODB.Recordset
 Set CatalogoQuery = ConsultaOperacion("dbo.ConsultarCatalogosQuery", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Catalogo", adInteger, 0, Catalogo)), True)
End Function
Public Function ListaQueries(Optional ByVal Modulo As Integer = 6, Optional ByVal Buscar As String = "", Optional ByVal Estado As String = "", Optional ByVal Desde As Variant, Optional ByVal Hasta As Variant, Optional ByVal Usuario As Long = 0) As ADODB.Recordset
 If IsMissing(Desde) Then Desde = Null
 If IsMissing(Hasta) Then Hasta = Null
 Set ListaQueries = ConsultaOperacion("dbo.ConsultarQueries", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Buscar", adVarWChar, 100, Buscar), Array("@Estado", adVarWChar, 10, Estado), Array("@Desde", adDBTimeStamp, 0, Desde), Array("@Hasta", adDBTimeStamp, 0, Hasta), Array("@Usuario", adInteger, 0, Usuario)), True)
End Function
Public Function DatosQuery(ByVal Item As Long, Optional ByVal Seccion As Integer = 0, Optional ByVal Modulo As Integer = 6) As ADODB.Recordset
 Set DatosQuery = ConsultaOperacion("dbo.DetalleQuery", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Item", adInteger, 0, Item), Array("@Seccion", adInteger, 0, Seccion)), True)
End Function
Public Function ServiciosQuery(ByVal Documento As Long, Optional ByVal Modulo As Integer = 5) As ADODB.Recordset
 Set ServiciosQuery = ConsultaOperacion("dbo.ServiciosDocumentoQuery", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Modulo", adInteger, 0, Modulo), Array("@Documento", adInteger, 0, Documento)), True)
End Function
Public Function NotasQuery(ByVal Item As Long, ByVal Buscar As String) As ADODB.Recordset
 Set NotasQuery = ConsultaOperacion("dbo.BuscarNotasQuery", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Query", adInteger, 0, Item), Array("@Buscar", adVarWChar, 50, Buscar)), True)
End Function
Public Function GuardarQuery(ByVal Clave As String, ByVal XML As String) As ADODB.Recordset
 Set GuardarQuery = ConsultaOperacion("dbo.GuardarSolicitudQuery", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Clave", adVarChar, 36, Clave), Array("@Solicitud", adLongVarWChar, Len(XML), XML)), True)
End Function
Public Function GuardarCasoQuery(ByVal Item As Long, ByVal Version As Variant, ByVal Descripcion As String, ByVal Eliminar As Boolean) As ADODB.Recordset
 Set GuardarCasoQuery = ConsultaOperacion("dbo.GuardarOtroCasoQuery", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Item", adInteger, 0, Item), Array("@Version", adBinary, 8, Version), Array("@Descripcion", adVarWChar, 150, Descripcion), Array("@Eliminar", adBoolean, 0, Eliminar)), True)
End Function
Public Sub AnularSolicitudQuery(ByVal Item As Long, ByVal Version As Variant, ByVal Motivo As String)
 Dim R As ADODB.Recordset
 Set R = ConsultaOperacion("dbo.AnularQuery", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Item", adInteger, 0, Item), Array("@Version", adBinary, 8, Version), Array("@Motivo", adVarWChar, 1000, Motivo)), True)
 CerrarRecordset R
End Sub
Public Sub CerrarSolicitudQuery(ByVal Item As Long, ByVal Version As Variant, ByVal Nota As Long, ByVal VersionNota As Variant, ByVal Monto As Currency, ByVal Resolucion As String, ByVal XML As String)
 Dim R As ADODB.Recordset
 Set R = ConsultaOperacion("dbo.FinalizarQuery", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Item", adInteger, 0, Item), Array("@Version", adBinary, 8, Version), Array("@IdNota", adInteger, 0, Nota), Array("@VersionNota", adBinary, 8, VersionNota), Array("@Monto", adCurrency, 0, Monto), Array("@Resolucion", adVarWChar, 1000, Resolucion), Array("@Documentos", adLongVarWChar, Len(XML), XML)), True)
 CerrarRecordset R
End Sub
Public Sub ComboCatalogoQuery(ByVal Combo As ComboBox, ByVal Catalogo As Integer, Optional ByVal Modulo As Integer = 5)
 Dim R As ADODB.Recordset
 Set R = CatalogoQuery(Catalogo, Modulo)
 Combo.Clear
 Do While Not R.EOF
  Combo.AddItem CStr(R!Descripcion): Combo.ItemData(Combo.NewIndex) = CLng(R!Item): R.MoveNext
 Loop
 CerrarRecordset R
 If Combo.ListCount > 0 Then Combo.ListIndex = 0
End Sub
Public Sub SeleccionFilaQuery(ByVal G As MSFlexGrid, ByRef Ocupado As Boolean)
 Dim L As Long
 If Ocupado Or G.Rows < 2 Or G.Row < 1 Then Exit Sub
 Ocupado = True
 L = G.LeftCol
 G.Col = 0: G.RowSel = G.Row: G.ColSel = G.Cols - 1
 G.LeftCol = L: Ocupado = False
End Sub
Public Sub GrillaQuery(ByVal G As MSFlexGrid, ByVal Cabecera As String)
 PrepararGrillaDocumento G, Cabecera
 G.MergeCells = flexMergeNever
 G.BackColorSel = &H83693C: G.ForeColorSel = vbWhite
End Sub
Public Function XmlQuery(ByVal Texto As String) As Object
 Dim D As Object
 Set D = CreateObject("MSXML2.DOMDocument.6.0")
 D.async = False: D.resolveExternals = False: D.setProperty "ProhibitDTD", True
 If Len(Texto) > 0 Then
  If Not D.LoadXML(Texto) Then Err.Raise 5, , "Los datos de la solicitud no se pueden interpretar."
 End If
 Set XmlQuery = D
End Function
Public Function OpcionQuery(ByVal XML As Object, ByVal Nombre As String) As String
 Dim N As Object
 Set N = XML.selectSingleNode("/solicitud/@" & Nombre)
 If Not N Is Nothing Then OpcionQuery = N.Text
End Function
Public Sub VerHistorialQuery(ByVal Item As Long, Optional ByVal Modulo As Integer = 6)
 Dim Q As ADODB.Recordset, R As ADODB.Recordset, T As String, F As Variant
 On Error GoTo Fallo
 Set Q = DatosQuery(Item, 0, Modulo)
 If Q.EOF Then Err.Raise 5, , "La solicitud ya no está disponible."
 T = CStr(Q!Asient) & "  " & CStr(Q!TipQue) & vbCrLf & CStr(Q!CliDoc) & vbCrLf
 T = T & CStr(Q!TipoDocumento) & " " & CStr(Q!NumDoc) & "   " & CStr(Q!DivSol) & " " & Format$(Q!MonSol, "#,##0.00") & vbCrLf
 T = T & "Solicitud: " & FechaDocumento(Q!FecSol) & " / " & ValorTexto(Q!Creador) & vbCrLf & "Caso: " & CStr(Q!TipCas) & vbCrLf & ValorTexto(Q!Proble) & vbCrLf
 T = T & "Adicional: " & ValorTexto(Q!SolAdi) & vbCrLf & "Responsable: " & ValorTexto(Q!SolRes) & " / Servicio: " & ValorTexto(Q!ServicioSolicitud) & vbCrLf
 T = T & "Estado: " & CStr(Q!Estado) & " / " & CStr(Q!SitDoc) & vbCrLf & vbCrLf
 T = T & "SOLUCIÓN: " & ValorTexto(Q!Resolucion) & vbCrLf & "Cierre: " & FechaDocumento(Q!FechaCierre) & " / " & ValorTexto(Q!CerradoPor) & vbCrLf
 T = T & "Nota de crédito: " & ValorTexto(Q!NotaResolucion) & " / Aplicado: " & ValorTexto(Q!MontoAplicado) & vbCrLf & vbCrLf & "DOCUMENTOS DE REEMPLAZO" & vbCrLf
 Set R = DatosQuery(Item, 2, Modulo)
 Do While Not R.EOF
  T = T & CStr(R!Documento) & " / " & FechaDocumento(R!Emision) & " / " & CStr(R!Divisa) & " " & Format$(R!Importe, "#,##0.00") & vbCrLf: R.MoveNext
 Loop
 CerrarRecordset R
 T = T & vbCrLf & "HISTORIAL" & vbCrLf
 Set R = DatosQuery(Item, 1, Modulo)
 Do While Not R.EOF
  T = T & Format$(R!Fecha, "dd/mm/yyyy hh:nn:ss") & " / " & CStr(R!Usuario) & " / " & CStr(R!Accion) & vbCrLf & CStr(R!Detalle) & vbCrLf: R.MoveNext
 Loop
 CerrarRecordset R: CerrarRecordset Q
 Load FrmHistorialRP
 FrmHistorialRP.Caption = "Solución e historial de Query"
 FrmHistorialRP.TxtHistorial.Text = T
 FrmHistorialRP.Show vbModal, MdiPri
 Exit Sub
Fallo:
 T = Err.Description: CerrarRecordset R: CerrarRecordset Q
 MsgBox T, vbExclamation, App.Title
End Sub
Public Function NuevoReporteQuery() As ADODB.Recordset
 Dim R As New ADODB.Recordset, N As Variant
 For Each N In Split("Adicionales AnlDir AnlNc Asiento CamCLiente CamDireccion CamRuc CamTelefono ChkCon1 ChkCon2 ChkCon3 ChkDescuento ChkMal1 ChkMal2 ChkMal3 ChkOtros ChkServ1 ChkServ2 ChkServ3 ChkServ4 ChkServ5 ChkTar1 ChkTar2 ChkTar3 ChkTar4 ChkTar5 Cliente DescripOtros Descuento DivisaDoc Documento EmisionDoc Encargado Fecha FechaCaso FechaVisto Guia1 Guia2 Guia3 Guia4 Guia5 Guia6 ImporteDoc ModCliente ModDireccion ModRuc ModTelefono NoLibGuia OtrDescuento RecargoDoc Responsable Servicio SiLibGuia Territorio Tipo", " ")
  R.Fields.Append CStr(N), adVarWChar, 1000
 Next N
 R.Open
 Set NuevoReporteQuery = R
End Function
Public Sub FilaReporteQuery(ByVal R As ADODB.Recordset, ByVal Q As ADODB.Recordset, ByVal Modulo As Integer)
 Dim X As Object, N As Variant, I As Integer, G As ADODB.Recordset, Prefijo As String
 Set X = XmlQuery(ValorTexto(Q!SolicitudOriginal))
 R.AddNew
 For Each N In R.Fields: N.Value = "": Next N
 R!Asiento = Q!Asient: R!Fecha = FechaDocumento(Q!FecSol): R!FechaCaso = FechaDocumento(Q!FecIni): R!FechaVisto = FechaDocumento(Q!FecVis)
 R!Cliente = Q!CliDoc: R!Documento = Q!NumDoc: R!Tipo = Q!TipoDocumento: R!EmisionDoc = FechaDocumento(Q!FecEmi)
 R!DivisaDoc = Q!DivSol: R!ImporteDoc = Format$(Q!MonSol, "#,##0.00"): R!RecargoDoc = ValorTexto(Q!RcSol)
 R!Territorio = ValorTexto(Q!SolTer): R!Responsable = ValorTexto(Q!SolRes): R!Encargado = ValorTexto(Q!Creador): R!Servicio = ValorTexto(Q!ServicioSolicitud)
 R!Adicionales = ValorTexto(Q!SolAdi)
 If CStr(Q!SitDoc) = "ANL" Then R!Adicionales = "[SOLICITUD ANULADA] " & R!Adicionales
 If CStr(Q!TipQue) = "ANL ANL" Then R!AnlDir = "X"
 If CStr(Q!TipQue) = "ANL NC" Then R!AnlNc = "X"
 If Left$(CStr(Q!TipQue), 3) = "DES" Then R!Descuento = Format$(Q!SolDes, "#,##0.00")
 For Each N In Split("ChkMal1 ChkMal2 ChkMal3 ChkTar1 ChkTar2 ChkTar3 ChkTar4 ChkTar5 ChkServ1 ChkServ2 ChkServ3 ChkServ4 ChkServ5 ChkCon1 ChkCon2 ChkCon3 ChkDescuento SiLibGuia NoLibGuia", " ")
  If OpcionQuery(X, CStr(N)) = "1" Then R.Fields(CStr(N)).Value = "X"
 Next N
 If Val(OpcionQuery(X, "Otro")) > 0 Then R!ChkOtros = "X": R!DescripOtros = ValorTexto(Q!Proble)
 If OpcionQuery(X, "ChkDescuento") = "1" Then R!OtrDescuento = Format$(Q!SolDes, "#,##0.00")
 Prefijo = "Mod"
 If OpcionQuery(X, "CambioCliente") = "1" Then Prefijo = "Cam"
 R.Fields(Prefijo & "Cliente").Value = ValorTexto(Q!SolCli)
 R.Fields(Prefijo & "Ruc").Value = ValorTexto(Q!SolRuc)
 R.Fields(Prefijo & "Direccion").Value = ValorTexto(Q!SolDir)
 R.Fields(Prefijo & "Telefono").Value = ValorTexto(Q!SolTel)
 Set G = DatosQuery(CLng(Q!Item), 3, Modulo): I = 1
 Do While Not G.EOF And I <= 6
  R.Fields("Guia" & CStr(I)).Value = G!NumeroGuia: I = I + 1: G.MoveNext
 Loop
 CerrarRecordset G: R.Update
End Sub
Public Sub MostrarQueries(ByVal Items As Collection, Optional ByVal Modulo As Integer = 6)
 Dim Q As ADODB.Recordset, X As Object, Item As Variant, A As ADODB.Recordset, B As ADODB.Recordset, Mensaje As String
 On Error GoTo Fallo
 If Items.Count = 0 Then Err.Raise 5, , "Seleccione por lo menos una solicitud."
 If Items.Count > 200 Then Err.Raise 5, , "Se permiten hasta 200 solicitudes por reporte."
 Set A = NuevoReporteQuery: Set B = NuevoReporteQuery
 For Each Item In Items
  Set Q = DatosQuery(CLng(Item), 0, Modulo)
  If Q.EOF Then Err.Raise 5, , "Una solicitud ya no está disponible."
  Set X = XmlQuery(ValorTexto(Q!SolicitudOriginal))
  If OpcionQuery(X, "Formato") = "2" Then FilaReporteQuery B, Q, Modulo Else FilaReporteQuery A, Q, Modulo
  CerrarRecordset Q
 Next Item
 Unload RepQueries: Unload RepQueriesModelo
 CerrarRecordset ReporteQuery1: CerrarRecordset ReporteQuery2
 Set ReporteQuery1 = A: Set ReporteQuery2 = B
 If A.RecordCount > 0 Then
  A.MoveFirst: Set RepQueries.DataSource = A
  RepQueries.Show: CentrarReporteMDI RepQueries
 End If
 If B.RecordCount > 0 Then
  B.MoveFirst: Set RepQueriesModelo.DataSource = B
  RepQueriesModelo.Show: CentrarReporteMDI RepQueriesModelo
 End If
 Exit Sub
Fallo:
 Mensaje = Err.Description: CerrarRecordset Q
 MsgBox "La solicitud permanece guardada. No se pudo abrir el reporte: " & Mensaje, vbExclamation, App.Title
End Sub
Public Sub VerQuery(ByVal Item As Long, Optional ByVal Modulo As Integer = 6)
 Dim Items As New Collection
 Items.Add Item: MostrarQueries Items, Modulo
End Sub
Public Sub ExportarQueries(ByVal R As ADODB.Recordset)
 Dim E As Object, L As Object, H As Object, Campos As Variant, Titulos As Variant, I As Long, F As Long, V As Variant, Mensaje As String
 On Error GoTo Fallo
 If R.RecordCount = 0 Then Err.Raise 5, , "No hay solicitudes para exportar."
 Campos = Split("Asient|FecSol|TipoDocumento|NumDoc|CliDoc|DivSol|MonSol|TipQue|TipCas|SolDes|Creador|Estado|SitDoc|FechaCierre|NotaResolucion|Resolucion", "|")
 Titulos = Split("Solicitud|Fecha|Tipo|Documento|Cliente|Moneda|Importe|Solicitud de|Caso|Solicitado|Creado por|Estado|Vigencia|Cierre|Nota de crédito|Solución", "|")
 Set E = CreateObject("Excel.Application"): Set L = E.Workbooks.Add: Set H = L.Worksheets(1)
 H.Name = "Queries": H.Cells.NumberFormat = "@"
 For I = 0 To UBound(Campos)
  H.Cells(1, I + 1).Value2 = Titulos(I)
  If Campos(I) = "MonSol" Or Campos(I) = "SolDes" Then H.Columns(I + 1).NumberFormat = "#,##0.00"
 Next I
 R.MoveFirst: F = 2
 Do While Not R.EOF
  For I = 0 To UBound(Campos)
   V = R.Fields(CStr(Campos(I))).Value
   If Not IsNull(V) Then
    If VarType(V) = vbDate Then V = FechaDocumento(V)
    H.Cells(F, I + 1).Value2 = V
   End If
  Next I
  F = F + 1: R.MoveNext
 Loop
 H.Rows(1).Font.Bold = True: H.Columns.ColumnWidth = 20: H.Columns(5).ColumnWidth = 45: H.Columns(16).ColumnWidth = 70
 H.Rows(1).AutoFilter
 E.Visible = True: E.UserControl = True
 Set H = Nothing: Set L = Nothing: Set E = Nothing
 Exit Sub
Fallo:
 Mensaje = Err.Description
 On Error Resume Next
 If Not L Is Nothing Then L.Close False
 If Not E Is Nothing Then E.Quit
 On Error GoTo 0
 Err.Raise vbObjectError + 3600, , Mensaje
End Sub

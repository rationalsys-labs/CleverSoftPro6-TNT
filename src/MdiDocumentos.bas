Attribute VB_Name = "MdiDocumentos"
Option Explicit
Private DatosReporteCargo As ADODB.Recordset

Public Function ConsultarDocumentos(ByVal Tipo As String, ByVal Buscar As String, Optional ByVal Exacto As Boolean = False, Optional ByVal Modulo As Integer = 11) As ADODB.Recordset
 Set ConsultarDocumentos = ConsultaOperacion("dbo.ConsultarDocumentosGestion", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Tipo", adVarWChar, 2, Tipo), Array("@Buscar", adVarWChar, 100, Buscar), Array("@Exacto", adBoolean, 0, Exacto), Array("@Modulo", adInteger, 0, Modulo)), True)
End Function
Public Function DetalleDocumento(ByVal Tipo As String, ByVal Item As Long, ByVal Seccion As Integer, Optional ByVal Modulo As Integer = 11) As ADODB.Recordset
 Set DetalleDocumento = ConsultaOperacion("dbo.ConsultarDetalleDocumento", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Tipo", adVarWChar, 2, Tipo), Array("@IdDocumento", adInteger, 0, Item), Array("@Seccion", adInteger, 0, Seccion), Array("@Modulo", adInteger, 0, Modulo)), True)
End Function
Public Function EstadosCargos() As ADODB.Recordset
 Set EstadosCargos = ConsultaOperacion("dbo.ConsultarEstadosCargos", Array(Array("@Actor", adInteger, 0, IdUsuarioActual)), True)
End Function
Public Function RegistrosCargos(Optional ByVal Numero As Long = 0) As ADODB.Recordset
 Set RegistrosCargos = ConsultaOperacion("dbo.ConsultarRegistrosCargos", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Numero", adInteger, 0, Numero)), True)
End Function
Public Function DetalleCargo(ByVal Registro As Long) As ADODB.Recordset
 Set DetalleCargo = ConsultaOperacion("dbo.ConsultarDetalleCargo", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Registro", adInteger, 0, Registro)), True)
End Function
Public Function GuardarCargo(ByVal Estado As Long, ByVal Fecha As Date, ByVal Datos As ADODB.Recordset) As ADODB.Recordset
 Dim XML As String
 XML = "<documentos>"
 If Datos.RecordCount > 0 Then Datos.MoveFirst
 Do While Not Datos.EOF
  XML = XML & "<d" & AtributoXML("Tipo", CStr(Datos!Tipo)) & AtributoXML("Id", CStr(Datos!IdDocumento)) & AtributoXML("Version", CStr(Datos!Version)) & "/>"
  Datos.MoveNext
 Loop
 XML = XML & "</documentos>"
 Set GuardarCargo = ConsultaOperacion("dbo.RegistrarCargos", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@IdEstado", adInteger, 0, Estado), Array("@Fecha", adDBTimeStamp, 0, Fecha), Array("@Documentos", adLongVarWChar, Len(XML), XML)), True)
End Function
Public Sub CambiarCargo(ByVal Registro As Long, ByVal Version As Variant, ByVal Activo As Boolean)
 Dim R As ADODB.Recordset
 Set R = ConsultaOperacion("dbo.CambiarVigenciaCargo", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Registro", adInteger, 0, Registro), Array("@Version", adBinary, 8, Version), Array("@Activo", adBoolean, 0, Activo)), True)
 CerrarRecordset R
End Sub
Public Function CeldaDocumento(ByVal Valor As Variant) As String
 CeldaDocumento = Replace(Replace(Replace(ValorTexto(Valor), vbTab, " "), vbCr, " "), vbLf, " ")
End Function
Public Function FechaDocumento(ByVal Valor As Variant) As String
 If Not IsNull(Valor) Then FechaDocumento = Format$(CDate(Valor), "dd/mm/yyyy")
End Function
Public Sub PrepararGrillaDocumento(ByVal Grilla As MSFlexGrid, ByVal Cabecera As String)
 Grilla.Rows = 1
 Grilla.Cols = UBound(Split(Cabecera, "|")) + 1
 Grilla.FixedCols = 0
 Grilla.SelectionMode = flexSelectionByRow
 Grilla.HighLight = flexHighlightAlways
 Grilla.FocusRect = flexFocusNone
 Grilla.AllowUserResizing = flexResizeColumns
 Grilla.FormatString = Cabecera
End Sub
Public Function NuevaBandaCargos() As ADODB.Recordset
 Dim R As New ADODB.Recordset, Nombre As Variant
 R.Fields.Append "Tipo", adVarWChar, 2
 R.Fields.Append "IdDocumento", adInteger
 R.Fields.Append "Version", adVarWChar, 18
 R.Fields.Append "Documento", adVarWChar, 50
 R.Fields.Append "Emision", adDate
 R.Fields.Append "Cliente", adVarWChar, 300
 R.Fields.Append "Divisa", adVarWChar, 50
 R.Fields.Append "Importe", adCurrency
 For Each Nombre In Array("Situacion", "Estado", "Guias", "Servicios", "ContactoEntrega", "DireccionEntrega", "DistritoEntrega", "TelefonoEntrega")
  R.Fields.Append CStr(Nombre), adLongVarWChar, 65535
 Next Nombre
 R.Fields.Append "Anulado", adBoolean
 R.Open
 Set NuevaBandaCargos = R
End Function
Public Sub CopiarDocumentoBanda(ByVal Destino As ADODB.Recordset, ByVal Origen As ADODB.Recordset)
 Dim F As ADODB.Field
 For Each F In Destino.Fields
  If F.Name = "Version" Then
   F.Value = VersionXML(Origen.Fields("VersionFila").Value)
  Else
   F.Value = Origen.Fields(F.Name).Value
  End If
 Next F
 Destino.Update
End Sub
Public Function FechaGestionValida(ByVal Texto As String, ByRef Fecha As Date) As Boolean
 Dim Partes As Variant, D As Integer, M As Integer, A As Integer
 On Error GoTo Fallo
 If Not Trim$(Texto) Like "##/##/####" Then Exit Function
 Partes = Split(Trim$(Texto), "/")
 If UBound(Partes) <> 2 Then Exit Function
 If Len(Partes(2)) <> 4 Then Exit Function
 D = CInt(Partes(0)): M = CInt(Partes(1)): A = CInt(Partes(2))
 Fecha = DateSerial(A, M, D)
 If Day(Fecha) <> D Or Month(Fecha) <> M Or Year(Fecha) <> A Then Exit Function
 FechaGestionValida = (Fecha >= DateSerial(1900, 1, 1) And Fecha <= Date)
Fallo:
End Function
Public Sub ExportarDocumentosExcel(ByVal Datos As ADODB.Recordset, ByVal Titulo As String)
 Dim Excel As Object, Libro As Object, Hoja As Object, F As Long, C As Long, V As Variant
 Dim Mensaje As String, Columnas As New Collection, Nombre As Variant, Campo As ADODB.Field
 On Error GoTo Fallo
 If Datos.RecordCount = 0 Then Err.Raise 5, , "No hay documentos para exportar."
 Set Excel = CreateObject("Excel.Application")
 Set Libro = Excel.Workbooks.Add
 Set Hoja = Libro.Worksheets(1)
 Hoja.Name = "Documentos"
 Hoja.Cells.NumberFormat = "@"
 Hoja.Cells(1, 1).Value2 = Titulo
 For Each Nombre In Array("Tipo", "Documento", "Emision", "Cliente", "Divisa", "Importe", "Situacion", "Estado", "Guias", "Servicios", "ContactoEntrega", "DireccionEntrega", "DistritoEntrega", "TelefonoEntrega")
  For Each Campo In Datos.Fields
   If Campo.Name = Nombre Then Columnas.Add CStr(Nombre): Exit For
  Next Campo
 Next Nombre
 For C = 1 To Columnas.Count
  Hoja.Cells(3, C).Value2 = Columnas(C)
  If Columnas(C) = "Importe" Then Hoja.Columns(C).NumberFormat = "#,##0.00"
 Next C
 Datos.MoveFirst: F = 4
 Do While Not Datos.EOF
  For C = 1 To Columnas.Count
    V = Datos.Fields(CStr(Columnas(C))).Value
    If Not IsNull(V) Then
     If VarType(V) = vbDate Then V = Format$(CDate(V), "dd/mm/yyyy")
     If Len(CStr(V)) > 32767 Then Err.Raise 5, , "Un texto supera el tamaño de una celda de Excel."
     Hoja.Cells(F, C).Value2 = V
    End If
  Next C
  F = F + 1: Datos.MoveNext
 Loop
 Hoja.Rows(3).Font.Bold = True
 Hoja.Columns.AutoFit
 Hoja.Columns.ColumnWidth = 18
 Hoja.Columns(4).ColumnWidth = 45
 Excel.Visible = True
 Excel.UserControl = True
 Set Hoja = Nothing: Set Libro = Nothing: Set Excel = Nothing
 Exit Sub
Fallo:
 Mensaje = Err.Description
 On Error Resume Next
 If Not Libro Is Nothing Then Libro.Close False
 If Not Excel Is Nothing Then Excel.Quit
 On Error GoTo 0
 Err.Raise vbObjectError + 3500, , Mensaje
End Sub
Public Sub MostrarReporteCargos(ByVal Datos As ADODB.Recordset, ByVal Titulo As String, ByVal Registro As String, ByVal Correlativo As String, ByVal Fecha As Date)
 Dim R As ADODB.Recordset, Totales As Object, Moneda As Variant, Resumen As String
 If Datos.RecordCount = 0 Then Err.Raise 5, , "No hay documentos para el reporte."
 Unload RepCargos
 CerrarRecordset DatosReporteCargo
 Set R = New ADODB.Recordset
 R.Fields.Append "Tipo", adVarWChar, 2
 R.Fields.Append "Documento", adVarWChar, 50
 R.Fields.Append "Emision", adVarWChar, 10
 R.Fields.Append "Cliente", adVarWChar, 300
 R.Fields.Append "Divisa", adVarWChar, 50
 R.Fields.Append "Importe", adVarWChar, 40
 R.Open
 Set Totales = CreateObject("Scripting.Dictionary")
 Datos.MoveFirst
 Do While Not Datos.EOF
  R.AddNew Array("Tipo", "Documento", "Emision", "Cliente", "Divisa", "Importe"), Array(Datos!Tipo, Datos!Documento, FechaDocumento(Datos!Emision), Datos!Cliente, Datos!Divisa, Format$(Datos!Importe, "#,##0.00"))
  Moneda = CStr(Datos!Divisa)
  If Not Totales.Exists(Moneda) Then Totales.Add Moneda, CCur(0)
  Totales(Moneda) = CCur(Totales(Moneda)) + CCur(Datos!Importe)
  Datos.MoveNext
 Loop
 For Each Moneda In Totales.Keys
  Resumen = Resumen & CStr(Moneda) & ": " & Format$(Totales(Moneda), "#,##0.00") & vbCrLf
 Next Moneda
 R.MoveFirst
 Set DatosReporteCargo = R
 Set RepCargos.DataSource = DatosReporteCargo
 With RepCargos.Sections("Cabeza").Controls
  .Item("Titulo").Caption = Titulo
  .Item("Fecha").Caption = Format$(Fecha, "dd/mm/yyyy")
  .Item("Registro").Caption = Registro
  .Item("Correl").Caption = Correlativo
 End With
 With RepCargos.Sections("Pie").Controls
  .Item("Total").Caption = CStr(R.RecordCount)
  .Item("Soles").Visible = False: .Item("Dolares").Visible = False
  .Item("Etiqueta20").Visible = False
  .Item("Etiqueta16").Caption = Resumen
  .Item("Etiqueta16").Left = 7000: .Item("Etiqueta16").Top = 400
  .Item("Etiqueta16").Width = 3000: .Item("Etiqueta16").Height = 1000
  .Item("Etiqueta16").CanGrow = True
 End With
 RepCargos.Show
 CentrarReporteMDI RepCargos
End Sub

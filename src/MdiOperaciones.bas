Attribute VB_Name = "MdiOperaciones"
Option Explicit

Public Function TextoXML(ByVal Valor As String) As String
    Valor = Replace(Valor, "&", "&amp;")
    Valor = Replace(Valor, "<", "&lt;")
    Valor = Replace(Valor, ">", "&gt;")
    Valor = Replace(Valor, Chr$(34), "&quot;")
    Valor = Replace(Valor, vbCr, "&#13;")
    TextoXML = Replace(Valor, vbLf, "&#10;")
End Function

Public Function AtributoXML(ByVal Nombre As String, ByVal Valor As String) As String
    AtributoXML = " " & Nombre & "=" & Chr$(34) & TextoXML(Valor) & Chr$(34)
End Function

Public Function NumeroXML(ByVal Valor As String) As String
    If Not IsNumeric(Valor) Then Err.Raise vbObjectError + 3200, , "Ingrese un número válido."
    NumeroXML = Replace(Format$(CDbl(Valor), "0.0000"), Mid$(Format$(0.5, "0.0"), 2, 1), ".")
End Function

Public Function VersionXML(ByVal Valor As Variant) As String
    Dim I As Long, Resultado As String
    Resultado = "0x"
    For I = LBound(Valor) To UBound(Valor)
        Resultado = Resultado & Right$("0" & Hex$(Valor(I)), 2)
    Next I
    VersionXML = Resultado
End Function

Public Function ValorTexto(ByVal Valor As Variant) As String
    If Not IsNull(Valor) Then ValorTexto = CStr(Valor)
End Function

Public Function ConsultaOperacion(ByVal SQL As String, Optional ByVal Parametros As Variant, Optional ByVal Procedimiento As Boolean = False) As ADODB.Recordset
    Dim Cn As ADODB.Connection, Orden As ADODB.Command, Datos As ADODB.Recordset
    Dim I As Long, P As Variant, Mensaje As String
    On Error GoTo Fallo
    AbrirConexion Cn
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandText = SQL
    Orden.CommandTimeout = 120
    Orden.CommandType = IIf(Procedimiento, adCmdStoredProc, adCmdText)
    If Not IsMissing(Parametros) Then
        For I = LBound(Parametros) To UBound(Parametros)
            P = Parametros(I)
            Orden.Parameters.Append Orden.CreateParameter(CStr(P(0)), CLng(P(1)), adParamInput, CLng(P(2)), P(3))
        Next I
    End If
    Set Datos = New ADODB.Recordset
    Datos.CursorLocation = adUseClient
    Datos.Open Orden, , adOpenStatic, adLockReadOnly
    Set Datos.ActiveConnection = Nothing
    Cn.Close
    Set ConsultaOperacion = Datos
    Exit Function
Fallo:
    Mensaje = Err.Description
    CerrarRecordset Datos
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    Err.Raise vbObjectError + 3201, , Mensaje
End Function

Public Function CamposImportacion(ByVal Hoja As String) As String
    Select Case Hoja
        Case "Clientes": CamposImportacion = "Asient,CLIENT,RUCDNI,DIRECC,TELEF1,IdTerritorio,Contac,ContacEnt,DireccEnt,TelefoEnt,Recarg,IdComprobante,IdMoneda,IdNivel,IdCobranza,EsInternacional,IdPais,IdDistrito,IdDistritoEntrega,IdTipoPersona,IdTipoIdentificacion"
        Case "Guias": CamposImportacion = "NumeroGuia,Origen,FechaEnvio,Consignatario,Gestion,Producto"
        Case "Documentos": CamposImportacion = "Tipo,Documento,Emision,Vencimiento,Cliente,Moneda,Subtotal,Recargo,IGV,Total,Situacion,Referencia"
        Case "Servicios": CamposImportacion = "Documento,Linea,NumeroGuia,Servicio"
    End Select
End Function

Public Function ElegirLibroImportacion() As String
    Dim Excel As Object, Archivo As Variant, Mensaje As String
    On Error GoTo Fallo
    Set Excel = CreateObject("Excel.Application")
    Archivo = Excel.GetOpenFilename("Libros Excel (*.xls;*.xlsx),*.xls;*.xlsx", , "Seleccionar información externa")
    If VarType(Archivo) = vbString Then ElegirLibroImportacion = CStr(Archivo)
Salir:
    On Error Resume Next
    If Not Excel Is Nothing Then Excel.Quit
    Set Excel = Nothing
    On Error GoTo 0
    If Len(Mensaje) > 0 Then Err.Raise vbObjectError + 3202, , Mensaje
    Exit Function
Fallo:
    Mensaje = Err.Description
    Resume Salir
End Function

Public Function LeerLibroImportacion(ByVal Archivo As String, ByRef Resumen As String) As String
    Dim Excel As Object, Libro As Object, Hoja As Object, Nombres As Variant, Nombre As Variant
    Dim Campos As Variant, Fila As Long, Columna As Long, Ultima As Long, Valor As Variant
    Dim Texto As String, XML As String, Registro As String, HayDatos As Boolean, Mensaje As String, ContadorHoja As Long
    On Error GoTo Fallo
    If Len(Dir$(Archivo)) = 0 Then Err.Raise 53, , "No existe el archivo."
    Set Excel = CreateObject("Excel.Application")
    Excel.AutomationSecurity = 3
    Excel.AskToUpdateLinks = False
    Excel.DisplayAlerts = False
    Set Libro = Excel.Workbooks.Open(Archivo, 0, True)
    XML = "<carga>"
    Resumen = ""
    Nombres = Array("Clientes", "Guias", "Documentos", "Servicios")
    For Each Nombre In Nombres
        ContadorHoja = 0
        Set Hoja = Nothing
        For Each Hoja In Libro.Worksheets
            If Hoja.Name = CStr(Nombre) Then Exit For
        Next Hoja
        If Not Hoja Is Nothing Then
            If Hoja.Name = CStr(Nombre) Then
                Campos = Split(CamposImportacion(CStr(Nombre)), ",")
                For Columna = 0 To UBound(Campos)
                    If CStr(Hoja.Cells(1, Columna + 1).Value2) <> CStr(Campos(Columna)) Then Err.Raise vbObjectError + 3203, , CStr(Nombre) & ": encabezado esperado " & CStr(Campos(Columna))
                Next Columna
                Ultima = Hoja.UsedRange.Row + Hoja.UsedRange.Rows.Count - 1
                If Ultima > 50001 Then Err.Raise vbObjectError + 3204, , "Máximo 50.000 filas por hoja."
                For Fila = 2 To Ultima
                    Registro = "<r" & AtributoXML("hoja", CStr(Nombre)) & AtributoXML("fila", CStr(Fila))
                    HayDatos = False
                    For Columna = 0 To UBound(Campos)
                        If Hoja.Cells(Fila, Columna + 1).HasFormula Then Err.Raise vbObjectError + 3205, , "No se admiten fórmulas. Hoja " & CStr(Nombre) & ", fila " & CStr(Fila)
                        Valor = Hoja.Cells(Fila, Columna + 1).Value2
                        If IsError(Valor) Then Err.Raise vbObjectError + 3206, , "Celda con error."
                        Texto = Trim$(CStr(Valor))
                        If Len(Texto) > 0 Then HayDatos = True
                        If Campos(Columna) = "FechaEnvio" Or Campos(Columna) = "Emision" Or Campos(Columna) = "Vencimiento" Then
                            If Len(Texto) > 0 Then
                                If IsNumeric(Valor) Then
                                    Texto = Format$(CDate(Valor), "yyyy-mm-dd") & "T00:00:00"
                                Else
                                    If Not Texto Like "####-##-##" Then Err.Raise vbObjectError + 3207, , "Fecha: use AAAA-MM-DD o una fecha Excel."
                                End If
                            End If
                        ElseIf Campos(Columna) = "Subtotal" Or Campos(Columna) = "Recargo" Or Campos(Columna) = "IGV" Or Campos(Columna) = "Total" Or Campos(Columna) = "Recarg" Then
                            If Len(Texto) > 0 Then Texto = NumeroXML(Texto)
                        ElseIf Campos(Columna) = "NumeroGuia" Or Campos(Columna) = "Asient" Or Campos(Columna) = "RUCDNI" Or Campos(Columna) = "Documento" Or Campos(Columna) = "Cliente" Or Campos(Columna) = "Referencia" Then
                            If Len(Texto) > 0 And VarType(Valor) <> vbString Then Err.Raise vbObjectError + 3208, , "Códigos e identificadores deben ser celdas de texto. " & CStr(Nombre) & ", fila " & CStr(Fila)
                        End If
                        Registro = Registro & AtributoXML(CStr(Campos(Columna)), Texto)
                    Next Columna
                    If HayDatos Then
                        XML = XML & Registro & "/>"
                        ContadorHoja = ContadorHoja + 1
                    End If
                Next Fila
                Resumen = Resumen & CStr(Nombre) & ": " & CStr(ContadorHoja) & " registros" & vbCrLf
            End If
        End If
    Next Nombre
    XML = XML & "</carga>"
    LeerLibroImportacion = XML
Salir:
    On Error Resume Next
    If Not Libro Is Nothing Then Libro.Close False
    If Not Excel Is Nothing Then Excel.Quit
    Set Hoja = Nothing
    Set Libro = Nothing
    Set Excel = Nothing
    On Error GoTo 0
    If Len(Mensaje) > 0 Then Err.Raise vbObjectError + 3209, , Mensaje
    Exit Function
Fallo:
    Mensaje = Err.Description
    Resume Salir
End Function

Public Function ProcesarImportacion(ByVal Archivo As String, ByVal XML As String, ByVal Confirmar As Boolean) As ADODB.Recordset
    Set ProcesarImportacion = ConsultaOperacion("dbo.ImportarInformacion", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Archivo", adVarWChar, 255, Archivo), Array("@Datos", adLongVarWChar, Len(XML), XML), Array("@Confirmar", adBoolean, 0, Confirmar)), True)
End Function

Public Sub MostrarHistorialRP(ByVal Texto As String)
    FrmHistorialRP.TxtHistorial.Text = Texto
    FrmHistorialRP.Show vbModal, MdiPri
End Sub

Public Sub MostrarCartillaRP(ByVal Planilla As String)
    Dim Datos As ADODB.Recordset
    If Len(Planilla) = 0 Or Planilla = "Finalizado" Then Err.Raise vbObjectError + 3210, , "El registro no tiene cartilla de facturación."
    Set Datos = ConsultaOperacion("SELECT NroGui AS [REGISTRO 1],Cliente AS [REGISTRO 2],Territorio AS [REGISTRO 3],PesFis AS [REGISTRO 4],PesVol AS [REGISTRO 5],Neto AS [REGISTRO 6],Recargo AS [REGISTRO 7],IGV AS [REGISTRO 8],Total AS [REGISTRO 9],Moneda AS [REGISTRO 10],FechaEnvio AS [REGISTRO 11],Producto AS [REGISTRO 12] FROM dbo.DetallesCartillaRP WHERE Planilla=? ORDER BY NroGui", Array(Array("Planilla", adVarWChar, 20, Planilla)))
    If Datos.EOF Then CerrarRecordset Datos: Err.Raise vbObjectError + 3211, , "No se encontró la cartilla."
    Unload RepRP
    Set RepRP.DataSource = Datos
    RepRP.Sections("Sección4").Controls.Item("Etiqueta9").Caption = Format$(Date, "dd/mm/yyyy")
    RepRP.Sections("Sección4").Controls.Item("Etiqueta18").Caption = Planilla
    RepRP.Sections("Sección5").Controls.Item("REPCAN").Caption = CStr(Datos.RecordCount)
    RepRP.Show
    CentrarReporteMDI RepRP
End Sub

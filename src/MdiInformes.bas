Attribute VB_Name = "MdiInformes"
Option Explicit

Public Type FiltroInforme
 Campo As String
 Operador As Integer
 Valor As String
 OpcionO As Boolean
End Type
Private Type ARCHIVO_INFORME
 Tamano As Long
 Duenio As Long
 Instancia As Long
 Filtro As String
 FiltroPersonal As String
 MaxFiltroPersonal As Long
 IndiceFiltro As Long
 Archivo As String
 MaxArchivo As Long
 TituloArchivo As String
 MaxTituloArchivo As Long
 CarpetaInicial As String
 Titulo As String
 Opciones As Long
 PosArchivo As Integer
 PosExtension As Integer
 Extension As String
 Datos As Long
 Funcion As Long
 Plantilla As String
End Type
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef Datos As ARCHIVO_INFORME) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExW" (ByVal Origen As Long, ByVal Destino As Long, ByVal Opciones As Long) As Long

Public Function OrigenesInforme() As ADODB.Recordset
 Set OrigenesInforme = ConsultaOperacion("dbo.ConsultarOrigenesInforme", Array(Array("@Actor", adInteger, 0, IdUsuarioActual)), True)
End Function
Public Function CamposInforme(ByVal Origen As Long) As ADODB.Recordset
 Set CamposInforme = ConsultaOperacion("dbo.ConsultarCamposInforme", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Origen", adInteger, 0, Origen)), True)
End Function
Public Function FechaFiltroInforme(ByVal Texto As String, ByRef SoloDia As Boolean) As String
 Dim P As Variant, D As Variant, H As Variant, Fecha As Date, Anio As Integer, Mes As Integer, Dia As Integer, V As Variant
 On Error GoTo Fallo
 P = Split(Replace(Trim$(Texto), "T", " "), " ")
 If UBound(P) > 1 Then GoTo Fallo
 If InStr(P(0), "-") > 0 Then
  D = Split(P(0), "-")
  If UBound(D) <> 2 Then GoTo Fallo
  Anio = CInt(D(0)): Mes = CInt(D(1)): Dia = CInt(D(2))
 Else
  D = Split(P(0), "/")
  If UBound(D) <> 2 Then GoTo Fallo
  Dia = CInt(D(0)): Mes = CInt(D(1)): Anio = CInt(D(2))
 End If
 For Each V In D: If Len(CStr(V)) = 0 Or CStr(V) Like "*[!0-9]*" Then GoTo Fallo
 Next V
 If Anio < 1753 Or Anio > 9999 Or Mes < 1 Or Mes > 12 Or Dia < 1 Or Dia > 31 Then GoTo Fallo
 Fecha = DateSerial(Anio, Mes, Dia)
 If Day(Fecha) <> Dia Or Month(Fecha) <> Mes Or Year(Fecha) <> Anio Then GoTo Fallo
 SoloDia = UBound(P) = 0
 If Not SoloDia Then
  H = Split(P(1), ":")
  If UBound(H) <> 2 Then GoTo Fallo
  For Each V In H: If Len(CStr(V)) <> 2 Or CStr(V) Like "*[!0-9]*" Then GoTo Fallo
  Next V
  If CInt(H(0)) > 23 Or CInt(H(1)) > 59 Or CInt(H(2)) > 59 Then GoTo Fallo
  Fecha = Fecha + TimeSerial(CInt(H(0)), CInt(H(1)), CInt(H(2)))
 End If
 FechaFiltroInforme = Format$(Fecha, "yyyy-mm-dd") & "T" & Format$(Fecha, "hh:nn:ss")
 Exit Function
Fallo:
 Err.Raise 5, , "Fecha inválida. Use dd/mm/aaaa o aaaa-mm-dd, con hora hh:mm:ss opcional."
End Function
Public Function NumeroFiltroInforme(ByVal Texto As String) As String
 Dim Valor As String, Digitos As String, Punto As Long, DecimalLocal As String, Miles As String
 DecimalLocal = Mid$(Format$(0.5, "0.0"), 2, 1): Miles = Mid$(Format$(1000, "#,##0"), 2, 1)
 If Not IsNumeric(Texto) Then Err.Raise 5, , "Ingrese un valor numérico para el filtro."
 Valor = Trim$(Texto)
 If Miles <> DecimalLocal And Not Miles Like "[0-9]" Then Valor = Replace(Valor, Miles, "")
 Valor = Replace(Valor, DecimalLocal, "."): Digitos = Valor
 If Left$(Digitos, 1) = "+" Or Left$(Digitos, 1) = "-" Then Digitos = Mid$(Digitos, 2)
 If Len(Digitos) = 0 Or Digitos Like "*[!0-9.]*" Or Len(Replace(Digitos, ".", "")) = 0 Then Err.Raise 5, , "Use notación decimal en el filtro numérico."
 Punto = InStr(Digitos, ".")
 If Punto > 0 Then
  If InStr(Punto + 1, Digitos, ".") > 0 Then Err.Raise 5, , "Ingrese un valor numérico válido."
  If Len(Digitos) - Punto > 6 Then Err.Raise 5, , "El filtro numérico admite hasta seis decimales. Revise el valor; no se redondeó."
 End If
 NumeroFiltroInforme = Valor
End Function
Public Function ConsultarInforme(ByVal Origen As Long, ByVal Columnas As Collection, ByRef Filtros() As FiltroInforme, ByVal Cantidad As Integer) As ADODB.Recordset
 Dim R As ADODB.Recordset, Tipos As Object, N As Variant, S As String, I As Long, Tipo As String, Valor As String, SoloDia As Boolean, Mensaje As String
 On Error GoTo Fallo
 If Columnas.Count = 0 Then Err.Raise 5, , "Seleccione las columnas del informe."
 If Cantidad < 0 Or Cantidad > 4 Then Err.Raise 5, , "Use hasta cuatro filtros."
 Set Tipos = CreateObject("Scripting.Dictionary"): Set R = CamposInforme(Origen)
 Do Until R.EOF: Tipos.Add CStr(R!Campo), CStr(R!Tipo): R.MoveNext: Loop
 CerrarRecordset R: S = "<informe>": I = 0
 For Each N In Columnas
  If Not Tipos.Exists(CStr(N)) Then Err.Raise 5, , "Columna ajena al origen seleccionado."
  I = I + 1: S = S & "<c" & AtributoXML("Orden", CStr(I)) & AtributoXML("Campo", CStr(N)) & "/>"
 Next N
 For I = 0 To Cantidad - 1
  If Not Tipos.Exists(Filtros(I).Campo) Then Err.Raise 5, , "Campo de filtro ajeno al origen."
  Tipo = Tipos(Filtros(I).Campo): Valor = Filtros(I).Valor: SoloDia = False
  If Filtros(I).Operador < 7 Then
   Select Case Tipo
    Case "date", "datetime", "datetime2", "smalldatetime": Valor = FechaFiltroInforme(Valor, SoloDia)
    Case "int", "bigint", "smallint", "tinyint", "numeric", "decimal", "money", "smallmoney", "real", "float", "bit"
     Valor = NumeroFiltroInforme(Valor)
   End Select
  End If
  S = S & "<f" & AtributoXML("Orden", CStr(I + 1)) & AtributoXML("Campo", Filtros(I).Campo) & AtributoXML("Operador", CStr(Filtros(I).Operador)) & AtributoXML("Valor", Valor) & AtributoXML("Dia", IIf(SoloDia, "1", "0")) & AtributoXML("O", IIf(Filtros(I).OpcionO, "1", "0")) & "/>"
 Next I
 S = S & "</informe>"
 Set R = ConsultaOperacion("dbo.GenerarInforme", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Origen", adInteger, 0, Origen), Array("@Solicitud", adLongVarWChar, Len(S), S)), True)
 If R.RecordCount > 10000 Then Err.Raise 5, , "El resultado supera 10000 registros. Precise los filtros."
 Set ConsultarInforme = R: Exit Function
Fallo:
 Mensaje = Err.Description: CerrarRecordset R: Err.Raise 5, , Mensaje
End Function
Public Function TituloCampoInforme(ByVal Campo As String) As String
 Dim I As Integer, C As String
 For I = 1 To Len(Campo)
  C = Mid$(Campo, I, 1)
  If I > 1 Then
   If C Like "[A-Z]" And Mid$(Campo, I - 1, 1) Like "[a-z]" Then TituloCampoInforme = TituloCampoInforme & " "
  End If
  TituloCampoInforme = TituloCampoInforme & C
 Next I
End Function
Public Sub ExportarInformeExcel(ByVal Datos As ADODB.Recordset, ByVal Titulo As String, Optional ByVal Ruta As String = "", Optional ByVal Contexto As String = "")
 Dim E As Object, L As Object, H As Object, R As ADODB.Recordset, F As ADODB.Field, I As Long, J As Long, Mensaje As String
 On Error GoTo Fallo
 If Datos.RecordCount = 0 Then Err.Raise 5, , "El informe no contiene registros."
 Set E = CreateObject("Excel.Application"): Set L = E.Workbooks.Add: Set H = L.Worksheets(1): H.Name = "Informe"
 H.Cells.NumberFormat = "@": H.Cells(1, 1).Value2 = Titulo: H.Cells(2, 1).Value2 = Contexto
 For J = 0 To Datos.Fields.Count - 1
  Set F = Datos.Fields(J): H.Cells(4, J + 1).Value2 = TituloCampoInforme(F.Name)
  Select Case F.Type
   Case adCurrency, adDecimal, adNumeric, adDouble, adSingle: H.Range(H.Cells(5, J + 1), H.Cells(Datos.RecordCount + 4, J + 1)).NumberFormat = "#,##0.00####"
   Case adInteger, adSmallInt, adTinyInt, adUnsignedTinyInt, adBigInt: H.Range(H.Cells(5, J + 1), H.Cells(Datos.RecordCount + 4, J + 1)).NumberFormat = "0"
   Case adDate, adDBDate, adDBTimeStamp: H.Range(H.Cells(5, J + 1), H.Cells(Datos.RecordCount + 4, J + 1)).NumberFormat = "dd/mm/yyyy hh:mm:ss"
  End Select
 Next J
 Set R = Datos.Clone: R.MoveFirst: I = 5
 Do Until R.EOF
  For J = 0 To R.Fields.Count - 1
   If Not IsNull(R.Fields(J).Value) Then
    If VarType(R.Fields(J).Value) = vbDate Then H.Cells(I, J + 1).Value = CDate(R.Fields(J).Value) Else H.Cells(I, J + 1).Value2 = R.Fields(J).Value
   End If
  Next J
  I = I + 1: R.MoveNext
 Loop
 CerrarRecordset R: H.Rows(4).Font.Bold = True: H.Rows(4).Interior.Color = &H83693C: H.Rows(4).Font.Color = vbWhite
 H.Columns.AutoFit: H.Rows(1).Font.Bold = True
 For J = 1 To Datos.Fields.Count: If H.Columns(J).ColumnWidth > 55 Then H.Columns(J).ColumnWidth = 55
 Next J
 H.Range(H.Cells(4, 1), H.Cells(I - 1, Datos.Fields.Count)).AutoFilter
 If Len(Ruta) > 0 Then L.SaveAs Ruta, 51: L.Close False: E.Quit Else E.Visible = True: E.UserControl = True
 Set H = Nothing: Set L = Nothing: Set E = Nothing: Exit Sub
Fallo:
 Mensaje = Err.Description
 On Error Resume Next
 CerrarRecordset R
 If Not L Is Nothing Then L.Close False
 If Not E Is Nothing Then E.Quit
 On Error GoTo 0
 Err.Raise 5, , Mensaje
End Sub
Public Function ElegirArchivoInforme(ByVal Duenio As Long) As String
 Dim A As ARCHIVO_INFORME
 A.Tamano = Len(A): A.Duenio = Duenio: A.Filtro = "Informe de texto (*.txt)" & vbNullChar & "*.txt" & vbNullChar & vbNullChar
 A.Archivo = String$(1024, vbNullChar): A.MaxArchivo = Len(A.Archivo): A.Titulo = "Guardar informe de texto": A.Extension = "txt": A.Opciones = &H80002 Or &H800 Or &H8
 If GetSaveFileName(A) <> 0 Then
  ElegirArchivoInforme = Left$(A.Archivo, InStr(A.Archivo, vbNullChar) - 1)
 ElseIf CommDlgExtendedError() <> 0 Then
  Err.Raise 5, , "No se pudo abrir el diálogo de guardado."
 End If
End Function
Public Sub GuardarInformeTexto(ByVal Datos As ADODB.Recordset, ByVal Ruta As String, Optional ByVal Sobrescribir As Boolean = False)
 Dim FSO As Object, A As Object, R As ADODB.Recordset, F As ADODB.Field, Linea As String, Temporal As String, Mensaje As String, Columna As Integer
 On Error GoTo Fallo
 If Datos.RecordCount = 0 Then Err.Raise 5, , "El informe no contiene registros."
 Set FSO = CreateObject("Scripting.FileSystemObject")
 If FSO.FileExists(Ruta) And Not Sobrescribir Then Err.Raise 5, , "El archivo ya existe."
 Temporal = Ruta & "." & NuevaClaveQuery & ".tmp": Set A = FSO.CreateTextFile(Temporal, False, True)
 For Each F In Datos.Fields: If Len(Linea) > 0 Then Linea = Linea & vbTab
  Linea = Linea & CeldaDocumento(TituloCampoInforme(F.Name))
 Next F
 A.WriteLine Linea: Set R = Datos.Clone: R.MoveFirst
 Do Until R.EOF
  Linea = "": Columna = 0
  For Each F In R.Fields
   If Columna > 0 Then Linea = Linea & vbTab
   If IsNull(F.Value) Then Linea = Linea & " " Else If VarType(F.Value) = vbDate Then Linea = Linea & Format$(F.Value, "dd/mm/yyyy hh:nn:ss") Else Linea = Linea & CeldaDocumento(F.Value)
   Columna = Columna + 1
  Next F
  A.WriteLine Linea: R.MoveNext
 Loop
 CerrarRecordset R: A.Close: Set A = Nothing
 If MoveFileEx(StrPtr(Temporal), StrPtr(Ruta), IIf(Sobrescribir, 9, 8)) = 0 Then Err.Raise 5, , "No se pudo publicar el archivo. El destino anterior se conserva."
 Exit Sub
Fallo:
 Mensaje = Err.Description
 On Error Resume Next
 CerrarRecordset R
 If Not A Is Nothing Then A.Close
 If Len(Temporal) > 0 Then If FSO.FileExists(Temporal) Then FSO.DeleteFile Temporal, True
 On Error GoTo 0
 Err.Raise 5, , Mensaje
End Sub
Private Sub ImprimirLineaInforme(ByVal Texto As String, ByRef Arriba As Single, ByVal Ancho As Single)
 Dim Largo As Long, Corte As Long
 Do While Len(Texto) > 0
  Largo = Len(Texto)
  Do While Largo > 1 And Printer.TextWidth(Left$(Texto, Largo)) > Ancho: Largo = Largo - 1: Loop
  If Largo < Len(Texto) Then Corte = InStrRev(Left$(Texto, Largo), " "): If Corte > 1 Then Largo = Corte
  If Arriba + 6 > Printer.ScaleHeight - 15 Then Printer.NewPage: Arriba = 15
  Printer.CurrentX = 15: Printer.CurrentY = Arriba: Printer.Print Left$(Texto, Largo)
  Texto = Mid$(Texto, Largo + 1): Arriba = Arriba + 5
 Loop
End Sub
Public Sub ImprimirInforme(ByVal Datos As ADODB.Recordset, ByVal Titulo As String, Optional ByVal Contexto As String = "")
 Dim R As ADODB.Recordset, F As ADODB.Field, Arriba As Single, Valor As String, Mensaje As String
 On Error GoTo Fallo
 If Datos.RecordCount = 0 Then Err.Raise 5, , "El informe no contiene registros."
 Set R = Datos.Clone: R.MoveFirst: Printer.ScaleMode = vbMillimeters: Printer.FontName = "Arial": Printer.FontSize = 9: Arriba = 15
 ImprimirLineaInforme Titulo, Arriba, Printer.ScaleWidth - 30
 ImprimirLineaInforme "Generado: " & Format$(Now, "dd/mm/yyyy hh:nn:ss"), Arriba, Printer.ScaleWidth - 30
 ImprimirLineaInforme Contexto, Arriba, Printer.ScaleWidth - 30
 Do Until R.EOF
  For Each F In R.Fields
   Valor = "": If Not IsNull(F.Value) Then If VarType(F.Value) = vbDate Then Valor = Format$(F.Value, "dd/mm/yyyy hh:nn:ss") Else Valor = CStr(F.Value)
   ImprimirLineaInforme TituloCampoInforme(F.Name) & ": " & Valor, Arriba, Printer.ScaleWidth - 30
  Next F
  Arriba = Arriba + 5: R.MoveNext
 Loop
 CerrarRecordset R: Printer.EndDoc: Exit Sub
Fallo:
 Mensaje = Err.Description
 On Error Resume Next
 CerrarRecordset R: Printer.KillDoc
 On Error GoTo 0
 Err.Raise 5, , Mensaje
End Sub

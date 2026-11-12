Attribute VB_Name = "MdiConect"
Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Const BASE_DATOS As String = "CleverSoftPro6"
Public Const MENSAJE_ACCESO As String = "Seleccione su usuario e indique la contraseña. F5 actualiza la lista."

Public Function UsuarioWindows() As String
    Dim Nombre As String
    Dim Longitud As Long
    Nombre = Space$(256)
    Longitud = Len(Nombre)
    If GetUserName(Nombre, Longitud) <> 0 Then UsuarioWindows = Left$(Nombre, Longitud - 1)
End Function

Public Sub CerrarRecordset(ByRef Datos As ADODB.Recordset)
    If Datos Is Nothing Then Exit Sub
    If Datos.State = adStateOpen Then Datos.Close
    Set Datos = Nothing
End Sub

Private Sub CerrarConexion(ByRef Conexion As ADODB.Connection)
    If Conexion Is Nothing Then Exit Sub
    If Conexion.State = adStateOpen Then Conexion.Close
    Set Conexion = Nothing
End Sub

Public Function AbrirConexion(ByRef Conexion As ADODB.Connection) As Boolean
    Dim Datos As ADODB.Recordset
    Set Conexion = New ADODB.Connection
    Conexion.ConnectionTimeout = 5
    Conexion.CommandTimeout = 15
    ' Windows autentica la conexión SQL; el usuario del formulario pertenece a la aplicación.
    Conexion.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=localhost\SQLEXPRESS;Initial Catalog=" & BASE_DATOS
    Set Datos = Conexion.Execute("SELECT DB_NAME()")
    If CStr(Datos.Fields(0).Value) <> BASE_DATOS Then
        CerrarRecordset Datos
        Err.Raise vbObjectError + 1711, , "La conexión no corresponde a la base del proyecto."
    End If
    CerrarRecordset Datos
    AbrirConexion = True
End Function

Public Function ConsultarUsuarios(ByRef Usuarios As ADODB.Recordset, ByRef Mensaje As String) As Boolean
    Dim Conexion As ADODB.Connection
    On Error GoTo ErrorUsuarios
    CerrarRecordset Usuarios
    Mensaje = ""
    If Not AbrirConexion(Conexion) Then GoTo ErrorUsuarios
    Set Usuarios = New ADODB.Recordset
    Usuarios.CursorLocation = adUseClient
    Usuarios.Open "SELECT Usuari FROM dbo.Int_Usuario WHERE ActPas=1 AND SitUsu=N'Activo' AND LEN(CONVERT(nvarchar(150),Paswor))=32 ORDER BY Usuari", Conexion, adOpenStatic, adLockReadOnly
    Set Usuarios.ActiveConnection = Nothing
    CerrarConexion Conexion
    ConsultarUsuarios = True
    Exit Function
ErrorUsuarios:
    CerrarRecordset Usuarios
    CerrarConexion Conexion
    Mensaje = "No se pudo conectar o cargar los usuarios. F5 o Aceptar permite reintentar."
End Function

Public Function ValidarUsuario(ByVal Usuario As String, ByVal ClaveMD5 As String, ByRef Permisos As ADODB.Recordset, ByRef Administrador As Boolean, ByRef Mensaje As String) As Boolean
    Dim Conexion As ADODB.Connection
    Dim Orden As ADODB.Command
    Dim Datos As ADODB.Recordset
    Dim Codigos As Variant
    Dim Permitidos(0 To 200) As Boolean
    Dim Leidos(0 To 200) As Boolean
    Dim Codigo As String
    Dim Valor As String
    Dim I As Integer
    Dim Numero As Long
    Dim CodigoUsuario As Long
    On Error GoTo ErrorAcceso
    CerrarRecordset Permisos
    Administrador = False
    Mensaje = "Usuario o contraseña incorrectos, o cuenta inactiva."
    If Len(Usuario) = 0 Or Len(Usuario) > 50 Then Exit Function
    If Len(ClaveMD5) <> 32 Or LCase$(ClaveMD5) Like "*[!0-9a-f]*" Then Exit Function
    If Not AbrirConexion(Conexion) Then GoTo ErrorAcceso
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Conexion
    Orden.CommandType = adCmdText
    Orden.CommandText = "SELECT ITEM,Usuari,Access FROM dbo.Int_Usuario WHERE Usuari=? AND ActPas=1 AND SitUsu=N'Activo' AND LEN(CONVERT(nvarchar(150),Paswor))=32 AND LOWER(CONVERT(nvarchar(150),Paswor))=?"
    Orden.Parameters.Append Orden.CreateParameter("Usuario", adVarWChar, adParamInput, 50, Usuario)
    Orden.Parameters.Append Orden.CreateParameter("Clave", adVarWChar, adParamInput, 32, LCase$(ClaveMD5))
    Set Datos = Orden.Execute
    If Datos.EOF Then GoTo AccesoRechazado
    CodigoUsuario = CLng(Datos!ITEM)
    If Not IsNull(Datos.Fields("Access").Value) Then Administrador = (UCase$(Trim$(CStr(Datos.Fields("Access").Value))) = "ADMINISTRADOR")
    CerrarRecordset Datos
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Conexion
    Orden.CommandType = adCmdText
    Orden.CommandType = adCmdStoredProc
    Orden.CommandText = "dbo.ObtenerAccesosUsuario"
    Orden.Parameters.Append Orden.CreateParameter("IdUsuario", adInteger, adParamInput, , CodigoUsuario)
    Set Datos = Orden.Execute
    Do While Not Datos.EOF
        If IsNull(Datos(0).Value) Or IsNull(Datos(1).Value) Then GoTo ErrorAcceso
        Codigo = Trim$(CStr(Datos(0).Value))
        Valor = CStr(Datos(1).Value)
        If Len(Codigo) <> 6 Or Codigo Like "*[!0-9]*" Then GoTo ErrorAcceso
        If Valor <> "0" And Valor <> "1" Then GoTo ErrorAcceso
        Numero = CLng(Codigo)
        ' Los estados y las acciones internas conservan su espacio de códigos original.
        If Numero >= 1 And Numero <= 200 Then
            If Leidos(Numero) Then GoTo ErrorAcceso
            Leidos(Numero) = True
            Permitidos(Numero) = (Valor = "1")
        End If
        Datos.MoveNext
    Loop
    CerrarRecordset Datos
    CerrarConexion Conexion
    ' Correspondencia comprobada por función entre el MDI 5.0 y los permisos 4.0.
    ' Cero significa opción nueva sin equivalencia: no se concede acceso por semejanza de número.
    Codigos = Array(103, 104, 1, 6, 21, 22, 101, 7, 24, 25, 8, 9, 88, 90, 98, 93, 100, 2, 11, 99, 83, 4, 105, 106, 107, 3, 108, 109, 110, 111, 112, 113, 14, 96, 114, 115, 116, 117, 102, 118)
    Set Permisos = New ADODB.Recordset
    Permisos.CursorLocation = adUseClient
    Permisos.Fields.Append "Destino", adVarWChar, 50
    Permisos.Fields.Append "Codigo", adVarChar, 6
    Permisos.Fields.Append "Acceso", adInteger
    Permisos.Fields.Append "IdUsuario", adInteger
    Permisos.Open
    For I = 1 To 40
        Numero = CLng(Codigos(I - 1))
        Permisos.AddNew
        Permisos(0).Value = "Menú " & CStr(I)
        Permisos(1).Value = Format$(I, "000000")
        Permisos(2).Value = 0
        Permisos(3).Value = CodigoUsuario
        If Numero > 0 Then
            If Permitidos(Numero) Then Permisos(2).Value = 1
        End If
        Permisos.Update
    Next I
    Permisos.MoveFirst
    Mensaje = ""
    ValidarUsuario = True
    Exit Function
AccesoRechazado:
    Administrador = False
    CerrarRecordset Datos
    CerrarConexion Conexion
    Exit Function
ErrorAcceso:
    Administrador = False
    CerrarRecordset Datos
    CerrarRecordset Permisos
    CerrarConexion Conexion
    Mensaje = "No se pudo verificar el acceso y sus permisos. Puede reintentar o cancelar."
End Function

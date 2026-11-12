Attribute VB_Name = "MdiPerfiles"
Option Explicit

Public Function PuedeAdministrarPerfiles() As Boolean
    PuedeAdministrarPerfiles = SesionActiva And ModUsuario And AccesoPermitido(27) And IdUsuarioActual > 0
End Function
Private Function OrdenPerfil(ByVal Procedimiento As String, ByRef Cn As ADODB.Connection) As ADODB.Command
    Dim Orden As ADODB.Command
    If Not PuedeAdministrarPerfiles Then Err.Raise vbObjectError + 3001, , "Solo el administrador autorizado puede gestionar perfiles."
    AbrirConexion Cn
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandType = adCmdStoredProc
    Orden.CommandText = Procedimiento
    Orden.Parameters.Append Orden.CreateParameter("@Actor", adInteger, adParamInput, , IdUsuarioActual)
    Set OrdenPerfil = Orden
End Function
Private Function DatosPerfil(ByVal Orden As ADODB.Command, ByVal Cn As ADODB.Connection) As ADODB.Recordset
    Dim Datos As ADODB.Recordset, Mensaje As String
    On Error GoTo ErrorConsulta
    Set Datos = New ADODB.Recordset
    Datos.CursorLocation = adUseClient
    Datos.Open Orden, , adOpenStatic, adLockReadOnly
    Set Datos.ActiveConnection = Nothing
    Cn.Close
    Set DatosPerfil = Datos
    Exit Function
ErrorConsulta:
    Mensaje = Err.Description
    CerrarRecordset Datos
    If Cn.State = adStateOpen Then Cn.Close
    Err.Raise vbObjectError + 3002, , Mensaje
End Function
Public Function ListarPerfilesAcceso(Optional ByVal IncluirEliminados As Boolean = False, Optional ByVal Filtro As String = "", Optional ByVal Codigo As Long = 0) As ADODB.Recordset
    Dim Cn As ADODB.Connection, Orden As ADODB.Command
    Set Orden = OrdenPerfil("dbo.ConsultarPerfilesAcceso", Cn)
    Orden.Parameters.Append Orden.CreateParameter("@IncluyeEliminados", adBoolean, adParamInput, , IncluirEliminados)
    Orden.Parameters.Append Orden.CreateParameter("@Filtro", adVarWChar, adParamInput, 50, Filtro)
    Orden.Parameters.Append Orden.CreateParameter("@Item", adInteger, adParamInput, , Codigo)
    Set ListarPerfilesAcceso = DatosPerfil(Orden, Cn)
End Function
Public Function ListarPermisosPerfil(ByVal Codigo As Long) As ADODB.Recordset
    Dim Cn As ADODB.Connection, Orden As ADODB.Command
    Set Orden = OrdenPerfil("dbo.ConsultarPermisosPerfil", Cn)
    Orden.Parameters.Append Orden.CreateParameter("@Item", adInteger, adParamInput, , Codigo)
    Set ListarPermisosPerfil = DatosPerfil(Orden, Cn)
End Function
Public Function GuardarDatosPerfil(ByVal Codigo As Long, ByVal Version As Variant, ByVal Nombre As String, ByVal Permisos As String, Optional ByVal Eliminar As Boolean = False) As Long
    Dim Cn As ADODB.Connection, Orden As ADODB.Command, Datos As ADODB.Recordset, Mensaje As String
    On Error GoTo ErrorGuardado
    Set Orden = OrdenPerfil("dbo.GuardarPerfilAcceso", Cn)
    Orden.Parameters.Append Orden.CreateParameter("@Item", adInteger, adParamInput, , Codigo)
    Orden.Parameters.Append Orden.CreateParameter("@Version", adBinary, adParamInput, 8, Version)
    Orden.Parameters.Append Orden.CreateParameter("@Nombre", adVarWChar, adParamInput, 50, Trim$(Nombre))
    Orden.Parameters.Append Orden.CreateParameter("@Permisos", adVarChar, adParamInput, 40, Permisos)
    Orden.Parameters.Append Orden.CreateParameter("@Eliminar", adBoolean, adParamInput, , Eliminar)
    Set Datos = Orden.Execute
    GuardarDatosPerfil = CLng(Datos!Item)
    CerrarRecordset Datos
    Cn.Close
    Exit Function
ErrorGuardado:
    Mensaje = Err.Description
    CerrarRecordset Datos
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    Err.Raise vbObjectError + 3003, , Mensaje
End Function

Attribute VB_Name = "MdiUsuarios"
Option Explicit

Public Function PuedeAdministrarUsuarios() As Boolean
    PuedeAdministrarUsuarios = SesionActiva And ModUsuario And AccesoPermitido(34) And IdUsuarioActual > 0
End Function

Private Function PrepararOrden(ByVal Procedimiento As String, ByRef Cn As ADODB.Connection) As ADODB.Command
    Dim Orden As ADODB.Command
    If Not PuedeAdministrarUsuarios Then Err.Raise vbObjectError + 2901, , "Solo el administrador autorizado puede gestionar usuarios."
    AbrirConexion Cn
    Set Orden = New ADODB.Command
    Set Orden.ActiveConnection = Cn
    Orden.CommandType = adCmdStoredProc
    Orden.CommandText = Procedimiento
    Orden.Parameters.Append Orden.CreateParameter("@Actor", adInteger, adParamInput, , IdUsuarioActual)
    Set PrepararOrden = Orden
End Function

Private Function ObtenerDatos(ByVal Orden As ADODB.Command, ByVal Cn As ADODB.Connection) As ADODB.Recordset
    Dim Datos As ADODB.Recordset, Mensaje As String
    On Error GoTo ErrorConsulta
    Set Datos = New ADODB.Recordset
    Datos.CursorLocation = adUseClient
    Datos.Open Orden, , adOpenStatic, adLockReadOnly
    Set Datos.ActiveConnection = Nothing
    Cn.Close
    Set ObtenerDatos = Datos
    Exit Function
ErrorConsulta:
    Mensaje = Err.Description
    CerrarRecordset Datos
    If Cn.State = adStateOpen Then Cn.Close
    Err.Raise vbObjectError + 2902, , Mensaje
End Function

Public Function ListarUsuariosAdministracion(Optional ByVal IncluirEliminados As Boolean = False, Optional ByVal Filtro As String = "", Optional ByVal Codigo As Long = 0) As ADODB.Recordset
    Dim Cn As ADODB.Connection, Orden As ADODB.Command
    Set Orden = PrepararOrden("dbo.ConsultarUsuarios", Cn)
    Orden.Parameters.Append Orden.CreateParameter("@IncluyeEliminados", adBoolean, adParamInput, , IncluirEliminados)
    Orden.Parameters.Append Orden.CreateParameter("@Filtro", adVarWChar, adParamInput, 100, Filtro)
    Orden.Parameters.Append Orden.CreateParameter("@Item", adInteger, adParamInput, , Codigo)
    Set ListarUsuariosAdministracion = ObtenerDatos(Orden, Cn)
End Function

Public Function ConsultarPermisosUsuario(ByVal Codigo As Long) As ADODB.Recordset
    Dim Cn As ADODB.Connection, Orden As ADODB.Command
    Set Orden = PrepararOrden("dbo.ConsultarPermisosUsuario", Cn)
    Orden.Parameters.Append Orden.CreateParameter("@Item", adInteger, adParamInput, , Codigo)
    Set ConsultarPermisosUsuario = ObtenerDatos(Orden, Cn)
End Function

Public Function GuardarDatosUsuario(ByVal Accion As Integer, ByVal Codigo As Long, ByVal Version As Variant, ByVal Usuario As String, ByVal Nombre As String, ByVal Apellido As String, ByVal Correo As String, ByVal Perfil As Integer, ByVal Clave As String, ByVal Confirmacion As String, ByVal Permisos As String, Optional ByVal IdPerfilAcceso As Long = 0, Optional ByVal VersionPerfil As Variant) As Long
    Dim Cn As ADODB.Connection, Orden As ADODB.Command, Datos As ADODB.Recordset
    Dim Resumen As Variant, Digest As MD5, Mensaje As String, PerfilAcceso As Variant
    On Error GoTo ErrorGuardado
    Resumen = Null
    If Accion = 1 Or Accion = 3 Then
        If Len(Clave) = 0 Or Clave <> Confirmacion Then Err.Raise vbObjectError + 2903, , "Ingrese una clave y su confirmación coincidente."
        Set Digest = New MD5
        Resumen = LCase$(Digest.DigestStrToHexStr(Clave))
    End If
    Set Orden = PrepararOrden("dbo.GuardarUsuario", Cn)
    Orden.Parameters.Append Orden.CreateParameter("@Accion", adUnsignedTinyInt, adParamInput, , Accion)
    Orden.Parameters.Append Orden.CreateParameter("@Item", adInteger, adParamInput, , Codigo)
    Orden.Parameters.Append Orden.CreateParameter("@Version", adBinary, adParamInput, 8, Version)
    Orden.Parameters.Append Orden.CreateParameter("@Usuario", adVarWChar, adParamInput, 50, Trim$(Usuario))
    Orden.Parameters.Append Orden.CreateParameter("@Nombre", adVarWChar, adParamInput, 50, Trim$(Nombre))
    Orden.Parameters.Append Orden.CreateParameter("@Apellido", adVarWChar, adParamInput, 50, Trim$(Apellido))
    Orden.Parameters.Append Orden.CreateParameter("@Correo", adVarWChar, adParamInput, 100, Trim$(Correo))
    Orden.Parameters.Append Orden.CreateParameter("@Perfil", adUnsignedTinyInt, adParamInput, , Perfil)
    Orden.Parameters.Append Orden.CreateParameter("@Clave", adVarChar, adParamInput, 32, Resumen)
    Orden.Parameters.Append Orden.CreateParameter("@Permisos", adVarChar, adParamInput, 40, Permisos)
    PerfilAcceso = Null
    If IdPerfilAcceso > 0 Then PerfilAcceso = IdPerfilAcceso Else VersionPerfil = Null
    Orden.Parameters.Append Orden.CreateParameter("@IdPerfilAcceso", adInteger, adParamInput, , PerfilAcceso)
    Orden.Parameters.Append Orden.CreateParameter("@VersionPerfil", adBinary, adParamInput, 8, VersionPerfil)
    Set Datos = Orden.Execute
    GuardarDatosUsuario = CLng(Datos!Item)
    CerrarRecordset Datos
    Cn.Close
    Exit Function
ErrorGuardado:
    Mensaje = Err.Description
    CerrarRecordset Datos
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then Cn.Close
    End If
    Err.Raise vbObjectError + 2904, , Mensaje
End Function

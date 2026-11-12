Attribute VB_Name = "ModSesion"
Option Explicit
Public SesionActiva As Boolean
Public VistaPrevia As Boolean
Public VarUsuario As String
Public ModUsuario As Boolean
Public IdUsuarioActual As Long
Private Accesos(1 To 40) As Boolean

Public Sub LimpiarSesion()
    Dim I As Integer
    SesionActiva = False
    VistaPrevia = False
    VarUsuario = ""
    ModUsuario = False
    IdUsuarioActual = 0
    For I = 1 To 40
        Accesos(I) = False
    Next I
End Sub

Public Function IniciarSesion(ByVal Usuario As String, ByVal Administrador As Boolean, ByVal Permisos As ADODB.Recordset, ByRef Mensaje As String) As Boolean
    Dim NuevosAccesos(1 To 40) As Boolean
    Dim Codigo As String
    Dim Valor As String
    Dim I As Integer
    Dim Numero As Integer
    On Error GoTo ErrorPermisos
    LimpiarSesion
    Mensaje = ""
    If Len(Trim$(Usuario)) = 0 Then GoTo ErrorPermisos
    If Permisos Is Nothing Then GoTo ErrorPermisos
    If Permisos.State <> adStateOpen Then GoTo ErrorPermisos
    If Permisos.Fields.Count < 4 Then GoTo ErrorPermisos
    If Permisos.BOF And Permisos.EOF Then GoTo ErrorPermisos
    Permisos.MoveFirst
    IdUsuarioActual = CLng(Permisos.Fields("IdUsuario").Value)
    If IdUsuarioActual <= 0 Then GoTo ErrorPermisos
    Do While Not Permisos.EOF
        If IsNull(Permisos.Fields(1).Value) Or IsNull(Permisos.Fields(2).Value) Then GoTo ErrorPermisos
        Codigo = Trim$(CStr(Permisos.Fields(1).Value))
        Valor = Trim$(CStr(Permisos.Fields(2).Value))
        If Len(Codigo) = 0 Or Codigo Like "*[!0-9]*" Then GoTo ErrorPermisos
        If Len(Codigo) > 6 Then GoTo ErrorPermisos
        Numero = CInt(CLng(Codigo))
        If Numero < 1 Or Numero > 40 Then GoTo ErrorPermisos
        If Valor <> "0" And Valor <> "1" Then GoTo ErrorPermisos
        NuevosAccesos(Numero) = (Valor = "1")
        Permisos.MoveNext
    Loop
    ' Se publica la sesión únicamente después de validar todos los accesos.
    For I = 1 To 40
        Accesos(I) = NuevosAccesos(I)
    Next I
    VarUsuario = Trim$(Usuario)
    ModUsuario = Administrador
    SesionActiva = True
    MdiPri.ActualizarAccesos
    IniciarSesion = True
    Exit Function
ErrorPermisos:
    LimpiarSesion
    Mensaje = "No se pudieron verificar los permisos del usuario. No se abrió la sesión."
    MdiPri.ActualizarAccesos
End Function

Public Function AccesoPermitido(ByVal Codigo As Integer) As Boolean
    Dim Padre As Integer
    If Not SesionActiva Then Exit Function
    If Codigo < 1 Or Codigo > 40 Then Exit Function
    If Codigo >= 36 And Codigo <= 38 Then Exit Function
    If Codigo >= 22 And Codigo <= 25 Then Exit Function
    If Codigo = 17 Then Exit Function
    If Not Accesos(Codigo) Then Exit Function
    Padre = Codigo
    Do
        Select Case Padre
            Case 2: Padre = 1
            Case 4, 8, 11 To 17: Padre = 3
            Case 5 To 7: Padre = 4
            Case 9, 10: Padre = 8
            Case 19 To 21: Padre = 18
            Case 23 To 25: Padre = 22
            Case 27, 33, 39: Padre = 26
            Case 28, 32: Padre = 27
            Case 29 To 31: Padre = 28
            Case 34, 35: Padre = 33
            Case Else: Padre = 0
        End Select
        If Padre = 0 Then Exit Do
        If Not Accesos(Padre) Then Exit Function
    Loop
    AccesoPermitido = True
End Function

Public Function FuncionDisponible(ByVal Codigo As Integer) As Boolean
    ' Se habilitarán los módulos al incorporarlos, uno por uno.
    Select Case Codigo
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 19, 21, 20, 39, 34, 18, 26, 27, 28, 33, 40
            FuncionDisponible = True
    End Select
End Function

Public Sub AbrirModulo(ByVal Codigo As Integer)
    If Codigo < 1 Or Codigo > 40 Then Exit Sub
    If Not AccesoPermitido(Codigo) Then
        MdiPri.MostrarEstado "No tiene acceso a esta opción."
        Exit Sub
    End If
    If Codigo = 40 Then
        MostrarModuloCentrado FrmInformes
        Exit Sub
    End If
    If Codigo = 12 Then
        MostrarModuloCentrado FrmQuejas
        Exit Sub
    End If
    If Codigo = 13 Then
        MostrarModuloCentrado FrmCheqRech
        Exit Sub
    End If
    If Codigo = 14 Then
        MostrarModuloCentrado FrmRecojos
        Exit Sub
    End If
    If Codigo = 15 Then
        MostrarModuloCentrado FrmCartas
        Exit Sub
    End If
    If Codigo = 16 Then
        MostrarModuloCentrado FrmDevolucion
        Exit Sub
    End If
    If Codigo = 5 Then
        MostrarModuloCentrado FrmQueries
        Exit Sub
    End If
    If Codigo = 6 Then
        MostrarModuloCentrado FrmDatQuerie
        Exit Sub
    End If
    If Codigo = 7 Then
        MostrarModuloCentrado FrmFinQuerie
        Exit Sub
    End If
    If Codigo = 11 Then
        MostrarModuloCentrado FrmFacturas
        Exit Sub
    End If
    If Codigo = 19 Then
        MostrarModuloCentrado FrmCargos
        Exit Sub
    End If
    If Codigo = 21 Then
        MostrarModuloCentrado FrmImportacion
        Exit Sub
    End If
    If Codigo = 9 Then
        MostrarModuloCentrado FrmDiarRp
        Exit Sub
    End If
    If Codigo = 10 Then
        MostrarModuloCentrado FrmCompRp
        Exit Sub
    End If
    If Codigo = 27 Then
        If PuedeAdministrarPerfiles Then MostrarModuloCentrado FrmPerfiles
        Exit Sub
    End If
    If Codigo = 34 Then
        If PuedeAdministrarUsuarios Then MostrarModuloCentrado FrmUsuarios
        Exit Sub
    End If
    If Codigo = 39 Then
        MostrarModuloCentrado FrmCatalogos
        Exit Sub
    End If
    If Codigo = 20 Then
        MostrarModuloCentrado FrmClientes
        Exit Sub
    End If
    If Not FuncionDisponible(Codigo) Then
        MdiPri.MostrarEstado "Esta función no está disponible en esta instalación."
    End If
End Sub

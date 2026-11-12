Attribute VB_Name = "MdiGruposRP"
Option Explicit
Public Sub CargarComboGruposRP(ByVal Combo As ComboBox)
 Dim R As ADODB.Recordset
 On Error GoTo Fallo
 Combo.Clear
 Set R = ConsultaOperacion("dbo.ConsultarGruposRP", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@IncluyeGeneral", adBoolean, 0, True)), True)
 Do While Not R.EOF
  Combo.AddItem CStr(R!Nombre)
  If IsNull(R!IdCliente) Then
   Combo.ItemData(Combo.NewIndex) = -1
  Else
   Combo.ItemData(Combo.NewIndex) = CLng(R!IdCliente)
  End If
  R.MoveNext
 Loop
 Combo.AddItem "Por asignar"
 Combo.ItemData(Combo.NewIndex) = 0
 CerrarRecordset R
 Exit Sub
Fallo:
 CerrarRecordset R
 Err.Raise vbObjectError + 3501, , Err.Description
End Sub
Public Function ConsultarGruposDiario() As ADODB.Recordset
 Set ConsultarGruposDiario = ConsultaOperacion("dbo.ConsultarGruposRP", Array(Array("@Actor", adInteger, 0, IdUsuarioActual)), True)
End Function
Public Function BuscarClientesParaGrupo(ByVal Nombre As String) As ADODB.Recordset
 Set BuscarClientesParaGrupo = ConsultaOperacion("dbo.BuscarClientesGrupoRP", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Nombre", adVarWChar, 100, Left$(Nombre, 100))), True)
End Function
Public Function ConfirmarGruposDiario(ByVal XML As String) As ADODB.Recordset
 Set ConfirmarGruposDiario = ConsultaOperacion("dbo.GuardarGruposRP", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Cambios", adLongVarWChar, Len(XML), XML)), True)
End Function

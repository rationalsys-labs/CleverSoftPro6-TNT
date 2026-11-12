VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmVistaInforme
 Caption="Resultado del informe"
 BorderStyle=1
 ClientWidth=14400
 ClientHeight=7800
 MaxButton=0
 MinButton=0
 ShowInTaskbar=0
 BeginProperty Font
  Name="Tahoma"
  Size=8.25
 EndProperty
 Begin VB.Label LblTitulo
  Left=180
  Top=180
  Width=14000
  Height=300
 End
 Begin VB.Label LblContexto
  Left=180
  Top=540
  Width=14000
  Height=660
 End
 Begin MSFlexGridLib.MSFlexGrid GridInforme
  Left=180
  Top=1260
  Width=14000
  Height=5580
  FixedCols=0
  BackColorBkg=&H00FFFFFF&
  BackColorSel=&H0083693C&
  ForeColorSel=&H00FFFFFF&
 End
 Begin VB.Label LblEstado
  Left=180
  Top=7020
  Width=11600
  Height=540
 End
 Begin VB.CommandButton CmdCerrar
  Caption="Cerrar"
  Left=12400
  Top=7140
  Width=1800
  Height=420
 End
End
Attribute VB_Name = "FrmVistaInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Datos As ADODB.Recordset, Titulo As String, Contexto As String
Private Seleccionando As Boolean
Public Sub Preparar(ByVal R As ADODB.Recordset, ByVal Nombre As String, ByVal Descripcion As String)
 Dim F As Long, C As Long, S As String, V As Variant, Texto As String, Ancho As Long
 CerrarRecordset Datos: Set Datos = R.Clone: Titulo = Nombre: Contexto = Descripcion
 LblTitulo.Caption = Nombre: LblContexto.Caption = Descripcion: GridInforme.Redraw = False
 GridInforme.Rows = 1: GridInforme.Cols = Datos.Fields.Count: GridInforme.FixedCols = 0
 For C = 0 To Datos.Fields.Count - 1
  GridInforme.TextMatrix(0, C) = TituloCampoInforme(Datos.Fields(C).Name): GridInforme.ColWidth(C) = Me.TextWidth(GridInforme.TextMatrix(0, C)) + 420
  If GridInforme.ColWidth(C) < 1200 Then GridInforme.ColWidth(C) = 1200
 Next C
 If Datos.RecordCount > 0 Then Datos.MoveFirst
 Do Until Datos.EOF
  S = "": F = F + 1
  For C = 0 To Datos.Fields.Count - 1
   If C > 0 Then S = S & vbTab
   V = Datos.Fields(C).Value: Texto = ""
   If Not IsNull(V) Then
    Select Case Datos.Fields(C).Type
     Case adDate, adDBDate, adDBTimeStamp
      Texto = Format$(V, "dd/mm/yyyy")
      If TimeValue(CDate(V)) <> 0 Then Texto = Texto & " " & Format$(V, "hh:nn:ss")
     Case adCurrency, adNumeric, adDecimal, adDouble, adSingle: Texto = Format$(V, "#,##0.00####")
     Case Else: Texto = CeldaDocumento(V)
    End Select
   End If
   S = S & Texto
   If F <= 100 Then
    Ancho = Me.TextWidth(Texto) + 420: If Ancho > 6000 Then Ancho = 6000
    If Ancho > GridInforme.ColWidth(C) Then GridInforme.ColWidth(C) = Ancho
   End If
  Next C
  GridInforme.AddItem S: Datos.MoveNext
 Loop
 GridInforme.SelectionMode = flexSelectionByRow: GridInforme.HighLight = flexHighlightAlways: GridInforme.FocusRect = flexFocusNone: GridInforme.AllowUserResizing = flexResizeColumns
 If GridInforme.Rows > 1 Then GridInforme.Row = 1: SeleccionFilaQuery GridInforme, Seleccionando
 GridInforme.Redraw = True
 LblEstado.Caption = CStr(Datos.RecordCount) & " registros. Los importes se muestran en su moneda, sin sumar monedas distintas."
 If Datos.RecordCount > 0 Then Datos.MoveFirst
End Sub
Private Sub GridInforme_Click()
 SeleccionFilaQuery GridInforme, Seleccionando
End Sub
Private Sub GridInforme_RowColChange()
 SeleccionFilaQuery GridInforme, Seleccionando
End Sub
Private Sub CmdCerrar_Click()
 Unload Me
End Sub
Private Sub Form_Activate()
 CentrarVentana Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
 CerrarRecordset Datos
End Sub

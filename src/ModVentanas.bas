Attribute VB_Name = "ModVentanas"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef Area As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef Area As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal Relacion As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal Padre As Long, ByVal HijoAnterior As Long, ByVal Clase As String, ByVal Titulo As String) As Long
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal Indicador As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal Monitor As Long, ByRef Info As MONITORINFO) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal Despues As Long, ByVal X As Long, ByVal Y As Long, ByVal Ancho As Long, ByVal Alto As Long, ByVal Opciones As Long) As Long
Private Pendientes As Collection
Private Const GW_OWNER As Long = 4
Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const SWP_NOSIZE As Long = 1
Private Const SWP_NOZORDER As Long = 4
Private Const SWP_NOACTIVATE As Long = 16

Public Sub CentrarVentana(ByVal Ventana As Form)
    Dim Marco As RECT, Limites As RECT, Tamano As RECT
    Dim Info As MONITORINFO, Padre As Long, Ancho As Long, Alto As Long
    Dim X As Long, Y As Long, Monitor As Long
    If Ventana.WindowState <> vbNormal Then Exit Sub
    If GetWindowRect(Ventana.hWnd, Tamano) = 0 Then Exit Sub
    Ancho = Tamano.Right - Tamano.Left
    Alto = Tamano.Bottom - Tamano.Top
    If Ventana.MDIChild Then
        ' El padre nativo es el area MDI, sin menus, herramientas ni barra de estado.
        Padre = GetParent(Ventana.hWnd)
        If GetClientRect(Padre, Marco) = 0 Then Exit Sub
        Limites = Marco
    Else
        ' La propiedad modal ya esta establecida al activarse el formulario.
        Padre = GetWindow(Ventana.hWnd, GW_OWNER)
        If Padre = 0 Then Padre = Ventana.hWnd
        Monitor = MonitorFromWindow(Padre, MONITOR_DEFAULTTONEAREST)
        Info.cbSize = Len(Info)
        If GetMonitorInfo(Monitor, Info) = 0 Then Exit Sub
        Limites = Info.rcWork
        Marco = Limites
        If Padre <> Ventana.hWnd Then
            If GetWindowRect(Padre, Marco) = 0 Then Marco = Limites
        End If
    End If
    X = Marco.Left + (Marco.Right - Marco.Left - Ancho) \ 2
    Y = Marco.Top + (Marco.Bottom - Marco.Top - Alto) \ 2
    ' Si no cabe, se mantiene accesible el borde superior izquierdo, sin redimensionar.
    If X + Ancho > Limites.Right Then X = Limites.Right - Ancho
    If Y + Alto > Limites.Bottom Then Y = Limites.Bottom - Alto
    If X < Limites.Left Then X = Limites.Left
    If Y < Limites.Top Then Y = Limites.Top
    If Ventana.MDIChild Then
        Ventana.Move X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
    Else
        SetWindowPos Ventana.hWnd, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    End If
End Sub

Public Sub MostrarModuloCentrado(ByVal Ventana As Form)
    Ventana.Show
    Ventana.ZOrder 0
    Ventana.SetFocus
    PrepararPosicionVentana Ventana
End Sub

Public Sub CentrarReporteMDI(ByVal Informe As Object)
    Dim Area As RECT, Cliente As Long, X As Single, Y As Single
    If Informe.WindowState <> vbNormal Then Exit Sub
    Cliente = FindWindowEx(MdiPri.hWnd, 0, "MDIClient", vbNullString)
    If Cliente = 0 Then Exit Sub
    If GetClientRect(Cliente, Area) = 0 Then Exit Sub
    X = ((Area.Right - Area.Left) * Screen.TwipsPerPixelX - Informe.Width) / 2
    Y = ((Area.Bottom - Area.Top) * Screen.TwipsPerPixelY - Informe.Height) / 2
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    Informe.Left = X
    Informe.Top = Y
End Sub

Public Sub PrepararPosicionVentana(ByVal Ventana As Form)
    Dim F As Form
    CentrarVentana Ventana
    If Not Ventana.MDIChild Then Exit Sub
    ' VB6 termina de ajustar las barras MDI despues de activar el hijo.
    If Pendientes Is Nothing Then Set Pendientes = New Collection
    For Each F In Pendientes
        If F Is Ventana Then Exit Sub
    Next F
    Pendientes.Add Ventana
    MdiPri.TmrCentrado.Enabled = True
End Sub

Public Sub AplicarCentradosPendientes()
    Dim Lista As Collection, F As Form, Abierta As Form
    MdiPri.TmrCentrado.Enabled = False
    Set Lista = Pendientes
    Set Pendientes = Nothing
    If Lista Is Nothing Then Exit Sub
    For Each F In Lista
        For Each Abierta In Forms
            If Abierta Is F Then
                If F.Visible Then CentrarVentana F
                Exit For
            End If
        Next Abierta
    Next F
End Sub

Public Sub CancelarCentradosPendientes()
    Set Pendientes = Nothing
End Sub

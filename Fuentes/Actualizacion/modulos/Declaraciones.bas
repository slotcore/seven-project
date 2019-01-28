Attribute VB_Name = "Declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES.BAS
'* Tipo             : MODULO
'* Descripcion      : MODULO DONDE SE DECLARAN LAS VARIABLES PUBLICAS QUE SE UTILIZARAN EN LA CLASE
'*                    ASI COMO LA DEFINICION DE ALGUNAS FUNCIONES PROPIAS DE LA CLASE
'* DISEÑADO POR     : JOSE CAHCON
'* ULTIMA REVISION  : 17/11/2011
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Global AP_RUTAORIGEN As String ' DIRECCION DE DONDE SE VAN A COPIAR LOS ARCHIVOS
Global AP_RUTADESTINO As String ' DIRECCION HACIA DONDE SE VAN A COPIAR LOS ARCHIVOS

Global AP_RUTAPROGRAMA As String ' RUTA DEL PROGRAMA QUE SE VA A EJECUTAR DESPUES DE LA ACTUALIZACION
Global AP_RUTAREGISTRAR As String ' RUTA DEL ARCHIVO .BAT PARA REGISTRAR LAS LIBRERIAS NUEVAS
Global AP_IDVERSION As Integer ' ID DE LA VERSION A INSTALAR
Global AP_VERSIONTIPO As Integer ' TIPO DE VERSION A INSTALAR 0: TODOS, 1: SEGUN LISTA
Global AP_MOTIVO As String ' MOTIVOS DE LA ACTUALIZACION

Global xCon As New ADODB.Connection

Dim F_BASEDATOS As String
Dim F_GRUPOTRABAJO As String
Dim F_PASSWORD As String
Dim F_USUARIO As String
Dim F_PROVEEDOR As String


Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Const SC_SIZE = &HF000
Const SC_MOVE = &HF010
Const SC_MINIMIZE = &HF020
Const SC_MAXIMIZE = &HF030
Const SC_CLOSE = &HF060
Const SC_RESTORE = &HF120

Const MF_SEPARATOR = &H800
Const MF_BYPOSITION = &H400
Const MF_BYCOMMAND = &H0

Public Sub KillProcess2(ByVal processName As String)
    On Error GoTo ErrHandler
        
    Dim oWMI
    Dim ret
    Dim sService
    Dim oWMIServices
    Dim oWMIService
    Dim oServices
    Dim oService
    Dim servicename

    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")

    For Each oService In oServices
        servicename = _
            LCase(Trim(CStr(oService.Name) & ""))

        If InStr(1, servicename, _
            LCase(processName), vbTextCompare) > 0 Then
            ret = oService.Terminate
        End If
    Next

    Set oServices = Nothing
    Set oWMI = Nothing
    Exit Sub
ErrHandler:
    Err.Clear
End Sub

Public Sub FrmOcultarBoton(FrmhWnd As Long, mTipo As Integer)
    '--mTipo
    Dim hwnd&, hMenu&, Success&
    Dim i%

    hwnd = FrmhWnd
    hMenu = GetSystemMenu(hwnd, 0)
    
    'Usa esto para quitar los menús que te interesen:
    Select Case mTipo
        Case 0: Success = DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)
        Case 1: Success = DeleteMenu(hMenu, SC_MOVE, MF_BYCOMMAND)
        Case 2: Success = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
        Case 3: Success = DeleteMenu(hMenu, SC_MINIMIZE, MF_BYCOMMAND)
        Case 4: Success = DeleteMenu(hMenu, SC_MAXIMIZE, MF_BYCOMMAND)
        Case 6: Success = DeleteMenu(hMenu, SC_RESTORE, MF_BYCOMMAND)
    End Select
    
End Sub

Public Sub CentrarFrm(frm As Object)
    '--frm formulario
    On Error Resume Next
    If frm.WindowState <> 2 Then
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
    ' frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
    End If
    Err.Clear
End Sub

Function LeerNumeroDisco(Unidad As String) As Variant
    'unidad = "c:"   formato para pasarle a la funcion
    Dim fs As New Scripting.FileSystemObject
    Dim d As Scripting.Drive
    Dim s As Variant

    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(Trim(Unidad))))
    s = d.DriveType
    s = d.DriveType
    If s <> 2 Then
        LeerNumeroDisco = 0
    Else
        LeerNumeroDisco = d.SerialNumber
    End If
    Set fs = Nothing
    Set d = Nothing
End Function

Function obtenerUsuario() As String
    Dim lpBuff As String * 25
    Dim ret As Long
    Dim UserName As String
    
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    
    obtenerUsuario = UserName
End Function

Function xLeerLineaINI(RutaArchivoIni As String, TextoLinea As String, Posicion As String) As String
    'RutaArchivoIni = Ruta del Archivo Ini Incluiyendo el nombre del archivo
    'TextoLinea = Cadena que se buscar en el archivo INI
    'Posicion = titulo del archivo ini
    Dim L1 As Long
    Dim xRuta As String * 150
    L1 = GetPrivateProfileString(Posicion, TextoLinea, "", xRuta, Len(xRuta), RutaArchivoIni)
    xLeerLineaINI = Trim(UCase(Trim(Left(xRuta, InStr(xRuta, Chr(0)) - 1))))
End Function

Function F_NulosN(valor_nulo As Variant) As Double
    If Trim(valor_nulo) = "" Or IsNull(valor_nulo) Then
        F_NulosN = 0
    Else
        'Si el valor no es nulo retorna el valor original
        If IsNumeric(valor_nulo) Then
            F_NulosN = valor_nulo
        Else
            F_NulosN = 0
        End If
    End If
End Function

Function F_NulosC(valor_nulo As Variant) As String
    If Trim(valor_nulo) = "" Or IsNull(valor_nulo) Then
        F_NulosC = ""
    Else
        'Si el valor no es nulo retorna el valor original
        If IsNumeric(valor_nulo) Then
            F_NulosC = valor_nulo
        Else
            F_NulosC = Trim(valor_nulo)
        End If
    End If
End Function

Sub F_RST_Busq(rstBusq As ADODB.Recordset, TxtSQLoTabla As String, xConeccion As Connection)
On Error GoTo LaCague:
    If rstBusq.State = adStateOpen Then
        rstBusq.Close
    End If
    
    rstBusq.CursorLocation = adUseClient
    rstBusq.CursorType = adOpenForwardOnly
    rstBusq.LockType = adLockOptimistic
    
    rstBusq.ActiveConnection = xConeccion
    rstBusq.Open F_NulosC(TxtSQLoTabla), , , , adAsyncFetch
    'adAsyncFetch
    Exit Sub
LaCague:
    'MsgBox "No se pudo guardar el recorset por el siguiente motivo :" + Trim(Err.Description)
    Set rstBusq = Nothing
End Sub

Sub AbrirConeccion(ByRef ConeccOpen)
                        
    Dim xCadConeccion As String
    
    xCadConeccion = "Provider=" + Trim(F_PROVEEDOR) _
                    & ";Password=" & Trim(F_PASSWORD) & ";Persist Security Info=true" _
                    & ";User ID=" & Trim(F_USUARIO) & ";Data Source=" + Trim(F_BASEDATOS) _
                    & ";Jet OLEDB:System database=" + Trim(F_GRUPOTRABAJO)
    
    ConeccOpen.ConnectionString = xCadConeccion
    ConeccOpen.Open
End Sub

Sub Main()
    Dim xCad As String
    Dim xRs As New ADODB.Recordset
    
    ' Se llenan los datos necesarios
    F_BASEDATOS = xLeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS") + "data.mdb"
    F_GRUPOTRABAJO = xLeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS") + "seven.mdw"
    F_PASSWORD = "010419762005"
    F_USUARIO = "cav2005sialp"
    F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    AP_RUTAPROGRAMA = xLeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS") + "seven.exe"
    AP_RUTAREGISTRAR = xLeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS") + "registrar.bat"
    AP_RUTADESTINO = xLeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    
    ' Se abre al conexion
    AbrirConeccion xCon
    
    xCad = "SELECT mae_version.id, mae_version.fchreg, mae_version.descripcion, mae_version.glosa, mae_version.origen, mae_version.estado, mae_version.tipo " _
            + vbCr + "From mae_version " _
            + vbCr + "WHERE (((mae_version.estado)=-1));"
    
    F_RST_Busq xRs, xCad, xCon
    If xRs.State = 0 Then KillProcess2 ("seven.exe"): Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
        
    AP_RUTAORIGEN = F_NulosC(xRs("origen"))
    ' Verfifica que la direccion este bien escrita
    If Mid(AP_RUTAORIGEN, Len(AP_RUTAORIGEN), 1) <> "\" Then AP_RUTAORIGEN = AP_RUTAORIGEN + "\"
    
    AP_IDVERSION = F_NulosN(xRs("id"))
    AP_VERSIONTIPO = F_NulosN(xRs("tipo"))
    AP_MOTIVO = F_NulosC(xRs("glosa"))
        
End Sub

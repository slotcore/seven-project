Attribute VB_Name = "Module1"
Option Explicit
Public xCon As New ADODB.Connection
Public xNumRuc As String
Public xNomEmp As String
Public SeEjecutoEmp As Boolean
Public xIdEmpresa As Integer
Public xIdUsuario As Integer
Public xTitulo As String
Public T_ToolTipText(20) As String
'Global BdSIG As Database
Public xNivelUsuario As Integer
Global AP_RUTASY As String
Global AP_RUTABD As String
Global AP_RUTABM As String

Public xConDATA As New ADODB.Connection 'esta coneccion abre la data del antiguo sgi

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'Constante que le indica que es para maximizar la ventana
Public Const SHOWMAXIMIZED_eSW = 3&
Public Const SW_RESTORE As Integer = 9 ' Restaura la ventana a su tamaño y posición original
Public Const SW_HIDE = 0

Sub CargarToolTipText()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    RST_Busq Rst, "SELECT  * FROM mae_toolbar ORDER BY id", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            T_ToolTipText(A) = NulosC(Rst("descripcion"))
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    Set Rst = Nothing
End Sub

Sub Main()
    AP_RUTASY = RutaSY
    AP_RUTABD = RutaBD
    AP_RUTABM = RutaBM
'
    SeEjecutoEmp = False
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCad As String
    Dim xRutaData As String
'
    xTitulo = "Sistema Gestion Informati2                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        ca"
    
    xRutaData = "J:\seven\data\2012\0002\data.mdb"
    'xRutaData = "C:\mantenimiento\mantenimiento.mdb"
    'xRutaData = "e:\bd\savar\data\2008\0001\data.mdb"
    'xRutaData = "w:\seven\data\2008\0001\data.mdb"


    xFun.F_BASEDATOS = xRutaData
    xFun.F_GRUPOTRABAJO = Trim(App.Path) + "\seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    Set xCon = xFun.AbrirConeccion
 
   
'    xCon.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};DATABASE=0001;SERVER=192.168.1.108;UID=kike;password=010419762005;PORT=3306;"

    'xCon.Open

    FrmMenuRapido.Show
    'MDIForm1.Show
    'Form1.Show
    'FrmAcceso.Show
End Sub


Function RutaExe() As String
    Dim xRuta As String
    xRuta = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    RutaExe = xRuta
End Function

Function RutaSY() As String
    Dim L1 As Long
    Dim xRuta As String * 150
    L1 = GetPrivateProfileString("RUTAS", "RUTASY", "", xRuta, Len(xRuta), RutaExe & "sgi.ini")
    RutaSY = Trim(UCase(Trim(Left(xRuta, InStr(xRuta, Chr(0)) - 1))))
End Function

Function RutaBD() As String
    Dim L1 As Long
    Dim xRuta As String * 150
    L1 = GetPrivateProfileString("RUTAS", "RUTABD", "", xRuta, Len(xRuta), RutaExe & "sgi.ini")
    RutaBD = Trim(UCase(Trim(Left(xRuta, InStr(xRuta, Chr(0)) - 1))))
End Function

Function RutaBM() As String
    Dim L1 As Long
    Dim xRuta As String * 150
    L1 = GetPrivateProfileString("RUTAS", "RUTABM", "", xRuta, Len(xRuta), RutaExe & "sgi.ini")
    RutaBM = Trim(Left(xRuta, InStr(xRuta, Chr(0)) - 1))
End Function

Function ChequeEmpresa() As Boolean
    If xIdEmpresa = 0 Then
        MsgBox "No ha especificado ninguna empresa a procesar, seleccione " & Chr(13) _
            & "una empresa y vuelva a intentar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        ChequeEmpresa = False
        Exit Function
    End If
    ChequeEmpresa = True
End Function

Sub CargarEmpresa()
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT * FROM  m_empresa", xCon
    xNumRuc = Rst("numruc")
    xNomEmp = Rst("nomemp")
End Sub




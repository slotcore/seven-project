Attribute VB_Name = "FuncionesModule"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function mLeerLineaINI(RutaArchivoIni As String, TextoLinea As String, Posicion As String) As String
    'RutaArchivoIni = Ruta del Archivo Ini Incluiyendo el nombre del archivo
    'TextoLinea = Cadena que se buscar en el archivo INI
    'Posicion = titulo del archivo ini
    Dim L1 As Long
    Dim xRuta As String * 150
    L1 = GetPrivateProfileString(Posicion, TextoLinea, "", xRuta, Len(xRuta), RutaArchivoIni)
    mLeerLineaINI = Trim(UCase(Trim(Left(xRuta, InStr(xRuta, Chr(0)) - 1))))
End Function

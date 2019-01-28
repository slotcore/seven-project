VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncionesES"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function LeerLineaINI(RutaArchivoIni As String, TextoLinea As String, TextoMarcador As String) As String
    LeerLineaINI = xLeerLineaINI(RutaArchivoIni, TextoLinea, TextoMarcador)
End Function

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

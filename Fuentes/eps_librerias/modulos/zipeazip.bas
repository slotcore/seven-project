Attribute VB_Name = "zipeazip"
Public Type ZIPUSERFUNCTIONS
    DLLPrnt As Long
    DLLPassword As Long
    DLLComment As Long
    DLLService As Long
End Type


Public Type ZPOPT
fSuffix As Long
fEncrypt As Long
fSystem As Long
fVolume As Long
fExtra As Long
fNoDirEntries As Long
fExcludeDate As Long
fIncludeDate As Long
fVerbose As Long
fQuiet As Long
fCRLF_LF As Long
fLF_CRLF As Long
fJunkDir As Long
fRecurse As Long
fGrow As Long
fForce As Long
fMove As Long
fDeleteEntries As Long
fUpdate As Long
fFreshen As Long
fJunkSFX As Long
fLatestTime As Long
fComment As Long
fOffsets As Long
fPrivilege As Long
fEncryption As Long
fRepair As Long
flevel As Byte
date As String
szRootDir As String
End Type

Public Type ZIPnames
    s(0 To 99) As String
End Type

Public Type CBChar
    ch(4096) As Byte
End Type

Public Declare Function ZpInit Lib "zip32.dll" (ByRef Zipfun As ZIPUSERFUNCTIONS) As Long
Public Declare Function ZpSetOptions Lib "zip32.dll" (ByRef Opts As ZPOPT) As Long
Public Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long


Sub xZip(Archivos() As String, ArchivoZip As String)
    Dim Resultado As Long
    Dim intContadorFicheros As Integer
    
    Dim FuncionesZip As ZIPUSERFUNCTIONS
    Dim OpcionesZip As ZPOPT
    
    Dim NombresFicherosZip As ZIPnames
    
    FuncionesZip.DLLComment = DevolverDireccionMemoria(AddressOf FuncionParaProcesarComentarios)
    FuncionesZip.DLLPassword = DevolverDireccionMemoria(AddressOf FuncionParaProcesarPassword)
    FuncionesZip.DLLPrnt = DevolverDireccionMemoria(AddressOf FuncionParaProcesarMensajes)
    FuncionesZip.DLLService = DevolverDireccionMemoria(AddressOf FuncionParaProcesarServicios)
    
    'For intContadorFicheros = 0 To File1.ListCount
    '    NombresFicherosZip.s(intContadorFicheros) = File1.List(intContadorFicheros)
    'Next
    For intContadorFicheros = 0 To UBound(Archivos)
        NombresFicherosZip.s(intContadorFicheros) = Archivos(intContadorFicheros)
    Next
    
    Resultado = ZpInit(FuncionesZip)
    Resultado = ZpSetOptions(OpcionesZip)
    Resultado = ZpArchive(intContadorFicheros - 1, ArchivoZip, NombresFicherosZip)
End Sub


Function FuncionParaProcesarPassword(ByRef B1 As Byte, L As Long, ByRef B2 As Byte, ByRef B3 As Byte) As Long
    FuncionParaProcesarPassword = 0
End Function

Function FuncionParaProcesarServicios(ByRef fname As CBChar, ByVal x As Long) As Long
    FuncionParaProcesarServicios = 0
End Function

Function FuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal x As Long) As Long
    FuncionParaProcesarMensajes = 0
End Function

Function FuncionParaProcesarComentarios(Comentario As CBChar) As CBChar
    Comentario.ch(0) = vbNullString
    FuncionParaProcesarComentarios = Comentario
End Function

Public Function DevolverDireccionMemoria(Direccion As Long) As Long
    DevolverDireccionMemoria = Direccion
End Function



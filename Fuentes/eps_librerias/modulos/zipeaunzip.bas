Attribute VB_Name = "zipeaunzip"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   MODULO UNZIP, REALIZADO POR EDUARDO GONZALEZ MUÑIZ
'            7 - SEPTIEMBRE - 2001
'               egonzalez@ocs.es
'   TODOS LOS DERECHOS RESERVADOS (ES BROMA ;->)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Type CBChar
    ch(4096) As Byte
End Type


Private Type ZIPnames
    s(0 To 99) As String
End Type


Private Type UNZIPUSERFUNCTION
    UNZIPPrntFunction As Long
    UNZIPSndFunction As Long
    UNZIPReplaceFunction  As Long
    UNZIPPassword As Long
    UNZIPMessage  As Long
    UNZIPService  As Long
    TotalSizeComp As Long
    TotalSize As Long
    CompFactor As Long
    NumFiles As Long
    Comment As Integer
End Type


Private Type UNZIPOPTIONS
    ExtractOnlyNewer  As Long
    SpaceToUnderScore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nUflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    fPrivilege As Long
    Zip As String
    extractdir As String
End Type

Public Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As ZIPnames, ByVal xfnc As Long, ByRef xfnv As ZIPnames, dcll As UNZIPOPTIONS, Userf As UNZIPUSERFUNCTION) As Long

Public Sub xUnZip(Zip As String, extractdir As String)
    On Error GoTo err_Unzip
    
    Dim Resultado As Long
    Dim intContadorFicheros As Integer
    
    Dim FuncionesUnZip As UNZIPUSERFUNCTION
    Dim OpcionesUnZip As UNZIPOPTIONS
    
    Dim NombresFicherosZip As ZIPnames, NombresFicheros2Zip As ZIPnames
    
    NombresFicherosZip.s(0) = vbNullChar
    NombresFicheros2Zip.s(0) = vbNullChar
    FuncionesUnZip.UNZIPMessage = 0&
    FuncionesUnZip.UNZIPPassword = 0&
    FuncionesUnZip.UNZIPPrntFunction = DevolverDireccionMemoria(AddressOf UNFuncionParaProcesarMensajes)
    FuncionesUnZip.UNZIPReplaceFunction = DevolverDireccionMemoria(AddressOf UNFuncionReplaceOptions)
    FuncionesUnZip.UNZIPService = 0&
    FuncionesUnZip.UNZIPSndFunction = 0&
    OpcionesUnZip.C_flag = 1
    OpcionesUnZip.fQuiet = 2
    OpcionesUnZip.noflag = 1
    OpcionesUnZip.Zip = Zip
    OpcionesUnZip.extractdir = extractdir
    
    Resultado = Wiz_SingleEntryUnzip(0, NombresFicherosZip, 0, NombresFicheros2Zip, OpcionesUnZip, FuncionesUnZip)
    
    Exit Sub
err_Unzip:
    MsgBox "Unzip: " + Err.Description, vbExclamation
    Err.Clear
End Sub


Private Function UNFuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal x As Long) As Long
On Error GoTo err_UNFuncionParaProcesarMensajes

    UNFuncionParaProcesarMensajes = 0

Exit Function
err_UNFuncionParaProcesarMensajes:
    MsgBox "UNFuncionParaProcesarMensajes: " + Err.Description, vbExclamation
    Err.Clear
End Function

Private Function UNFuncionReplaceOptions(ByRef p As CBChar, ByVal L As Long, ByRef m As CBChar, ByRef Name As CBChar) As Integer
On Error GoTo err_UNFuncionReplaceOptions

    UNFuncionReplaceOptions = 0

Exit Function
err_UNFuncionReplaceOptions:
    MsgBox "UNFuncionReplaceOptions: " + Err.Description, vbExclamation
    Err.Clear
End Function

Public Function DevolverDireccionMemoria(Direccion As Long) As Long
On Error GoTo err_DevolverDireccionMemoria

    DevolverDireccionMemoria = Direccion

Exit Function
err_DevolverDireccionMemoria:
    MsgBox "DevolverDireccionMemoria: " + Err.Description, vbExclamation
    Err.Clear
End Function



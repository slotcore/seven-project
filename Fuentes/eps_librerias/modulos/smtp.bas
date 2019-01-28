Attribute VB_Name = "smtp"
Option Explicit

Public Enum CONEXION
    CONECTED = 0
    MailFrom = 1
    RCPTTO = 2
    DATAC = 3
    MESSAGGE = 4
    QUIT = 5
End Enum

Public SendStatus As CONEXION
Public Respuesta As String
Public Code As Integer
Public DServer As String
Public DHelo As String
Public DMailFrom As String
Public DRcptTo As String
Public DSubject As String
Public DMensaje As String
Public DFrom As String
Public exCaption As String

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public UUfiles(0 To 9) As String
Public indexUUfiles As Byte

Public Function Conectado() As Boolean
    Dim ret As Long
    ret = InternetGetConnectedStateEx(ret, sConnType, 254, 0)
    If ret = 1 Then
        Conectado = True
    Else
        Conectado = False
    End If
End Function

Public Function TempDirectory() As String
Dim TempPath As String
Dim temp
TempPath = String(145, Chr(0))
temp = GetTempPath(145, TempPath)
TempDirectory = Left(TempPath, InStr(TempPath, Chr(0)) - 1)
End Function

Public Function UUEncodeFile(strFilePath As String) As String
    Dim intFile         As Integer      'file handler
    Dim intTempFile     As Integer      'temp file
    Dim TempFileName    As String
    Dim lFileSize       As Long         'size of the file
    Dim strFileName     As String       'name of the file
    Dim strFileData     As String       'file data chunk
    Dim lEncodedLines   As Long         'number of encoded lines
    Dim strTempLine     As String       'temporary string
    Dim i               As Long         'loop counter
    Dim j               As Integer      'loop counter
    
    'Get file name
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    
    'Insert first marker: "begin 664 ..."
    intTempFile = FreeFile
    TempFileName = TempDirectory & GenerateCode(8)
    Open TempFileName For Binary As intTempFile
    Put intTempFile, , "begin 664 " + strFileName + vbCrLf
    
    'Get file size
    lFileSize = FileLen(strFilePath)
    lEncodedLines = lFileSize \ 45 + 1
    
    'Prepare buffer to retrieve data from the file by 45 symbols chunks
    strFileData = Space(45)
    
    intFile = FreeFile
    
    Open strFilePath For Binary As intFile
        For i = 1 To lEncodedLines
            DoEvents
            
            On Error Resume Next
            FrmMensaje.PB.Value = i * 100 / lEncodedLines
            
            'Read file data by 45-bytes cnunks
            If i = lEncodedLines Then strFileData = Space(lFileSize Mod 45)

            'Retrieve data chunk from file to the buffer
            Get intFile, , strFileData
            
            'Add first symbol to encoded string that informs
            'about quantity of symbols in encoded string.
            'More often "M" symbol is used.
            strTempLine = Chr(Len(strFileData) + 32)
            
                
            'If the last line is processed and length of
            'source data is not a number divisible by 3, add one or two
            'blankspace symbols
            If i = lEncodedLines And (Len(strFileData) Mod 3) Then strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
            
            For j = 1 To Len(strFileData) Step 3
                'Breake each 3 (8-bits) bytes to 4 (6-bits) bytes
                '1 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
                '2 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                               + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
                '3 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
                '4 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
            Next j
            
            'replace " " with "`"
            strTempLine = Replace(strTempLine, " ", "`")
            'add encoded line to result buffer
            
            'strResult = strResult + strTempLine + vbCrLf
            Put intTempFile, , strTempLine + vbCrLf
            
            strTempLine = ""
        Next i
    Close intFile

    'add the end marker
    'strResult = strResult & "`" & vbCrLf + "end" + vbCrLf
    Put intTempFile, , "`" & vbCrLf + "end" + vbCrLf
    Close intTempFile
    
    
    Dim fso, f, ts
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(TempFileName)
    Set ts = f.OpenAsTextStream(1, -2)
    UUEncodeFile = ts.ReadAll
    ts.Close
    
    Kill TempFileName
End Function

'conectar al servidor
Sub Conectar()
    With FrmMensaje
        .cmdEnviar.Enabled = False
        .cmdCancel.Visible = True
        .Refresh
        .Caption = "Enviando..."
        .sck.Close
        .sck.Connect DServer, 25
    End With
    AddStatus ("Conectando a " & DServer & "... " & Now)
End Sub

'cerrar coneccion
Sub DesConectar()
    SendStatus = CONECTED
    Call AddStatus("Desconectado")
    With FrmMensaje
        .cmdCancel.Visible = False
        .Caption = exCaption
        .cmdEnviar.Enabled = True
        .sck.Close
    End With
End Sub

'agregar status
Sub AddStatus(Texto As String)
    FrmMensaje.txtStatus = FrmMensaje.txtStatus & vbCrLf & Texto
    FrmMensaje.txtStatus.SelStart = Len(FrmMensaje.txtStatus.Text)
    FrmMensaje.txtStatus.Refresh
End Sub

'generador de codigos alfanumericos
Function GenerateCode(NumChar As Integer)
    Randomize Timer
    Dim Code As String
    Dim Chars As Integer
    Dim Alfa As Integer
    Code = ""
    For Chars = 1 To NumChar
        Alfa = Int(Rnd * 2 + 1)
        If Alfa = 2 Then
            Code = Chr(Int((Rnd * 25 + 1) + 97)) & Code
        Else
            Code = Int((Rnd * 9 + 1)) & Code
        End If
    Next
    GenerateCode = Code
End Function

Public Function Enviar(From As String, MailFrom As String, MailTo As String, subject As String, Mensaje As String)
    If Not Conectado Then
        AddStatus "No conectado a Internet para mandar: " & From & " " & Now
        Exit Function
    End If
        
    DHelo = GenerateCode(8)
    DMailFrom = MailFrom
    DFrom = From
    DSubject = subject
    DMensaje = Mensaje
    DRcptTo = MailTo
    Call Conectar
End Function

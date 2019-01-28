VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActualizador 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2430
   ClientLeft      =   3105
   ClientTop       =   2610
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "FrmActualizador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2430
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDetalle 
      Height          =   1335
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "FrmActualizador.frx":0442
      Top             =   1020
      Width           =   5775
   End
   Begin VB.CheckBox ckDetalle 
      Caption         =   "Mostrar Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   780
      Width           =   1485
   End
   Begin MSComctlLib.ProgressBar pgBar 
      Height          =   405
      Left            =   30
      TabIndex        =   1
      Top             =   330
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label LblPorc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "LblPorc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5280
      TabIndex        =   2
      Top             =   90
      Width           =   540
   End
   Begin VB.Label lblDetalle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "lblDetalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   645
   End
End
Attribute VB_Name = "frmActualizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400

Dim cSQL As String
Dim IDPC_ As Double
Dim MENSAJEERROR As String

Private Function ShellandWait(ExeFullPath As String, Optional TimeOutValue As Long = 0) As Boolean
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessId As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean

    ' Se carga el mensaje por si hay algun Error
    MENSAJEERROR = "Error al verificar el Archivo de Registro: " & AP_RUTAREGISTRAR
    
    lStart = CLng(Timer)
    sExeName = ExeFullPath

    ' Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If
    lInst = Shell(sExeName, vbMinimizedNoFocus)
    lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst) 'Optenemos el ProcessID
    
    Do 'Aqui se genera un ciclo hasta que el proceso sea distinto de pendiente, o sea, Alla terminado.
        Call GetExitCodeProcess(lProcessId, lExitCode) ' Optenemos el si hay exits code o todavia esta en ejecucion (pending)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                If Timer < lStart Then Exit Do
            Else
                Exit Do ' Se sale del ciclo si se acavo el tiemo de espera
            End If
        End If
    Loop While lExitCode = STATUS_PENDING
    ShellandWait = True
End Function

Private Sub procesarActualizacion()
    Dim ret As Double
    Dim TER As Boolean
    Dim result As Variant
           
    ret = 0
    result = xfilecopy(AP_RUTAORIGEN, AP_RUTADESTINO, "*.*", lblDetalle)
    LblPorc.Caption = ""
    lblDetalle.Caption = "Aplicando cambios, espere por Favor"
    
    TER = ShellandWait(AP_RUTAREGISTRAR)
End Sub

Private Function xfilecopy(origen As String, destino As String, archivo As String, informa As Label)
    Dim n As String
    Dim cuenta As Double
    Dim result As Double
    Dim pcent As String
    
    ' Se carga el mensaje por si hay algun Error
    MENSAJEERROR = "Error al copiar los archivos de Actualizacion, verifique los Programas abiertos"
    
    ' cuenta los archivos a copiar
    cuenta = 0
    n = Dir(origen & archivo)
    
    While (n <> "")
         cuenta = cuenta + 1
         n = Dir
    Wend

    ' Copia
    result = 0
    n = Dir(origen & archivo)
    
    pgBar.Max = cuenta
    pgBar.Min = 0
    
    While (n <> "") And (result > -1)
        DoEvents
        'Me.Refresh
        
        ' Se cambia de atributo a normal
        SetAttr destino & n, vbNormal
        ' Se copia el Archivo
        FileCopy origen & n, destino & n
        
        ' Se cambia de atributo a Solo lectura
        SetAttr destino & n, vbReadOnly
        
        result = result + 1
        ' Detalle de la copia
        lblDetalle.Caption = "Actualizando Programa... : "
        
        'pcent = result & "/" & cuenta & " "
        'lblDetalle.Caption = "Actualizando Programa... : " & pcent & " " & n
        
        ' Detalle del Porcentaje
        LblPorc.Caption = Format(100 * result / cuenta, "#0.0") & "%"
        
        pgBar.Value = result
        n = Dir
    Wend
    
    informa.Caption = ""
    xfilecopy = result
End Function

Public Function ArchivoExiste(Path As String) As Boolean
    Dim X As VbFileAttribute
    On Error GoTo Fallo
    X = GetAttr(Path)
    ArchivoExiste = True
    Exit Function
Fallo:
    ArchivoExiste = False
End Function

Private Function verificarVersion() As Boolean
    Dim VERIFICO_  As Boolean
    Dim VERIFICOPERSONAL_ As Boolean
    Dim xRs As New ADODB.Recordset
    
    ' Si no hay ninguna version activa
    If F_NulosN(AP_IDVERSION) = 0 Then verificarVersion = True: Exit Function
    
    If Not ArchivoExiste(AP_RUTAORIGEN) Then
        MsgBox "No se encontro la Ruta de Origen de la actualizacion, verifiquela", vbExclamation, ""
        verificarVersion = True
        Exit Function
    End If
    
    ' Se carga el mensaje por si hay algun Error
    MENSAJEERROR = "Error al verificar los componentes necesarios para la Actualizacion del Programa"
    
    ' Por defecto se analiza a todos los Usuarios
    VERIFICOPERSONAL_ = True
    ' Por defecto todos los usuarios ya tienen la ultima version
    VERIFICO_ = True
    
    ' Se encuentra el Id de la PC
    cSQL = "SELECT * FROM mae_pc WHERE serdis = '" & Trim(Str(LeerNumeroDisco("c:"))) & "'"
    Set xRs = Nothing
    F_RST_Busq xRs, cSQL, xCon
    
    IDPC_ = F_NulosN(xRs("id"))
    
    ' Se verifica el tipo de Actualizacion
    If AP_VERSIONTIPO = 1 Then ' Actualizar a grupo
        cSQL = "SELECT mae_versiondet.idver, mae_versiondet.idpc " _
            + vbCr + "From mae_versiondet " _
            + vbCr + "WHERE (((mae_versiondet.idpc)=" & IDPC_ & "));"
        Set xRs = Nothing
        F_RST_Busq xRs, cSQL, xCon
        If xRs.RecordCount = 0 Then
            VERIFICOPERSONAL_ = False
        End If
    End If
    
    If VERIFICOPERSONAL_ Then
        ' Se verifica si tiene instalada la version
        cSQL = "SELECT mae_versionact.idver, mae_versionact.idpc, mae_versionact.fchact, mae_versionact.horact, mae_versionact.userpc " _
            + vbCr + "From mae_versionact " _
            + vbCr + "WHERE (((mae_versionact.idver)=" & AP_IDVERSION & ") AND ((mae_versionact.idpc)=" & IDPC_ & "));"
        Set xRs = Nothing
        F_RST_Busq xRs, cSQL, xCon
        
        If xRs.RecordCount = 0 Then
            VERIFICO_ = False
        End If
    End If
    
    verificarVersion = VERIFICO_
End Function

Public Sub KillProcess(ByVal processName As String)
    On Error GoTo ErrHandler
    
    ' Se carga el mensaje por si hay algun Error
    MENSAJEERROR = "Error al tratar de cerrar el Programa para su Actualizacion"
    
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

Private Sub grabarOperacion()
    Dim xRs As New ADODB.Recordset
    Set xRs = Nothing
    
    ' Se carga el mensaje por si hay algun Error
    MENSAJEERROR = "Error al grabar el registro de la Actualizacion"
    
    cSQL = "SELECT mae_versionact.* " _
        + vbCr + "From mae_versionact;"
        
    F_RST_Busq xRs, cSQL, xCon
        
    xCon.BeginTrans
    xRs.AddNew
    xRs("idver") = F_NulosN(AP_IDVERSION)
    xRs("idpc") = F_NulosN(IDPC_)
    xRs("fchact") = Format(Now, "dd/mm/yyyy")
    xRs("horact") = Format(Now, "HH:mm")
    xRs("userpc") = F_NulosC(obtenerUsuario)
    xRs.Update
    xCon.CommitTrans
End Sub

Private Sub iniciarCampos()
    lblDetalle.Caption = ""
    LblPorc.Caption = ""
    Me.Height = 1380
    MENSAJEERROR = "Error al iniciar los valores de los componentes necesarios para la Actualizacion"
End Sub

Private Sub ckDetalle_Click()
    If ckDetalle.Value = 0 Then
        Me.Height = 1380
    Else
        Me.Height = 2745
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Activate()
    On Error GoTo SALIR
    ' Se cierra seven si esta abierto
    KillProcess ("seven.exe")
    ' Se actualiza Librerias
    procesarActualizacion
    ' Se graba el registro
    grabarOperacion
    ' Se abre seven
    Call Shell(AP_RUTAPROGRAMA, vbNormalFocus)
    Unload Me
    Exit Sub
SALIR:
    MsgBox MENSAJEERROR, vbExclamation, ""
    'Call Shell(AP_RUTAPROGRAMA, vbNormalFocus)
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo SALIR
    ' Se inicializan todas las variables
    iniciarCampos
    Main
    txtDetalle.Text = AP_MOTIVO
    FrmOcultarBoton Me.hwnd, 2 ' se dessactiva el boton cerrar
    CentrarFrm Me ' se centra el formulario
    ' Se verifica si la version actual ya esta instalada
    If verificarVersion Then
        Unload Me
    End If
    Exit Sub
SALIR:
    MsgBox MENSAJEERROR, vbExclamation, ""
    Call Shell(AP_RUTAPROGRAMA, vbNormalFocus)
    Unload Me
End Sub

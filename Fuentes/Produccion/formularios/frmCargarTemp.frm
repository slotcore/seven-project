VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form frmCargarTemp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Actualizar Asistencia"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "[ Rango de Actualización ]"
      Height          =   675
      Left            =   30
      TabIndex        =   4
      Top             =   840
      Width           =   6285
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchAsisIni 
         Height          =   300
         Left            =   660
         TabIndex        =   5
         Top             =   285
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchAsisFin 
         Height          =   300
         Left            =   2700
         TabIndex        =   6
         Top             =   285
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   345
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fin"
         Height          =   195
         Left            =   2160
         TabIndex        =   7
         Top             =   345
         Width           =   210
      End
   End
   Begin VB.CheckBox ckAvanzado 
      Caption         =   "Avanzado"
      Height          =   195
      Left            =   5250
      TabIndex        =   3
      Top             =   540
      Width           =   1065
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cancelar"
      Height          =   330
      Index           =   1
      Left            =   1380
      TabIndex        =   2
      Top             =   450
      Width           =   1305
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Iniciar"
      Height          =   330
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Width           =   1305
   End
   Begin MSComctlLib.ProgressBar pgBar 
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmCargarTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conSQLS As ADODB.Connection       ' Base de datos del control de asistencia
Dim CONSASISTENCIA As String
Dim RstAsis As New ADODB.Recordset
Dim RstErr As New ADODB.Recordset
Dim cSQL As String

Private Sub ckAvanzado_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ckAvanzado.Value = 0 Then
        Me.Height = 1185
    Else
        Me.Height = 1890
    End If
End Sub

Public Sub cmd_Click(Index As Integer)
    Dim MENSAJE_ As String
    
    Select Case Index
        Case 0
            procesarAsistencia
            If Grabar Then
                MsgBox "Se procesó correctamente las marcaciones", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            
            RstErr.Filter = adFilterNone
            If RstErr.RecordCount <> 0 Then
                RstErr.MoveFirst
                While Not RstErr.EOF
                    MENSAJE_ = MENSAJE_ + vbCr + RstErr("numdoc")
                    RstErr.MoveNext
                Wend
                
                MsgBox "No se procesaron los siguientes Documentos : " + MENSAJE_ + vbCr + "Verifique la existencia de los mismos en el Sistema", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            Set conSQLS = Nothing
            Unload Me
            
        Case 1
            Set conSQLS = Nothing
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()
    QueHace = 3
    If conectarBD("TEMPUS") Then
        iniciarCampos
    Else
        Unload Me
    End If
End Sub

Private Sub iniciarCampos()
    TxtFchAsisIni.valor = Date
    TxtFchAsisFin.valor = Date
    ckAvanzado.Value = 0
    Me.Height = 1185
    preparaRST RstErr
End Sub

Private Function conectarBD(nombre_BD As String) As Boolean
    Dim AP_PROVIDER As String
    Dim AP_INITIALCATALOG As String
    Dim AP_DATASOURCE As String
    Dim AP_USER As String
    Dim AP_PASSWORD As String
    
On Error GoTo HORROR_
    ' La conexión a la base de datos
    Set conSQLS = New ADODB.Connection
    
    AP_PROVIDER = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "PROVIDER", "DATOS")
    AP_INITIALCATALOG = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "INITIALCATALOG", "DATOS")
    AP_DATASOURCE = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "DATASOURCE", "DATOS")
    AP_USER = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "USER", "DATOS")
    AP_PASSWORD = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "PASSWORD", "DATOS")
    
    conSQLS.Open "Provider=" & AP_PROVIDER & "; " & _
             "Initial Catalog=" & AP_INITIALCATALOG & "; " & _
             "Data Source=" & AP_DATASOURCE & "; " & _
             "user id = " & AP_USER & "; " & _
             "password = " & AP_PASSWORD & ""
             
    conectarBD = True
    Exit Function
    
HORROR_:
    MsgBox "Ocurrio un Error al tratar de conectarse al Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    conectarBD = False
End Function

Private Sub hallarConsulta()
    If ckAvanzado.Value = 0 Then
        CONSASISTENCIA = "(TEMPUS.MARCACIONES.FECHA >= CAST('" & Format(Date, "dd/mm/yyyy") & "' AS datetime)) " _
                                & "AND (TEMPUS.MARCACIONES.FECHA <= CAST('" & Format(Date, "dd/mm/yyyy") & "' AS datetime))"
    Else
        CONSASISTENCIA = "(TEMPUS.MARCACIONES.FECHA >= CAST('" & CDate(TxtFchAsisIni.valor) & "' AS datetime)) " _
                                & "AND (TEMPUS.MARCACIONES.FECHA <= CAST('" & CDate(TxtFchAsisFin.valor) & "' AS datetime))"
    End If
End Sub

Private Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    With Rst
        If .State <> 0 Then
            If TODO Then .Filter = adFilterNone
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
        End If
    End With
End Sub

Private Sub procesarAsistencia()
    Dim rs As New ADODB.Recordset
    
    hallarConsulta
    ' Para una base de datos normal:
    cSQL = "SELECT TEMPUS.MARCACIONES.* " _
            + vbCr + "FROM TEMPUS.MARCACIONES " _
            + vbCr + "WHERE " & CONSASISTENCIA & " " _
            + vbCr + "ORDER BY TEMPUS.MARCACIONES.FECHA"
    
    ' Abrir el recordset de forma estática, no vamos a cambiar datos
    RST_Busq rs, cSQL, conSQLS
    
    If RstAsis.State <> 0 Then
        RstAsis.Filter = adFilterNone
        If RstAsis.RecordCount <> 0 Then limpiarRST RstAsis
    End If
    
    Set RstAsis = rs
End Sub

Sub preparaRST(ByRef xRs As Recordset)
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(1, 3) As String

    xCampos(0, 0) = "numdoc":      xCampos(0, 1) = "C":      xCampos(0, 2) = "100"
    Set xRs = xFun.CrearRstTMP(xCampos)
    xRs.Open
End Sub

Function Grabar() As Boolean
    Dim RstCab As New ADODB.Recordset
    Dim IDEMP_ As Double
    Dim IDEMPTEMP_ As Double
    
'On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    
    If ckAvanzado.Value = 0 Then
        cSQL = "DELETE * FROM pla_recmarcacion " _
            + vbCr + "WHERE (dia >= CDate('" & Date - 1 & "') And dia <=CDate('" & Date & "'))"
    Else
        cSQL = "DELETE * FROM pla_recmarcacion " _
            + vbCr + "WHERE (dia >= CDate('" & TxtFchAsisIni.valor & "') And dia <=CDate('" & TxtFchAsisFin.valor & "'))"
    End If
            
    xCon.Execute cSQL
    RST_Busq RstCab, "SELECT top 1 * FROM pla_recmarcacion ", xCon
            
    If RstAsis.State = 0 Then Grabar = False: Exit Function
    RstAsis.Filter = adFilterNone
    If RstAsis.RecordCount = 0 Then Grabar = False: Exit Function
    
    pgBar.Min = 0
    pgBar.Max = RstAsis.RecordCount
    
    RstAsis.MoveFirst
    While Not RstAsis.EOF
        pgBar.Value = pgBar.Value + 1
        
        IDEMP_ = NulosN(Busca_Codigo(NulosC(RstAsis("NUMERO_TARJETA")), "numdoc", "id", "pla_empleados", "C", xCon))
        
        If IDEMP_ = 0 Then
            RstErr.Filter = "numdoc = " & NulosC(RstAsis("NUMERO_TARJETA"))
            If RstErr.RecordCount = 0 Then
                RstErr.AddNew
                RstErr("numdoc") = NulosC(RstAsis("NUMERO_TARJETA"))
                RstErr.Update
            End If
        Else
            RstCab.AddNew
            RstCab("dia") = Format(RstAsis("FECHA"), "dd/mm/yyyy")
            RstCab("hora") = Format(RstAsis("HORA"), FORMAT_HORA_AL_SEGUNDO)
            RstCab("numdoc") = NulosC(RstAsis("NUMERO_TARJETA"))
            RstCab("idemp") = IDEMP_
            RstCab.Update
        End If
        
        IDEMPTEMP_ = IDEMP_
        RstAsis.MoveNext
    Wend
    
    xCon.CommitTrans
    Grabar = True
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Exit Function

LaCague:
    Resume
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstIns = Nothing:    Set RstTar = Nothing:    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function


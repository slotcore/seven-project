VERSION 5.00
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCargarTemp2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produccion - Mantenimiento de Personal"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ckCriPro 
      Caption         =   "Exportar Registros a Servidor de Asistencia"
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   12
      Top             =   1560
      Width           =   3435
   End
   Begin VB.CheckBox ckCriPro 
      Caption         =   "Procesar Bajas en el Sistema"
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   11
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Criterios de Procesamiento ]"
      Height          =   1185
      Left            =   2310
      TabIndex        =   9
      Top             =   840
      Width           =   4005
      Begin VB.CheckBox ckCriPro 
         Caption         =   "Dar Mantenimiento a Servidor de Asistencia"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   13
         Top             =   960
         Width           =   3435
      End
      Begin VB.CheckBox ckCriPro 
         Caption         =   "Procesar Asistencia en el Sistema"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Rango de Actualizacion ]"
      Height          =   1185
      Left            =   30
      TabIndex        =   4
      Top             =   840
      Width           =   2205
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   690
         TabIndex        =   5
         Top             =   360
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   690
         TabIndex        =   6
         Top             =   705
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
         Top             =   435
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fin"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   765
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
Attribute VB_Name = "frmCargarTemp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conSQLS As New ADODB.Connection       ' Base de datos del control de asistencia
Dim CONSASISTENCIA As String
Dim cSQL As String
Dim RstReg As New ADODB.Recordset
Dim Agregando As Boolean
Dim NUMERODIASBAJA_ As Integer
Dim RstAsis As New ADODB.Recordset
Dim RstErr As New ADODB.Recordset
Dim RstExp As New ADODB.Recordset
Dim RstBajSer As New ADODB.Recordset
Dim RstAltSer As New ADODB.Recordset

Private Sub ckAvanzado_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ckAvanzado.Value = 0 Then
        Me.Height = 1185
    Else
        Me.Height = 2445
    End If
End Sub

Public Sub cmd_Click(Index As Integer)
    Dim MENSAJE_ As String
    
    Select Case Index
        Case 0 ' Btn Iniciar
            ' Se procesa la asistencia
            If ckCriPro(0).Value = 1 Then
                procesarAsistencia
            End If
            ' Se procesa la baja de Personal
            If ckCriPro(1).Value = 1 Then
                procesarBajaPersonal
            End If
            ' Se exporta al sevidor de asistencia
            If ckCriPro(2).Value = 1 Then
                procesarMantenimientoMarcador 0
            End If
            ' Se da mantenimiento al servidor de asistencia
            If ckCriPro(3).Value = 1 Then
                procesarMantenimientoMarcador 1
            End If
            
            Unload Me
            
        Case 1 ' Btn Cancelar
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()
    QueHace = 3
    iniciarCampos
End Sub

Private Sub iniciarCampos()
    TxtFchIni.valor = Date
    TxtFchFin.valor = Date
    ckAvanzado.Value = 0
    Me.Height = 1185
    preparaRST RstErr
    NUMERODIASBAJA_ = 3
    ckCriPro(0).Value = 1
    ckCriPro(1).Value = 1
    ckCriPro(2).Value = 1
    ckCriPro(3).Value = 1
End Sub

Private Function conectarBD(nombre_BD As String) As Boolean
    Dim AP_PROVIDER As String
    Dim AP_INITIALCATALOG As String
    Dim AP_DATASOURCE As String
    Dim AP_USER As String
    Dim AP_PASSWORD As String

On Error GoTo HORROR_

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
    MsgBox "Ocurrio un error al tratar de conectarse al Servidor de Asistencia", vbError + vbOKOnly + vbDefaultButton1, xTitulo
    conectarBD = False
End Function

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
    Dim xRs As New ADODB.Recordset
        
    If Not conectarBD("TEMPUS") Then Exit Sub
    
    CONSASISTENCIA = hallarConsulta(0)
    
    cSQL = "SELECT TEMPUS.MARCACIONES.* " _
            + vbCr + "FROM TEMPUS.MARCACIONES " _
            + vbCr + "WHERE " & CONSASISTENCIA & " " _
            + vbCr + "ORDER BY TEMPUS.MARCACIONES.FECHA"
    
    ' Abrir el recordset de forma estática, no vamos a cambiar datos
    Set xRs = Nothing
    RST_Busq xRs, cSQL, conSQLS
                
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then
        MsgBox "No se han encontrado registros para la fecha especificada; " _
            + vbCr + "no se procesara la recolección de marcaciones", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    ' Se procede a grabar lo registros
    If Grabar(xRs) Then
        MsgBox "Se procesó correctamente las marcaciones", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        pExportar 0
    End If
End Sub

Private Sub procesarMantenimientoMarcador(TIPO_ As Integer)
    Select Case TIPO_
        Case 0
            If exportarMarcador Then
                MsgBox "Se procesó correctamente la exportación de registros al Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                pExportar 2
            End If
            
        Case 1
            If darBajaServidorAsistencia(Format(Date, FORMAT_DATE)) Then
                MsgBox "Se procesó correctamente la baja de registros al Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                pExportar 3
            End If
            
            If darAltaServidorAsistencia Then
                MsgBox "Se procesó correctamente el alta de registros al Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                pExportar 4
            End If
            
    End Select
End Sub

Private Function hallarConsulta(TIPO_ As Integer, Optional FECH_ As String) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    Dim FCHINI_ As Date
    Dim FCHFIN_ As Date
    Dim CONSULTA_ As String
     
    Select Case TIPO_
        Case 0
            If ckAvanzado.Value = 0 Then
                CONSULTA_ = "(TEMPUS.MARCACIONES.FECHA = CAST('" & Format(Date, "dd/mm/yyyy") & "' AS datetime))"
            Else
                CONSULTA_ = "(TEMPUS.MARCACIONES.FECHA >= CAST('" & CDate(TxtFchIni.valor) & "' AS datetime)) " _
                                & "AND (TEMPUS.MARCACIONES.FECHA <= CAST('" & CDate(TxtFchFin.valor) & "' AS datetime))"
            End If
        
        Case 1
            FCHINI_ = CDate(FECH_) - NUMERODIASBAJA_
            FCHFIN_ = CDate(FECH_)
            
            cSQL = "SELECT pla_recmarcacion.idemp, pla_recmarcacion.numdoc " _
                + vbCr + "FROM pla_recmarcacion " _
                + vbCr + "WHERE (((pla_recmarcacion.dia)>=CDate('" & FCHINI_ & "') And (pla_recmarcacion.dia)<=CDate('" & FCHFIN_ & "'))) " _
                + vbCr + "GROUP BY pla_recmarcacion.idemp, pla_recmarcacion.numdoc;"
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            If xRs.State = 0 Then CONSULTA_ = "": Exit Function
            If xRs.RecordCount = 0 Then CONSULTA_ = "": Exit Function
            
            CONSULTA_ = cSQL
    End Select
    
    hallarConsulta = CONSULTA_
End Function

Private Function GENERAR_SQL_ID_RST(Rst As ADODB.Recordset, nDesc As String, _
                            nCampo As String, Optional nTipoIn As String = "IN", _
                            Optional fEsNumero As Boolean = True) As String
    Dim nSQL As String
    Dim k&
    nSQL = ""
    
    If Rst.RecordCount = 0 Then Exit Function Else Rst.MoveFirst
    While Not Rst.EOF
        If Trim(CStr(Rst("" & nDesc & ""))) <> "" Then
            If fEsNumero = True Then
                nSQL = nSQL & NulosN(Rst("" & nDesc & "")) & ","
            Else
                nSQL = nSQL & "'" & NulosC(Rst("" & nDesc & "")) & "',"
            End If
        End If
        Rst.MoveNext
    Wend
    
    If nSQL <> "" Then nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") "
        
    GENERAR_SQL_ID_RST = nSQL
End Function

Private Sub procesarBajaPersonal()
    Dim A As Integer
    Dim ERROR_ As Boolean
    Dim D_ As Date
    
    If ckAvanzado.Value = 0 Then
        If darBaja(Format(Date, "dd/mm/yyyy")) Then
            MsgBox "El proceso de matenimiento de Personal se procesó correctamente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            ERROR_ = True
        End If
    Else
        For D_ = CDate(TxtFchIni.valor) To CDate(TxtFchFin.valor)
            If darBaja(Format(D_, "dd/mm/yyyy")) Then
                If D_ = CDate(TxtFchFin.valor) Then
                    MsgBox "El proceso de matenimiento de Personal se proceso correctamente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                End If
            Else
                ERROR_ = True
                Exit For
            End If
        Next D_
    End If
    
    ' Se Exporta el listado de personal dado de baja
    If ERROR_ Then
        MsgBox "No se han procesado las bajas del personal; no hay registros para procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Else
        pExportar 1
    End If
End Sub

Private Sub pExportar(TIPO_ As Integer)
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim xRs  As New ADODB.Recordset
    Dim xCampos() As String
    Dim TITULO_ As String
        
    Select Case TIPO_
        Case 0
            ReDim xCampos(0, 3) As String
            xCampos(0, 0) = "DNI":                   xCampos(0, 1) = "numdoc":        xCampos(0, 2) = 0:      xCampos(0, 3) = "500"
            
            TITULO_ = "ERRORES DE RECOLECCION DE MARCACION DE PERSONAL"
            RstErr.Filter = adFilterNone
            Set xRs = RstErr
        
        Case 1
            ReDim xCampos(2, 3) As String
            
            xCampos(0, 0) = "Id":                   xCampos(0, 1) = "idemp":        xCampos(0, 2) = 2:      xCampos(0, 3) = "500"
            xCampos(1, 0) = "DNI":                  xCampos(1, 1) = "numdoc":       xCampos(1, 2) = 0:      xCampos(1, 3) = "900"
            xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":       xCampos(2, 2) = 0:      xCampos(2, 3) = "2500"
            
            TITULO_ = "PERSONAL DADO DE BAJA"
            RstReg.Filter = adFilterNone
            Set xRs = RstReg
        
        Case 2
            ReDim xCampos(2, 3) As String
            
            xCampos(0, 0) = "Id":                   xCampos(0, 1) = "id":        xCampos(0, 2) = 2:      xCampos(0, 3) = "500"
            xCampos(1, 0) = "DNI":                  xCampos(1, 1) = "numdoc":       xCampos(1, 2) = 0:      xCampos(1, 3) = "900"
            xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":       xCampos(2, 2) = 0:      xCampos(2, 3) = "2500"
            
            TITULO_ = "PERSONAL EXPORTADO AL SISTEMA DE ASISTENCIA"
            RstExp.Filter = adFilterNone
            Set xRs = RstExp
        
        Case 3
            ReDim xCampos(4, 3) As String
            
            xCampos(0, 0) = "Codigo":               xCampos(0, 1) = "CODIGO":               xCampos(0, 2) = 2:      xCampos(0, 3) = "500"
            xCampos(1, 0) = "DNI":                  xCampos(1, 1) = "DNI":                  xCampos(1, 2) = 0:      xCampos(1, 3) = "500"
            xCampos(2, 0) = "Apellido Paterno":     xCampos(2, 1) = "APELLIDO_PATERNO":     xCampos(2, 2) = 0:      xCampos(2, 3) = "2500"
            xCampos(3, 0) = "Apellido Materno":     xCampos(3, 1) = "APELLIDO_MATERNO":     xCampos(3, 2) = 0:      xCampos(3, 3) = "2500"
            xCampos(4, 0) = "Nombres":              xCampos(4, 1) = "NOMBRES":              xCampos(4, 2) = 0:      xCampos(4, 3) = "2500"
            
            TITULO_ = "PERSONAL DADO DE BAJA EN SERVIDOR DE ASISTENCIA"
            RstBajSer.Filter = adFilterNone
            Set xRs = RstBajSer
        
        Case 4
            ReDim xCampos(4, 3) As String
            
            xCampos(0, 0) = "Codigo":               xCampos(0, 1) = "CODIGO":               xCampos(0, 2) = 2:      xCampos(0, 3) = "500"
            xCampos(1, 0) = "DNI":                  xCampos(1, 1) = "DNI":                  xCampos(1, 2) = 0:      xCampos(1, 3) = "500"
            xCampos(2, 0) = "Apellido Paterno":     xCampos(2, 1) = "APELLIDO_PATERNO":     xCampos(2, 2) = 0:      xCampos(2, 3) = "2500"
            xCampos(3, 0) = "Apellido Materno":     xCampos(3, 1) = "APELLIDO_MATERNO":     xCampos(3, 2) = 0:      xCampos(3, 3) = "2500"
            xCampos(4, 0) = "Nombres":              xCampos(4, 1) = "NOMBRES":              xCampos(4, 2) = 0:      xCampos(4, 3) = "2500"
            
            TITULO_ = "PERSONAL DADO DE ALTA EN SERVIDOR DE ASISTENCIA"
            RstAltSer.Filter = adFilterNone
            Set xRs = RstAltSer
            
    End Select
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , TITULO_, "", "", TITULO_, xRs, xCampos
    Set oExport = Nothing
    Set xRs = Nothing
End Sub

Private Function darBaja(FECH_ As String) As Boolean
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim RstPerLab As New ADODB.Recordset
    Dim nSQLId As String
    Dim CORRELATIVO_ As Double
    Dim A As Integer
    Dim xHorIni As Date
    Dim AGREGARNUEVO_ As Boolean

On Error GoTo HORROR_
    
    xHorIni = Time
    
    cSQL = hallarConsulta(1, FECH_)
    If cSQL = "" Then darBaja = False: Exit Function
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon

    nSQLId = GENERAR_SQL_ID_RST(xRs, "idemp", " And pla_empleados.id", "NOT IN", True)

    cSQL = "SELECT pla_empleados.id AS idemp, pla_empleados.numdoc, pla_empleados.nombre " _
        + vbCr + "FROM pla_empleados " _
        + vbCr + "WHERE (((pla_empleados.fchcese) Is Null)) " & nSQLId

    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then darBaja = False: Exit Function
    If xRs.RecordCount = 0 Then darBaja = False: Exit Function
    
    If RstReg.State = 0 Then DEFINIR_RST_TMP RstReg, xRs
    ' Se cargan los registros dados de baja
    If xRs.RecordCount > 0 Then CARGAR_RST_TMP RstReg, xRs
    
    RST_Busq RstPerLab, "SELECT * FROM pla_periodolaboral ; ", xCon
    
    xCon.BeginTrans
    xRs.MoveFirst
    
    PgBar.Min = 0
    PgBar.Max = xRs.RecordCount
    
    For A = 1 To xRs.RecordCount
        Me.Refresh
        PgBar.Value = A
        ' Hallamos el correlativo
        cSQL = "SELECT pla_periodolaboral.idemp, Max(pla_periodolaboral.corr) AS MáxDecorr " _
        + vbCr + "FROM pla_periodolaboral " _
        + vbCr + "GROUP BY pla_periodolaboral.idemp " _
        + vbCr + "HAVING (((pla_periodolaboral.idemp)=" & NulosN(xRs("idemp")) & "));"

        RST_Busq xRsAux, cSQL, xCon
        
        AGREGARNUEVO_ = False
        If xRsAux.State = 0 Then CORRELATIVO_ = 1: AGREGARNUEVO_ = True
        If xRsAux.RecordCount = 0 Then CORRELATIVO_ = 1: AGREGARNUEVO_ = True
        
        If AGREGARNUEVO_ Then
            RstPerLab.AddNew
            RstPerLab("idemp") = NulosN(xRs("idemp"))
            RstPerLab("fchini") = CDate(FECH_)
            RstPerLab("corr") = CORRELATIVO_
            RstPerLab("idcat") = 6
            RstPerLab("idfinper") = 1
        Else
            CORRELATIVO_ = NulosN(xRsAux("MáxDecorr"))
            RstPerLab.Filter = "idemp = " & NulosN(xRs("idemp")) & " And corr = " & CORRELATIVO_
        End If
        
        RstPerLab("fchfin") = CDate(FECH_)
        RstPerLab.Update
        
        ' Se actualiza la fecha de cese
        cSQL = "UPDATE pla_empleados SET pla_empleados.fchcese = CDate('" & FECH_ & "') " _
        + vbCr + "WHERE (((pla_empleados.id)=" & NulosN(xRs("idemp")) & "));"

        xCon.Execute cSQL
        
        ' Grabamos el movimiento en la Tabla de Empleados
        GrabarOperacion xIdUsuario, 59, 7, xHorIni, Time, Date, xCon, NulosN(xRs("idemp"))
    
        xRs.MoveNext
     Next A
    
    xCon.CommitTrans
    Set xRs = Nothing
    darBaja = True
    Exit Function
    
HORROR_:
    'Resume
    xCon.RollbackTrans
    MsgBox "Ocurrió un error al tratar de procesar las bajas del Sistema", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set xRs = Nothing
    darBaja = True
End Function

Sub preparaRST(ByRef xRs As Recordset)
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(1, 3) As String

    xCampos(0, 0) = "numdoc":      xCampos(0, 1) = "C":      xCampos(0, 2) = "100"
    Set xRs = xFun.CrearRstTMP(xCampos)
    xRs.Open
End Sub

Private Function exportarMarcador() As Boolean
    Dim xRs As New ADODB.Recordset
    Dim RstPer As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstEstTra As New ADODB.Recordset
    Dim RstHorTra As New ADODB.Recordset
    Dim RstPerLec As New ADODB.Recordset
    Dim RstPerHue As New ADODB.Recordset
    Dim nSQLId As String
    Dim IDPER_ As Double
    Dim IDHUE_ As Double
    Dim IDEMP_ As String
   
'On Error GoTo ERROR_

    If conSQLS.State = 0 Then conectarBD ("TEMPUS")
    
    ' Se halla el codigo de empleado y de huella
    cSQL = "SELECT TOP 1 * FROM TEMPUS.PERSONAL ORDER BY (CODIGO + 0) DESC "
    RST_Busq xRs, cSQL, conSQLS
    
    IDPER_ = NulosN(xRs("CODIGO")) + 1
    IDHUE_ = HallaCodigoTabla("TEMPUS.PERSONAL_HUELLA", conSQLS, "INDICE_HUELLA")
    IDEMP_ = "03"
    
    ' Se busca el personal del Servidor de asistencia
    cSQL = "SELECT TEMPUS.PERSONAL.* " _
            + vbCr + "FROM TEMPUS.PERSONAL"
    Set xRs = Nothing
    RST_Busq xRs, cSQL, conSQLS
    
    If xRs.State = 0 Then GoTo ERROR_
    If xRs.RecordCount = 0 Then GoTo ERROR_
        
    nSQLId = GENERAR_SQL_ID_RST(xRs, "DNI", " And pla_empleados.numdoc", "NOT IN", False)
    
    cSQL = "SELECT * FROM pla_empleados " _
        + vbCr + "WHERE ((fchcese) Is Null)" & nSQLId
        
    Set RstExp = Nothing
    RST_Busq RstExp, cSQL, xCon
    
    If RstExp.State = 0 Then GoTo ERROR_
    If RstExp.RecordCount = 0 Then
        MsgBox "No se encontró personal para exportar al Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        exportarMarcador = False
        Exit Function
    End If
    
    conSQLS.BeginTrans
    Me.MousePointer = vbHourglass
    
    RST_Busq RstPer, "SELECT top 1 * FROM TEMPUS.PERSONAL ", conSQLS
    RST_Busq RstTar, "SELECT top 1 * FROM TEMPUS.TARJETAS ", conSQLS
    RST_Busq RstEstTra, "SELECT top 1 * FROM TEMPUS.ESTADO_TRABAJADORES ", conSQLS
    RST_Busq RstHorTra, "SELECT top 1 * FROM TEMPUS.HORARIO_TRABAJADORES ", conSQLS
    RST_Busq RstPerLec, "SELECT top 1 * FROM TEMPUS.PERSONAL_LECTORAS ", conSQLS
    RST_Busq RstPerHue, "SELECT top 1 * FROM TEMPUS.PERSONAL_HUELLA ", conSQLS
        
    RstExp.MoveFirst
    
    PgBar.Min = 0
    PgBar.Max = RstExp.RecordCount
    
    For A = 1 To RstExp.RecordCount
        Me.Refresh
        PgBar.Value = A
        
        ' TEMPUS.PERSONAL
        RstPer.AddNew
        RstPer("EMPRESA") = IDEMP_
        RstPer("CODIGO") = NulosC(IDPER_)
        RstPer("CENTRO_DE_RESPONSABILIDAD") = "000"
        RstPer("OFICINA") = "001"
        RstPer("GRUPOSAN") = "0"
        RstPer("IDORG1") = "000"
        RstPer("DIVISION") = "000"
        RstPer("CARGO") = "037"
        RstPer("IDORG2") = "000"
        RstPer("CATEGORIA") = "002"
        RstPer("CENTRO_DE_COSTO") = "001"
        RstPer("APELLIDO_PATERNO") = UCase(NulosC(RstExp("apepat")))
        RstPer("APELLIDO_MATERNO") = UCase(NulosC(RstExp("apemat")))
        RstPer("NOMBRES") = UCase(NulosC(RstExp("nom")))
        RstPer("FECHA_DE_NACIMIENTO") = RstExp("fchnac")
        RstPer("FECHA_DE_INGRESO") = RstExp("fching")
        RstPer("TARJETA_TMP") = NulosC(RstExp("numdoc"))
        RstPer("TOPE_G") = 0
        RstPer("ESTADO_TMP") = "001"
        RstPer("TIPO") = 0
        RstPer("DNI") = NulosC(RstExp("numdoc"))
        RstPer("REUNIF") = 1
        RstPer("SUELDO") = 0
        RstPer("MAIL_ALIAS") = 0
        RstPer("FLGBLOQUEO") = 0
        RstPer("BLOQUEADO") = 0
        RstPer("ESTADOTB") = 1
        RstPer("JEFEGUARDIA") = 0
        RstPer.Update
        ' TEMPUS.TARJETAS
        RstTar.AddNew
        RstTar("EMPRESA") = IDEMP_
        RstTar("CODIGO") = NulosC(IDPER_)
        RstTar("IDDIA") = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00")
        RstTar("NUMERO_TARJETA") = NulosC(RstExp("numdoc"))
        RstTar("HORA_DE_VIGENCIA") = Format(Time, FORMAT_HORA_SIN_SEGUNDO)
        RstTar("FECHA_DE_VIGENCIA") = Date
        RstTar("TMP_LISTAR") = False
        RstTar("SITEMPORAL") = False
        RstTar.Update
        ' TEMPUS.ESTADO_TRABAJADORES
        RstEstTra.AddNew
        RstEstTra("IDDIA") = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00")
        RstEstTra("ESTADO") = "001"
        RstEstTra("EMPRESA") = IDEMP_
        RstEstTra("CODIGO") = NulosC(IDPER_)
        RstEstTra("FECHA_DE_VIGENCIA") = Date
        RstEstTra("ANULADO") = 0
        RstEstTra("DOCUMENTO") = NulosC(IDPER_)
        RstEstTra.Update
        ' TEMPUS.HORARIO_TRABAJADORES
        RstHorTra.AddNew
        RstHorTra("HORARIO") = "001"
        RstHorTra("TIPO_HORARIO") = "S"
        RstHorTra("EMPRESA") = IDEMP_
        RstHorTra("CODIGO") = NulosC(IDPER_)
        RstHorTra("IDDIA") = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00")
        RstHorTra("INTERVALO") = 1
        RstHorTra("FECHA_DE_VIGENCIA") = Date
        RstHorTra("FECHA_VIGENCIA_CICLO") = Date
        RstHorTra.Update
        ' TEMPUS.PERSONAL_LECTORAS
        For B = 1 To 2
            RstPerLec.AddNew
            RstPerLec("EMPRESA") = IDEMP_
            RstPerLec("CODIGO") = NulosC(IDPER_)
            RstPerLec("IDTERMINAL") = 1
            RstPerLec("IDLECTORA") = B
            RstPerLec("ID_HORARIO_SEM") = 0
            RstPerLec("RESTRINGIDO") = False
            RstPerLec("BLOQUEADO") = False
            RstPerLec("SKIP_CLAVE") = True
            RstPerLec("RESERVA1") = False
            RstPerLec("RESERVA2") = False
            RstPerLec("SKIP_FOTO") = False
            RstPerLec("SKIP_HUELLA") = True
            RstPerLec("NIVEL_SEG") = 0
            RstPerLec.Update
        Next B
        ' TEMPUS.PERSONAL_HUELLA
        RstPerHue.AddNew
        RstPerHue("INDICE_HUELLA") = IDHUE_
        RstPerHue("EMPRESA") = IDEMP_
        RstPerHue("CODIGO") = IDPER_
        RstPerHue("ENVIADO") = 0
        RstPerHue("REGISTRADO") = 0
        RstPerHue("MODIFICADO") = 0
        RstPerHue.Update
            
        RstExp.MoveNext
        IDPER_ = IDPER_ + 1
        IDHUE_ = IDHUE_ + 1
    Next A
    
    conSQLS.CommitTrans
    exportarMarcador = True
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Exit Function
    
ERROR_:
    conSQLS.RollbackTrans
    Me.MousePointer = vbDefault
    exportarMarcador = False
    MsgBox "Ocurrio un error al tratar de exportar personal al Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Private Function darAltaServidorAsistencia() As Boolean
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    Dim IDPER_ As String
 
On Error GoTo ERROR_

    If conSQLS.State = 0 Then conectarBD ("TEMPUS")
    
    ' Se busca todos los activos en el sistema
    cSQL = "SELECT * FROM pla_empleados " _
        + vbCr + "WHERE ((fchcese) Is Null)"
        
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    nSQLId = GENERAR_SQL_ID_RST(xRs, "numdoc", " And TEMPUS.PERSONAL.DNI", "IN", False)
    
    ' Se halla el listado de personal inactivo del servidor
    cSQL = "SELECT TEMPUS.PERSONAL.APELLIDO_PATERNO, TEMPUS.PERSONAL.APELLIDO_MATERNO, TEMPUS.PERSONAL.NOMBRES, TEMPUS.PERSONAL.FECHA_DE_CESE, TEMPUS.PERSONAL.CODIGO, TEMPUS.PERSONAL.DNI " _
        + vbCr + "FROM TEMPUS.PERSONAL " _
        + vbCr + "WHERE (((TEMPUS.PERSONAL.FECHA_DE_CESE) Is Not Null)) " & nSQLId
    
    Set RstAltSer = Nothing
    RST_Busq RstAltSer, cSQL, conSQLS
        
    If RstAltSer.State = 0 Then GoTo ERROR_
    If RstAltSer.RecordCount = 0 Then
        MsgBox "No se encontró personal para dar mantenimiento en el Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        darAltaServidorAsistencia = False
        Exit Function
    End If
          
    conSQLS.BeginTrans
    Me.MousePointer = vbHourglass
    
    PgBar.Min = 0
    PgBar.Max = RstAltSer.RecordCount
    
    RstAltSer.MoveFirst
    For A = 1 To RstAltSer.RecordCount
        Me.Refresh
        PgBar.Value = A
        
        IDPER_ = NulosC(RstAltSer("CODIGO"))
        
        ' TEMPUS.ESTADO_TRABAJADORES
        cSQL = "DELETE FROM TEMPUS.ESTADO_TRABAJADORES " _
            + vbCr + "WHERE (((TEMPUS.ESTADO_TRABAJADORES.CODIGO)='" & IDPER_ & "') AND ((TEMPUS.ESTADO_TRABAJADORES.ESTADO)='002'))"
        
        conSQLS.Execute cSQL
        
        ' TEMPUS.PERSONAL
        cSQL = "UPDATE TEMPUS.PERSONAL SET TEMPUS.PERSONAL.FECHA_DE_CESE = Null " _
            + vbCr + "WHERE (((TEMPUS.PERSONAL.CODIGO)='" & IDPER_ & "'))"
        
        conSQLS.Execute cSQL
        
        RstAltSer.MoveNext
    Next A
    
    conSQLS.CommitTrans
    darAltaServidorAsistencia = True
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Exit Function
    
ERROR_:
    Resume
    conSQLS.RollbackTrans
    Me.MousePointer = vbDefault
    darAltaServidorAsistencia = False
    MsgBox "Ocurrió un error al tratar de realizar mantenimiento al Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Private Function darBajaServidorAsistencia(FECH_ As String) As Boolean
    Dim xRs As New ADODB.Recordset
    Dim RstPer As New ADODB.Recordset
    Dim RstEstTra As New ADODB.Recordset
    Dim nSQLId As String
    Dim IDPER_ As String
    Dim IDEMP_ As String
   
On Error GoTo ERROR_

    If conSQLS.State = 0 Then conectarBD ("TEMPUS")
    
    ' Se halla el personal activo del sistema
    cSQL = "SELECT * FROM pla_empleados " _
        + vbCr + "WHERE ((fchcese) Is Null)"
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then GoTo ERROR_
    If xRs.RecordCount = 0 Then GoTo ERROR_
        
    nSQLId = GENERAR_SQL_ID_RST(xRs, "numdoc", " And TEMPUS.PERSONAL.DNI", "NOT IN", False)
        
    ' Se halla el listado de personal activo del servidor
    cSQL = "SELECT TEMPUS.PERSONAL.APELLIDO_PATERNO, TEMPUS.PERSONAL.APELLIDO_MATERNO, TEMPUS.PERSONAL.NOMBRES, TEMPUS.PERSONAL.FECHA_DE_CESE, TEMPUS.PERSONAL.CODIGO, TEMPUS.PERSONAL.DNI " _
        + vbCr + "FROM TEMPUS.PERSONAL " _
        + vbCr + "WHERE (((TEMPUS.PERSONAL.FECHA_DE_CESE) Is Null)) " & nSQLId
    
    Set RstBajSer = Nothing
    RST_Busq RstBajSer, cSQL, conSQLS
    
    If RstBajSer.State = 0 Then GoTo ERROR_
    If RstBajSer.RecordCount = 0 Then
        MsgBox "No se encontró personal para dar mantenimiento en el Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        darBajaServidorAsistencia = False
        Exit Function
    End If
          
    conSQLS.BeginTrans
    Me.MousePointer = vbHourglass
    
    RST_Busq RstEstTra, "SELECT top 1 * FROM TEMPUS.ESTADO_TRABAJADORES ", conSQLS
    
    PgBar.Min = 0
    PgBar.Max = RstBajSer.RecordCount
    
    RstBajSer.MoveFirst
    For A = 1 To RstBajSer.RecordCount
        Me.Refresh
        PgBar.Value = A
        
        IDPER_ = NulosC(RstBajSer("CODIGO"))
        IDEMP_ = Busca_Codigo(IDPER_, "CODIGO", "EMPRESA", "TEMPUS.PERSONAL", "C", conSQLS)
        
        ' TEMPUS.PERSONAL
        cSQL = "UPDATE TEMPUS.PERSONAL SET TEMPUS.PERSONAL.FECHA_DE_CESE = '" & FECH_ & "' " _
            + vbCr + "WHERE (((TEMPUS.PERSONAL.CODIGO)='" & IDPER_ & "'))"
        conSQLS.Execute cSQL
        
        ' TEMPUS.ESTADO_TRABAJADORES
        cSQL = "SELECT TEMPUS.ESTADO_TRABAJADORES.* " _
            + vbCr + "From TEMPUS.ESTADO_TRABAJADORES " _
            + vbCr + "WHERE (((TEMPUS.ESTADO_TRABAJADORES.CODIGO)='" & IDPER_ & "') AND ((TEMPUS.ESTADO_TRABAJADORES.ESTADO)='002'))"
            
        Set xRs = Nothing
        RST_Busq xRs, cSQL, conSQLS
        
        If xRs.RecordCount = 0 Then
            RstEstTra.AddNew
            RstEstTra("IDDIA") = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00")
            RstEstTra("ESTADO") = "002"
            RstEstTra("EMPRESA") = IDEMP_
            RstEstTra("CODIGO") = NulosC(IDPER_)
            RstEstTra("FECHA_DE_VIGENCIA") = Date
            RstEstTra("ANULADO") = 0
            RstEstTra("DOCUMENTO") = NulosC(IDPER_)
            RstEstTra.Update
        End If
            
        RstBajSer.MoveNext
        IDPER_ = IDPER_ + 1
        IDHUE_ = IDHUE_ + 1
    Next A
    
    conSQLS.CommitTrans
    darBajaServidorAsistencia = True
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Exit Function
ERROR_:
    'Resume
    conSQLS.RollbackTrans
    Me.MousePointer = vbDefault
    darBajaServidorAsistencia = False
    MsgBox "Ocurrio un error al tratar de realizar mantenimiento al Servidor de Asistencia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Function Grabar(ByRef RST_ As ADODB.Recordset) As Boolean
    Dim RstCab As New ADODB.Recordset
    Dim IDEMP_ As Double
    Dim IDEMPTEMP_ As Double
    
On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
      
    RST_Busq RstCab, "SELECT top 1 * FROM pla_recmarcacion ", xCon
    
    If ckAvanzado.Value = 0 Then
        cSQL = "DELETE * FROM pla_recmarcacion " _
            + vbCr + "WHERE (dia = CDate('" & Date - 1 & "'))"
    Else
        cSQL = "DELETE * FROM pla_recmarcacion " _
            + vbCr + "WHERE (dia >= CDate('" & TxtFchIni.valor & "') And dia <=CDate('" & TxtFchFin.valor & "'))"
    End If
     
    xCon.Execute cSQL
    
    PgBar.Min = 0
    PgBar.Max = RST_.RecordCount
    
    RST_.MoveFirst
    While Not RST_.EOF
        PgBar.Value = PgBar.Value + 1
        
        IDEMP_ = NulosN(Busca_Codigo(NulosC(RST_("NUMERO_TARJETA")), "numdoc", "id", "pla_empleados", "C", xCon))
        
        If IDEMP_ = 0 Then
            RstErr.Filter = "numdoc = " & NulosC(RST_("NUMERO_TARJETA"))
            If RstErr.RecordCount = 0 Then
                RstErr.AddNew
                RstErr("numdoc") = NulosC(RST_("NUMERO_TARJETA"))
                RstErr.Update
            End If
        Else
            RstCab.AddNew
            RstCab("dia") = Format(RST_("FECHA"), "dd/mm/yyyy")
            RstCab("hora") = Format(RST_("HORA"), FORMAT_HORA_AL_SEGUNDO)
            RstCab("numdoc") = NulosC(RST_("NUMERO_TARJETA"))
            RstCab("idemp") = IDEMP_
            RstCab.Update
        End If
        
        IDEMPTEMP_ = IDEMP_
        RST_.MoveNext
    Wend
    
    xCon.CommitTrans
    Grabar = True
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Exit Function

LaCague:
    'Resume
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstIns = Nothing:    Set RstTar = Nothing:    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    ' Se cierra la conexion
    If conSQLS.State = 1 Then conSQLS.Close
End Sub

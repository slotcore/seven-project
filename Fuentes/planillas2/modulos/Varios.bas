Attribute VB_Name = "Varios"
Public Enum e_Marcacion
    e_Asist_Permiso = 1
    e_Asist_Licencia = 2
    e_Asist_Vacaciones = 3
    e_Asist_Todos = 4 '--todas las marcaciones
    e_Asist_DomingoDiaFestivos = 5
End Enum
'*****************************************************************************************************

'********** INICIO MARCACION DE ASISTENCIA ********************
Public Sub pMacacionDia(dFecha As Date, IdMarcacion As e_Marcacion, Optional IdEmp = -1)
    '--este proceso generara la marcacion en automatico de vacaciones, permisos, licencias, dias festivos
    '--tambien contemplara los dias domingos(pues se considera como feriado)
    '--OBS:pendiente: falta contemplar los sabados, pues los sabados se trabaja 1/2 tiempo, se tendra que completar para considerar las 8 horas
    Dim nSQL As String
    Dim mEmp&, mIdEmp&, mIdMarca&
    Dim RstEmp As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim RstMarca As New ADODB.Recordset '--para la marcacion * dia
    Dim RstMarcaDet As New ADODB.Recordset
    Dim RstMarcaHora As New ADODB.Recordset
    Dim nSQLEmp As String
    On Error GoTo error
    
    If IdEmp <> -1 Then
        nSQLEmp = " AND pla_empleados.id=" & IdEmp
    End If
    
    '--consultar todos los empleados
    nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres,mae_horario.id AS IdHorario, mae_horario.tolerancia, mae_horariohora.hingreso, mae_horariohora.hsalida, 'Marcación' AS origen " _
        + vbCr + " FROM ((mae_horario INNER JOIN (pla_empleados INNER JOIN mae_horarioemp ON pla_empleados.id = mae_horarioemp.idemp) ON mae_horario.id = mae_horarioemp.idhor) INNER JOIN mae_horariohora ON mae_horario.id = mae_horariohora.idhor) INNER JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp " _
        + vbCr + " WHERE (((mae_horariohora.idhora) = 1) And ((mae_horarioemp.vigencia) = -1)) AND pla_periodolaboral.fchfin Is Null " & nSQLEmp _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]; "
        
    RST_Busq RstEmp, nSQL, xCon
    If RstEmp.RecordCount = 0 Then
        Set RstEmp = Nothing
        Exit Sub
    End If
    RstEmp.MoveFirst
    '************************
    '--ver si esta el registro del dia en pla_marcacion
    nSQL = "SELECT pla_marcacion.id FROM pla_marcacion WHERE pla_marcacion.dia = cdate('" & dFecha & "');"
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        mIdMarca = RstTmp.Fields("id")
        '--eliminar datos en auto con respecto a tiempos
        '--origen:indica quien genero el origen =>> ver tabla pla_origenes
        '--idori=1: Asistencia
        If IdMarcacion = e_Asist_Todos Then
            xCon.Execute "Delete from pla_marcaciondet where idori <> 1 AND idmarca= " & mIdMarca & "; "
            '--eliminar datos en auto con respecto a horas
            '-- idhora=1: Hora Normal
            xCon.Execute "Delete from pla_marcacionhora where idhora <> 1 AND idmarca= " & mIdMarca & "; "
        '--dias festivos
        ElseIf IdMarcacion = e_Asist_DomingoDiaFestivos Then
            xCon.Execute "Delete from pla_marcaciondet where idori in (6,7) AND idmarca= " & mIdMarca & "; "
            xCon.Execute "Delete from pla_marcacionhora where idhora in (9,10) AND idmarca= " & mIdMarca & "; "
        '--vacaciones
        ElseIf IdMarcacion = e_Asist_Todos Or IdMarcacion = e_Asist_Vacaciones Then
            xCon.Execute "Delete from pla_marcaciondet where idori =4 AND idmarca= " & mIdMarca & "; "
            xCon.Execute "Delete from pla_marcacionhora where idhora =4 AND idmarca= " & mIdMarca & "; "
        '--
        ElseIf IdMarcacion = e_Asist_Permiso Then
            xCon.Execute "Delete from pla_marcaciondet where idori =2 AND idmarca= " & mIdMarca & "; "
            xCon.Execute "Delete from pla_marcacionhora where idhora in (5,6) AND idmarca= " & mIdMarca & "; "
        '--
        ElseIf IdMarcacion = e_Asist_Licencia Then
            xCon.Execute "Delete from pla_marcaciondet where idori =3 AND idmarca= " & mIdMarca & "; "
            xCon.Execute "Delete from pla_marcacionhora where idhora in (7,8) AND idmarca= " & mIdMarca & "; "
        End If
        '--
    Else
        mIdMarca = HallaCodigoTabla("pla_marcacion", xCon, "id") 'nuevo id
        RST_Busq RstMarca, "SELECT TOP 1 * FROM pla_marcacion", xCon
        RstMarca.AddNew
        RstMarca("id") = mIdMarca
        RstMarca("dia") = CDate(dFecha)
        RstMarca.Update
    End If
    RST_Busq RstMarcaDet, "SELECT TOP 1 * FROM pla_marcaciondet", xCon
    RST_Busq RstMarcaHora, "SELECT TOP 1 * FROM pla_marcacionhora", xCon
    '************************
    Do While Not RstEmp.EOF
        DoEvents
        mIdEmp = RstEmp.Fields("id")
        If IdMarcacion = e_Asist_Todos Or IdMarcacion = e_Asist_DomingoDiaFestivos Then
            '--domingo
            If LCase(Format(dFecha, "dddd")) = "domingo" Then
                '--registrar en intervalo de horas
                nSQL = "SELECT pla_empleados.id as idemp, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, mae_horario.tolerancia, mae_horariohora.hingreso AS hini, mae_horariohora.hsalida AS hfin, 'Marcación' AS origen, 9 as IdTipoHora " _
                    + vbCr + " FROM pla_empleados INNER JOIN ((mae_horario INNER JOIN mae_horarioemp ON mae_horario.id = mae_horarioemp.idhor) INNER JOIN mae_horariohora ON mae_horario.id = mae_horariohora.idhor) ON pla_empleados.id = mae_horarioemp.idemp " _
                    + vbCr + " WHERE (((pla_empleados.id) = " & mIdEmp & ") And ((mae_horariohora.idhora) = 1) And ((mae_horarioemp.vigencia) = -1))"
                fMacacionDiaDet mIdMarca, mIdEmp, RstMarcaDet, RstMarcaHora, nSQL, 6
                
            Else
                '--ver si tiene vacaciones (si no lo tiene => insertar los dias festivos)
                nSQL = "SELECT pla_marcaciondet.hingreso, pla_marcaciondet.hsalida " _
                    + vbCr + " FROM pla_marcacion INNER JOIN pla_marcaciondet ON pla_marcacion.id = pla_marcaciondet.idmarca " _
                    + vbCr + " WHERE (((pla_marcacion.dia)=CDate('" & dFecha & "')) AND ((pla_marcaciondet.idori)=4) AND ((pla_marcaciondet.idemp)=" & mIdEmp & "));"
                RST_Busq RstTmp, nSQL, xCon
                If RstTmp.RecordCount = 0 Then
                    '--dias festivos
                    nSQL = "SELECT pla_empleados.id as idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, IIf(CDate('" & dFecha & "')=[mae_diasfestivos].[fchini] And ([mae_diasfestivos].[horini]>[mae_horariohora].[hingreso]),[mae_diasfestivos].[horini],[mae_horariohora].[hingreso]) AS hini, IIf(CDate('" & dFecha & "')=[mae_diasfestivos].[fchfin] And ([mae_diasfestivos].[horfin]<[mae_horariohora].[hsalida]),[mae_diasfestivos].[horfin],[mae_horariohora].[hsalida]) AS hfin, 'Dia Festivo' AS Origen, mae_diasfestivos.descripcion AS motivo, 10 as IdTipoHora " _
                        + vbCr + " FROM mae_diasfestivos, pla_empleados INNER JOIN ((mae_horario INNER JOIN mae_horarioemp ON mae_horario.id = mae_horarioemp.idhor) INNER JOIN mae_horariohora ON mae_horario.id = mae_horariohora.idhor) ON pla_empleados.id = mae_horarioemp.idemp " _
                        + vbCr + " WHERE (((pla_empleados.id)=" & mIdEmp & ") AND ((CDate('" & dFecha & "')) Between [mae_diasfestivos].[fchini] And [mae_diasfestivos].[fchfin]) AND ((mae_horarioemp.vigencia)=-1)) "
                    fMacacionDiaDet mIdMarca, mIdEmp, RstMarcaDet, RstMarcaHora, nSQL, 7
                End If
            End If
            
        End If
        If LCase(Format(dFecha, "dddd")) = "domingo" Then GoTo Ir_Siguiente:
        '--vacaciones
        If IdMarcacion = e_Asist_Todos Or IdMarcacion = e_Asist_Vacaciones Then
            nSQL = "SELECT pla_empleados.id as idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_horariohora.hingreso as hini, mae_horariohora.hsalida as hfin, 'Vacaciones' AS origen, 4 as IdTipoHora  " _
                + vbCr + " FROM (pla_empleados INNER JOIN (mae_horarioemp INNER JOIN mae_horariohora ON mae_horarioemp.idhor = mae_horariohora.idhor) ON pla_empleados.id = mae_horarioemp.idemp) INNER JOIN (pla_vacaciones INNER JOIN pla_vacacionesdet ON pla_vacaciones.id = pla_vacacionesdet.idvac) ON pla_empleados.id = pla_vacaciones.idemp " _
                + vbCr + " WHERE (((pla_empleados.id)=" & mIdEmp & ") AND ((mae_horarioemp.vigencia)=-1) AND ((mae_horariohora.idhora)=1) AND ((CDate('" & dFecha & "')) Between [pla_vacacionesdet].[fchini] And [pla_vacacionesdet].[fchfin]))"
            fMacacionDiaDet mIdMarca, mIdEmp, RstMarcaDet, RstMarcaHora, nSQL, 4
        End If
        '--permiso
        If IdMarcacion = e_Asist_Todos Or IdMarcacion = e_Asist_Permiso Then
            nSQL = "SELECT pla_empleados.id as idemp, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, IIf((CDate('" & dFecha & "')=[pla_permiso].[fchini]) And ([pla_permiso].[horini]>[mae_horariohora].[hingreso]),[pla_permiso].[horini],[mae_horariohora].[hingreso]) AS hini, IIf((CDate('" & dFecha & "')=[pla_permiso].[fchfin]) And ([pla_permiso].[horfin]<[mae_horariohora].[hsalida]),[pla_permiso].[horfin],[mae_horariohora].[hsalida]) AS hfin, 'Permiso' AS Origen, mae_tipopermiso.descripcion AS motivo, mae_tipopermiso.gocehaber, IIF (mae_tipopermiso.gocehaber=0,6,5) as IdTipoHora " _
                + vbCr + " FROM (pla_empleados INNER JOIN (mae_horarioemp INNER JOIN mae_horariohora ON mae_horarioemp.idhor = mae_horariohora.idhor) ON pla_empleados.id = mae_horarioemp.idemp) INNER JOIN (mae_tipopermiso INNER JOIN pla_permiso ON mae_tipopermiso.id = pla_permiso.idper) ON pla_empleados.id = pla_permiso.idemp " _
                + vbCr + " WHERE (((pla_empleados.id)=" & mIdEmp & ") AND ((CDate('" & dFecha & "')) Between [pla_permiso].[fchini] And [pla_permiso].[fchfin]) AND ((mae_horarioemp.vigencia)=-1) AND ((mae_horariohora.idhora)=1))"
            fMacacionDiaDet mIdMarca, mIdEmp, RstMarcaDet, RstMarcaHora, nSQL, 2
        End If
        '--licencia
        If IdMarcacion = e_Asist_Todos Or IdMarcacion = e_Asist_Licencia Then
            nSQL = "SELECT pla_empleados.id as idemp, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, IIf((CDate('" & dFecha & "')=pla_licencia.fchini) And (pla_licencia.horini>mae_horariohora.hingreso),pla_licencia.horini,mae_horariohora.hingreso) AS hini, IIf((CDate('" & dFecha & "')=pla_licencia.fchfin) And (pla_licencia.horfin<mae_horariohora.hsalida),pla_licencia.horfin,mae_horariohora.hsalida) AS hfin, 'Licencia' AS Origen, mae_tipolicencia.descripcion AS motivo, 7 as IdTipoHora " _
                + vbCr + " FROM (pla_empleados INNER JOIN (mae_horarioemp INNER JOIN mae_horariohora ON mae_horarioemp.idhor = mae_horariohora.idhor) ON pla_empleados.id = mae_horarioemp.idemp) INNER JOIN (mae_tipolicencia INNER JOIN pla_licencia ON mae_tipolicencia.id = pla_licencia.idlic) ON pla_empleados.id = pla_licencia.idemp " _
                + vbCr + " WHERE (((pla_empleados.id)=" & mIdEmp & ") AND (((CDate('" & dFecha & "')) Between [pla_licencia].[fchini] And [pla_licencia].[fchfin]))) "
            fMacacionDiaDet mIdMarca, mIdEmp, RstMarcaDet, RstMarcaHora, nSQL, 3
        End If
        
''            '--del horario por defecto
''            nSQL = "SELECT pla_empleados.id as idemp, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, mae_horario.tolerancia, mae_horariohora.hingreso AS hini, mae_horariohora.hsalida AS hfin, 'Marcación' AS origen, 1 as IdTipoHora  " _
''                + vbCr + " FROM pla_empleados INNER JOIN ((mae_horario INNER JOIN mae_horarioemp ON mae_horario.id=mae_horarioemp.idhor) INNER JOIN mae_horariohora ON mae_horario.id=mae_horariohora.idhor) ON pla_empleados.id=mae_horarioemp.idemp " _
''                + vbCr + " WHERE (((pla_empleados.id) = " & mIdEmp & ") And ((mae_horariohora.idhora) = 1) And ((mae_horarioemp.vigencia) = -1)) "
''            fMacacionDiaDet xCon, mIdMarca, mIdEmp, RstMarcaDet, RstMarcaHora, nSQL, 1

Ir_Siguiente:
        
        RstEmp.MoveNext
    Loop
    Exit Sub
error:
    SHOW_ERROR "Error", "pMarcacionDia"
End Sub

Public Function fMarcacionDefault(dDia As Date) As ADODB.Recordset
                                  
    '--cargar empleados que tienen horario(solo hora normal)
    '--recorrer c/u de los usuarios, revizar si tienen movimientos en marcacion.
    '--     si tienen dejas sin efecto
    '--     si no tienen mov. => crear una marcacion temporal con origen=1 asistencia

    Dim nSQL As String
    Dim RstHorario As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim RstMarca As New ADODB.Recordset
    '--cargar datos iniciales segun fecha, si ya tiene movimiento
    nSQL = "SELECT -1 as IdGrid,pla_empleados.id AS idemp, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, pla_marcaciondet.hingreso AS hini, pla_marcaciondet.hsalida AS hfin, pla_marcaciondet.idori, pla_origenes.descripcion AS origen, pla_marcaciondet.tiporegistro AS tipreg " _
        + vbCr + " FROM pla_empleados INNER JOIN (pla_marcacion INNER JOIN (pla_marcaciondet INNER JOIN pla_origenes ON pla_marcaciondet.idori = pla_origenes.id) ON pla_marcacion.id = pla_marcaciondet.idmarca) ON pla_empleados.id = pla_marcaciondet.idemp " _
        + vbCr + " WHERE (((pla_marcacion.dia)=CDate('" & dDia & "'))) " _
        + vbCr + " ORDER BY pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom, pla_marcaciondet.hingreso; "
        
    RST_Busq RstTmp, nSQL, xCon
    DEFINIR_RST_TMP RstMarca, RstTmp
    CARGAR_RST_TMP RstMarca, RstTmp
    Set RstTmp = Nothing
    '----
    '--cargar solo los que esten activos ver pla_periodolaboral
    nSQL = "SELECT pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_horario.id AS IdHorario, mae_horario.tolerancia, mae_horariohora.hingreso AS hini, mae_horariohora.hsalida AS hfin, 'Horario' AS Origen, [mae_horariohora].[hingreso]+[mae_horario].[tolerancia] AS hinitol " _
        + vbCr + " FROM ((mae_horario INNER JOIN (pla_empleados INNER JOIN mae_horarioemp ON pla_empleados.id = mae_horarioemp.idemp) ON mae_horario.id = mae_horarioemp.idhor) INNER JOIN mae_horariohora ON mae_horario.id = mae_horariohora.idhor) INNER JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp " _
        + vbCr + " Where (((mae_horariohora.idhora) = 1) And ((mae_horarioemp.vigencia) = -1) And ((pla_periodolaboral.fchfin) Is Null)) " _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"

    RST_Busq RstHorario, nSQL, xCon
    If RstHorario.RecordCount <> 0 Then RstHorario.MoveFirst
    Do While Not RstHorario.EOF
        '--ver si tiene movimientos
        nSQL = "SELECT pla_marcacion.id, pla_empleados.id AS idemp, pla_marcaciondet.corr, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, pla_marcaciondet.hingreso AS hini, pla_marcaciondet.hsalida AS hfin, pla_origenes.descripcion AS origen " _
            + vbCr + " FROM pla_empleados INNER JOIN (pla_marcacion INNER JOIN (pla_marcaciondet INNER JOIN pla_origenes ON pla_marcaciondet.idori = pla_origenes.id) ON pla_marcacion.id = pla_marcaciondet.idmarca) ON pla_empleados.id = pla_marcaciondet.idemp " _
            + vbCr + " WHERE (((pla_marcacion.dia) = CDate('" & dDia & "')))  and pla_marcaciondet.idemp = " & RstHorario("idemp") & " " _
            + vbCr + " ORDER BY pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom, pla_marcaciondet.hingreso; "
        RST_Busq RstTmp, nSQL, xCon
        
        If RstTmp.RecordCount <> 0 Then
            RstTmp.MoveFirst
            If TimeValue(RstTmp.Fields("hini")) > TimeValue(RstHorario.Fields("hinitol")) Then
                RstMarca.AddNew
                RstMarca("idemp") = RstHorario("idemp")
                RstMarca("nombres") = RstHorario("nombres")
                RstMarca("hini") = TimeValue(RstHorario.Fields("hini"))
                RstMarca("hfin") = TimeValue(RstTmp.Fields("hini"))
                RstMarca("idori") = "-1"
                RstMarca("origen") = "Asistencia"
                RstMarca("tipreg") = "-1"
                RstMarca.Update
            End If
            If RstTmp.RecordCount <> 1 Then RstTmp.MoveLast
            If TimeValue(RstTmp.Fields("hfin")) < TimeValue(RstHorario.Fields("hfin")) Then
                RstMarca.AddNew
                RstMarca("idemp") = RstHorario("idemp")
                RstMarca("nombres") = RstHorario("nombres")
                RstMarca("hini") = TimeValue(RstTmp.Fields("hfin"))
                RstMarca("hfin") = TimeValue(RstHorario.Fields("hfin"))
                RstMarca("idori") = "-1"
                RstMarca("origen") = "Asistencia"
                RstMarca("tipreg") = "-1"
                RstMarca.Update
            End If
            
        Else
            RstMarca.AddNew
            RstMarca("idemp") = RstHorario("idemp")
            RstMarca("nombres") = RstHorario("nombres")
            RstMarca("hini") = TimeValue(RstHorario.Fields("hini"))
            RstMarca("hfin") = TimeValue(RstHorario.Fields("hfin"))
            RstMarca("idori") = "-1"
            RstMarca("origen") = "Asistencia"
            RstMarca("tipreg") = "-1"
            RstMarca.Update
        End If
        RstHorario.MoveNext
    Loop
    '--ordenar el recordset
    RstMarca.Sort = "idemp,hini"
    '------
    Set RstHorario = Nothing
    Set RstTmp = Nothing
    Set fMarcacionDefault = RstMarca
    
End Function

Private Function fMacacionDiaDet(IdMarca, _
                            mIdEmp, _
                            RstMarcaDet As ADODB.Recordset, _
                            RstMarcaHora As ADODB.Recordset, _
                            nSQL As String, _
                            IdOrigen) As Boolean
                            
    Dim RstTmp As New ADODB.Recordset
    Dim mCorrDet&, mCorrHora&
    Dim mIdTipoHora&
    DoEvents
    '--determinar si hay registro
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount = 0 Then
        Set RstTmp = Nothing
        Exit Function
    End If
    '--obtener el correlativo del detalle de la marcacion
    RST_Busq RstTmp, "SELECT TOP 1 pla_marcaciondet.corr FROM pla_marcaciondet WHERE (((pla_marcaciondet.idmarca)=" & IdMarca & ")) AND pla_marcaciondet.idemp = " & mIdEmp & " ORDER BY pla_marcaciondet.corr DESC; ", xCon
    If RstTmp.RecordCount <> 0 Then
        mCorrDet = NulosN(RstTmp.Fields(0)) + 1
    Else
        mCorrDet = 1
    End If
    Set RstTmp = Nothing
    '--obtener el correlativo de los tipos de horas
    RST_Busq RstTmp, "SELECT TOP 1 pla_marcacionhora.corr FROM pla_marcacionhora WHERE (((pla_marcacionhora.idmarca)=" & IdMarca & ")) AND pla_marcacionhora.idemp = " & mIdEmp & "  ORDER BY pla_marcacionhora.corr DESC; ", xCon
    If RstTmp.RecordCount <> 0 Then
        mCorrHora = NulosN(RstTmp.Fields(0)) + 1
    Else
        mCorrHora = 1
    End If
    Set RstTmp = Nothing
    '------------
    RST_Busq RstTmp, nSQL, xCon
    If IdOrigen <> -1 Then
        RstMarcaDet.AddNew
        RstMarcaDet("idmarca") = IdMarca
        RstMarcaDet("idemp") = mIdEmp
        RstMarcaDet("corr") = mCorrDet
        RstMarcaDet("hingreso") = RstTmp("hini")
        RstMarcaDet("hsalida") = RstTmp("hfin")
        '--0=ingreso manual (cuando se ingreso o modifica los datos) ;-1=ingreso automático (lector barra, con tarjeta, ingreso individual)
        '-- -1= tambien cuando se haga el proceso de calculo de vacaciones,permiso,licencia en forma automatico
        RstMarcaDet("tiporegistro") = "1"
        RstMarcaDet("idori") = IdOrigen
        RstMarcaDet.Update
    End If
    mIdTipoHora = NulosN(RstTmp("IdTipoHora"))
    '--obtener el total de horas
    Dim dTotalHora As Date
    Dim dTotalHoraDescanso As Date
    Dim mTotalSegDescanso&
    Dim mTotalSeg&
    dTotalHora = Format(TimeValue(RstTmp("hfin")) - TimeValue(RstTmp("hini")), "hh:mm:ss")
    mIdTipoHora = RstTmp.Fields("IdTipoHora")
    
    '--ver si tiene descanso idhora=14=descanso
    Dim RstHorario As New ADODB.Recordset
    nSQL = "SELECT mae_horarioemp.idemp, mae_horariohora.hingreso AS hini, mae_horariohora.hsalida AS hfin " _
        + vbCr + " FROM mae_horariohora INNER JOIN mae_horarioemp ON mae_horariohora.idhor = mae_horarioemp.idhor " _
        + vbCr + " WHERE (((mae_horarioemp.idemp)=" & mIdEmp & ") AND ((mae_horarioemp.vigencia)=-1) AND ((mae_horariohora.idhora)=14));"
    RST_Busq RstHorario, nSQL, xCon
    
    If RstHorario.RecordCount <> 0 Then
        '*****************************************************************
        '--averiguar si se va utilizar el descuento por descanso
        Dim HIni, HFin  As Date
        Dim fCalcula As Boolean
        Do While Not RstTmp.EOF
            '--si la marcacion inicial es superior al final => avance
            If RstTmp.Fields("hfin") < RstHorario.Fields("hini") Then GoTo avance
            
            If (RstTmp.Fields("hini") < RstHorario.Fields("hini")) And (RstTmp.Fields("hfin") > RstHorario.Fields("hfin")) Then
                HIni = RstHorario.Fields("hini")
                HFin = RstHorario.Fields("hfin")
                fCalcula = True
            ElseIf (RstTmp.Fields("hfin") > RstHorario.Fields("hini")) And (RstTmp.Fields("hfin") < RstHorario.Fields("hfin")) Then
                HIni = RstHorario.Fields("hini")
                HFin = RstTmp.Fields("hfin")
                fCalcula = True
            ElseIf (RstTmp.Fields("hini") > RstHorario.Fields("hini")) And (RstTmp.Fields("hini") < RstHorario.Fields("hfin")) Then
                HIni = RstTmp.Fields("hfin")
                HFin = RstHorario.Fields("hfin")
                fCalcula = True
            Else
                fCalcula = False
            End If
            If fCalcula = True Then
                mTotalSegDescanso = mTotalSegDescanso + ConvertSeg(Format(HFin - HIni, "hh:mm:ss"))
            End If
                '--
avance:
            RstTmp.MoveNext
        Loop
    '*****************************************************************
    End If
    
    Set RstTmp = Nothing
    '--obtener cantidad de segundos
    mTotalSeg = ConvertSeg(dTotalHora & "") - mTotalSegDescanso
    '--insertando el registro en marcacion_hora
    RstMarcaHora.AddNew
    RstMarcaHora("IdMarca") = IdMarca
    RstMarcaHora("idemp") = mIdEmp
    RstMarcaHora("corr") = mCorrHora
    RstMarcaHora("idhora") = mIdTipoHora
    If mTotalSegDescanso <> 0 Then
        RstMarcaHora("tothor") = TimeValue(ConvertHora(mTotalSeg))
    Else
        RstMarcaHora("tothor") = dTotalHora
    End If
    RstMarcaHora("totseg") = mTotalSeg
    RstMarcaHora.Update
    
    fMacacionDiaDet = True
    
End Function


Public Function pCalculoHoras(RstMarca As ADODB.Recordset, dFecha As Date, IdEmp&, Optional mIdMarca& = -1) As ADODB.Recordset
    '--default= false: carga desde la base de datos (informacion calculada)
    '--         true:  calcula la informacion en funcion el recordset activo(uso temporal)
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    Dim RstHorario As New ADODB.Recordset
    Dim RstHoras As New ADODB.Recordset
    
    If mIdMarca <> -1 Then '--si hay marcacion
    
        nSQL = "SELECT pla_marcacionhora.idhora, mae_tipohora.descripcion, mae_tipohora.prioridad, pla_marcacionhora.tothor, pla_marcacionhora.totseg " _
            + vbCr + " FROM mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora " _
            + vbCr + " WHERE (((pla_marcacionhora.IdMarca) = " & mIdMarca & ") And ((pla_marcacionhora.IdEmp) = " & IdEmp & ")) " _
            + vbCr + " ORDER BY mae_tipohora.prioridad; "
        
        RST_Busq RstHoras, nSQL, xCon
        Set pCalculoHoras = RstHoras
        Exit Function
        
    End If
    
    '--lista de horas de que no sea asistencia,falta,tardanza
    nSQL = "SELECT pla_marcacionhora.idhora,mae_tipohora.descripcion, mae_tipohora.prioridad, pla_marcacionhora.tothor, pla_marcacionhora.totseg " _
        + vbCr + " FROM pla_marcacion LEFT JOIN (mae_tipohora RIGHT JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca " _
        + vbCr + " WHERE (((pla_marcacion.dia)=cdate('" & dFecha & "')) AND ((pla_marcacionhora.idemp)=" & IdEmp & ") AND ((pla_marcacionhora.idhora) In (4,5,6,7,8,10))) " _
        + vbCr + " ORDER BY mae_tipohora.prioridad; "
    
    RST_Busq RstTmp, nSQL, xCon
    
    DEFINIR_RST_TMP RstHoras, RstTmp    '--copiar estructura del recordset rsttmp a rsthoras
    CARGAR_RST_TMP RstHoras, RstTmp     '--copiar datos del recordset rsttmp a rsthoras
    Set RstTmp = Nothing
    '-------------
    RstMarca.Filter = "(idemp = " & IdEmp & " and idori =1) or (idemp = " & IdEmp & " and idori=-1) or (idemp = " & IdEmp & " and idori=5)" '--dif. 2:permiso, 3:licencia, 4:vacaciones
    If RstMarca.RecordCount = 0 Then
        RstHoras.Filter = ""
        Set pCalculoHoras = RstHoras
        Exit Function
    End If
    RstMarca.Sort = "hini asc " '--ordenar ascendente
        
    '--buscando las horas configuradas en horario
    nSQL = "SELECT mae_horariohora.idhora, mae_tipohora.descripcion, mae_horariohora.hingreso AS hini, mae_horariohora.hsalida AS hfin, mae_tipohora.prioridad, mae_horario.tolerancia, IIf([mae_horario].[tolerancia] Is Null,[mae_horariohora].[hingreso],[mae_horariohora].[hingreso]+[mae_horario].[tolerancia]) AS htol" _
        + vbCr + " FROM mae_horario INNER JOIN (mae_tipohora INNER JOIN (mae_horariohora INNER JOIN mae_horarioemp ON mae_horariohora.idhor = mae_horarioemp.idhor) ON mae_tipohora.id = mae_horariohora.idhora) ON mae_horario.id = mae_horariohora.idhor " _
        + vbCr + " WHERE (((mae_horarioemp.idemp)=" & IdEmp & ") AND ((mae_horarioemp.vigencia)= -1 ));"
    RST_Busq RstHorario, nSQL, xCon
    
    RstHorario.Filter = "idhora <> 14" '--tipos de horas que no sea igual al descanso
    If RstHorario.RecordCount = 0 Then
        MsgBox "El Personal no tiene Horario Configurado" + vbCr + "Configure el Hororario al Personal, Luego Proceda", vbExclamation, xTitulo
        Set RstHorario = Nothing
        Exit Function
    End If
    RstHorario.Sort = "hini asc"
    Dim HIni, HFin, HIniHorario As Date
    Dim nHorarioDescripcion As String '
    Dim mPrioridad& '--prioridad de tipo de hora, para ordenar
    Dim mIdTipoHora& '--codigo del tipo de hora igual a tabla mae_tipohora
    RstMarca.MoveFirst
    Do While Not RstMarca.EOF
        Do While Not RstHorario.EOF
            '--ver si el origen es asistencia
            If NulosN(RstHorario.Fields("idhora")) = 1 Then
                HIniHorario = RstHorario.Fields("hini")
            Else
                HIniHorario = RstHorario.Fields("hini")
            End If
            '---
            If IsNull(RstMarca.Fields("hini")) = True Or IsNull(RstMarca.Fields("hfin")) = True Then Exit Function
            '--si la marcacion inicial es superior al final => avance
            If RstMarca.Fields("hini") >= RstHorario.Fields("hfin") Then GoTo avance:
            
            Select Case RstMarca.Fields("idori")
                Case 5 '--hora falta
                    mIdTipoHora = 3
                    nHorarioDescripcion = "Hora Falta"
                    mPrioridad = 1
                Case 6 '--hora domingo
                    mIdTipoHora = 9
                    nHorarioDescripcion = "Hora Domingo"
                    mPrioridad = 1
                Case Else
                    mIdTipoHora = NulosN(RstHorario.Fields("idhora"))
                    nHorarioDescripcion = NulosC(RstHorario.Fields("descripcion"))
                    mPrioridad = NulosN(RstHorario.Fields("prioridad"))
            End Select
            '--hora de inicio
            If NulosN(RstHorario.Fields("idhora")) = 1 Then '--tipo hora normal
                If CDate(RstMarca.Fields("hini")) <= HIniHorario Then
                    HIni = RstHorario.Fields("hini")
                Else
                    HIni = RstMarca.Fields("hini")
                End If
            Else '--otros tipos de horas
                If RstMarca.Fields("hini") < HIniHorario Then
                    HIni = RstHorario.Fields("hini")
                Else
                    HIni = RstMarca.Fields("hini")
                End If
            End If
            
            '--hora de fin
            If RstMarca.Fields("hfin") > RstHorario.Fields("hfin") Then
                HFin = RstHorario.Fields("hfin")
            Else
                HFin = RstMarca.Fields("hfin")
            End If
            
            RstHoras.Filter = "idhora=" & mIdTipoHora
            If RstHoras.RecordCount = 0 Then
                RstHoras.AddNew
                RstHoras("idhora") = mIdTipoHora
                
                RstHoras("descripcion") = nHorarioDescripcion
                RstHoras("prioridad") = mPrioridad
            End If
            '--
            RstHoras("totseg") = NulosN(RstHoras("totseg")) + ConvertSeg(Format(HFin - HIni, "hh:mm:ss"))
            RstHoras("tothor") = ConvertHora(NulosN(RstHoras("totseg")))
            RstHoras.Update
            
            '--salir del bucle de horas
            If RstMarca.Fields("hfin") <= RstHorario.Fields("hfin") Then Exit Do
            '--si es falta salir
            '--obs: solo se considera como falta un registro, pues falta es todo el dia caso contrario se estaria hablando de tardanza
            '--la tardanza no se registra, su calculo es automatico
            If RstMarca.Fields("idori") = 5 Then GoTo avance1
avance:
            RstHorario.MoveNext
        Loop
        
        RstMarca.MoveNext
    Loop
avance1:
    
    RstHoras.Filter = ""
    
    '*******************************************************************************************************************
    
    RstHorario.Filter = "idhora = 14"   '--si tiene horario de descanso
    
    RstMarca.Filter = "(idemp = " & IdEmp & " and idori = -1) or (idemp = " & IdEmp & " and idori =1) or (idemp = " & IdEmp & " and idori =5) or (idemp = " & IdEmp & " and idori =6)"
    If RstMarca.RecordCount <> 0 Then
        RstMarca.Sort = "hini asc"
        RstMarca.MoveFirst
    End If
    
    Const TOTAL_SEG_DIA = 28800         '--total de segundos que se tienen que considerar por dia
    Dim mTotalSegDescanso&              '--indica el total de segundos descanso
    Dim mTotalSegTardanza&              '--indica el total de segundos de tardanza
    Dim mTotalSegAcumulado&
    Dim mTotalHoras&
    Dim fCalcula As Boolean
    Dim fHayFalta As Boolean            '--indica si la marcacion es de falta =>> se considera todo el dia de falta
    If RstHorario.RecordCount <> 0 Then
        '*****************************************************************
        '--averiguar si se va utilizar el descuento por descanso
        Do While Not RstMarca.EOF
            fCalcula = False
            '--si la marcacion inicial es superior al final => avance
            If RstMarca.Fields("idori") = 5 Then
                fHayFalta = True
                Exit Do
            End If
            '--
            If RstMarca.Fields("hini") > RstHorario.Fields("hfin") Then GoTo avance2
            
            If (RstMarca.Fields("hini") < RstHorario.Fields("hini")) And (RstMarca.Fields("hfin") > RstHorario.Fields("hfin")) Then
                '--ej.  hora 01:00 pm >> 02:00 pm
                '       marca  08:30 am >> 06:30 pm
                HIni = RstHorario.Fields("hini")
                HFin = RstHorario.Fields("hfin")
                fCalcula = True
                
            ElseIf (RstMarca.Fields("hini") <= RstHorario.Fields("hini")) And (RstMarca.Fields("hfin") >= RstHorario.Fields("hini") And RstMarca.Fields("hfin") <= RstHorario.Fields("hfin")) Then
                '--ej.  hora 01:00 pm >> 02:00 pm
                '       marca  12:30 pm >> 01:30 pm
                HIni = RstHorario.Fields("hini")
                HFin = RstMarca.Fields("hfin")
                fCalcula = True
                            
            ElseIf (RstMarca.Fields("hini") >= RstHorario.Fields("hini")) And (RstMarca.Fields("hfin") <= RstHorario.Fields("hfin")) Then
                '--ej.  hora 01:00 pm >> 02:00 pm
                '       marca  01:25 pm >> 01:50 pm
                HIni = RstMarca.Fields("hini")
                HFin = RstMarca.Fields("hfin")
                fCalcula = True
                
            ElseIf (RstMarca.Fields("hini") >= RstHorario.Fields("hini")) And (RstMarca.Fields("hfin") >= RstHorario.Fields("hfin")) Then
                '--ej.  hora 01:00 pm >> 02:00 pm
                '       marca  01:30 pm >> 03:45 pm
                HIni = RstMarca.Fields("hini")
                HFin = RstHorario.Fields("hfin")
                fCalcula = True
            Else
                fCalcula = False
            End If
            If fCalcula = True Then
                mTotalSegDescanso = mTotalSegDescanso + ConvertSeg(Format(HFin - HIni, "hh:mm:ss"))
            End If
                '--
avance2:
            RstMarca.MoveNext
        Loop
    End If
    
    '--si hay falta insertar el registro por falta
    If fHayFalta = True Then
        RstHoras.Filter = "idhora=3" '--idhora=3 ver tabla mae_tipohora
        If RstHoras.RecordCount = 0 Then
            RstHoras.AddNew
            RstHoras("idhora") = 3
            RstHoras("descripcion") = "Hora Falta"
            RstHoras("prioridad") = 1
        End If
        RstHoras("totseg") = TOTAL_SEG_DIA
        RstHoras("tothor") = ConvertHora(TOTAL_SEG_DIA)
        RstHoras.Update
        GoTo salir
    End If
    
    '------proceder a restar los tiempos si hay minutos de descanso
''    If mTotalSegDescanso <> 0 Then
        '--actualizando el tipo de hora asistencia
        
        
        RstHoras.Filter = "idhora=1" '
        If RstHoras.RecordCount <> 0 Then
            mTotalHoras = NulosN(RstHoras("totseg")) - mTotalSegDescanso
            RstHoras("totseg") = mTotalHoras
            RstHoras("tothor") = ConvertHora(mTotalHoras)
            RstHoras.Update
        End If
        
        mTotalSegTardanza = TOTAL_SEG_DIA - mTotalHoras
        '--acumulando los tolales de las marcaciones
        RstHoras.Filter = ""
        '--3:hora falta,11:HE 100%, 12:HE 50%,13:HE 25%
        RstHoras.Filter = "idhora<>3 and idhora<>11 and idhora<>12 and idhora<>13"
        mTotalHoras = 0
        If RstHoras.RecordCount <> 0 Then
            RstHoras.MoveFirst
            Do While Not RstHoras.EOF
                mTotalHoras = mTotalHoras + NulosN(RstHoras.Fields("totseg"))
                RstHoras.MoveNext
            Loop
        End If
       
        '--insertando las tardanzas
        If mTotalHoras < TOTAL_SEG_DIA Then
            mTotalSegTardanza = TOTAL_SEG_DIA - mTotalHoras
            RstHoras.AddNew
            RstHoras("idhora") = 2
            RstHoras("descripcion") = "Hora Tardanza"
            RstHoras("prioridad") = 2
            RstHoras("totseg") = mTotalSegTardanza
            RstHoras("tothor") = ConvertHora(NulosN(RstHoras("totseg")))
            RstHoras.Update
        Else
            '--
            RstHoras.Filter = "idhora<>3 and idhora<>11 and idhora<>12 and idhora<>13"
            If RstHoras.RecordCount = 1 Then
                RstHoras("totseg") = TOTAL_SEG_DIA
                RstHoras("tothor") = ConvertHora(TOTAL_SEG_DIA)
                RstHoras.Update
            End If
        End If
        
        
        
''    End If
    '*****************************************************************
salir:
    RstHoras.Filter = ""
    RstHoras.Sort = "prioridad asc"
    Set RstHorario = Nothing
    
    Set pCalculoHoras = RstHoras
    Set RstHoras = Nothing
End Function

'********** FIN MARCACION DE ASISTENCIA ********************

'*******************************************************************************************************************************
'*************** INICIO CONSULTAS ********************
Public Function pBuscarPersonal(RstTmp As ADODB.Recordset, _
                       Optional fSoloActivos As Boolean = True) As ADODB.Recordset
                       
                       
    Dim nSQL As String
    Dim xCampos(6, 4) As String
    Dim nSQLWhere As String
    
    xCampos(0, 0) = "TipDoc":               xCampos(0, 1) = "docabrev":   xCampos(0, 2) = "700":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Numero":               xCampos(1, 1) = "numdoc":     xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombres":    xCampos(2, 2) = "3200":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Sexo":                 xCampos(3, 1) = "sexo":       xCampos(3, 2) = "550":     xCampos(3, 3) = "C"
    xCampos(4, 0) = "Categoría":            xCampos(4, 1) = "categoria":  xCampos(4, 2) = "2000":    xCampos(4, 3) = "C"
    xCampos(5, 0) = "Estado":               xCampos(5, 1) = "estado":     xCampos(5, 2) = "700":     xCampos(5, 3) = "C"

    

'    If fSoloActivos = False Then
'        nSQL = "SELECT * FROM " _
'            + vbCr + " (SELECT pla_empleados.*, mae_dociden.abrev AS docabrev, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat]  & ' ' &  [pla_empleados].[nom]  AS nombres, mae_sexo.abrev AS sexo " _
'            + vbCr + " FROM mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex " _
'            + vbCr + " ORDER BY [pla_empleados].[nom] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] ) AS emp " _
'            + vbCr + " Left Join " _
'            + vbCr + " (SELECT pla_periodolaboral.idemp, Last(mae_categoria.descripcion) AS categoria, Last(pla_periodolaboral.fchini) AS ingreso, Last(pla_periodolaboral.fchfin) AS cese, IIf([cese] Is Not Null,'De Baja','Activo') AS estado " _
'            + vbCr + " FROM pla_periodolaboral INNER JOIN mae_categoria ON pla_periodolaboral.idcat = mae_categoria.id " & nSQLWhere _
'            + vbCr + " GROUP BY pla_periodolaboral.idemp " _
'            + vbCr + " ORDER BY pla_periodolaboral.idemp, Last(pla_periodolaboral.fchini), Last(pla_periodolaboral.fchfin)) AS periodo " _
'            + vbCr + " ON emp.id = periodo.idemp;"
'    Else
'        nSQL = "SELECT * FROM " _
'            + vbCr + " (SELECT pla_empleados.*, mae_dociden.abrev AS docabrev, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat]  & ' ' &  [pla_empleados].[nom]  AS nombres, mae_sexo.abrev AS sexo " _
'            + vbCr + " FROM mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex " _
'            + vbCr + " ORDER BY [pla_empleados].[nom] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] ) AS emp " _
'            + vbCr + " INNER JOIN " _
'            + vbCr + " (SELECT pla_periodolaboral.idemp, Last(mae_categoria.descripcion) AS categoria, Last(pla_periodolaboral.fchini) AS ingreso, Last(pla_periodolaboral.fchfin) AS cese, IIf([cese] Is Not Null,'De Baja','Activo') AS estado " _
'            + vbCr + " FROM pla_periodolaboral INNER JOIN mae_categoria ON pla_periodolaboral.idcat = mae_categoria.id " _
'            + vbCr + " WHERE pla_periodolaboral.fchfin IS NULL  " _
'            + vbCr + " GROUP BY pla_periodolaboral.idemp " _
'            + vbCr + " ORDER BY pla_periodolaboral.idemp, Last(pla_periodolaboral.fchini), Last(pla_periodolaboral.fchfin)) AS periodo " _
'            + vbCr + " ON emp.id = periodo.idemp;"
'    End If

    If fSoloActivos = True Then nSQLWhere = " WHERE pla_empleados.fchcese is null "
    nSQL = "SELECT pla_empleados.*, mae_dociden.abrev AS docabrev, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_sexo.abrev AS sexo, mae_categoria.descripcion AS categoria, pla_empleados.fching AS ingreso, pla_empleados.fchcese AS cese, IIf([pla_empleados].[fchcese] Is  Null,'Activo','De Baja') AS estado " _
        + vbCr + " FROM (mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex) LEFT JOIN mae_categoria ON pla_empleados.idcat = mae_categoria.id " _
        + vbCr + nSQLWhere & " ORDER BY [pla_empleados].[nom] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat];"




    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), "Buscando Personal", "nombres", "nombres", Principio

End Function


'********** INICIO PLANILLA DE PAGO ********************
Public Sub pCagarListaPersonal(Rst As ADODB.Recordset, mIdProceso As Long, mIdCategoria, anno, mes)

    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLTmp As String
    Dim fAsigFamiliar&, fEstaEnPlanilla&, mIdCatEmp&
    Dim mTotalSegundosMes As Long

    mTotalSegundosMes = HallaDiasMes(CDate("01/" & mes & "/" & anno))
    mTotalSegundosMes = mTotalSegundosMes * 8 * 60 * 60

    '******************************************************************************************************
    '--esta consulta es la union de la consulta de empleado + la categoria + la boleta
    nSQL = "SELECT  *, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora.totseg) AS totseg FROM mae_tipohora INNER JOIN (pla_marcacion INNER JOIN pla_marcacionhora ON pla_marcacion.id = pla_marcacionhora.idmarca) ON mae_tipohora.id = pla_marcacionhora.idhora WHERE (((pla_marcacionhora.idemp)=emp.idemp) AND ((Year([pla_marcacion].[dia]))=" & anno & ") AND ((Month([pla_marcacion].[dia]))=" & mes & ") AND ((mae_tipohora.hortrabajo)=-1)) GROUP BY pla_marcacionhora.idemp) AS totseg " _
        + vbCr + " FROM (SELECT * FROM " _
        + vbCr + " (SELECT pla_empleados.id AS idemp, mae_dociden.abrev AS docabrev, pla_empleados.numdoc AS docemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pla_empleados.fchnac, mae_sexo.abrev AS sexo, pla_empleados.idcargo, mae_cargo.descripcion AS cargo " _
        + vbCr + " FROM mae_sexo RIGHT JOIN (mae_cargo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_cargo.id = pla_empleados.idcargo) ON mae_sexo.id = pla_empleados.idsex " _
        + vbCr + " WHERE (((pla_empleados.idbolpag)=" & mIdProceso & "))" _
        + vbCr + " ORDER BY [pla_empleados].[nom] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat]) AS emp " _
        + vbCr + " INNER JOIN " _
        + vbCr + " (SELECT pla_periodolaboral.idemp AS idemp1, mae_categoria.descripcion AS categoria, mae_categoria.nomcor AS catabrev, Last(pla_periodolaboral.fchini) AS ingreso " _
        + vbCr + " FROM mae_categoria INNER JOIN pla_periodolaboral ON mae_categoria.id = pla_periodolaboral.idcat " _
        + vbCr + " Where (((pla_periodolaboral.fchfin) Is Null)) AND pla_periodolaboral.idcat=" & mIdCategoria & " " _
        + vbCr + " GROUP BY pla_periodolaboral.idemp, mae_categoria.descripcion, mae_categoria.nomcor " _
        + vbCr + " ORDER BY pla_periodolaboral.idemp, Last(pla_periodolaboral.fchini), Last(pla_periodolaboral.fchfin)) AS periodo " _
        + vbCr + " ON emp.idemp = periodo.idemp1) AS emp " _
        + vbCr + " LEFT JOIN " _
        + vbCr + " (SELECT pla_boleta.id AS idbol,  pla_boleta.idemp as idemp1,   pla_boleta.numreg, pla_boleta.idmon, pla_boleta.numser, pla_boleta.numdoc, pla_boleta.fchdoc, pla_boleta.fchpago, mae_moneda.simbolo, pla_boleta.impingr, pla_boleta.impapor, pla_boleta.impdesc, pla_boleta.imptot " _
        + vbCr + " FROM pla_proceso RIGHT JOIN (mae_moneda RIGHT JOIN pla_boleta ON mae_moneda.id = pla_boleta.idmon) ON pla_proceso.id = pla_boleta.idproc " _
        + vbCr + " WHERE pla_boleta.ano= " & anno & " and pla_boleta.idmes= " & mes & " and pla_boleta.idproc= " & mIdProceso & ") AS boleta ON emp.idemp = boleta.idemp1 " _
        + vbCr + " ORDER BY emp.nombres"

    '--cargar los datos
    RST_Busq Rst, nSQL, xCon

End Sub

'********** FIN PLANILLA DE PAGO ********************


Public Sub pCargarPersonal(frm As Form, Index As Integer)
    Dim xRs As New ADODB.Recordset
    pBuscarPersonal xRs, True
    If xRs.State = 1 Then
        frm.txt_cb(Index) = xRs.Fields("id") & "" '--TEXTO A MOSTRAR
        frm.lbl_cb(Index).Caption = xRs.Fields("nombres") & "" '--NOMBRE
        frm.lbl_cod(Index).Caption = xRs.Fields("id") & "" '--CODIGO
        frm.lbl_cb(Index).ToolTipText = xRs.Fields("nombres") & "" '--NOMBRE
    End If
    Set xRs = Nothing
End Sub







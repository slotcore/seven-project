Attribute VB_Name = "Varios"
''Private Function FunctionExist(FunctionName As String) As Boolean
''    'check if functionname is defined or not
''
''    Dim strTemp As String
''
''    On Error Resume Next
''    strTemp = vFunctions(UCase(FunctionName))
''    FunctionExist = (Err = 0)
''    On Error GoTo 0
''    Err.Clear
''End Function

Public Enum e_CategoriaConcepto
    e_Remuneracion = 1
    e_Aportacion = 2
    e_Descuento = 3
End Enum

'*****************************************************************************************************
'*****************************************************************************************************
'*****************************************************************************************************
'--para editor de formula

Public Sub eliminar_nodo(arb_obj As TreeView, Optional nd As Node, Optional key As String, Optional texto As String)
  
    If Not Nothing Is nd Then
        If MsgBox("Seguro desea eliminar " & nd.Text, vbQuestion + vbYesNo) = vbYes Then
            arb_obj.Nodes.Remove nd.key
        End If
        Exit Sub
    End If
    
    If Not IsMissing(key) Then
       If Trim(texto) <> "" Then
          If MsgBox(Trim("Seguro desea eliminar " & texto), vbQuestion + vbYesNo) = vbNo Then Exit Sub
       End If
       arb_obj.Nodes.Remove key
       Exit Sub
    End If

End Sub
'----------Agregar un node al treeview
Public Sub agregar_nodo(arb_obj As TreeView, _
                        relative As String, _
                        key As String, _
                        Text As String, _
                        Optional imagenA As String, _
                        Optional imagenB As String)
                        
    'On Error Resume Next
    Dim nds As Node
    'Set nds = arb_obj.Nodes.Add()
    If imagenA <> "" And imagenB = "" Then
       Set nds = arb_obj.Nodes.Add(relative, tvwChild, key, Text, imagenA)
    ElseIf imagenA <> "" And imagenB <> "" Then
       Set nds = arb_obj.Nodes.Add(relative, tvwChild, key, Text, imagenA, imagenB)
    ElseIf imagenA = "" And imagenB <> "" Then
       imagenA = imagenB
       Set nds = arb_obj.Nodes.Add(relative, tvwChild, key, Text, imagenA)
    Else
       Set nds = arb_obj.Nodes.Add(relative, tvwChild, key, Text)
    End If
    Err.Clear
End Sub



Public Sub SacarRepetidasArray(X() As String, cant As Long)
    Dim i As Integer
    Dim J As Integer
    Dim Y As Variant
    Dim z As Variant
    ReDim Y(cant)
    ReDim z(cant)
    Y = X
    For i = 0 To cant
        For J = i + 1 To cant - 1
            If Y(i) = Y(J) Then
               Y(J) = -1
            End If
        Next
    Next
    
    
    J = 0
    
    For i = 0 To cant - 1
        If Y(i) <> "-1" Then
           ReDim Preserve X(J)
           X(J) = Y(i)
           J = J + 1
        End If
    Next
    cant = J
End Sub


'********** FIN MARCACION DE ASISTENCIA ********************

'*******************************************************************************************************************************
'*************** INICIO CONSULTAS ********************
Public Function pBuscarPersonal(RstTmp As ADODB.Recordset, _
                       Optional fSoloActivos As Boolean = True, _
                       Optional fSoloPlanilla As Boolean = False) As ADODB.Recordset
                       
    Dim nSQL As String
    Dim xCampos(6, 4) As String
    Dim nSQLWhere As String
    
    xCampos(0, 0) = "TipDoc":               xCampos(0, 1) = "docabrev":   xCampos(0, 2) = "700":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Doc.":              xCampos(1, 1) = "numdoc":     xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombres":    xCampos(2, 2) = "3200":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch. Ing":             xCampos(3, 1) = "fching":     xCampos(3, 2) = "1000":     xCampos(3, 3) = "C"
    xCampos(4, 0) = "Categoría":            xCampos(4, 1) = "categoria":  xCampos(4, 2) = "1000":    xCampos(4, 3) = "C"
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
    If fSoloPlanilla = True Then nSQLWhere = IIf(nSQLWhere = "", " WHERE ", nSQLWhere & " and ") & " pla_empleados.aplanilla = -1 "
    
    nSQL = "SELECT pla_empleados.*, mae_dociden.abrev AS docabrev, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_sexo.abrev AS sexo, mae_categoria.descripcion AS categoria, pla_empleados.fching AS ingreso, pla_empleados.fchcese AS cese, IIf([pla_empleados].[fchcese] Is  Null,'Activo','De Baja') AS estado " _
        + vbCr + " FROM (mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex) LEFT JOIN mae_categoria ON pla_empleados.idcat = mae_categoria.id " _
        + vbCr + nSQLWhere & " ORDER BY [pla_empleados].[nom] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat];"




    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), "Buscando Personal", "nombres", "nombres", Principio

End Function

'*************** FIN CONSULTAS ********************


'********** INICIO PLANILLA DE PAGO ********************
Public Sub pCagarListaPersonal(Rst As ADODB.Recordset, mIdProceso As Long, mIdCategoria, anno, mes)

    Dim nSQL As String
    '******************************************************************************************************
    '--esta consulta es la union de la consulta de empleado + la categoria + la boleta
    nSQL = "SELECT  emp.*, boleta.*, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora.totseg) AS totseg FROM mae_tipohora INNER JOIN (pla_marcacion INNER JOIN pla_marcacionhora ON pla_marcacion.id = pla_marcacionhora.idmarca) ON mae_tipohora.id = pla_marcacionhora.idhora WHERE (((pla_marcacionhora.idemp)=emp.idemp) AND ((Year([pla_marcacion].[dia]))=" & anno & ") AND ((Month([pla_marcacion].[dia]))=" & mes & ") AND ((mae_tipohora.hortrabajo)=-1)) GROUP BY pla_marcacionhora.idemp) AS totseg " _
        + vbCr + " FROM " _
        + vbCr + " (SELECT pla_empleados.id AS idemp, mae_dociden.abrev AS docabrev, pla_empleados.numdoc AS docemp, pla_empleados.nombre AS nombres, pla_empleados.fchnac, mae_sexo.abrev AS sexo, pla_empleados.idcargo, mae_cargo.descripcion AS cargo, mae_categoria.descripcion AS categoria, mae_categoria.nomcor AS catabrev, pla_empleados.fching AS ingreso,mae_regimenpen.descripcion AS regpension " _
        + vbCr + " FROM mae_regimenpen RIGHT JOIN ((mae_sexo RIGHT JOIN ((mae_cargo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_cargo.id = pla_empleados.idcargo) LEFT JOIN mae_categoria ON pla_empleados.idcat = mae_categoria.id) ON mae_sexo.id = pla_empleados.idsex) LEFT JOIN pla_categoria1 ON pla_empleados.id = pla_categoria1.idemp) ON mae_regimenpen.id = pla_categoria1.idregpen " _
        + vbCr + " WHERE pla_empleados.idcat = " & mIdCategoria & " and pla_empleados.fchcese is null and pla_empleados.aplanilla = -1 and (((pla_empleados.idbolpag)=" & mIdProceso & ")) " _
        + vbCr + " ORDER BY pla_empleados.nombre ) AS emp " _
        + vbCr + " LEFT JOIN " _
        + vbCr + " (SELECT pla_boleta.id AS idbol,  pla_boleta.idemp as idemp1, pla_boleta.numreg, pla_boleta.idmon, pla_boleta.numser, pla_boleta.numdoc, pla_boleta.fchdoc, pla_boleta.fchpago, mae_moneda.simbolo, pla_boleta.impingr, pla_boleta.impapor, pla_boleta.impdesc, pla_boleta.imptot " _
        + vbCr + " FROM pla_proceso RIGHT JOIN (mae_moneda RIGHT JOIN pla_boleta ON mae_moneda.id = pla_boleta.idmon) ON pla_proceso.id = pla_boleta.idproc " _
        + vbCr + " WHERE pla_boleta.ano= " & anno & " and pla_boleta.idmes= " & mes & " and pla_boleta.idproc= " & mIdProceso & ") AS boleta ON emp.idemp = boleta.idemp1 " _
        + vbCr + " ORDER BY emp.nombres "

    '--cargar los datos
    RST_Busq Rst, nSQL, xCon
    
End Sub



Public Sub pConceptoSueldoAsignadoEmp(Rst As ADODB.Recordset, _
                                 mIdEmp As Long, mIdProceso As Long, _
                                 anno, mes, eTipo As e_CategoriaConcepto)

    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLTmp As String
    Dim fAsigFamiliar&, fEstaEnPlanilla&, mIdCatEmp&, IdESSALUDVIDA&
    Dim fEstaDeVacaciones As Boolean
    
    '******************************************************************************************************
    nSQLTmp = "SELECT pla_empleados.*,  mae_categoria.descripcion AS categoria, pla_periodolaboral.fchini AS fchingreso, pla_periodolaboral.fchfin AS fchsalida, pla_vacaciones.id AS idvac, pla_vacaciones.* " _
        + vbCr + " FROM (pla_empleados LEFT JOIN (mae_categoria RIGHT JOIN pla_periodolaboral ON mae_categoria.id = pla_periodolaboral.idcat) ON pla_empleados.id = pla_periodolaboral.idemp) LEFT JOIN pla_vacaciones ON pla_empleados.id = pla_vacaciones.idemp " _
        + vbCr + " WHERE (((pla_empleados.id)=" & mIdEmp & ") AND ((pla_periodolaboral.fchfin) Is Null));"

   
    RST_Busq RstTmp, nSQLTmp, xCon
    If RstTmp.RecordCount = 0 Then
        
        Exit Sub
    End If
    
    '--------------
    fAsigFamiliar = NulosN(RstTmp("asigfam"))
    fEstaEnPlanilla = NulosN(RstTmp("aplanilla"))
    mIdCatEmp = NulosN(RstTmp("idcat")) '--codigo de la categoria a la que pertenece
    IdESSALUDVIDA = NulosN(RstTmp("indessalud"))
    '--------------
    '******************************************************************************************************
    '--ver si esta de vacaciones para considerarlo horas de trabajo como optimo
    fEstaDeVacaciones = False
    If NulosN(RstTmp("idvac")) <> 0 Then
        If RstTmp("annopago") = ano And RstTmp("mespago") = mes Then fEstaDeVacaciones = True
    End If
    
    Set RstTmp = Nothing
    '******************************************************************************************************
    '-----
    
    If fEstaEnPlanilla = 0 Then '--no esta en planilla
        If mIdProceso <> 4 Then '--semanales,diario
            '--semana 1, semana 2, semana 3
            nSQL = "SELECT " & mIdEmp & " AS mIdEmp, pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.aplanilla, pla_concepto.formula, pla_conceptoemp.imptot & '' as imptot ,pla_concepto.nomcorto " _
                + vbCr + " FROM (pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN (pla_concepto INNER JOIN pla_conceptoemp ON pla_concepto.id = pla_conceptoemp.idcpto) ON pla_conceptotipo.id = pla_concepto.idtipo " _
                + vbCr + " WHERE (((pla_conceptoemp.anno)=" & anno & ") AND ((pla_conceptoemp.idmes)=" & mes & ") AND ((pla_conceptotipo.idcat)=" & eTipo & ") AND ((pla_conceptoemp.idproc)=" & mIdProceso & ") AND ((pla_conceptoemp.idemp)=" & mIdEmp & " ));"

        Else
            '--semana 4
            '--170::adelanto de sueldo
            '--175::total bonificacion
            nSQL = "SELECT " & mIdEmp & " AS mIdEmp,pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, pla_conceptoemp.imptot & '' as imptot ,pla_concepto.nomcorto " _
                + vbCr + " FROM ((pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) INNER JOIN pla_conceptoemp ON pla_concepto.id = pla_conceptoemp.idcpto " _
                + vbCr + " WHERE (((pla_conceptoemp.anno)=" & anno & ") AND ((pla_conceptoemp.idmes)=" & mes & ") AND ((pla_conceptotipo.idcat)= " & eTipo & " ) AND ((pla_conceptoemp.idproc)=" & mIdProceso & ") AND ((pla_conceptoemp.idemp)=" & mIdEmp & ") AND ((pla_concepto.id)<>170)); " _
                + vbCr + " UNION " _
                + vbCr + " SELECT " & mIdEmp & " AS mIdEmp, pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, pla_empleados.bono AS imptot,pla_concepto.nomcorto " _
                + vbCr + " FROM pla_empleados, pla_conceptocat INNER JOIN (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat " _
                + vbCr + " WHERE (((pla_concepto.id)=175) AND ((pla_empleados.id)= " & mIdEmp & " )) AND " & eTipo & " = 1 ; "

        End If
    Else
            '--todos los procesos estaran afecto a aportes del trabajador y aportes de empleador
            '--si no desea que se considere desactivar de conceptos
            Select Case eTipo
                Case 1 '--remuneraciones
                    '--todos los conceptos asignados (menos adelanto de sueldo id=170)
                    '--conceptos asignados
                    nSQL = "SELECT " & mIdEmp & " AS mIdEmp,pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, pla_conceptoemp.imptot & '' as imptot ,pla_concepto.nomcorto " _
                        + vbCr + " FROM ((pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) INNER JOIN pla_conceptoemp ON pla_concepto.id = pla_conceptoemp.idcpto " _
                        + vbCr + " WHERE (((pla_conceptoemp.anno)=" & anno & ") AND ((pla_conceptoemp.idmes)= " & mes & ") AND ((pla_conceptoemp.idproc)=" & mIdProceso & ") AND ((pla_conceptoemp.idemp)=" & mIdEmp & ") AND ((pla_concepto.id)<>170))  AND ((pla_conceptotipo.idcat)= 1 ) ; "
                    
                    '--todos los conceptos de ingresos menos los asignados
                    nSQL = nSQL + vbCr + " UNION " _
                        + vbCr + " SELECT " & mIdEmp & " AS mIdEmp, pla_concepto.id & '-0' AS Codigo, '0' AS Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, '' AS imptot, pla_concepto.nomcorto " _
                        + vbCr + " FROM (pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
                        + vbCr + " WHERE (((pla_conceptotipo.idcat)=1) AND ((pla_concepto.id) Not In (173)) AND ((pla_concepto.activo)=-1) AND ((pla_concepto.id)<>170 And (pla_concepto.id) Not In (SELECT pla_conceptoemp.idcpto " _
                        + vbCr + " FROM pla_conceptoemp " _
                        + vbCr + " WHERE (((pla_conceptoemp.anno)=" & anno & ") AND ((pla_conceptoemp.idmes)= " & mes & ") AND ((pla_conceptoemp.idemp)=" & mIdEmp & ") AND ((pla_conceptoemp.idproc)=" & mIdProceso & ")))));"
                    '--union concepto de bonificacion(171) que se encuentra en tabla pla_empleados campo bono
                    '--menos total remuneracion (id=174)
                    nSQL = nSQL + vbCr + " UNION " _
                        + vbCr + " SELECT " & mIdEmp & " AS mIdEmp,pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, pla_empleados.bono & '' AS imptot,pla_concepto.nomcorto " _
                        + vbCr + " FROM pla_empleados, pla_conceptocat INNER JOIN (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat " _
                        + vbCr + " WHERE (((pla_concepto.id)=171) AND ((pla_empleados.id)= " & mIdEmp & "));"
                
                Case 2 '--aportaciones
                    '--todas las aportaciones asignados UNION todas las aportaciones del empleador
                    nSQL = "SELECT " & mIdEmp & " AS mIdEmp, pla_concepto.id & '-0' AS Codigo, '0' AS Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, '' AS imptot, pla_concepto.nomcorto " _
                        + vbCr + " FROM (pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
                        + vbCr + " WHERE (((pla_conceptotipo.idcat)=2) AND ((pla_concepto.id) Not In (174)) AND ((pla_conceptotipo.id)=10) AND ((pla_concepto.activo)=-1));"

                Case 3 '--descuentos
                    '--todos los descuentos asignados UNION todas los descuentos del trabajador UNION aportes de empleador
                    nSQL = "SELECT " & mIdEmp & " AS mIdEmp,pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula,pla_concepto.aplanilla, pla_conceptoemp.imptot & '' as imptot ,pla_concepto.nomcorto " _
                        + vbCr + " FROM ((pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) INNER JOIN pla_conceptoemp ON pla_concepto.id = pla_conceptoemp.idcpto " _
                        + vbCr + " WHERE (((pla_conceptoemp.anno)=" & anno & ") AND ((pla_conceptoemp.idmes)=" & mes & ") AND (pla_conceptotipo.idcat= 3 OR (pla_conceptotipo.idcat= 2 AND pla_conceptotipo.id=9) ) AND ((pla_conceptoemp.idproc)=" & mIdProceso & ") AND ((pla_conceptoemp.idemp)=" & mIdEmp & ")); "

                    nSQL = nSQL + vbCr + " UNION " _
                        + vbCr + " SELECT " & mIdEmp & " AS mIdEmp, pla_concepto.id & '-0' AS Codigo, '0' AS Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, '' AS imptot, pla_concepto.nomcorto " _
                        + vbCr + " FROM (pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
                        + vbCr + " WHERE (((pla_conceptotipo.idcat) = 3) And ((pla_concepto.activo) = -1)) Or (((pla_conceptotipo.idcat) = 2) And ((pla_conceptotipo.id) = 9) And ((pla_concepto.activo) = -1)) " _
                        + vbCr + " AND pla_concepto.id NOT IN " _
                        + vbCr + " (SELECT pla_concepto.id AS idcpto " _
                        + vbCr + " FROM (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) INNER JOIN pla_conceptoemp ON pla_concepto.id = pla_conceptoemp.idcpto " _
                        + vbCr + " WHERE (((pla_conceptotipo.idcat)=3) AND ((pla_conceptoemp.anno)=" & anno & ") AND ((pla_conceptoemp.idmes)=" & mes & ") AND ((pla_conceptoemp.idemp)=" & mIdEmp & ") AND ((pla_conceptoemp.idproc)=" & mIdProceso & ")) OR (((pla_conceptotipo.idcat)=2) AND ((pla_conceptotipo.id)=9) AND ((pla_conceptoemp.anno)=" & anno & ") AND ((pla_conceptoemp.idmes)=" & mes & ") AND ((pla_conceptoemp.idemp)=" & mIdEmp & ") AND ((pla_conceptoemp.idproc)=" & mIdProceso & "));); "
                    
                    '--descuento por adelantos (id=155)
                    nSQL = nSQL + vbCr + " UNION " _
                        + vbCr + " SELECT " & mIdEmp & " AS mIdEmp, pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_conceptotipo.idcat,pla_concepto.id AS idcpto, pla_concepto.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, null as formula, pla_concepto.aplanilla, Sum(pla_boletadet.imptot) & '' AS imptot,pla_concepto.nomcorto " _
                        + vbCr + " FROM pla_proceso INNER JOIN (pla_boleta INNER JOIN pla_boletadet ON pla_boleta.id = pla_boletadet.idbol) ON pla_proceso.id = pla_boleta.idproc, pla_conceptocat INNER JOIN (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat " _
                        + vbCr + " WHERE (((pla_boleta.idemp)=" & mIdEmp & ") AND ((pla_boleta.ano)=" & anno & ") AND ((pla_boleta.idmes)=" & mes & ") AND ((pla_concepto.id)=155) AND ((pla_proceso.identificador) In (1,2,3))) " _
                        + vbCr + " GROUP BY pla_conceptotipo.idcat,pla_concepto.id, pla_concepto.descripcion, pla_concepto.descripcion, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla,pla_concepto.nomcorto;"


            End Select
            
        End If
'    End If
    

    '--cargar los datos
    RST_Busq RstTmp, nSQL, xCon
    '--si no tiene campos el recoordset => definir recordset temporal
    If Rst.State = 0 Then DEFINIR_RST_TMP Rst, RstTmp
    
    '--cargar los datos al recordset temporal
    CARGAR_RST_TMP Rst, RstTmp
    
    Set RstTmp = Nothing
    '******************************************************************************************************
    '--eliminando conceptos del temporal
    If IdESSALUDVIDA <> 1 Then '--si no tiene essalud + vida => eliminarlo
        RstRegistroEliminar Rst, "Codigo", "147-0", False '--eliminar EssaludVida
    End If
    If mIdProceso = 4 Then
        RstRegistroEliminar Rst, "Codigo", "170-0", False '--eliminar adelanto de sueldo
        
''
    ElseIf mIdProceso = 2 Then
        RstRegistroEliminar Rst, "Codigo", "21-0", False '--eliminar remueracion o jornal basico
''        RstRegistroEliminar Rst, "Codigo", "171-0", False '--eliminar bonificacion
''        RstRegistroEliminar Rst, "Codigo", "170-0", False '--eliminar adelanto de sueldo
''        RstRegistroEliminar Rst, "Codigo", "155-0", False '--eliminar descuento por adelanto de sueldo
        
    End If
    
    '******************************************************************************************************
    '--cuando esta en planilla eliminar ciertos conceptos segun condiciones
    If fEstaEnPlanilla = -1 Then
        '******************************************************************************************************
        If eTipo = e_Remuneracion Then
            '******************************************************************************************************
            If fAsigFamiliar = 0 Then RstRegistroEliminar Rst, "Codigo", "26-0", False                 '--eliminar asignacion familiar
            '******************************************************************************************************
            If fEstaDeVacaciones = True Then
                RstRegistroEliminar Rst, "Codigo", "21-0", False '--eliminar remuneracion o haber basico
            Else
                RstRegistroEliminar Rst, "Codigo", "18-0", False '--eliminar remuneracion vacacional
            End If
        Else
            '******************************************************************************************************
            '--Obtener el regimen pensionario y eliminar aquellos conceptos que no esten asociados a la categoria
            '-- regimen de pension
            nSQLTmp = "SELECT pla_categoria1.idemp, 1 AS idcat, pla_categoria1.idregpen, mae_regimenpen.descripcion AS regimen, pla_conceptoregpen.idcpto, pla_concepto.descripcion AS concepto " _
                + vbCr + " FROM mae_regimenpen RIGHT JOIN (pla_concepto RIGHT JOIN (pla_categoria1 LEFT JOIN pla_conceptoregpen ON pla_categoria1.idregpen = pla_conceptoregpen.idregpen) ON pla_concepto.id = pla_conceptoregpen.idcpto) ON mae_regimenpen.id = pla_conceptoregpen.idregpen " _
                + vbCr + " WHERE (((pla_categoria1.idemp)=" & mIdEmp & ")) AND 1 = " & mIdCatEmp & "; " _
                + vbCr + " UNION " _
                + vbCr + " SELECT pla_categoria2.idemp, 2 AS idcat, pla_categoria2.idregpen, mae_regimenpen.descripcion AS regimen, pla_conceptoregpen.idcpto, pla_concepto.descripcion AS concepto " _
                + vbCr + " FROM mae_regimenpen RIGHT JOIN (pla_concepto RIGHT JOIN (pla_conceptoregpen RIGHT JOIN pla_categoria2 ON pla_conceptoregpen.idregpen = pla_categoria2.idregpen) ON pla_concepto.id = pla_conceptoregpen.idcpto) ON mae_regimenpen.id = pla_categoria2.idregpen " _
                + vbCr + " WHERE (((pla_categoria2.idemp)=" & mIdEmp & ")) AND 2 = " & mIdCatEmp & ";"
            
            RST_Busq RstTmp, nSQLTmp, xCon
            
            If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
            '--generar la cadena de conceptos que no se van eliminar del regimen pensionario
            nSQLTmp = ""
            Do While Not RstTmp.EOF
                nSQLTmp = nSQLTmp & NulosN(RstTmp("idcpto")) & ","
                RstTmp.MoveNext
            Loop
            If nSQLTmp <> "" Then nSQLTmp = " WHERE pla_conceptoregpen.idcpto NOT IN (" + Left(nSQLTmp, Len(nSQLTmp) - 1) + ") "
            
            nSQLTmp = "SELECT pla_conceptoregpen.idcpto, pla_concepto.descripcion FROM pla_concepto INNER JOIN pla_conceptoregpen ON pla_concepto.id = pla_conceptoregpen.idcpto " & nSQLTmp & " GROUP BY pla_conceptoregpen.idcpto, pla_concepto.descripcion;"
            
            RST_Busq RstTmp, nSQLTmp, xCon
            If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
            Do While Not RstTmp.EOF
                RstRegistroEliminar Rst, "Codigo", NulosN(RstTmp("idcpto")) & "-0", False '--eliminar
                RstTmp.MoveNext
            Loop
            '******************************************************************************************************
        End If
        
    End If
    '--
    
End Sub

Public Sub pConceptoCagarDetalleFormula(Rst As ADODB.Recordset, _
                                    mIdEmp As Long, mIdProceso As Long, _
                                    anno, mes)
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQLTmp As String
    Dim nSQL As String
    Dim fAsigFamiliar&, fEstaEnPlanilla&, mIdCatEmp&
    Dim fEstaDeVacaciones As Boolean
    Dim mIdCargo&, mIdRegPen&
    Dim mTotalSegundosMes As Long
    Dim dFchIngreso As String
    '--------------------------------------------------------------------
    mTotalSegundosMes = HallaDiasMes(CDate("01/" & mes & "/" & anno))
    mTotalSegundosMes = mTotalSegundosMes * 8 * 60 * 60
    '--------------------------------------------------------------------
    nSQLTmp = "SELECT pla_empleados.*,  mae_categoria.descripcion AS categoria, pla_periodolaboral.fchini AS fchingreso, pla_periodolaboral.fchfin AS fchsalida, pla_vacaciones.id AS idvac, pla_vacaciones.* " _
        + vbCr + " FROM (pla_empleados LEFT JOIN (mae_categoria RIGHT JOIN pla_periodolaboral ON mae_categoria.id = pla_periodolaboral.idcat) ON pla_empleados.id = pla_periodolaboral.idemp) LEFT JOIN pla_vacaciones ON pla_empleados.id = pla_vacaciones.idemp " _
        + vbCr + " WHERE (((pla_empleados.id)=" & mIdEmp & ") AND ((pla_periodolaboral.fchfin) Is Null));"
   
    RST_Busq RstTmp, nSQLTmp, xCon
    
    '--------------
    fAsigFamiliar = NulosN(RstTmp("asigfam"))
    fEstaEnPlanilla = NulosN(RstTmp("aplanilla"))
    mIdCatEmp = NulosN(RstTmp("idcat")) '--codigo de la categoria a la que pertenece
    mIdCargo = NulosN(RstTmp("idcargo"))
    dFchIngreso = NulosC(RstTmp("fchingreso"))
    
    '--ver si esta de vacaciones para considerarlo horas de trabajo como optimo
    fEstaDeVacaciones = False
    If NulosN(RstTmp("idvac")) <> 0 Then
        If RstTmp("annopago") = ano And RstTmp("mespago") = mes Then fEstaDeVacaciones = True
    End If
    
    Set RstTmp = Nothing
    '******************************************************************************************************
    '--obtener el regimen pensionario
    nSQLTmp = "SELECT 1 AS idcat, pla_categoria1.idregpen, mae_regimenpen.descripcion AS regimen " _
        + vbCr + " FROM mae_regimenpen RIGHT JOIN pla_categoria1 ON mae_regimenpen.id = pla_categoria1.idregpen " _
        + vbCr + " WHERE (((pla_categoria1.idemp)=" & mIdEmp & ") AND ((1)=" & mIdCatEmp & ")); " _
        + vbCr + " UNION " _
        + vbCr + " SELECT 2 AS idcat, pla_categoria2.idregpen, mae_regimenpen.descripcion AS regimen " _
        + vbCr + " FROM mae_regimenpen RIGHT JOIN pla_categoria2 ON mae_regimenpen.id = pla_categoria2.idregpen " _
        + vbCr + " WHERE (((pla_categoria2.idemp)=" & mIdEmp & ") AND ((2)= " & mIdCatEmp & " )); "

    RST_Busq RstTmp, nSQLTmp, xCon
    If RstTmp.RecordCount <> 0 Then
        mIdRegPen = NulosN(RstTmp("idregpen"))
    End If
    Set RstTmp = Nothing
    '******************************************************************************************************
    
    '******************************************************************************************************
    '--conceptos que tiene valores asignados tabla:: pla_conceptovarios
    '--ejm. essalud, %SNP, %AFP, ETC
    '--OBS: si hubiera otros origenes
    '--valores fijos  UNION regimen pensionario UNION cargo UNION meses
    nSQL = "SELECT pla_conceptovarios.id & '-2' as Codigo,'2' as Origen, pla_conceptovarios.id AS idcpto, '' AS categoria, pla_conceptovarios.descripcion AS concepto, pla_conceptovarios.variable, '' AS formula, 0 AS aplanilla, pla_conceptovariosdet.imptot " _
        + vbCr + " FROM pla_conceptovarios INNER JOIN pla_conceptovariosdet ON pla_conceptovarios.id = pla_conceptovariosdet.idcptov " _
        + vbCr + " WHERE (((pla_conceptovarios.esfijo)=0) AND ((pla_conceptovariosdet.idref)=" & mIdRegPen & ") AND ((pla_conceptovariosdet.anno)=" & anno & ") AND ((pla_conceptovariosdet.idmes)=" & mes & " ) AND ((pla_conceptovarios.entgen)=1)); " _
        + vbCr + " UNION " _
        + vbCr + " SELECT pla_conceptovarios.id & '-2' as Codigo,'2' as Origen, pla_conceptovarios.id AS idcpto, '' AS categoria, pla_conceptovarios.descripcion AS concepto, pla_conceptovarios.variable, '' AS formula, 0 AS aplanilla, pla_conceptovarios.formula AS imptot " _
        + vbCr + " FROM pla_conceptovarios " _
        + vbCr + " WHERE (((pla_conceptovarios.esfijo)=-1)); " _
        + vbCr + " UNION " _
        + vbCr + " SELECT pla_conceptovarios.id & '-2' as Codigo,'2' as Origen, pla_conceptovarios.id AS idcpto, '' AS categoria, pla_conceptovarios.descripcion AS concepto, pla_conceptovarios.variable, '' AS formula, 0 AS aplanilla, pla_conceptovariosdet.imptot " _
        + vbCr + " FROM pla_conceptovarios INNER JOIN pla_conceptovariosdet ON pla_conceptovarios.id = pla_conceptovariosdet.idcptov " _
        + vbCr + " WHERE (((pla_conceptovarios.esfijo)=0) AND ((pla_conceptovariosdet.idref)=" & mIdCargo & ") AND ((pla_conceptovariosdet.anno)=" & anno & ") AND ((pla_conceptovariosdet.idmes)=" & mes & ") AND ((pla_conceptovarios.entgen)=2)); " _
        + vbCr + " UNION " _
        + vbCr + " SELECT pla_conceptovarios.id & '-2' as Codigo,'2' as Origen, pla_conceptovarios.id AS idcpto, '' AS categoria, pla_conceptovarios.descripcion AS concepto, pla_conceptovarios.variable, '' AS formula, 0 AS aplanilla, pla_conceptovariosdet.imptot " _
        + vbCr + " FROM pla_conceptovarios INNER JOIN pla_conceptovariosdet ON pla_conceptovarios.id = pla_conceptovariosdet.idcptov " _
        + vbCr + " WHERE (((pla_conceptovarios.esfijo)=0) AND ((pla_conceptovariosdet.idref)=" & mes & ") AND ((pla_conceptovariosdet.anno)=" & anno & ") AND ((pla_conceptovariosdet.idmes)=" & mes & ") AND ((pla_conceptovarios.entgen)=3)); "

    '--no se considera fondo de pensiones, snp,
    '--basico(177),bonificacion(171),total descuento(172),remuneracion(173),aportes(174)
    '--acumulado de gratificacion, vacaciones (PENDIENTE)
    nSQL = nSQL + vbCr + " UNION " _
        + vbCr + " SELECT pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, NULL AS imptot " _
        + vbCr + " FROM pla_conceptocat INNER JOIN (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat " _
        + vbCr + " WHERE (((pla_concepto.id) Not In (177,171,172,173,174) And (pla_concepto.id) Not In (SELECT pla_conceptoregpen.idcpto FROM pla_conceptoregpen))) and pla_concepto.activo = -1; "

    '--horas de asistencia
    '--obs: idhora = 15:: horas trabajadas
    nSQL = nSQL + vbCr + " UNION " _
        + vbCr + " SELECT mae_tipohora.id & '-1' AS Codigo, '1' AS Origen, mae_tipohora.id AS idcpto, '' AS categoria, mae_tipohora.descripcion AS concepto, mae_tipohora.variable, '' AS formula, 0 AS aplanilla, pla_marcaresumhor.totseg AS imptot " _
        + vbCr + " FROM pla_marcaresum INNER JOIN (mae_tipohora INNER JOIN pla_marcaresumhor ON mae_tipohora.id = pla_marcaresumhor.idhora) ON pla_marcaresum.id = pla_marcaresumhor.idres " _
        + vbCr + " WHERE (((pla_marcaresumhor.idemp)=" & mIdEmp & ") AND ((pla_marcaresumhor.idres)=1)) "

''
''    nSQL = nSQL + vbCr + " UNION " _
''        + vbCr + " SELECT * FROM ( " _
''        + vbCr + " SELECT mae_tipohora.id & '-1' as Codigo,'1' as Origen, mae_tipohora.id AS idcpto,'' as categoria,  mae_tipohora.descripcion as concepto, mae_tipohora.variable,'' as formula, 0 AS aplanilla , Sum(pla_marcacionhora.totseg) AS imptot " _
''        + vbCr + " FROM pla_marcacion INNER JOIN (mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca " _
''        + vbCr + " WHERE pla_marcacionhora.IdEmp = " & mIdEmp & " And (((Month([pla_marcacion].[dia])) = " & mes & ") And ((Year([pla_marcacion].[dia])) = " & anno & ")) " _
''        + vbCr + " GROUP BY mae_tipohora.id & '-1',pla_marcacionhora.idemp, mae_tipohora.id, mae_tipohora.descripcion, mae_tipohora.variable " _
''        + vbCr + " UNION " _
''        + vbCr + " SELECT '15-1' as Codigo,'1' as Origen, mae_tipohora_1.id AS  idcpto,'' as categoria, mae_tipohora_1.descripcion as concepto, mae_tipohora_1.variable,'' as formula, 0 AS aplanilla , Sum(pla_marcacionhora.totseg) AS imptot " _
''        + vbCr + " FROM mae_tipohora AS mae_tipohora_1, pla_marcacion INNER JOIN (mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca " _
''        + vbCr + " WHERE pla_marcacionhora.IdEmp = " & mIdEmp & " And (((Month([pla_marcacion].[dia])) = " & mes & ") And ((Year([pla_marcacion].[dia])) = " & anno & ") And ((mae_tipohora.hortrabajo) = -1)) " _
''        + vbCr + " GROUP BY pla_marcacionhora.idemp, mae_tipohora_1.id, mae_tipohora_1.descripcion, mae_tipohora_1.variable, mae_tipohora_1.id " _
''        + vbCr + " HAVING (((mae_tipohora_1.id) = 15)) " _
''        + vbCr + " ) AS hora " '--) AS hora Order By  hora.idcpto; "
    
    
    '--total de horas del mes
    '--si esta en
    nSQL = nSQL + vbCr + " UNION " _
        + vbCr + " SELECT '16-1' AS Codigo, '1' AS Origen, mae_tipohora.id AS idcpto, '' AS categoria, mae_tipohora.descripcion AS concepto, mae_tipohora.variable, '' AS formula, 0 AS aplanilla, " _
        + vbCr + " IIF ('" & dFchIngreso & "' = '' , " & mTotalSegundosMes & " , IIF(month(cdate('" & dFchIngreso & "'))=" & mes & ",day(cdate('" & dFchIngreso & "'))*8*60, " & mTotalSegundosMes & " )) AS imptot " _
        + vbCr + " From mae_tipohora " _
        + vbCr + " GROUP BY mae_tipohora.id, mae_tipohora.descripcion, mae_tipohora.variable " _
        + vbCr + " HAVING (((mae_tipohora.id)=16)); "
              
    '--total bonificacion(171) ::origen mantenimiento de empleados
    nSQL = nSQL + vbCr + " UNION " _
        + vbCr + " SELECT pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, pla_empleados.bono AS imptot " _
        + vbCr + " FROM pla_empleados, pla_conceptocat INNER JOIN (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat " _
        + vbCr + " WHERE (((pla_concepto.id)=171) AND ((pla_empleados.id)= " & mIdEmp & " )); "

     '--total sueldo basico(177) ::origen mantenimiento de empleados
    nSQL = nSQL + vbCr + " UNION " _
        + vbCr + " SELECT pla_concepto.id & '-0' as Codigo,'0' as Origen, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, pla_empleados.basico AS imptot " _
        + vbCr + " FROM pla_empleados, pla_conceptocat INNER JOIN (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat " _
        + vbCr + " WHERE (((pla_concepto.id)=177) AND ((pla_empleados.id)= " & mIdEmp & " )); "
    
    '--falta los descuentos
    
    
    
    RST_Busq Rst, nSQL, xCon
    
End Sub


'************** CALCULOS
Public Sub pConceptoEfectuarCalculo(rst_remun As ADODB.Recordset, rst_descu As ADODB.Recordset, rst_aport As ADODB.Recordset, _
                            TotRemuneracion As Double, TotDescuento As Double, TotAportacion As Double, _
                            RstCptoValores As ADODB.Recordset, _
                            RstCptoEmp As ADODB.Recordset)
    DoEvents
    
    pCalculoAplicarFormula rst_remun, TotRemuneracion, TotDescuento, TotAportacion, RstCptoValores, RstCptoEmp
    TotRemuneracion = RstRegistroSumar(rst_remun, "imptot", "aplanilla", "-1", "N", True)
    
    '*********************************************************************************************************************************

    '--obtener el total de remuneracion considerando la tabla de sunat
    '--ingresos  con aportes de empleador y a vez armar la formula por cada concepto de aporte y descuento
    Dim nSQL As String
    Dim RstTmpAporte As New ADODB.Recordset
    Dim RstTmpCpto As New ADODB.Recordset
    Dim nFormula As String
    '--consulta de conceptos de aportes del trabajador y empleador
    nSQL = "SELECT pla_concepto.id AS idcpto, pla_conceptotipo.idcat, pla_conceptotipo.id AS idtipo, pla_concepto.descripcion, pla_concepto.variable, pla_concepto.formula " _
        + vbCr + " FROM pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + vbCr + " WHERE (((pla_conceptotipo.idcat)=2) AND ((pla_concepto.activo)=-1) AND ((pla_concepto.formula) Is Not Null And (pla_concepto.formula)<>''));"
    
    RST_Busq RstTmpAporte, nSQL, xCon
    
    Do While Not RstTmpAporte.EOF
        Set RstTmpCpto = Nothing
        '--conceptos relacionados al impuesto
        nSQL = "SELECT pla_concepto.id AS idcpto, pla_concepto.descripcion, pla_concepto.variable, pla_conceptoapo.idcptoref " _
            + vbCr + " FROM pla_concepto INNER JOIN pla_conceptoapo ON pla_concepto.id = pla_conceptoapo.idcpto " _
            + vbCr + " WHERE (((pla_concepto.activo)=-1) AND ((pla_conceptoapo.idcptoref)=" & NulosN(RstTmpAporte("idcpto")) & "));"
        RST_Busq RstTmpCpto, nSQL, xCon
        If RstTmpCpto.RecordCount <> 0 Then
            nFormula = "("
            If rst_remun.RecordCount <> 0 Then rst_remun.MoveFirst
            Do While Not rst_remun.EOF
                If NulosN(rst_remun.Fields("imptot")) <> 0 Then '--solo los que tienen monto distinto a cero
                    RstTmpCpto.Filter = "idcpto=" & NulosN(rst_remun("idcpto"))
                    If RstTmpCpto.RecordCount <> 0 Then
                        nFormula = nFormula & " " & RstTmpCpto("variable") & " +"
                    End If
                End If
                
                rst_remun.MoveNext
            Loop
            
            '---------------------------------------------------------------------------
            If nFormula <> "(" Then
            
                nFormula = Left(nFormula, Len(nFormula) - 2) & " )"
                nFormula = Replace(RstTmpAporte("formula"), "var_TotalIngreso", nFormula)
                
                If NulosN(RstTmpAporte.Fields("idtipo")) = 9 Then '--descuento
                    RstRegistroReemplazar rst_descu, "idcpto", NulosN(RstTmpAporte("idcpto")), True, "formula", nFormula
                ElseIf NulosN(RstTmpAporte.Fields("idtipo")) = 10 Then '--aporte
                    RstRegistroReemplazar rst_aport, "idcpto", NulosN(RstTmpAporte("idcpto")), True, "formula", nFormula
                End If
            
            End If
            '---------------------------------------------------------------------------
            
            
            
        End If
        RstTmpAporte.MoveNext
    Loop
    
    TotRemuneracion = 0
    '*********************************************************************************************************************************
    
    DoEvents
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    pCalculoAplicarFormula rst_descu, TotRemuneracion, TotDescuento, TotAportacion, RstCptoValores, RstCptoEmp
    TotDescuento = RstRegistroSumar(rst_descu, "imptot", "aplanilla", "-1", "N", True)
    DoEvents
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    pCalculoAplicarFormula rst_aport, TotRemuneracion, TotDescuento, TotAportacion, RstCptoValores, RstCptoEmp
    TotAportacion = RstRegistroSumar(rst_aport, "imptot", "aplanilla", "-1", "N", True)
    DoEvents
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rst_remun.Filter = ""
    rst_descu.Filter = ""
    rst_aport.Filter = ""
End Sub


Private Sub pCalculoAplicarFormula(Rst As ADODB.Recordset, _
                           TotRemuneracion As Double, _
                           TotDescuento As Double, _
                           TotAportacion As Double, _
                           RstCptoValores As ADODB.Recordset, _
                           RstCptoEmp As ADODB.Recordset)
                           
    Dim monto As String
    Dim xmark As Variant
    
    Dim n&
    Dim txt
    If Rst.RecordCount <> 0 Then Rst.MoveFirst
'    Do While Not Rst.EOF
'        For n = 0 To Rst.Fields.Count - 1
'            txt = txt & " ** " & NulosC(Rst.Fields(n))
'        Next
'        txt = txt & vbCrLf
'        Rst.MoveNext
'    Loop
        
    With Rst
'        If .RecordCount = 0 Then Exit Sub
        .Filter = "imptot = null or imptot = 0 "
        While Not .EOF
            '--almaceno en un temporal la posicion del registro
            xmark = .Bookmark
            If IsNumeric(.Fields("formula")) Then
               monto = .Fields("formula")
            ElseIf IsNull(.Fields("formula")) = True Then
                monto = ""
            Else
                monto = fCalculoPonerMonto(.Fields("Codigo"), NulosC(.Fields("formula")), _
                                    TotRemuneracion, TotDescuento, TotAportacion, _
                                     RstCptoValores, RstCptoEmp)
                
            End If
            If monto <> "error" Then
               .Bookmark = xmark    '--me ubico en la posicion temporal
               .Fields("imptot") = NulosN(monto)
               
               '--actualizar el recordset RstCptoEmp
               If NulosN(monto) <> 0 Then RstRegistroReemplazar RstCptoEmp, "IdCpto", .Fields("Idcpto"), True, "imptot", NulosN(monto)
            End If
            .MoveNext
        Wend
    End With
    Rst.Filter = ""
    
End Sub

Private Function fCalculoPonerMonto(ByVal CodConcepto As String, _
                                ByVal ConceptoFormula As String, _
                                ByVal TotRemuneracion As Double, _
                                ByVal TotDescuento As Double, _
                                ByVal TotAportacion As Double, _
                                ByVal RstCptoValores As ADODB.Recordset, _
                                ByVal RstCptoEmp As ADODB.Recordset) As String
                                
    Dim Formula As New CProcessor
    Dim Valor As String
    Dim NomVariable As String
    Dim CodConceptoRef As String
    Dim Xbookmark As Variant
    Dim Old_Filter As String
    Dim i&
    Dim RstTmpFormulas As New ADODB.Recordset
    If ConceptoFormula & "" = "" Then
       fCalculoPonerMonto = ""
       Exit Function
    Else 'si tiene formula
        '--busco en el recordset del detalle de formulas
        
        pCalculoObtenerCptoEnFormula RstTmpFormulas, CodConcepto, ConceptoFormula
        
'        RstCptoSoloFormula.Filter = "Codigo ='" & CodConcepto & "'"
        
        
        
        For i = 1 To RstTmpFormulas.RecordCount
            CodConceptoRef = RstTmpFormulas.Fields("CodigoRef")
            Valor = ""
            
            Select Case NulosC(RstTmpFormulas.Fields("CodigoRef"))
                Case "173-0" 'total aportacion
                    Valor = CStr(TotRemuneracion)
                Case "172-0" 'total descuentp
                    Valor = CStr(TotDescuento)
                Case "174-0" 'total aportacion
                    Valor = CStr(TotAportacion)
            End Select
            
            '--busco si el concepto tiene valor asignado un valor en la union de [ingresos + descuentos + aportaciones = RstCptoEmp]
            If Valor = "" Then Valor = fCalculoBuscarValor(RstCptoEmp, CodConceptoRef, "imptot")
            
            '*************************************************************************************************
            If Valor = "" Then Valor = fCalculoBuscarValor(RstCptoValores, CodConceptoRef, "imptot")
            If Valor = "" Then Valor = fCalculoBuscarValor(RstCptoValores, CodConceptoRef, "formula")
            '*************************************************************************************************
            'la variable a utilizar es otra formula ES UNA RECURSIVA!!!!!!!!!AUXILIO!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            If IsNumeric(Valor) = False And Valor <> "" Then
                Xbookmark = RstTmpFormulas.Bookmark
                Old_Filter = RstTmpFormulas.Filter

                Valor = fCalculoPonerMonto(CodConceptoRef, Valor, TotRemuneracion, TotDescuento, _
                          TotAportacion, RstCptoValores, RstCptoEmp)
                 'revizar
                 If Old_Filter <> 0 Then RstTmpFormulas.Filter = Old_Filter
                 'revizar
                 RstTmpFormulas.Bookmark = Xbookmark
            End If
            If Valor = "" Then Valor = "0"
            '*************************************************************************************************
            '--agregando a la clase la variable con su respectivo valor
            NomVariable = RstTmpFormulas.Fields("variableRef")
            Formula.DeclareConstant(NomVariable) = Valor
            RstTmpFormulas.MoveNext
        Next
        
        If Valor <> "" Then
             Formula.BaseCalculation = 1
             fCalculoPonerMonto = Format(Formula.Calculate(ConceptoFormula), "#####0.00")
        Else
             fCalculoPonerMonto = "error"
        End If
        
        Set Formula = Nothing
           
    End If
    
    Set RstTmpFormulas = Nothing
    
End Function

Private Sub pCalculoObtenerCptoEnFormula(ByVal Rst As ADODB.Recordset, CodConcepto As String, nFormula As String)
    Dim nSQL As String
    Dim nCpto As String
    nCpto = Replace(CodConcepto, "-0", "")
    nCpto = Replace(nCpto, "-1", "")
    nCpto = Replace(nCpto, "-2", "")
    
    '--obtener concepto de ingresos
    nSQL = "SELECT '" & nCpto & "-0' AS codigo, [pla_concepto].[id] & '-0' AS codigoRef, " & nCpto & " AS IdCpto, pla_concepto.id AS IdCptoRef, pla_concepto.variable AS variableRef, pla_concepto.descripcion, 0 AS Origen, 0 AS OrigenRef " _
        + vbCr + " From pla_concepto " _
        + vbCr + " WHERE ((('" & nFormula & "') Like '%' & [pla_concepto].[variable] & '%') AND ((pla_concepto.variable) Is Not Null)); "
    
    '--obtener concepto de tipos de horas
    nSQL = nSQL + vbCr + "UNION" _
        + vbCr + " SELECT '" & nCpto & "-0' AS Codigo, [mae_tipohora].[id] & '-1' AS CodigoRef, " & nCpto & " AS IdCpto, mae_tipohora.id AS IdCptoRef, mae_tipohora.variable AS variableRef, mae_tipohora.descripcion, 0 AS Origen, 1 AS OrigenRef " _
        + vbCr + " From mae_tipohora " _
        + vbCr + " WHERE ((('" & nFormula & "') Like '%' & [mae_tipohora].[variable] & '%') AND ((mae_tipohora.variable) Is Not Null));"
    
    '--obtener  conceptos de varios
    nSQL = nSQL + vbCr + "UNION" _
        + vbCr + " SELECT '" & nCpto & "-0' AS Codigo, [pla_conceptovarios].[id] & '-2' AS CodigoRef, " & nCpto & " AS IdCpto, pla_conceptovarios.id AS IdCptoRef, pla_conceptovarios.variable AS VariableRef, pla_conceptovarios.descripcion, 0 AS Origen, 2 AS OrigenRef " _
        + vbCr + " From pla_conceptovarios " _
        + vbCr + " WHERE ((('" & nFormula & "') Like '%' & [pla_conceptovarios].[variable] & '%') AND ((pla_conceptovarios.variable) Is Not Null));"
    
    RST_Busq Rst, nSQL, xCon
    
End Sub


Private Function fCalculoBuscarValor(ByVal Rst As ADODB.Recordset, CodConceptoRef As String, CampoValor As String) As String
    Dim xmark As Variant '--posicion inicial
    With Rst
        If .RecordCount > 0 Then
            xmark = .Bookmark
           .MoveFirst
           .Find "Codigo='" & CodConceptoRef & "'"
           If Not .EOF Then
              fCalculoBuscarValor = NulosC(.Fields(CampoValor))
             .Bookmark = xmark
              Exit Function
           End If
        End If
    End With
    fCalculoBuscarValor = ""
    Rst.Bookmark = xmark
End Function

'****************************************************

Public Sub pConceptoDocumentoEmp(Rst As ADODB.Recordset, _
                                 mIdBol As String, _
                                 eTipo As e_CategoriaConcepto, Optional fVariosRegistros As Boolean = False)

    Dim RstTmp As New ADODB.Recordset
    Dim nSQLListaIDS As String
    
    Dim nSQL As String
    
    If fVariosRegistros = False Then
        nSQLListaIDS = "pla_boleta.id = " & mIdBol & " "
    Else
        nSQLListaIDS = mIdBol
    End If
    
    '******************************************************************************************************
    Select Case eTipo
        Case 1 '--remuneraciones
            nSQL = "SELECT pla_boleta.idemp AS mIdEmp, [pla_concepto].[id] & '-0' AS codigo, 0 AS Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, pla_boletadet.imptot, pla_concepto.nomcorto " _
                + vbCr + " FROM pla_boleta LEFT JOIN (pla_conceptocat RIGHT JOIN (pla_conceptotipo RIGHT JOIN (pla_concepto RIGHT JOIN pla_boletadet ON pla_concepto.id = pla_boletadet.idcpto) ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat) ON pla_boleta.id = pla_boletadet.idbol " _
                + vbCr + " WHERE " & nSQLListaIDS & " And ((pla_conceptotipo.idcat) = 1) " _
                + vbCr + " ORDER BY pla_boleta.id, pla_concepto.orden; "
        
        Case 2 '--aportaciones
            nSQL = "SELECT pla_boleta.idemp AS mIdEmp, [pla_concepto].[id] & '-0' AS codigo, 0 AS Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, pla_boletadet.imptot, pla_concepto.nomcorto " _
                + vbCr + " FROM pla_boleta LEFT JOIN (pla_conceptocat RIGHT JOIN (pla_conceptotipo RIGHT JOIN (pla_concepto RIGHT JOIN pla_boletadet ON pla_concepto.id = pla_boletadet.idcpto) ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat) ON pla_boleta.id = pla_boletadet.idbol " _
                + vbCr + " WHERE ( " & nSQLListaIDS & " And ((pla_conceptotipo.idcat) = 2)) AND pla_conceptotipo.id=10  " _
                + vbCr + " ORDER BY pla_boleta.id, pla_concepto.orden; "
            
        Case 3 '--descuentos
            nSQL = "SELECT pla_boleta.idemp AS mIdEmp, [pla_concepto].[id] & '-0' AS codigo, 0 AS Origen, pla_conceptotipo.idcat, pla_concepto.id AS idcpto, pla_conceptocat.descripcion AS categoria, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula, pla_concepto.aplanilla, pla_boletadet.imptot, pla_concepto.nomcorto " _
                + vbCr + " FROM pla_boleta LEFT JOIN (pla_conceptocat RIGHT JOIN (pla_conceptotipo RIGHT JOIN (pla_concepto RIGHT JOIN pla_boletadet ON pla_concepto.id = pla_boletadet.idcpto) ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat) ON pla_boleta.id = pla_boletadet.idbol " _
                + vbCr + " WHERE " & nSQLListaIDS & " And (pla_conceptotipo.idcat = 3 OR (pla_conceptotipo.idcat= 2 AND pla_conceptotipo.id=9))  " _
                + vbCr + " ORDER BY pla_boleta.id, pla_concepto.orden; "
            
    End Select
    
    If fVariosRegistros = False Then
        '--cargar los datos
        RST_Busq RstTmp, nSQL, xCon
        '--si no tiene campos el recoordset => definir recordset temporal
        If Rst.State = 0 Then DEFINIR_RST_TMP Rst, RstTmp
        '--cargar los datos al recordset temporal
        CARGAR_RST_TMP Rst, RstTmp
        Set RstTmp = Nothing
    Else
        RST_Busq Rst, nSQL, xCon
    End If
    
    '--
    
End Sub

'********** FIN PLANILLA DE PAGO ********************


Public Sub pCargarPersonal(frm As Form, Index As Integer, Optional fSoloPlanilla As Boolean = False)
    Dim xRs As New ADODB.Recordset
    pBuscarPersonal xRs, True, fSoloPlanilla
    If xRs.State = 1 Then
        frm.txt_cb(Index) = xRs.Fields("id") & "" '--TEXTO A MOSTRAR
        frm.lbl_cb(Index).Caption = xRs.Fields("nombres") & "" '--NOMBRE
        frm.lbl_cod(Index).Caption = xRs.Fields("id") & "" '--CODIGO
        frm.lbl_cb(Index).ToolTipText = xRs.Fields("nombres") & "" '--NOMBRE
    End If
    Set xRs = Nothing
End Sub







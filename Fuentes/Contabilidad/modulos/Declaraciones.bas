Attribute VB_Name = "Declaraciones"
Option Explicit

Public xCon As New ADODB.Connection
Public xTitulo As String

Public NomEmp As String
Public NumRUC As String
Public CONTABILIZAR As Boolean
Public xMes As Integer
Public AnoTra As String
Public xIdEmpresa As Integer
Public MostrarValorizado  As Boolean
Public xIdUsuario As Integer

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)

Public Const FORMAT_IMPORTEKARDEX = "###,###,##0.0000"

Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    Set Rst = Nothing
End Sub

Function KardexMovimientoSQL(xIdItem As Double, IdAlmacen As Integer, xFchIni As Date, xFchFin As Date) As String
     '===================================================================================================
    'Creado : 19/01/12 Por Johan Castro
    'Propósito: Generar Sentencia SQL para mostrar los movimientos al detalle de un item en funcion al periodo indicado
    '
    'Entradas:  xIdItem= Código del item(ver tabla mae_inventario.id)
    '           xFchIni= Fecha Inicial
    '           xFchFin= Fecha Final
    '
    'Resultados: Sentencia SQL listo para ejecutar
    '
    'Nota:  Descripción de los estados (ver tabla mae_estados)
    '--1=Pendiente, 2=Procesado, 3=Aprobado,4=Rechazado,5=Anulado
    '
    'Modificado: 21/01/12 Johan Castro; Se extrae la Sentencia SQL de evento FrmVerKardex.MuestraKardexProm y FrmResuMov.HallarPrecioPromedio que son lo mismo y se centraliza
    '           No mostrar registros que esten con el siguiente estado(1=Pendiente, 4=Rechazado y 5=Anulado)
    '           Agregar filtro en Almacen Ingreso - Salida "AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 "
    '           Agregar filtro en Produccion Insumos, "AND pro_producciondet.estado Not In (1,4,5) AND pro_producciondetins.canutil<>0 AND alm_inventario.tippro=3"
    '           Agregar filtro en Produccion Producto Terminado, "AND pro_producciondet.estado Not In (1,4,5)"
    '           Agregar Variable xSQLFiltroPS para mostrar solo materia prima a partir del 2012 en adelante en Produccion Insumos

    '===================================================================================================

    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    'GR = GUIAS DE REMISION
    'PS =
    
    '--ALMACEN INGRESO
    '--ALMACEN SALIDA
    '--COMPRAS
    '--GUIA REMISION => SALIDA
    '--PRODUCCION DETALLE INSUMOS,PRODUCTOS INTERMEDIOS, MATERIA PRIMA => SALIDA
    '--PRODUCCION PRODUCTOS TERMINADOS => SALIDA
    '--VENTAS => SALIDA
    '--VENTAS => DEVOLUCIONES INGRESO

    Dim xCadSQL As String
    Dim xSQLFiltroPS As String '--Util para aplicar un filtro adicional que mostrará solo materia prima en sentencia de "produccion insumos salida"
    Dim F As New SistemaLogica.Funciones
    Dim mFechaInicioMovimientos As Date
    Dim IdTipoInvInicial As Integer


    mFechaInicioMovimientos = F.FechaInicioMovimientos(CInt(IdAlmacen), xCon)
    IdTipoInvInicial = F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", xCon))
    If NulosN(AnoTra) >= 2012 Then
        '--Aplicar filtro en produccion de salida para mostrar solo materia prima del 2012 en adelante
        xSQLFiltroPS = " AND alm_inventario.tippro=3  "
    End If

    ' PREPARAMOS LA SELECT PARA ARMAR EL KARDEX
    xCadSQL = "SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, alm_ingreso.numser, alm_ingreso.numdoc,  alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, IIf(alm_ingreso.idtipdocref=" & IdTipoInvInicial & ", 'II','AI') AS tipo, alm_ingreso.tipmov, alm_ingreso.nombre AS entidad, 0 AS aa,  (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos ,'Almacén' & iif(cstr(numdocumentos) <>'0', ' - Compras','') as modulo, '' AS registro, '' AS ctanum, '' AS ctanom, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin  " _
        + vbCr + " FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id " _
        + vbCr + " WHERE (((alm_ingresodet.iditem)=" & xIdItem & ") AND ((alm_ingreso.fching)>=CDate('" & xFchIni & "') " _
                & " And (alm_ingreso.fching)<=CDate('" & xFchFin & "') AND (alm_ingreso.fching)>=CDate('" & mFechaInicioMovimientos & "')) AND ((alm_ingreso.tipmov)=-1)) " _
                & " AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " _
        + vbCr + " UNION ALL " _
        + vbCr + " SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, alm_ingreso.numser, alm_ingreso.numdoc, " _
                & " alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AS' AS tipo, alm_ingreso.tipmov, alm_ingreso.nombre AS entidad, 0 AS aa, " _
                & " (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos ,'Almacén' & iif(cstr(numdocumentos) <>'0', ' - Compras','') as modulo, '' AS registro,'' AS ctanum, '' AS ctanom, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin  " _
        + vbCr + " FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id " _
        + vbCr + " WHERE (((alm_ingresodet.iditem)=" & xIdItem & ") AND ((alm_ingreso.fching)>=CDate('" & xFchIni & "') " _
                & " And (alm_ingreso.fching)<=CDate('" & xFchFin & "') AND (alm_ingreso.fching)>=CDate('" & mFechaInicioMovimientos & "')) AND ((alm_ingreso.tipmov)=0)) " _
                & " AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " _
        + vbCr + " UNION ALL " _
        + vbCr + " SELECT com_compras.id, com_comprasdet.iditem, alm_inventario.descripcion, com_compras.fchdoc, com_compras.numser, com_compras.numdoc, " _
                & " com_comprasdet.canpro, IIf([com_compras]![idmon]=2,[com_comprasdet]![preuni]*[con_tc]![impcom],[com_comprasdet]![preuni]) AS preuni, mae_documento.abrev AS descdoc, " _
                & " 'C' AS Tipo, -1 AS tipmov, mae_prov.nombre AS entidad, 0 AS aa, 0 AS numdocumentos,'Compras' as modulo,com_compras.numreg AS registro,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS desope, '' AS horini, '' AS horfin " _
        + vbCr + " FROM (alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc)  " _
                & " LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
        + vbCr + " WHERE (((com_comprasdet.iditem)=" & xIdItem & ") AND " _
                & " ((com_compras.fchdoc)>=CDate('" & xFchIni & "') And (com_compras.fchdoc)<=CDate('" & xFchFin & "') AND (com_compras.fchdoc)>=CDate('" & mFechaInicioMovimientos & "')) AND ((com_compras.tipcom)=1))"

    xCadSQL = xCadSQL _
        + vbCr + "  Union All" _
        + vbCr + " SELECT vta_guia.id, vta_guiadet.iditem, alm_inventario.descripcion, vta_guia.fecgiro, vta_guia.numser, vta_guia.numdoc, vta_guiadet.canpro, " _
                & " 0 AS preuni, mae_documento.abrev AS desdoc, 'GR' AS tipo, 0 AS tipmov, mae_cliente.nombre AS entidad, 0 AS aa, IIf([vta_guia]![iddocven]<>0,1,0) AS numdocumentos,'Guia de Remisión' as modulo, '' AS registro,'' AS ctanum, '' AS ctanom, '' AS desope, '' AS horini, '' AS horfin  " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN vta_guia ON mae_cliente.id = vta_guia.idcli) LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id) LEFT JOIN (vta_guiadet " _
                & " LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id) ON vta_guia.id = vta_guiadet.idgui " _
        + vbCr + " WHERE (((vta_guiadet.iditem)=" & xIdItem & ") " _
                & " AND ((vta_guia.fecgiro)>=CDate('" & xFchIni & "') And (vta_guia.fecgiro)<=CDate('" & xFchFin & "') AND (vta_guia.fecgiro)>=CDate('" & mFechaInicioMovimientos & "'))) " _

    xCadSQL = xCadSQL + " UNION All " _
        + vbCr + " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, vta_ventas.numser, vta_ventas.numdoc, " _
                    & " vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, 0 AS tipmov, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, " _
                    & " 'Ventas' as modulo, vta_ventas.numreg AS registro,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS desope, '' AS horini, '' AS horfin " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet  " _
                    & " LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
        + vbCr + " WHERE (((vta_ventasdet.iditem)=" & xIdItem & ") " _
                    & " AND ((vta_ventas.fchdoc)>=CDate('" & xFchIni & "') And (vta_ventas.fchdoc)<=CDate('" & xFchFin & "') AND (vta_ventas.fchdoc)>=CDate('" & mFechaInicioMovimientos & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) " _
                    & " AND ((vta_ventas.iddocref) Is Null Or (vta_ventas.iddocref)=0) )" _
        + vbCr + " UNION All " _
        + vbCr + " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, vta_ventas.numser, vta_ventas.numdoc, " _
                    & " vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, 0 AS tipmov, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, " _
                    & " 'Ventas NC' AS modulo, vta_ventas.numreg AS registro,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS desope, '' AS horini, '' AS horfin " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet " _
                    & " LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
        + vbCr + " WHERE (((vta_ventasdet.iditem)=" & xIdItem & ") AND ((vta_ventas.fchdoc)>=CDate('" & xFchIni & "') And (vta_ventas.fchdoc)<=CDate('" & xFchFin & "') AND (vta_ventas.fchdoc)>=CDate('" & mFechaInicioMovimientos & "')) " _
                    & " AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref)<>0) AND ((vta_ventas.idmotnotcre)=4))"

    KardexMovimientoSQL = xCadSQL

End Function

Public Function pHallarPrecioInicial(IDTEM_ As Integer, FECHA_ As String, ANIOTRABAJO_ As Integer) As Double
    Dim xRs As New ADODB.Recordset
    Dim cSQL As String
    Dim SALDOINICIAL_ As Double
    Dim INGRESOCANTIDAD_ As Double
    Dim INGRESOIMPORTE_ As Double
    Dim SALIDACANTIDAD_ As Double
    Dim SALIDAIMPORTE_ As Double
    Dim PRECIOPROMEDIO_ As Double
    
    ' SI ES EL PRIMER DIA DEL AÑO
    If CDate(FECHA_) = CDate("01/01/" & ANIOTRABAJO_) Then
        pHallarPrecioInicial = NulosN(Busca_Codigo("id", NulosC(IDTEM_), "preini", "alm_inventario", "N", xCon))
    Else
        pHallarDatosMovimientos IDTEM_, "01/01/" & ANIOTRABAJO_, CDate(FECHA_), SALDOINICIAL_, INGRESOCANTIDAD_, INGRESOIMPORTE_, SALIDACANTIDAD_, SALIDAIMPORTE_, PRECIOPROMEDIO_
        If INGRESOCANTIDAD_ = SALIDACANTIDAD_ And SALDOINICIAL_ = 0 Then
            pHallarPrecioInicial = 0: Exit Function
        Else
            pHallarPrecioInicial = PRECIOPROMEDIO_: Exit Function
        End If
    End If
End Function

Public Function PrecioUni(IdDocumento, IdItem As Double, DondeBuscar As String) As Double
    '===================================================================================================
    'Creado:     01/07/11 Johan Castro
    'Propósito:  Obtener el Precio unitario del registro de compras vinculado con documentos (de ingreso de almacen, Guia Remision)
    '
    'Entradas:   IdDocumento = Código de Libro
    '            IdItem = Código del Item (Producto, Materia prima, Insumo, etc)
    '            DondeBuscar = Indica el origen del registro
    '
    'Resultados: Precio unitario del item segun el documento ingresado
    '===================================================================================================
    
    Dim xRst As New ADODB.Recordset
    Dim F As New SistemaLogica.Funciones
    Dim nSQL As String
    
    If DondeBuscar = "AI" Then
        nSQL = "SELECT Sum(IIf([com_compras].[idmon]=2,[com_comprasdet].[imptot]*[con_tc].[impcom],[com_comprasdet].[imptot])) AS importetot, Sum(com_comprasdet.canpro) AS cantidadtot " _
            + vbCr + "FROM (com_compras INNER JOIN (com_comprasdet INNER JOIN alm_ingresodoc ON com_comprasdet.idcom = alm_ingresodoc.iddoc) ON com_compras.id = com_comprasdet.idcom) INNER JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
            + vbCr + "WHERE (((alm_ingresodoc.id)=" & IdDocumento & ") AND ((com_comprasdet.iditem)=" & IdItem & "));"
    
        Set xRst = Nothing
        RST_Busq xRst, nSQL, xCon
        If xRst.State = 0 Then PrecioUni = 0: Exit Function
        If xRst.RecordCount = 0 Then PrecioUni = 0: Exit Function
        
        If NulosN(xRst("cantidadtot")) = 0 Then
            PrecioUni = 0
        Else
            PrecioUni = NulosN(xRst("importetot")) / NulosN(xRst("cantidadtot"))
        End If
        Exit Function
        
    ElseIf DondeBuscar = "II" Then
        nSQL = "SELECT alm_inventarioinicialdet.costo " _
            + vbCr + "FROM (alm_inventarioinicial INNER JOIN alm_inventarioinicialdet ON alm_inventarioinicial.idinventarioinicial = alm_inventarioinicialdet.idinventarioinicial) INNER JOIN alm_ingreso ON alm_inventarioinicial.idinventarioinicial = alm_ingreso.iddocref " _
            + vbCr + "WHERE (((alm_ingreso.id)=" & IdDocumento & ") AND ((alm_inventarioinicial.idestado)=" & F.NuloNumeric(F.KeyValue("EstadoAprobadoInventarioInicial", xCon)) & ") AND ((alm_inventarioinicialdet.iditem)=" & IdItem & "))"
    
        Set xRst = Nothing
        RST_Busq xRst, nSQL, xCon
        If xRst.State = 0 Then PrecioUni = 0: Exit Function
        If xRst.RecordCount = 0 Then PrecioUni = 0: Exit Function
        
        If NulosN(xRst("costo")) = 0 Then
            PrecioUni = 0
        Else
            PrecioUni = NulosN(xRst("costo"))
        End If
        Exit Function
        
    ElseIf DondeBuscar = "GR" Then
        nSQL = "SELECT vta_guia.id, vta_ventasdet.iditem, Avg(vta_ventasdet.preuni) AS preuniprom " _
            + vbCr + " FROM vta_guia INNER JOIN vta_ventasdet ON vta_guia.iddocven = vta_ventasdet.idvta " _
            + vbCr + " GROUP BY vta_guia.id, vta_ventasdet.iditem " _
            + vbCr + " HAVING (((vta_guia.id)=" & IdDocumento & ") AND ((vta_ventasdet.iditem)=" & IdItem & ")); "
       
        Set xRst = Nothing
        RST_Busq xRst, nSQL, xCon
        If xRst.State = 0 Then PrecioUni = 0: Exit Function
        If xRst.RecordCount = 0 Then PrecioUni = 0: Exit Function
        
        PrecioUni = NulosN(xRst("preuniprom"))
        Exit Function
    Else
        PrecioUni = 0
        Exit Function
    End If
End Function

Public Sub pHallarDatosMovimientos(IDITEM_ As Integer, FCHINI_ As Date, FCHFIN_ As Date, ByRef SALDOINICIAL_ As Double, ByRef INGRESOCANTIDAD_ As Double, ByRef INGRESOIMPORTE_ As Double, _
                            ByRef SALIDACANTIDAD_ As Double, ByRef SALIDAIMPORTE_ As Double, ByRef PRECIOPROMEDIO_ As Double)
    Dim xCadSQL As String
    Dim UltPreCosto As Double
    Dim mInicioGrupo As Long '--indica la fila inicial de un grupo, cambia cuando cambia de item
    Dim xPrecioUni As Double '--Indica el precio unitario de cada registro
    Dim xPrecioIni As Double
    Dim StockIni As Double
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim TIPOPRODUCTO_ As Integer
    Dim xSaldo As Double
    Dim xSaldoImp As Double
    Dim A&
    Dim xFila As Integer
    Dim xTotSal, xTotEnt As Double
    Dim xImpSal, xImpEnt As Double
    Dim cSQL As String
    
    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    ' GR = GUIAS DE REMISION
    ' PS =
    
    xTotSal = 0
    xImpSal = 0
    xTotEnt = 0
    xImpEnt = 0
    xPrecioUni = 0
    
    '--Generar la consulta SQL para obtener el detalle de movimientos del kardex
    xCadSQL = KardexMovimientoSQL(CDbl(IDITEM_), 0, FCHINI_, FCHFIN_)
            
    RST_Busq xRs, xCadSQL, xCon
    xRs.Sort = "fchdoc, Tipo, numdoc"
          
    StockIni = 0
    xPrecioIni = 0
    xSaldo = StockIni
    xSaldoImp = xSaldo * xPrecioIni
    xPrecioUni = xPrecioIni
    
    If xRs.RecordCount = 0 Then GoTo SALIR_
    xRs.MoveFirst

    While Not xRs.EOF
        ' ----------------------------------------------INGRESOS
        If xRs("tipo") = "C" Or xRs("tipo") = "AI" Or xRs("tipo") = "P" Then
            If xRs("descdoc") = "NC" Then
                xSaldo = xSaldo - NulosN(xRs("canpro"))
                xTotSal = xTotSal + NulosN(xRs("canpro"))
            Else
                xSaldo = xSaldo + NulosN(xRs("canpro"))
                xTotEnt = xTotEnt + NulosN(xRs("canpro"))
            End If
            
            '--obtener el precio
            If xRs("tipo") = "AI" And xRs("numdocumentos") <> 0 Then
                xPrecioUni = PrecioUni(xRs("id"), CDbl(IDITEM_), NulosC(xRs("tipo")))
            Else
                xPrecioUni = NulosN(xRs("preuni"))
            End If
            
            If xRs("tipo") = "P" Then
                xPrecioUni = 0
                
                cSQL = "SELECT con_librocostodet.impmprima, con_librocostodet.impmanobr, con_librocostodet.impgasfab, con_librocostodet.cantidad " _
                    + vbCr + "FROM con_librocostodet " _
                    + vbCr + "WHERE (((con_librocostodet.idprod)=" & NulosN(xRs("id")) & ") AND ((con_librocostodet.iditem)=" & IDITEM_ & "));"
                
                Set xRsAux = Nothing
                RST_Busq xRsAux, cSQL, xCon
                
                If xRsAux.State = 0 Then Exit Sub
                If xRsAux.RecordCount > 0 Then
                    xPrecioUni = (NulosN(xRsAux("impmprima")) + NulosN(xRsAux("impmanobr")) + NulosN(xRsAux("impgasfab"))) / NulosN(xRs("canpro"))
'                        xPrecioUni = (NulosN(xRs("impmprima")) + NulosN(xRs("impmanobr"))) / NulosN(xRs("cantidad"))
                End If
            End If
                             
            If xRs("descdoc") = "NC" Then
                xSaldoImp = xSaldoImp - (NulosN(xRs("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(xRs("canpro")) * xPrecioUni)
            Else
                xSaldoImp = xSaldoImp + (NulosN(xRs("canpro")) * xPrecioUni)
                xImpEnt = xImpEnt + (NulosN(xRs("canpro")) * xPrecioUni)
            End If
                                
            UltPreCosto = xPrecioUni
        ' ----------------------------------------------------------SALIDAS
        Else
            If xSaldo = 0 Then
                xPrecioUni = 0
            Else
                xPrecioUni = xSaldoImp / xSaldo
            End If
            UltPreCosto = xPrecioUni
            
            If xRs("descdoc") = "NC" Then
                xSaldo = xSaldo + NulosN(xRs("canpro"))
                xTotEnt = xTotEnt + NulosN(xRs("canpro"))
            Else
                xSaldo = xSaldo - NulosN(xRs("canpro"))
                xTotSal = xTotSal + NulosN(xRs("canpro"))
            End If
                                        
            If xRs("descdoc") = "NC" Then
                xSaldoImp = xSaldoImp + (NulosN(xRs("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(xRs("canpro")) * xPrecioUni)
            Else
                xSaldoImp = xSaldoImp - (NulosN(xRs("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(xRs("canpro")) * xPrecioUni)
            End If
        End If
        
        xRs.MoveNext
    Wend

SALIR_:
    SALDOINICIAL_ = StockIni
    SALIDACANTIDAD_ = xTotSal
    SALIDAIMPORTE_ = xImpSal
    INGRESOCANTIDAD_ = xTotEnt
    INGRESOIMPORTE_ = xImpEnt
    PRECIOPROMEDIO_ = xPrecioUni
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Varias"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TIPO_CAMBIO
    'VALORES PARA BUSCAR EL TIPO DE CAMBIO
    compra = 1
    Venta = 2
End Enum

Public Enum TIPO_MOVIMIENTO
    'VALORES PARA BUSCAR EL TIPO DE CAMBIO
    origen = 1
    destino = 2
End Enum


Function DetallarModulo(IdOriDes As Integer, TipoMov As TIPO_MOVIMIENTO, Con As ADODB.Connection) As Boolean
    Dim Rst As New ADODB.Recordset
    If TipoMov = origen Then
        RST_Busq Rst, "SELECT * FROM tes_origen WHERE id = " & IdOriDes & "", Con
    Else
        RST_Busq Rst, "SELECT * FROM tes_destino WHERE id = " & IdOriDes & "", Con
    End If
    
    If Rst.RecordCount <> 0 Then
        DetallarModulo = NulosN(Rst("detallar"))
    End If
    Set Rst = Nothing
End Function

Function SeleccionaMes(Coneccion As ADODB.Connection) As Integer
    Set xCon = Coneccion
    FrmSeteaMes.Show vbModal
    SeleccionaMes = xMes
End Function

Function NomMes(nummes As Integer) As String
    'Valor Requerido un dato de tipo integer
    'Devuelve el nombre del Mes
    Select Case nummes
        Case 1
            NomMes = "Enero"
        Case 2
            NomMes = "Febrero"
        Case 3
            NomMes = "Marzo"
        Case 4
            NomMes = "Abril"
        Case 5
            NomMes = "Mayo"
        Case 6
            NomMes = "Junio"
        Case 7
            NomMes = "Julio"
        Case 8
            NomMes = "Agosto"
        Case 9
            NomMes = "Setiembre"
        Case 10
            NomMes = "Octubre"
        Case 11
            NomMes = "Noviembre"
        Case 12
            NomMes = "Diciembre"
    End Select
End Function

Function HallaTipoCambio(Dia As String, Moneda As Integer, Tipo As TIPO_CAMBIO, Coneccion As ADODB.Connection) As Double
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT con_tc.fecha, con_tc.idmon, con_tc.impcom,con_tc.impven From con_tc " _
        & " WHERE (((con_tc.fecha)=CDate('" & Dia & "')) AND ((con_tc.idmon)=" & Moneda & "))", Coneccion
    If Rst.State = 0 Then Exit Function
    If Rst.RecordCount <> 0 Then
        If Tipo = compra Then
            HallaTipoCambio = NulosN(Rst("impcom"))
        Else
            HallaTipoCambio = NulosN(Rst("impven"))
        End If
    Else
        HallaTipoCambio = 0
    End If
    Set Rst = Nothing
End Function

Function DevuelveNumAsiento(Libro As Integer, IdMovimiento As Double, Mes As Integer, Coneccion As ADODB.Connection) As String
    Dim RstDia As New ADODB.Recordset
    
    RST_Busq RstDia, "SELECT * FROM con_diario WHERE idmes = " & Mes & " AND " _
        & " idlib = " & Libro & " AND idmov = " & IdMovimiento & "", Coneccion
            
    If RstDia.RecordCount <> 0 Then
        DevuelveNumAsiento = RstDia("numasi") & ""
    Else
        DevuelveNumAsiento = ""
    End If
    Set RstDia = Nothing
End Function

Function NuevoNumAsiento(Libro As Integer, Mes As Integer, Coneccion As ADODB.Connection) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT TOP 1 con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=" & Libro & ")) ORDER BY numasi DESC ", Coneccion
    
    If Rst.RecordCount = 0 Then
        NuevoNumAsiento = "0001"
    Else
        Rst.MoveLast
        NuevoNumAsiento = Format(NulosN(Rst("numasi")) + 1, "0000")
    End If
    Set Rst = Nothing
    
End Function

Function HallaNumdocVenta(TipoDoc As Integer, NumSerie As String, Coneccion As ADODB.Connection) As String
    Dim Rst As New ADODB.Recordset
    Dim xNum As Double
    RST_Busq Rst, "SELECT top 1 vta_ventas.tipdoc, vta_ventas.numser, vta_ventas.numdoc From vta_ventas " _
        & " WHERE (((vta_ventas.tipdoc)=" & TipoDoc & ") AND ((vta_ventas.numser)='" & NumSerie & "')) " _
        & " ORDER BY vta_ventas.numdoc DESC ", Coneccion

    If Rst.RecordCount <> 0 Then
        Rst.MoveLast
        xNum = NulosN(Rst("numdoc"))
        xNum = xNum + 1
        HallaNumdocVenta = Format(xNum, "0000000000")
    Else
        xNum = 1
        HallaNumdocVenta = Format(xNum, "0000000000")
    End If
    Set Rst = Nothing
End Function

Function HallaNumGuia(NumSerie As String, Coneccion As ADODB.Connection) As String
    Dim Rst As New ADODB.Recordset
    Dim xNum As Double
    RST_Busq Rst, "SELECT top 1 numdoc FROM vta_guia WHERE numser =  '" & NumSerie & "' ORDER BY numdoc DESC", Coneccion

    If Rst.RecordCount <> 0 Then
        Rst.MoveLast
        xNum = Val(Rst("numdoc"))
        xNum = xNum + 1
        HallaNumGuia = Format(xNum, "0000000000")
    Else
        xNum = 1
        HallaNumGuia = Format(xNum, "0000000000")
    End If
    Set Rst = Nothing
End Function

Sub CargaDatosEmpresa(xCon As ADODB.Connection, NomEmp As String, NumRUC As String)
On Error Resume Next
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    
    Set Rst = Nothing
Err.Clear
End Sub

Function TotalProducido(IdProducto As Integer, FchInicio As String, FchFinal As String, xCon As ADODB.Connection) As Double
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim xTotal As Double
    Dim xCad As String
    
    'CARGAMOS TODAS LAS ENTRADAS
    xCad = "SELECT Sum(com_comprasdet.canpro) AS total FROM alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN " _
        & " (com_compras LEFT JOIN  con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet " _
        & " ON com_compras.id = com_comprasdet.idcom)  ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem  " _
        & " WHERE (((com_comprasdet.iditem)=" & IdProducto & ") AND ((com_compras.fchdoc)>=CDate('" & FchInicio & "') And (com_compras.fchdoc)<=CDate('" & FchFinal & "')) " _
        & " AND ((con_tc.idmon)=2))" _
        & " Union " _
        & " SELECT alm_ingresodet.cantidad AS total FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN " _
        & " (alm_ingresodet  LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id  " _
        & " WHERE (((alm_ingresodet.iditem)=" & IdProducto & ") AND ((alm_ingreso.fching)>=CDate('" & FchInicio & "') And (alm_ingreso.fching)<=CDate('" & FchFinal & "')) " _
        & " AND ((alm_ingreso.tipmov)=-1)) " _
        & " Union " _
        & " SELECT Sum([cantidad]) AS total FROM pro_produccion LEFT JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        & " WHERE (((pro_producciondet.iditem)=" & IdProducto & ") AND ((pro_produccion.dia)>=CDate('" & FchInicio & "') And (pro_produccion.dia)<=CDate('" & FchFinal & "')))"

    RST_Busq Rst, xCad, xCon
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            xTotal = xTotal + NulosN(Rst("total"))
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        
        TotalProducido = xTotal
    Else
        TotalProducido = 0
    End If
End Function

Function CambioOpcionLiro(idFormato As Integer, Con As ADODB.Connection, TipoFormatos As Integer) As Boolean
    Set xCon = Con
    xIdFormatos = TipoFormatos
    FrmConfFormato.LblIdLibro.Caption = idFormato
    FrmConfFormato.Show vbModal
    If Grabo = True Then
        CambioOpcionLiro = True
    Else
        CambioOpcionLiro = False
    End If
End Function

Function GrabarOperacion(IdUsuario As Integer, IdFormulario As Integer, IdOperacion As Integer, HorIni As Date, HorFin As Date, Fecha As Date, xCon As ADODB.Connection, IdMovimiento As Double)
    Dim xId As Double
    xId = HallaCodigoTabla("var_edicion", xCon, "id")
    
    xCon.Execute "INSERT INTO var_edicion ( id, idusuario, idform, idoperacion, horini, horfin, fchope, idmov )" _
        & " SELECT " & xId & " AS Expr1, " & IdUsuario & " AS Expr2, " & IdFormulario & " AS Expr3, " & IdOperacion & " AS Expr4, " _
        & " '" & HorIni & "' AS Expr5, '" & HorFin & "' AS Expr6, '" & CDate(Fecha) & "' AS Expr7, " & IdMovimiento & " as Expr8 "
End Function

Sub VerMovimientos(CadenaSQL As String, Con As ADODB.Connection)
    Set xCon = Con
    FrmVerMovimiento.TxtSQL.Text = CadenaSQL
    FrmVerMovimiento.Show vbModal
End Sub

Sub VerMovimientos1(IdMenu As Integer, idmov As Double, Con As ADODB.Connection)
    '===================================================================================================
    'Creado 30/04/2009 Por: Johan Castro
    'Propósito: Mostrar una ventana historial del registro, es una mejora de funcion VerMovimientos()
    '
    'Entradas:  IdMenu = Codigo del menu (Formulario utilizado por el usuario)
    '           idmov = Codigo del movimiento
    '           Con = Conexion activa a la base de datos
    '
    'Resultados: Formulario con historial del registro
    '
    'Modificado  13/01/11 por Johan Castro
    '            Se modifica la consulta, inicialmente estaba ralacionada a tabla mae_formularios,
    '            se modifica para que este relacionada a tabla mae_menu
    '            15/01/11 por Johan Castro
    '            Cambiar tipo de dato de parametro IdMenu antes double
    '===================================================================================================
    
    
    Set xCon = Con
    Dim CadenaSQL As String
    
    'antes no se utiliza
    'CadenaSQL = "SELECT var_edicion.*, UCase(mae_usuarios!ape) & ', ' & mae_usuarios!nom AS apenom, mae_formularios.descripcion AS formulario, " _
            & " mae_tipooperacion.descripcion AS operacion FROM ((var_edicion LEFT JOIN mae_usuarios ON var_edicion.idusuario = mae_usuarios.id) " _
            & " LEFT JOIN mae_formularios ON var_edicion.idform = mae_formularios.id) LEFT JOIN mae_tipooperacion ON var_edicion.idoperacion = mae_tipooperacion.id " _
            & " WHERE (((var_edicion.idform)=" & IdMenu & ") AND ((var_edicion.idmov)=" & idmov & "))"
            
    CadenaSQL = "SELECT var_edicion.*, UCase(mae_usuarios!ape) & ', ' & mae_usuarios!nom AS apenom, mae_menu.descripcion AS formulario, mae_tipooperacion.descripcion AS operacion " _
    & " FROM ((var_edicion LEFT JOIN mae_usuarios ON var_edicion.idusuario = mae_usuarios.id) LEFT JOIN mae_tipooperacion ON var_edicion.idoperacion = mae_tipooperacion.id) LEFT JOIN mae_menu ON var_edicion.idform = mae_menu.id " _
    & " WHERE (((var_edicion.idform)=" & IdMenu & ") AND ((var_edicion.idmov)=" & idmov & "))"
    
    FrmVerMovimiento.TxtSQL.Text = CadenaSQL
    FrmVerMovimiento.Show vbModal
End Sub

Function OpcionBusquedaForm(IdFormulario As Integer, IdBusqueda As Integer, Con As ADODB.Connection) As Integer
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM var_busquedaformulario WHERE idform = " & IdFormulario & " AND idbusform = " & IdBusqueda & "", Con
    If Rst.RecordCount <> 0 Then
        OpcionBusquedaForm = Rst("opcion")
    Else
        OpcionBusquedaForm = 0
    End If
    Set Rst = Nothing
End Function

Sub OpcionesUsuario(xIdUsuario As Integer, xidMenu As Integer, xToolbar As Object, xCon As ADODB.Connection)

'--Modificado 06/11/11 Johan Castro
'             Ocultar opciones cuando usuario no este en relaciòn de accesos
'             Si usuario es administrador e intenta entrar a menu [Configurar Opciones de Usuario]
'             se le dara privilegios de Modificar por defecto
    On Error GoTo LaCague
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT mae_menuusuario.idusuario, mae_menuusuario.idmenu, mae_menuusuario.opcion1, mae_menuusuario.opcion2, mae_menuusuario.opcion3 " _
        & " From mae_menuusuario WHERE (((mae_menuusuario.idusuario)=" & xIdUsuario & ") AND ((mae_menuusuario.idmenu)=" & xidMenu & "))", xCon
    If Rst.RecordCount <> 0 Then
        '--Agregar
        If Rst("opcion1") = -1 Then
            xToolbar.Buttons(1).Visible = True
        Else
            xToolbar.Buttons(1).Visible = False
        End If
        
        '--Modificar
        If Rst("opcion2") = -1 Then
            xToolbar.Buttons(2).Visible = True
        Else
            xToolbar.Buttons(2).Visible = False
        End If
        
        '--Eliminar
        If Rst("opcion3") = -1 Then
            xToolbar.Buttons(3).Visible = True
        Else
            xToolbar.Buttons(3).Visible = False
        End If

        If Rst("opcion1") = 0 And Rst("opcion2") = 0 Then
            xToolbar.Buttons(4).Visible = False
            xToolbar.Buttons(5).Visible = False
            xToolbar.Buttons(6).Visible = False
            xToolbar.Buttons(7).Visible = False
        End If

        If Rst("opcion3") = -1 Then
            xToolbar.Buttons(7).Visible = True
        Else
            xToolbar.Buttons(7).Visible = False
        End If
    Else
        xToolbar.Buttons(1).Visible = False '--Nuevo
        xToolbar.Buttons(2).Visible = False '--Modificar
        xToolbar.Buttons(3).Visible = False '--Eliminar
        xToolbar.Buttons(4).Visible = False '--Separado
        xToolbar.Buttons(5).Visible = False '--Grabar
        xToolbar.Buttons(6).Visible = False '--Cancelar
        xToolbar.Buttons(7).Visible = False '--Sepaador
    End If
    
    '--Si usuario es administrador
    If xIdUsuario = 1 And xidMenu = 101 Then
        xToolbar.Buttons(2).Visible = True '--Modificar
        xToolbar.Buttons(4).Visible = True '--Separado
        xToolbar.Buttons(5).Visible = True '--Grabar
        xToolbar.Buttons(6).Visible = True '--Cancelar
    End If
    
    
    Set Rst = Nothing
    Exit Sub

LaCague:
    Err.Clear
    Set Rst = Nothing
    
End Sub

Sub GrabarOperacionCtaCteDocRef(idlib As Integer, IdOperacion As Double, NumDocRef As String, IdCli As Integer, _
                                IdtipDoc As Integer, NumDoc As String, FchEmi As String, idmon As Integer, ImpTC As Double, _
                                DebeSol As Double, HaberSol As Double, DebeDol As Double, HaberDol As Double, NumReg As String, _
                                xCon As ADODB.Connection, Optional xCorr As Integer, Optional Glosa As String, Optional xIdTipDocRef As Integer, _
                                Optional xIdDocRef As Double, Optional xTipo As Integer, Optional xIdCuenta As Double)
    
    Dim Rst As New ADODB.Recordset
    
    '--27/06/11 Johan Castro
    '--Cambiar tipo de dato a parametro xIdDocRef antes Integer ahora Double



    If xCorr = 0 Then
        xCon.Execute "DELETE * FROM var_analisisctacte WHERE idlib = " & idlib & " AND idope = " & IdOperacion & ""
    Else
        xCon.Execute "DELETE * FROM var_analisisctacte WHERE idlib = " & idlib & " AND idope = " & IdOperacion & " AND corr = " & xCorr & ""
    End If
    
    
    
    If NulosC(NumDocRef) = "" Then
        Dim xNumDocRef
        xNumDocRef = "."    ' CDate('27/2/2009') AS Expr9,
        xCon.Execute "INSERT INTO var_analisisctacte ( idlib, idope, idtipdocref, idnumdocref, numdocref, idcli, idtipdoc, numdoc, fchemi, idmon, imptc, " _
            & " debesol, habersol, debedol, haberdol, numreg, corr, glosa, idtipo) " _
            & " SELECT " & idlib & " AS Expr1, " & IdOperacion & " AS Expr2, " & xIdTipDocRef & " AS Expr3, " & xIdDocRef & " AS Expr4, '" & NulosC(NumDocRef) & "' AS Expr5, " & IdCli & " AS Expr6, " & IdtipDoc & " AS Expr7, " _
            & " '" & NumDoc & "' AS Expr8, cdate('" & FchEmi & "') AS Expr9,  " & idmon & " AS Expr10, " & ImpTC & " AS Expr11, " & DebeSol & " AS Expr12, " & HaberSol & " AS Expr13, " _
            & " " & DebeDol & " AS Expr14, " & HaberDol & " AS Expr15, '" & NumReg & "' as Expr16, " & xCorr & " as Expr17, '" & Glosa & "' AS Expr18, " & xTipo & " as Expr19 " _
            & " "
    Else
        xCon.Execute "INSERT INTO var_analisisctacte ( idlib, idope, idtipdocref, idnumdocref, numdocref, idcli, idtipdoc, numdoc, fchemi, idmon, imptc, " _
            & " debesol, habersol, debedol, haberdol, numreg, corr, glosa, idtipo ) " _
            & " SELECT " & idlib & " AS Expr1, " & IdOperacion & " AS Expr2, " & xIdTipDocRef & " AS Expr3, " & xIdDocRef & " AS Expr4, '" & NulosC(NumDocRef) & "' AS Expr5, " & IdCli & " AS Expr6, " & IdtipDoc & " AS Expr7, " _
            & " '" & NumDoc & "' AS Expr8, cdate('" & FchEmi & "') AS Expr9,  " & idmon & " AS Expr10, " & ImpTC & " AS Expr11, " & DebeSol & " AS Expr12, " & HaberSol & " AS Expr13, " _
            & " " & DebeDol & " AS Expr14, " & HaberDol & " AS Expr15, '" & NumReg & "' as Expr16, " & xCorr & " as Expr17, '" & Glosa & "' AS Expr18, " & xTipo & " as Expr19 " _
            & " "
    End If
End Sub

Function SaldoActual(IdProducto As Double, FchInicio As String, FchFinal As String, xCon As ADODB.Connection, Optional xTipo As Integer = 0) As Double
     '===================================================================================================
    'Creado : xx/xx/10 Por: Enrique Pollongo
    'Propósito: Muestra el stock actual del item
    '
    'Entradas:  IdProducto= Codigo del item
    '           FchInicio= Fecha Inicial
    '           FchFinal=  Fecha Final
    '           xCon= Conexion a la base de datos
    '           xTipo =Tipo de consulta. 1 = Entradas; 2= Salida; 0=Todos(x Defecto)
    '
    'Resultados: Stock Actual o Stock en un intervalo de periodo
    '
    'Nota:       1.- Buscar todos los ingresos + Buscar todas las salidas - Buscar el stock incial
    '
    'Modificado: 09/11/10 Por: Johan Castro
    '           Agregar notas de credito por devolucion (entradas)
    '           10/02/11 Por Johan Castro
    '           Agregar parametro xTipo que indica el tipo de consulta
    '           15/02/11 Por Johan Castro
    '           Modificar el envio del valor a la funcion segun el parametro xTipo
    'Modificado: 06/05/11 Johan Castro
    '           Cambiar tipo de dato a parametro IdProducto a Double
    '           En Ingreso de compras filtrar documentos distintos a notas de credito
    '           En Salidas agregar documentos de compras solo con notas de credito
    'Modificado: 21/01/12 Johan Castro
    '           Agregar Variable xSQLFiltroPS para mostrar solo materia prima a partir del 2012 en adelante en Produccion Insumos
    '           No mostrar registros que esten con el siguiente estado(1=Pendiente, 4=Rechazado y 5=Anulado)
    '           Agregar filtro en Almacen Ingreso - Salida "AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 "
    '           Agregar filtro en Produccion Insumos, "AND pro_producciondet.estado Not In (1,4,5) AND pro_producciondetins.canutil<>0 AND alm_inventario.tippro=3"
    '           Agregar filtro en Produccion Producto Terminado, "AND pro_producciondet.estado Not In (1,4,5)"
    'Nota:      Descripcion de los estados (ver tabla mae_estados)
    '           1=Pendiente, 2=Procesado, 3=Aprobado,4=Rechazado,5=Anulado
    '===================================================================================================

    Dim Rst As New ADODB.Recordset
    Dim xCad As String
    Dim xTotEnt, xTotSal As Double
    Dim A As Integer
    Dim xSQLFiltroPS As String '--Util para aplicar un filtro adicional que mostrará solo materia prima en sentencia de "produccion insumos salida" esto a partir del 2012 en adelante
    
    If NulosN(AnoTra) = 0 Then CargaDatosEmpresa xCon, "", ""
    If NulosN(AnoTra) >= 2012 Then
        xSQLFiltroPS = " AND alm_inventario.tippro=3  "
    End If
    
    ' CARGAMOS TODAS LAS ENTRADAS

    '--CArgar solo cuando xTipo sea todos o Entradas
    If xTipo = 0 Or xTipo = 1 Then
        xCad = "SELECT Sum([canpro]) AS total " _
            + vbCr + " FROM alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem " _
            + vbCr + " WHERE (((com_comprasdet.iditem)=" & IdProducto & ") " _
                    & " AND ((com_compras.fchdoc)>=CDate('" & FchInicio & "') And (com_compras.fchdoc)<=CDate('" & FchFinal & "')) AND ((com_compras.tipcom)=1)) and com_compras.tipdoc <> 7 " _
            + vbCr + " Union " _
            + vbCr + " SELECT Sum([cantidad]) AS total " _
            + vbCr + " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id " _
            + vbCr + " WHERE (((alm_ingresodet.iditem)=" & IdProducto & ") AND ((alm_ingreso.fching)>=CDate('" & FchInicio & "') And (alm_ingreso.fching)<=CDate('" & FchFinal & "')) AND ((alm_ingreso.tipmov)=-1)) " _
                    & " AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " _
            + vbCr + " Union " _
            + vbCr + " SELECT Sum([cantidad]) AS total " _
            + vbCr + " FROM pro_produccion LEFT JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
            + vbCr + " WHERE (((pro_producciondet.iditem)=" & IdProducto & ") AND ((pro_produccion.dia)>=CDate('" & FchInicio & "') And (pro_produccion.dia)<=CDate('" & FchFinal & "'))) " _
                    & " AND pro_producciondet.estado Not In (1,4,5) " _
            + vbCr + " Union " _
            + vbCr + " SELECT Sum(vta_ventasdet.canpro) AS total " _
            + vbCr + " FROM (mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta " _
            + vbCr + " WHERE (((vta_ventasdet.iditem)=" & IdProducto & ") AND ((vta_ventas.fchdoc)>=CDate('" & FchInicio & "') And (vta_ventas.fchdoc)<=CDate('" & FchFinal & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref)<>0) AND ((vta_ventas.idmotnotcre)=4)) "
        
        '--Cargar rst
        RST_Busq Rst, xCad, xCon
        If Rst.RecordCount <> 0 Then
            xTotEnt = 0
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                xTotEnt = xTotEnt + NulosN(Rst("total"))
                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next A
        End If
    
    End If
    Set Rst = Nothing
    '--CArgar solo cuando xTipo sea todos o Entradas
    If xTipo = 0 Or xTipo = 2 Then
        'CARGAMOS TODAS LAS SALIDAS
        xCad = " SELECT Sum(vta_ventasdet.canpro) AS total " _
            + vbCr + " FROM (mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN (vta_ventasdet LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta " _
            + vbCr + " WHERE (((vta_ventasdet.iditem)=" & IdProducto & ") AND ((vta_ventas.fchdoc)>=CDate('" & FchInicio & "') " _
                        & " And (vta_ventas.fchdoc)<=CDate('" & FchFinal & "')) AND ((vta_ventas.oriitem)=1) AND ((vta_ventas.iddocref)=0 Or (vta_ventas.iddocref) Is Null))" _
            + vbCr + " Union " _
            + vbCr + " SELECT Sum(alm_ingresodet.cantidad) AS total  " _
            + vbCr + " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario  ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id " _
            + vbCr + " WHERE (((alm_ingresodet.iditem)=" & IdProducto & ") AND ((alm_ingreso.fching)>=CDate('" & FchInicio & "') And (alm_ingreso.fching)<=CDate('" & FchFinal & "')) AND ((alm_ingreso.tipmov)=0)) " _
                        & " AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " _
            + vbCr + " Union " _
            + vbCr + " SELECT Sum(vta_guiadet.canpro) AS total " _
            + vbCr + " FROM vta_guia LEFT JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui " _
            + vbCr + " WHERE (((vta_guiadet.iditem)=" & IdProducto & ") " _
                        & " AND ((vta_guia.fecgiro)>=CDate('" & FchInicio & "') And (vta_guia.fecgiro)<=CDate('" & FchFinal & "')))" _
            + vbCr + " UNION " _
            + vbCr + " SELECT Sum(pro_producciondetins.canutil) AS total " _
            + vbCr + " FROM ((pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) LEFT JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id " _
            + vbCr + " WHERE (((pro_producciondetins.iditem)=" & IdProducto & ") AND ((pro_produccion.dia)>=CDate('" & FchInicio & "') And (pro_produccion.dia)<=CDate('" & FchFinal & "'))) " _
                        & " AND pro_producciondet.estado Not In (1,4,5) AND pro_producciondetins.canutil<>0 " & xSQLFiltroPS _
            + vbCr + " Union " _
            + vbCr + " SELECT Sum(com_comprasdet.canpro) AS total " _
            + vbCr + " FROM alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem " _
            + vbCr + " WHERE (((com_comprasdet.iditem)=" & IdProducto & ") " _
                        & " AND ((com_compras.fchdoc)>=CDate('" & FchInicio & "') And (com_compras.fchdoc)<=CDate('" & FchFinal & "')) AND ((com_compras.tipcom)=1)) and com_compras.tipdoc = 7 " _

        RST_Busq Rst, xCad, xCon
        If Rst.RecordCount <> 0 Then
            xTotSal = 0
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                xTotSal = xTotSal + NulosN(Rst("total"))
                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next A
        End If
    End If
    
    Set Rst = Nothing
    
    Dim xSalIni As Double
    
    If xTipo = 0 Then
        xSalIni = NulosN(Busca_Codigo(IdProducto, "id", "stckini", "alm_inventario", "N", xCon))
        
        SaldoActual = (xSalIni + NulosN(xTotEnt)) - NulosN(xTotSal)
        
    ElseIf xTipo = 1 Then '--Solo entradas
        SaldoActual = NulosN(xTotEnt)
        
    ElseIf xTipo = 2 Then '--Solo Salidas
        SaldoActual = NulosN(xTotSal)
    End If
End Function


Function EsDemo() As Boolean
     '===================================================================================================
    'Creado : 12/09/11 Por: Enrique Pollongo
    'Propósito: Averiguar si la version del seven tiene clave de distribucion o solo es demostracion
    '           True = Es Demostracion;
    '           False=Se registro con clave
    'Resultados: Devuelve verdadero si es demostracion o falso
    '
    '===================================================================================================
    Dim xNumSerDis As String
    Dim xRst As New ADODB.Recordset
    Dim AP_RUTABD As String   ' ruta de la base de datos de enlace
    Dim AP_RUTASY As String   ' ruta de la base de datos de enlace
    Dim xFun As New eps_librerias.FuncionesData
    Dim xSql As String
    Dim xCon2 As New ADODB.Connection
    
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")         ' LEEMOS LA RUTA DE LA BASE DE DATOS
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")          ' LEEMOS LA RUTA DEL SISTEMA
    
    xFun.F_BASEDATOS = AP_RUTABD + "data.mdb"                                           ' PASAMOS LA RUTA DE LA BASE DE DATOS PARA ABRIR LA CONECCION
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"                                       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABJO DE LA BASE DE DATOS
    xFun.F_PASSWORD = Eps_Pass                                                          ' PASAMOS EL PASWORD DE LA BASE DE DATOS
    xFun.F_USUARIO = Eps_User                                                           ' PASAMOS EL USUARIO DE LA BASE DE DATOS
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"                                        ' PASAMOS EL NOMBRE DEL PROVEEDORE DE DATOS PARA ADO 2.5
    
    Set xCon2 = xFun.AbrirConeccion                                                      ' ABRIMOS LA CONECCION DE DATOS
    Set xFun = Nothing
    
    xNumSerDis = Trim(Str(LeerNumeroDisco("c:")))
    
    xSql = "SELECT mae_pc.* From mae_pc WHERE (((mae_pc.serdis)='" & xNumSerDis & "'))"

    RST_Busq xRst, xSql, xCon2
    
    If xRst.RecordCount <> 0 Then
        If xRst("registro") = 0 Then
            EsDemo = True
        Else
            EsDemo = False
        End If
    End If
    
    Set xCon2 = Nothing
End Function

Function PuedeAgregarRegistro(NombreModulo As String, xCon As ADODB.Connection) As Boolean
     '===================================================================================================
    'Creado : 12/09/11 Por: Enrique Pollongo
    '
    'Propósito: Averigurar si la version del seven es una demostracion, permitira registrar una
    '           cantidad determinada de registros por modulo
    '           Si Version es Demostracion(True: Permitira registrar Operacion
    '                                      False: No Permite registrar Operacion)
    'Entradas:  NombreModulo= Indica el nombre del módulo
    '           xCon= Conexion
    '
    'Resultados: Si version es demostracion permitira registrar operacion.
    '
    '===================================================================================================
    
    Dim xRst As New ADODB.Recordset
    Dim xSql As String
    
    If EsDemo() = True Then
        If NombreModulo = "COMPRAS" Then
            xSql = "SELECT * FROM com_compras"
        End If
        
        If NombreModulo = "VENTAS" Then
            xSql = "SELECT * FROM vta_ventas"
        End If
        
        If NombreModulo = "PROVISIONES" Then
            xSql = "SELECT * FROM con_proviciones"
        End If
        
        RST_Busq xRst, xSql, xCon
        
        If xRst.RecordCount >= 50 Then
            PuedeAgregarRegistro = False
        Else
            PuedeAgregarRegistro = True
        End If
        
    Else
        PuedeAgregarRegistro = True
    End If
    Set xRst = Nothing
    
End Function

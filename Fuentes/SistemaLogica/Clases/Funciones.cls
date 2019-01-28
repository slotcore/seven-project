VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Funciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function GetMachineName() As String
    GetMachineName = mGetMachineName()
End Function

Public Function GenerarSQLInRst(Rst As ADODB.Recordset, _
                                nDesc As String, _
                                nCampo As String, _
                                Optional ArmaSQL As Boolean = True, _
                                Optional nTipoIn As String = "IN", _
                                Optional fEsNumero As Boolean = True) As String
                            
    GenerarSQLInRst = mGenerarSQLInRst(Rst, nDesc, nCampo, ArmaSQL, nTipoIn, fEsNumero)
    
End Function

Public Function GenerarSQLInGrid(GRID As Object, _
                                mCol As Integer, _
                                nCampo As String, _
                                Optional ArmaSQL As Boolean = True, _
                                Optional nTipoIn As String = "IN", _
                                Optional fEsNumero As Boolean = True) As String
                                
    GenerarSQLInGrid = mGenerarSQLInGrid(GRID, mCol, _
                                nCampo, ArmaSQL, _
                                nTipoIn, fEsNumero)
End Function

Sub PrintText(Printer As Object, text As String, Alignment As AlignmentConstants)
    Select Case Alignment
    Case vbCenter
        If Printer.CurrentX - (Printer.TextWidth(text) / 2) > 0 Then
            Printer.CurrentX = Printer.CurrentX - (Printer.TextWidth(text) / 2)
        Else
            Printer.CurrentX = 0
        End If
    Case vbLeftJustify
        Printer.CurrentX = Printer.CurrentX
    Case vbRightJustify
        If Printer.CurrentX - Printer.TextWidth(text) > 0 Then
            Printer.CurrentX = Printer.CurrentX - Printer.TextWidth(text)
        Else
            Printer.CurrentX = 0
        End If
    End Select
    Printer.Print text
End Sub

Sub PintarGrid(Grilla As Object, NumberColumn As Long, ColorResultPos As String, ColorResultNeg As String)
    Dim A As Long
    
    With Grilla
        For A = Grilla.FixedRows To .Rows - 1
            .Select A, NumberColumn
            If NulosN(.TextMatrix(A, NumberColumn)) >= 0 Then
                .CellForeColor = ColorResultPos
            Else
                .CellForeColor = ColorResultNeg
            End If
        Next A
    End With
End Sub

Function CentrarFrame(mFrame As Object, mForm As Object) As Boolean
    CentrarFrame = mCentrarFrame(mFrame, mForm)
End Function

Function NuloNumeric(value As Variant) As Double
    NuloNumeric = mNuloNumeric(value)
End Function

Function NuloString(value As Variant) As String
    NuloString = mNuloString(value)
End Function

Function StringFormat(value As Variant, formato As String) As String
    StringFormat = mStringFormat(value, formato)
End Function

Function ConvertirFechaANumero(value As Date) As Double
    ConvertirFechaANumero = mConvertirFechaANumero(value)
End Function

Function ConvertirHoraANumero(value As Date) As Double
    ConvertirHoraANumero = mConvertirHoraANumero(value)
End Function

Function ConvertirNumeroAFecha(value As Double) As Date
    ConvertirNumeroAFecha = mConvertirNumeroAFecha(value)
End Function

Function HallaInsumosParteProduccion(IdParteProduccion As Long, Optional mConexion As ADODB.Connection) As ADODB.Recordset
    Set HallaInsumosParteProduccion = mHallaInsumosParteProduccion(IdParteProduccion, mConexion)
End Function

Public Function HallaNumeroDocumento(nombreTabla As String, condicion1 As String, _
                                            campo1 As String, _
                                            Optional mConexion As ADODB.Connection, _
                                            Optional condicion2 As String = "", _
                                            Optional campo2 As String = "", _
                                            Optional condicion3 As String = "", _
                                            Optional campo3 As String = "", _
                                            Optional campoOrden As String = "numdoc", _
                                            Optional FORMATO_ As String = "0000000000") As String
    HallaNumeroDocumento = mHallaNumeroDocumento(nombreTabla, condicion1, campo1, mConexion, condicion2, campo2, condicion3, campo3, campoOrden, FORMATO_)
End Function

Function KeyValue(key As String, Optional mConexion As ADODB.Connection) As String
    KeyValue = mKeyValue(key, mConexion)
End Function

Public Function KeyValue_(key As String, DataSource As String) As String
    Dim xRs As New ADODB.Recordset
    Dim xNum As Double
    Dim cSQL As String
    Dim nSQL As String
    Dim database As New SistemaData.EDataBase
    
    database.OpenConnection DataSource
    database.CommandText = "SELECT TOP 1 * " _
                    + vbCr + "FROM mae_configuracion " _
                    + vbCr + "WHERE ((llave) = '" & key & "')"
    
    Set xRs = database.GetRecordset()
    
    If xRs.State = 0 Then Exit Function
    If xRs.RecordCount = 0 Then
        MsgBox "No se encontro la llave especificada: " & key, vbInformation + vbOKOnly + vbDefaultButton1, "Funciones - KeyValue"
        KeyValue_ = ""
    Else
        KeyValue_ = NulosC(xRs("valor"))
    End If
    
    Set xRs = Nothing
    Set database = Nothing
End Function

Sub MostrarMensajeError(Mensaje As String, Titulo As String, Optional Pila As String, Optional mErrorNumber)
    mMostrarMensajeError Mensaje, Titulo, Pila, mErrorNumber
End Sub

Function BuscaCodigoTabla(valorBusca As Variant, campoCodigo As String, campoDescripcion As String, Tabla As String, Tipo As String, Optional mConexion As ADODB.Connection)
     BuscaCodigoTabla = mBuscaCodigoTabla(valorBusca, campoCodigo, campoDescripcion, Tabla, Tipo, mConexion)
End Function
    
Function FechaInicioMovimientos(IdAlmacen As Long, Optional mConexion As ADODB.Connection) As Date
    FechaInicioMovimientos = mFechaInicioMovimientos(IdAlmacen, mConexion)
End Function

Public Function FechaInicioMovimientos_(IdAlmacen As Long, DataSource As String) As Date
    Dim xRs As New ADODB.Recordset
    Dim database As New SistemaData.EDataBase
    
On Error GoTo BloqueError
    database.OpenConnection DataSource
    database.CommandText = "SELECT alm_inventarioinicial.fechavigencia " _
        + vbCr + "FROM alm_inventarioinicial " _
        + vbCr + "WHERE (((alm_inventarioinicial.idalm)=" & IdAlmacen & ") AND ((alm_inventarioinicial.idestado)=" & mNuloNumeric(KeyValue_("EstadoAprobadoInventarioInicial", DataSource)) & "))"
    
    Set xRs = database.GetRecordset
    If xRs.State = 0 Then Err.Raise &HFFFFFF01, , "Error en el estado del Recordset"
    If xRs.RecordCount = 0 Then
        FechaInicioMovimientos_ = Empty
        Set database = Nothing
        Exit Function
    End If
    
    FechaInicioMovimientos_ = mConvertirNumeroAFecha(xRs("fechavigencia"))
    Set database = Nothing
    Exit Function
    
BloqueError:
    FechaInicioMovimientos_ = Empty
    Set database = Nothing
    mMostrarMensajeError Err.Description, "[FechaInicioMovimientos] " & Err.Source
End Function

Function SaldoActual(IdItem As Long, IdAlmacen As Long, FchInicio As Date, FchFinal As Date, Optional mConexion As ADODB.Connection, Optional xTipo As Long = 0) As Double
     SaldoActual = mSaldoActual(IdItem, IdAlmacen, FchInicio, FchFinal, mConexion, xTipo)
End Function

Function SaldoActual_(IdItem As Long, IdAlmacen As Long, _
                                FchInicio As Date, _
                                FchFinal As Date, _
                                DataSource As String, _
                                Optional xTipo As Long = 0) As Double
    Dim database As New SistemaData.EDataBase
    Dim Rst As New ADODB.Recordset
    Dim mSQL As String
    Dim xTotEnt As Double
    Dim xTotSal As Double
    Dim A As Long
    
    ' Se abre la conexion a la base de datos
    database.OpenConnection DataSource
    '--CArgar solo cuando xTipo sea todos o Entradas
    If xTipo = 0 Or xTipo = 1 Then
        database.CommandText = "SELECT SUM(M.canpro) AS total " _
            + vbCr + "FROM (" _
            + vbCr + KardexMovimientoSQL_(IdItem, IdAlmacen, FchInicio, FchFinal, DataSource, 1) _
            + vbCr + ") AS M " _
            + vbCr + "WHERE (((M.tipo)<>'II'))"
        '--Cargar rst
        Set Rst = database.GetRecordset
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
        database.ClearParameter
        database.CommandText = "SELECT SUM(M.canpro) AS total " _
            + vbCr + "FROM (" _
            + vbCr + KardexMovimientoSQL_(IdItem, IdAlmacen, FchInicio, FchFinal, DataSource, 2) _
            + vbCr + ") AS M " _
            + vbCr + "WHERE (((M.tipo)<>'II'))"
            
        Set Rst = database.GetRecordset
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
    
    If xTipo = 0 Then
        SaldoActual_ = NulosN(xTotEnt) - NulosN(xTotSal) + SaldoInicial_(IdItem, IdAlmacen, DataSource)
        
    ElseIf xTipo = 1 Then '--Solo entradas
        SaldoActual_ = NulosN(xTotEnt)
        
    ElseIf xTipo = 2 Then '--Solo Salidas
        SaldoActual_ = NulosN(xTotSal)
    End If
End Function

Function ErrorDescriptionDLL(Optional ByVal lLastDLLError As Long) As String
    ErrorDescriptionDLL = mErrorDescriptionDLL(lLastDLLError)
End Function

Public Function MachineName() As String
    MachineName = mMachineName()
End Function

Public Function ExisteDocumento(Tabla As String, Valor As String, _
                                        Optional Conexion As ADODB.Connection, _
                                        Optional Campo As String = "numdoc", _
                                        Optional Valor2 As String, _
                                        Optional campo2 As String = "numser", _
                                        Optional Valor3 As String, _
                                        Optional campo3 As String, _
                                        Optional IdDocumento As Long, _
                                        Optional CampoIdDocumento As String) As Boolean
    ExisteDocumento = mExisteDocumento(Tabla, Valor, _
                                        Conexion, Campo, Valor2, _
                                        campo2, Valor3, campo3, _
                                        IdDocumento, CampoIdDocumento)
End Function

Public Function AlmacenaEn(IdItem As Long, Optional Conexion As ADODB.Connection) As Long
    AlmacenaEn = mAlmacenaEn(IdItem, Conexion)
End Function

Public Function DespachaEn(IdItem As Long, Optional Conexion As ADODB.Connection) As Long
    DespachaEn = mDespachaEn(IdItem, Conexion)
End Function

Public Function CostoInicial(IdItem As Long, IdAlmacen As Long, Optional mConexion As ADODB.Connection) As Double
    CostoInicial = mCostoInicial(IdItem, IdAlmacen, mConexion)
End Function

Public Function CostoActual(IdItem As Long, IdAlmacen As Long, _
                            FchInicio As Date, FchFinal As Date, _
                            Optional mConexion As ADODB.Connection, _
                            Optional xTipo As Long = 0) As Double
    CostoActual = mCostoActual(IdItem, IdAlmacen, FchInicio, FchFinal, mConexion, xTipo)
End Function

Public Function SaldoInicial(IdItem As Long, IdAlmacen As Long, Optional mConexion As ADODB.Connection) As Double
    SaldoInicial = mSaldoInicial(IdItem, IdAlmacen, mConexion)
End Function

Public Function SaldoInicial_(IdItem As Long, IdAlmacen As Long, _
                                DataSource As String) As Double
    Dim database As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones
    Dim mRecord As New ADODB.Recordset
    Dim QUERYA As String
    Dim QUERYB As String
    Dim QUERYC As String
    
On Error GoTo BloqueError
    database.OpenConnection DataSource
    database.CommandText = "SELECT alm_inventarioinicialdet.cantidad " _
            + vbCr + "FROM alm_inventarioinicial INNER JOIN alm_inventarioinicialdet ON alm_inventarioinicial.idinventarioinicial = alm_inventarioinicialdet.idinventarioinicial " _
            + vbCr + "WHERE (((alm_inventarioinicialdet.iditem)=" & IdItem & ") AND ((alm_inventarioinicial.idalm)=" & IdAlmacen & ") AND ((alm_inventarioinicial.idestado)=" & F.NuloNumeric(KeyValue_("EstadoAprobadoInventarioInicial", DataSource)) & "))"
        
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount = 0 Then SaldoInicial_ = 0: Exit Function
    mRecord.MoveFirst
    SaldoInicial_ = F.NuloNumeric(mRecord("cantidad"))
    Set database = Nothing
    Exit Function
    
BloqueError:
    SaldoInicial_ = 0
    mMostrarMensajeError Err.Description, "[SaldoInicial] " & Err.Source
    Set database = Nothing
End Function

Function KardexMovimientoSQL(IdItem As Long, IdAlmacen As Long, xFchIni As Date, xFchFin As Date, Optional mConexion As ADODB.Connection, Optional xTipo As Long = 0) As String
    KardexMovimientoSQL = mKardexMovimientoSQL(IdItem, IdAlmacen, xFchIni, xFchFin, mConexion, xTipo)
End Function

Function KardexMovimientoSQL_(IdItem As Long, _
                                IdAlmacen As Long, _
                                xFchIni As Date, _
                                xFchFin As Date, _
                                DataSource As String, _
                                Optional xTipo As Long = 0, _
                                Optional MuestraInvIni As Boolean = True) As String
    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    'GR = GUIAS DE REMISION
    'PS =

    Dim xCadSQL As String
    Dim xSQLFiltroPS As String '--Util para aplicar un filtro adicional que mostrará solo materia prima en sentencia de "produccion insumos salida"
    Dim F As New SistemaLogica.Funciones
    Dim mFechaInicioMovimientos As Date
    Dim mSQLFechaMov As String
    Dim mSQLIdItem As String
    Dim mSQLIdAlmacen As String
    Dim mSQLInvInicial As String
    Dim mSQLExclude As String
    Dim IdTipoInvInicial As Long

    mFechaInicioMovimientos = FechaInicioMovimientos_(IdAlmacen, DataSource)
    IdTipoInvInicial = F.NuloNumeric(KeyValue_("IdDocumentoInventarioInicial", DataSource))

    ' PREPARAMOS LA SELECT PARA ARMAR EL KARDEX
    If xTipo = 0 Or xTipo = 1 Then
        ' Ingresos de Almacen
        If IdItem = 0 Then
            mSQLIdItem = ""
        Else
            mSQLIdItem = "AND ((alm_ingresodet.iditem)=" & IdItem & ") "
        End If
        If IdAlmacen = 0 Then
            mSQLIdAlmacen = ""
        Else
            mSQLIdAlmacen = "AND ((alm_ingreso.idalm)=" & IdAlmacen & ") "
        End If
        If Not IsDate(mFechaInicioMovimientos) Or mFechaInicioMovimientos = Empty Then
            mSQLFechaMov = ""
        Else
            mSQLFechaMov = "AND (alm_ingreso.fching)>=CDate('" & mFechaInicioMovimientos & "') "
        End If
        If MuestraInvIni Then
            mSQLInvInicial = ""
        Else
            mSQLInvInicial = "AND ((alm_ingreso.idtipdocref)<>" & F.NuloNumeric(KeyValue_("IdDocumentoInventarioInicial", DataSource)) & ") "
        End If
        
        xCadSQL = "SELECT alm_ingreso.id, alm_ingresodet.idmovdet, alm_ingresodet.iditem, alm_inventario.codpro, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, alm_ingreso.numser, alm_ingreso.numdoc, alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, IIf(alm_ingreso.idtipdocref=121,'II','AI') AS tipo, alm_ingreso.tipmov, alm_ingreso.nombre AS entidad, 0 AS aa, (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos, 'Almacén' & IIf(CStr(numdocumentos)<>'0',' - Compras','') AS modulo, '' AS registro, '' AS ctanum, '' AS ctanom, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, " _
                    & "alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, (con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS costo, con_librocostotemp.costounitariopromedio  " _
            + vbCr + "FROM (alm_ingreso LEFT JOIN ((((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id " _
            + vbCr + "WHERE (((alm_ingreso.fching)>=CDate('" & xFchIni & "') And (alm_ingreso.fching)<=CDate('" & xFchFin & "')) AND ((alm_ingreso.tipmov)=-1)) AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad <> 0 " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov & mSQLInvInicial
        
        ' Ingreso por Compras
        If IdItem = 0 Then
            mSQLIdItem = ""
        Else
            mSQLIdItem = "AND ((com_comprasdet.iditem)=" & IdItem & ") "
        End If
        If IdAlmacen = 0 Then
            mSQLIdAlmacen = ""
        Else
            mSQLIdAlmacen = "AND ((com_compras.idalm)=" & IdAlmacen & ") "
        End If
        If Not IsDate(mFechaInicioMovimientos) Or mFechaInicioMovimientos = Empty Then
            mSQLFechaMov = ""
        Else
            mSQLFechaMov = " AND (com_compras.fchdoc)>=CDate('" & mFechaInicioMovimientos & "')"
        End If
        '**********************************************
        ' Exlusione de compras
        mSQLExclude = mSQLExclude & " AND ((alm_inventario.tippro)<> 5) "
        '**********************************************
        xCadSQL = xCadSQL _
            + vbCr + "UNION ALL " _
            + vbCr + "SELECT com_compras.id, 0 As idmovdet, com_comprasdet.iditem, alm_inventario.codpro, alm_inventario.descripcion, com_compras.fchdoc, com_compras.numser, com_compras.numdoc, com_comprasdet.canpro, IIf([com_compras]![idmon]=2,[com_comprasdet]![preuni]*[con_tc]![impcom],[com_comprasdet]![preuni]) AS preuni, mae_documento.abrev AS descdoc, 'C' AS Tipo, -1 AS tipmov, mae_prov.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, 'Compras' AS modulo, com_compras.numreg AS registro, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS desope, '' AS horini, '' AS horfin, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, com_comprasdet.imptot AS costo, com_comprasdet.preuni AS costounitariopromedio " _
            + vbCr + "FROM (((alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((com_compras.fchdoc)>=CDate('" & xFchIni & "') And (com_compras.fchdoc)<=CDate('" & xFchFin & "')) AND ((com_compras.tipcom)=1)) " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov & mSQLExclude
    
    End If
    
    If xTipo = 0 Or xTipo = 2 Then
        If xTipo = 0 Then xCadSQL = xCadSQL + vbCr + "UNION ALL "
        ' Salidas de Almacen
        If IdItem = 0 Then
            mSQLIdItem = ""
        Else
            mSQLIdItem = "AND ((alm_ingresodet.iditem)=" & IdItem & ") "
        End If
        If IdAlmacen = 0 Then
            mSQLIdAlmacen = ""
        Else
            mSQLIdAlmacen = "AND ((alm_ingreso.idalm)=" & IdAlmacen & ") "
        End If
        If Not IsDate(mFechaInicioMovimientos) Or mFechaInicioMovimientos = Empty Then
            mSQLFechaMov = ""
        Else
            mSQLFechaMov = " AND (alm_ingreso.fching)>=CDate('" & mFechaInicioMovimientos & "')"
        End If
        xCadSQL = xCadSQL _
            + vbCr + "SELECT alm_ingreso.id, alm_ingresodet.idmovdet, alm_ingresodet.iditem, alm_inventario.codpro, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, alm_ingreso.numser, alm_ingreso.numdoc, alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AS' AS tipo, alm_ingreso.tipmov, alm_ingreso.nombre AS entidad, 0 AS aa, (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos, 'Almacén' & IIf(CStr(numdocumentos)<>'0',' - Compras','') AS modulo, '' AS registro, '' AS ctanum, '' AS ctanom, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, " _
                    & "alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, (con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS costo, con_librocostotemp.costounitariopromedio  " _
            + vbCr + "FROM (alm_ingreso LEFT JOIN ((((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id " _
            + vbCr + "WHERE (((alm_ingreso.fching)>=CDate('" & xFchIni & "') And (alm_ingreso.fching)<=CDate('" & xFchFin & "')) AND ((alm_ingreso.tipmov)=0)) AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov
        
    End If
    
    KardexMovimientoSQL_ = xCadSQL
End Function

Public Function GetItemCollection(mCollection As Collection, mKey As String) As Object
    Set GetItemCollection = mGetItemCollection(mCollection, mKey)
End Function

Public Sub ExportarExcelRecordSet(Record As ADODB.Recordset)
    mExportarExcelRecordSet Record
End Sub

Public Function MesCerradoOpcion(IdMes As Integer, IdOpcion As Long, Optional mConexion As ADODB.Connection) As Boolean
    MesCerradoOpcion = mMesCerradoOpcion(IdMes, IdOpcion, mConexion)
End Function

Public Function RetornarMesFecha(Fecha As Date) As Integer
    RetornarMesFecha = mRetornarMesFecha(Fecha)
End Function

Public Function RetornarAnhoFecha(Fecha As Date) As Integer
    RetornarAnhoFecha = mRetornarAnhoFecha(Fecha)
End Function

Public Function RetornarPrimerDiaMes(Fecha As Date) As Date
    RetornarPrimerDiaMes = mRetornarPrimerDiaMes(Fecha)
End Function

Public Function RetornarUltimoDiaMes(Fecha As Date) As Date
    RetornarUltimoDiaMes = mRetornarUltimoDiaMes(Fecha)
End Function

Public Sub LlenarCombo(cb, Rst As ADODB.Recordset, ValueColumn As String, BoundColumn As String)
    mLlenarCombo cb, Rst, ValueColumn, BoundColumn
End Sub

Public Function CompararConCriterio(Numero1 As Double, Numero2 As Double, Optional NumeroDecimales As Integer = 4) As Boolean
    CompararConCriterio = mCompararConCriterio(Numero1, Numero2, NumeroDecimales)
End Function

Public Function SQL_MovHistoricoTotalizado(IdAlmacen As Long, _
                                    FechaFin As Date, _
                                    Optional Cad_IdItem_In As String = "", _
                                    Optional mConexion As ADODB.Connection, _
                                    Optional FiltraAlmacen As Boolean = False) As String
                                    
    SQL_MovHistoricoTotalizado = mSQL_MovHistoricoTotalizado(IdAlmacen, FechaFin, Cad_IdItem_In, _
                                                        mConexion, FiltraAlmacen)
End Function

Public Function SQL_MovTotalizado(Cad_IdItem_In As String, _
                                    IdAlmacen As Long, _
                                    FechaInicio As Date, _
                                    FechaFin As Date, _
                                    Optional mConexion As ADODB.Connection, _
                                    Optional All_Items As Boolean = False, _
                                    Optional FiltraAlmacen As Boolean = False) As String
                                    
    SQL_MovTotalizado = mSQL_MovTotalizado(Cad_IdItem_In, IdAlmacen, FechaInicio, _
                                    FechaFin, mConexion, All_Items, FiltraAlmacen)
End Function


Public Function GeneraRstSQL(CadSQL As String, _
                                Optional Cnn As ADODB.Connection) As ADODB.Recordset
    Set GeneraRstSQL = mGeneraRstSQL(CadSQL, Cnn)
End Function

Public Function GeneraRstSQL_(CadSQL As String, _
                                DataSource As String) As ADODB.Recordset
    Dim mDataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
    mDataBase.OpenConnection DataSource
    mDataBase.ClearParameter
    mDataBase.CommandText = CadSQL
    Set mRecord = mDataBase.GetRecordset
    Set mDataBase = Nothing
    
    Set GeneraRstSQL_ = mRecord
End Function

Function SQL_MovDetallado(Optional Cad_IdItem_In As String = "", _
                                Optional IdAlmacen As Long = 0, _
                                Optional FechaInicio As Date = Empty, _
                                Optional FechaFin As Date = Empty, _
                                Optional mConexion As ADODB.Connection, _
                                Optional FiltraTipoMovimiento As Integer = 0, _
                                Optional MuestraInvIni As Boolean = True, _
                                Optional MuestraCompras As Boolean = False, _
                                Optional IdMovimientoDetalle As Long = 0, _
                                Optional FiltraAlmacen As Boolean = False) As String

    SQL_MovDetallado = mSQL_MovDetallado(Cad_IdItem_In, IdAlmacen, FechaInicio, _
                                    FechaFin, mConexion, FiltraTipoMovimiento, _
                                    MuestraInvIni, MuestraCompras, IdMovimientoDetalle, FiltraAlmacen)
End Function

Function ConvertirHoraADecimal(Hora As Date) As Double
    ConvertirHoraADecimal = Hour(Hora) + (Minute(Hora) / 60) + (Second(Hora) / 3600)
End Function

Attribute VB_Name = "FuncionesModule"
Option Explicit

''' <summary>
''' Halla el importe unitario promedio del item especificado y el de sus movimientos
''' </summary>
Public Function CosteaItem(ByRef LErrorCosto As ContabilidadEntidad.LEErrorCosto, _
                            IdAlmacenProceso As Long, _
                            IdItem As Long, _
                            IdAlmacen As Long, _
                            FechaInicio As Date, _
                            FechaFin As Date, _
                            Optional mConexion As ADODB.Connection) As Boolean
                            
    Dim F As New SistemaLogica.Funciones
    Dim mLMovimientoItem As New ContabilidadEntidad.LEMovimientoItem
    Dim mCantidadAcumulada As Double
    Dim mCantidadAcumuladaEntrada As Double
    Dim mCantidadAcumuladaSalida As Double
    Dim mCostoAcumulado As Double
    Dim mCostoUnitarioPromedio As Double
    Dim mCostoUnitarioMovimiento As Double
    Dim mCostoMovimiento As Double
    Dim mErrorCosto As New ContabilidadEntidad.EErrorCosto
    
'    ' ******************************************
'    ' Usado para pruebas
'    If IdItem = 2080 Then
'        MsgBox "Entro"
'    End If
'    ' ******************************************
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set mLMovimientoItem.Conexion = mConexion
    mLMovimientoItem.Fetch IdItem, IdAlmacen, FechaInicio, FechaFin
    If mLMovimientoItem.Count <= 0 Then
        mErrorCosto.CodigoError = "Err001"
        mErrorCosto.DetalleError = "Inconsistencia en cantidad, el Item no cuenta con movimientos a la fecha." _
                            + vbCr + " Codigo de Item: " & F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "codpro", "alm_inventario", "N", mConexion)) _
                            + vbCr + " Item: " & F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "descripcion", "alm_inventario", "N", mConexion)) _
                            + vbCr + " Fecha: " & F.NuloString(FechaFin)
        mErrorCosto.Origen = "CosteaItem"
        mErrorCosto.SolucionError = ""
        ' Se agrega a la lista de errores
        LErrorCosto.Add mErrorCosto
        CosteaItem = False
        Exit Function
        
'        Err.Raise &HFFFFFF01, , "Inconsistencia en cantidad, el Item no cuenta con movimientos a la fecha." _
'                            + vbCr + " Codigo de Item: " & F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "codpro", "alm_inventario", "N", mConexion)) _
'                            + vbCr + " Item: " & F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "descripcion", "alm_inventario", "N", mConexion)) _
'                            + vbCr + " Fecha: " & F.NuloString(FechaFin)
    End If
    
    mCantidadAcumulada = mLMovimientoItem.SaldoInicial
    mCantidadAcumuladaEntrada = mLMovimientoItem.SaldoInicial
    mCostoAcumulado = mLMovimientoItem.CostoInicial
    mCostoUnitarioPromedio = mLMovimientoItem.CostoUnitarioPromedioInicial
    Dim mMovimientoItem As ContabilidadEntidad.EMovimientoItem
    For Each mMovimientoItem In mLMovimientoItem
        ' Se busca errores de busqueda de items
        If mMovimientoItem.IdMovimientoDet = 0 Then
            Set mErrorCosto = New ContabilidadEntidad.EErrorCosto
            mErrorCosto.CodigoError = "Err00-Load"
            mErrorCosto.DetalleError = "Error de data - se ha encontrado un movimiento inconsistente. Consultar con el administrador del sistema " _
                        + vbCr + " Id de Item: " & F.NuloString(IdItem) _
                        + vbCr + " Id de Almacen: " & F.NuloString(IdAlmacen) _
                        + vbCr + " Fecha de inicio de consulta: " & F.NuloString(FechaInicio) _
                        + vbCr + " Fecha de fin de consulta: " & F.NuloString(FechaFin) _
                        + vbCr + " Id almacen de proceso: " & F.NuloString(IdAlmacenProceso)
            mErrorCosto.Origen = "CosteaItem"
            mErrorCosto.SolucionError = "Consultar con el administrador del sistema"
            ' Se agrega a la lista de errores
            LErrorCosto.Add mErrorCosto
            CosteaItem = False
            Exit Function
        End If
        If mMovimientoItem.Cantidad <= 0 Then
            Set mErrorCosto = New ContabilidadEntidad.EErrorCosto
            mErrorCosto.CodigoError = "Err002"
            mErrorCosto.DetalleError = "Movimiento con cantidad igual (o menor) a cero. " _
                        + vbCr + " Codigo de Item: " & mMovimientoItem.CodigoItem _
                        + vbCr + " Item: " & mMovimientoItem.Item _
                        + vbCr + " Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                        + vbCr + " Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                        + vbCr + " Documento: " & mMovimientoItem.NumeroDocumento
            mErrorCosto.Origen = "CosteaItem"
            mErrorCosto.SolucionError = ""
            ' Se agrega a la lista de errores
            LErrorCosto.Add mErrorCosto
            CosteaItem = False
            Exit Function
'            Err.Raise &HFFFFFF01, , "Movimiento con cantidad igual (o menor) a cero. " _
'                        + vbCr + " Codigo de Item: " & mMovimientoItem.CodigoItem _
'                        + vbCr + " Item: " & mMovimientoItem.Item _
'                        + vbCr + " Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
'                        + vbCr + " Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
'                        + vbCr + " Documento: " & mMovimientoItem.NumeroDocumento
        End If
        '**********
        ' INGRESOS
        '**********
        If mMovimientoItem.TipoMovimiento = "I" Then
            mCantidadAcumulada = mCantidadAcumulada + mMovimientoItem.Cantidad
            mCantidadAcumuladaEntrada = mCantidadAcumuladaEntrada + mMovimientoItem.Cantidad
            ' MOVIMIENTO SIN COSTEAR
            If mMovimientoItem.Costo = 0 Then
                mMovimientoItem.CostoUnitarioPromedio = mCostoUnitarioPromedio
                mCostoMovimiento = CosteaMovimientoDetalle(LErrorCosto, IdAlmacenProceso, FechaInicio, mMovimientoItem, mConexion)
                mCostoUnitarioMovimiento = mCostoMovimiento / mMovimientoItem.Cantidad
                                
                ' Validamos el costo unitario del movimiento
                If mCostoUnitarioMovimiento <= 0 Then
                    Set mErrorCosto = New ContabilidadEntidad.EErrorCosto
                    mErrorCosto.CodigoError = "Err003"
                    mErrorCosto.DetalleError = "Costo Unitario de movimiento igual (o menor) a cero. " _
                                + vbCr + " Codigo de Item: " & mMovimientoItem.CodigoItem _
                                + vbCr + " Item: " & mMovimientoItem.Item _
                                + vbCr + " Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                                + vbCr + " Costo Unitario: " & F.NuloString(mCostoUnitarioMovimiento) _
                                + vbCr + " Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                + vbCr + " Documento: " & mMovimientoItem.NumeroDocumento
                    mErrorCosto.Origen = "CosteaItem"
                    mErrorCosto.SolucionError = ""
                    ' Se agrega a la lista de errores
                    LErrorCosto.Add mErrorCosto
                    CosteaItem = False
                    Exit Function
'                    Err.Raise &HFFFFFF01, , "Costo Unitario de movimiento igual (o menor) a cero. " _
'                                + vbCr + " Codigo de Item: " & mMovimientoItem.CodigoItem _
'                                + vbCr + " Item: " & mMovimientoItem.Item _
'                                + vbCr + " Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
'                                + vbCr + " Costo Unitario: " & F.NuloString(mCostoUnitarioMovimiento) _
'                                + vbCr + " Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
'                                + vbCr + " Documento: " & mMovimientoItem.NumeroDocumento
                End If
                
                mCostoAcumulado = mCostoAcumulado + mCostoMovimiento
                mCostoUnitarioPromedio = mCostoAcumulado / mCantidadAcumulada
                ' Costeamos el movimiento
                mMovimientoItem.Costo = mCostoMovimiento
                mMovimientoItem.CostoPrimo = mCostoMovimiento
                mMovimientoItem.CostoUnitario = mCostoUnitarioMovimiento
                mMovimientoItem.CostoUnitarioPromedio = mCostoUnitarioPromedio
                
                GrabaCostoMovimiento IdAlmacenProceso, mMovimientoItem, mConexion
            ' MOVIMIENTO COSTEADO
            Else
                ' Buscamos el importe del movimiento
                mCostoAcumulado = mCostoAcumulado + mMovimientoItem.Costo
                mCostoMovimiento = mMovimientoItem.Costo
                mCostoUnitarioMovimiento = mMovimientoItem.CostoUnitario
                mCostoUnitarioPromedio = mCostoAcumulado / mCantidadAcumulada
                
                ' Se valida que este bien costeado
                If mCostoUnitarioPromedio <> mMovimientoItem.CostoUnitarioPromedio Then ' Si esta mal costeado
                    ' Actualizamos el costo unitario promedio
                    mMovimientoItem.CostoUnitarioPromedio = mCostoUnitarioPromedio
                    
                    GrabaCostoMovimiento IdAlmacenProceso, mMovimientoItem, mConexion
                End If
            End If
        '**********
        ' SALIDAS
        '**********
        Else
            mCostoMovimiento = mCostoUnitarioPromedio * mMovimientoItem.Cantidad
            mCantidadAcumulada = mCantidadAcumulada - mMovimientoItem.Cantidad
            mCantidadAcumuladaSalida = mCantidadAcumuladaSalida + mMovimientoItem.Cantidad
            mCostoUnitarioMovimiento = mCostoUnitarioPromedio
            
            ' Se valida que no hayan saldos negativos
            If CDbl(Format(mCantidadAcumulada, "0.00")) < 0 Then
                Set mErrorCosto = New ContabilidadEntidad.EErrorCosto
                mErrorCosto.CodigoError = "Err004"
                mErrorCosto.DetalleError = "Cantidad acumulada menor a cero. " _
                                        + vbCr + " Codigo de Item: " & mMovimientoItem.CodigoItem _
                                        + vbCr + " Item: " & mMovimientoItem.Item _
                                        + vbCr + " Cantidad: " & Format(mCantidadAcumulada, "0.00") _
                                        + vbCr + " Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                        + vbCr + " Documento: " & mMovimientoItem.NumeroDocumento
                mErrorCosto.Origen = "CosteaItem"
                mErrorCosto.SolucionError = ""
                ' Se agrega a la lista de errores
                LErrorCosto.Add mErrorCosto
                CosteaItem = False
                Exit Function
'                Err.Raise &HFFFFFF01, , "Cantidad acumulada menor a cero. " _
'                                        + vbCr + " Codigo de Item: " & mMovimientoItem.CodigoItem _
'                                        + vbCr + " Item: " & mMovimientoItem.Item _
'                                        + vbCr + " Cantidad: " & F.NuloString(mCantidadAcumulada) _
'                                        + vbCr + " Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
'                                        + vbCr + " Documento: " & mMovimientoItem.NumeroDocumento
            End If
            
            ' MOVIMIENTO NO COSTEADO
            If mMovimientoItem.Costo = 0 Then
                ' Costeamos el movimiento
                mMovimientoItem.Costo = mCostoMovimiento
                mMovimientoItem.CostoPrimo = mCostoMovimiento
                mMovimientoItem.CostoUnitario = mCostoUnitarioMovimiento
                mMovimientoItem.CostoUnitarioPromedio = mCostoUnitarioPromedio
                
                GrabaCostoMovimiento IdAlmacenProceso, mMovimientoItem, mConexion
            ' MOVIMIENTO COSTEADO
            Else
                ' Si ha cambiado el costo unitario promedio
                If mCostoUnitarioPromedio <> mMovimientoItem.CostoUnitarioPromedio Then
                    ' Costeamos el movimiento
                    mMovimientoItem.Costo = mCostoMovimiento
                    mMovimientoItem.CostoPrimo = mCostoMovimiento
                    mMovimientoItem.CostoUnitario = mCostoUnitarioMovimiento
                    mMovimientoItem.CostoUnitarioPromedio = mCostoUnitarioPromedio
                    
                    GrabaCostoMovimiento IdAlmacenProceso, mMovimientoItem, mConexion
                End If
            End If
            
            ' Actualizamos el costo acumulado
            mCostoAcumulado = mCostoAcumulado - mCostoMovimiento
        End If
    Next
    
    Set mLMovimientoItem = Nothing
    CosteaItem = True
    Exit Function

BloqueError:
    'Resume
    Set mLMovimientoItem = Nothing
    Set mErrorCosto = New ContabilidadEntidad.EErrorCosto
    mErrorCosto.CodigoError = "Err00x"
    mErrorCosto.DetalleError = Err.Description
    mErrorCosto.Origen = "CosteaItem"
    mErrorCosto.SolucionError = ""
    ' Se agrega a la lista de errores
    LErrorCosto.Add mErrorCosto
    CosteaItem = False
End Function

''' <summary>
''' Graba el registro de kardex de un movimiento
''' </summary>
Public Sub GrabaCostoMovimiento(IdAlmacenProceso As Long, _
                                    MovimientoItem As ContabilidadEntidad.EMovimientoItem, _
                                    Optional mConexion As ADODB.Connection = Nothing)
    
    Dim mCostoTemp As New ContabilidadEntidad.ELibroCostoTemp
    Dim F As New SistemaLogica.Funciones
            
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set mCostoTemp.Conexion = mConexion
    ' Se trae el registro si existiese
    mCostoTemp.Fetch MovimientoItem.IdMovimientoDet
    ' Se modifican los datos
    'mCostoTemp
    mCostoTemp.IdAlmacenProceso = IdAlmacenProceso
    mCostoTemp.IdMovimientoDetalle = MovimientoItem.IdMovimientoDet
    mCostoTemp.TipoMovimiento = MovimientoItem.TipoMovimiento
    mCostoTemp.FechaMovimiento = MovimientoItem.FechaMovimiento
    mCostoTemp.Cantidad = MovimientoItem.Cantidad
    mCostoTemp.CostoUnitario = MovimientoItem.CostoUnitario
    mCostoTemp.CostoUnitarioPromedio = MovimientoItem.CostoUnitarioPromedio
    mCostoTemp.CostoPrimo = MovimientoItem.CostoPrimo
    mCostoTemp.CostoMOD = MovimientoItem.CostoMOD
    mCostoTemp.CostoCIF = MovimientoItem.CostoCIF
    ' Se graba el registro
    If Not mConexion Is Nothing Then Set mCostoTemp.Conexion = mConexion
    If Not mCostoTemp.Save(0, "") Then
        Set mCostoTemp = Nothing
        Err.Raise &HFFFFFF01, , "Error al intentar grabar el costo del movimiento. " _
                                    + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                    + vbCr + " Item: " & MovimientoItem.Item _
                                    + vbCr + " Movimiento: " & MovimientoItem.NumeroDocumento _
                                    + vbCr + " Fecha: " & MovimientoItem.FechaMovimiento
    End If
    Set mCostoTemp = Nothing
    Exit Sub
    
BloqueError:
    Set mCostoTemp = Nothing
    Err.Raise Err.Number, "[GrabaCostoMovimiento] " & Err.Source, Err.Description
End Sub

Private Function CosteaParteDetalle(ByRef LErrorCosto As ContabilidadEntidad.LEErrorCosto, _
                                        IdAlmacenProceso As Long, _
                                        IdParteProdDet As Long, _
                                        FechaInicioProceso As Date, _
                                        Optional mConexion As ADODB.Connection) As Double
    Dim F As New SistemaLogica.Funciones
    Dim mParteProdDet As New ProduccionEntidad.EParteProdDet
    Dim mCostoTotalInsumo As Double
    Dim mCostoTotalMovimientos As Double
    Dim mNumParteProd As String
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set mParteProdDet.Conexion = mConexion
    mParteProdDet.Fetch IdParteProdDet
    
    If mParteProdDet.LParteProduccionDetIns.Count = 0 Then
        mNumParteProd = F.NuloString(F.BuscaCodigoTabla(mParteProdDet.IdParteProduccion, "id", "numser", "pro_produccion", "N", mConexion))
        mNumParteProd = mNumParteProd & "-" & F.NuloString(F.BuscaCodigoTabla(mParteProdDet.IdParteProduccion, "id", "numdoc", "pro_produccion", "N", mConexion))
        Err.Raise &HFFFFFF01, , "Los insumos del detalle del Parte de Produccion: " & mNumParteProd & " no cuentan con movimientos en el almacén " _
            + vbCr + " Fecha: " & F.NuloString(mParteProdDet.Fecha) _
            + vbCr + " Codigo de Item: " & mParteProdDet.CodigoItem _
            + vbCr + " Item: " & mParteProdDet.Item _
            + vbCr + " Cantidad: " & F.NuloString(mParteProdDet.CantidadProducida)
    End If
    
    ' Insumos del Parte
    mCostoTotalInsumo = 0
    Dim mParteProdDetIns As New ProduccionEntidad.EParteProdDetIns
    For Each mParteProdDetIns In mParteProdDet.LParteProduccionDetIns
        ' Se valida que el item insumo no sea igual al padre
        If mParteProdDetIns.IdItem = mParteProdDet.IdItem Then
            mNumParteProd = F.NuloString(F.BuscaCodigoTabla(mParteProdDet.IdParteProduccion, "id", "numser", "pro_produccion", "N", mConexion))
            mNumParteProd = mNumParteProd & "-" & F.NuloString(F.BuscaCodigoTabla(mParteProdDet.IdParteProduccion, "id", "numdoc", "pro_produccion", "N", mConexion))
            Err.Raise &HFFFFFF01, , "Movimiento de item registrado como parte del mismo item" _
                + vbCr + " Parte de Produccion: " & mNumParteProd _
                + vbCr + " Fecha: " & F.NuloString(mParteProdDet.Fecha) _
                + vbCr + " Codigo de Item: " & mParteProdDet.CodigoItem _
                + vbCr + " Item: " & mParteProdDet.Item _
                + vbCr + " Cantidad: " & F.NuloString(mParteProdDet.CantidadProducida)
        End If
        
        ' Movimientos del Insumo
        mCostoTotalMovimientos = 0
        Dim mParteProdDetInsMov As New ProduccionEntidad.EParteProdDetInsMov
        For Each mParteProdDetInsMov In mParteProdDetIns.LParteProduccionDetInsMov
            ' Se obtiene el item movimiento
            Dim mMovimientoItem As New ContabilidadEntidad.EMovimientoItem
            If Not mConexion Is Nothing Then Set mMovimientoItem.Conexion = mConexion
            mMovimientoItem.Fetch mParteProdDetInsMov.IdMovimientoDetalle
            If mMovimientoItem.Costo = 0 Then
                ' Se valida movimientos inconsistentes
                If mMovimientoItem.IdMovimientoDet = 0 Then
                    mNumParteProd = F.NuloString(F.BuscaCodigoTabla(mParteProdDet.IdParteProduccion, "id", "numser", "pro_produccion", "N", mConexion))
                    mNumParteProd = mNumParteProd & "-" & F.NuloString(F.BuscaCodigoTabla(mParteProdDet.IdParteProduccion, "id", "numdoc", "pro_produccion", "N", mConexion))
                    Err.Raise &HFFFFFF01, , "Error de data - se ha encontrado un movimiento inconsistente. Consultar con el administrador del sistema " _
                        + vbCr + " Parte de Produccion: " & mNumParteProd _
                        + vbCr + " Fecha: " & F.NuloString(mParteProdDet.Fecha) _
                        + vbCr + " Codigo de Item: " & mParteProdDet.CodigoItem _
                        + vbCr + " Item: " & mParteProdDet.Item _
                        + vbCr + " Cantidad: " & F.NuloString(mParteProdDet.CantidadProducida) _
                        + vbCr + " Id del detalle del movimiento: " & F.NuloString(mParteProdDetInsMov.IdMovimientoDetalle)
                End If
                mCostoTotalMovimientos = mCostoTotalMovimientos + CosteaMovimientoDetalle(LErrorCosto, IdAlmacenProceso, mParteProdDet.Fecha, mMovimientoItem, mConexion)
            Else
                mCostoTotalMovimientos = mCostoTotalMovimientos + mMovimientoItem.Costo
            End If
        Next
        mCostoTotalInsumo = mCostoTotalInsumo + mCostoTotalMovimientos
    Next
    Set mParteProdDet = Nothing
    Set mParteProdDetIns = Nothing
    Set mParteProdDetInsMov = Nothing
    CosteaParteDetalle = mCostoTotalInsumo
    Exit Function

BloqueError:
'Resume
    Set mParteProdDet = Nothing
    Set mParteProdDetIns = Nothing
    Set mParteProdDetInsMov = Nothing
    Err.Raise Err.Number, "[CosteaParteDetalle] " & Err.Source, Err.Description
End Function

Public Function CosteaMovimientoDetalle(ByRef LErrorCosto As ContabilidadEntidad.LEErrorCosto, _
                                    IdAlmacenProceso As Long, _
                                    FechaInicioProceso As Date, _
                                    MovimientoItem As ContabilidadEntidad.EMovimientoItem, _
                                    Optional mConexion As ADODB.Connection) As Double
    Dim F As New SistemaLogica.Funciones
    Dim mRecord As New ADODB.Recordset
    Dim mSQL As String
    Dim mCostoMovimientoDetalle As Double
    
On Error GoTo BloqueError
    If MovimientoItem.IdAlmacen = 0 Then
        Err.Raise &HFFFFFF01, , "El movimiento no cuenta con almacen de referencia: " _
                        + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                        + vbCr + " Item: " & MovimientoItem.Item _
                        + vbCr + " Cantidad: " & F.NuloString(MovimientoItem.Cantidad) _
                        + vbCr + " Fecha de Movimiento: " & F.NuloString(MovimientoItem.FechaMovimiento) _
                        + vbCr + " Documento: " & MovimientoItem.NumeroDocumento _
                        + vbCr + " Id del Movimiento: " & F.NuloString(MovimientoItem.IdMovimiento) _
                        + vbCr + " Id detalle del Movimiento: " & F.NuloString(MovimientoItem.IdMovimientoDet)
    End If
    mCostoMovimientoDetalle = 0
    ' INGRESOS
    If MovimientoItem.TipoMovimiento = "I" Then
        Select Case MovimientoItem.IdTipoDocumentoReferencia
            ' PARTE DE PRODUCCION
            Case F.NuloNumeric(F.KeyValue("ParteProduccion", mConexion))
                If MovimientoItem.IdDocumentoAnexado = 0 Then
                    MovimientoItem.IdDocumentoAnexado = HallaDocumentoAnexado(MovimientoItem, , mConexion)
                End If
                mCostoMovimientoDetalle = CosteaParteDetalle(LErrorCosto, IdAlmacenProceso, MovimientoItem.IdDocumentoAnexado, FechaInicioProceso, mConexion)
                               
            ' SOLICITUD DE MATERIALES
            Case F.NuloNumeric(F.KeyValue("SolictudMateriales", mConexion))
                ' Validamos el costo promedio
                If MovimientoItem.CostoUnitarioPromedio = 0 Then
                    Err.Raise &HFFFFFF01, , "Costo Unitario Promedio igual a cero al intentar costear Factura. " _
                                + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                + vbCr + " Item: " & MovimientoItem.Item _
                                + vbCr + " Cantidad: " & F.NuloString(MovimientoItem.Cantidad) _
                                + vbCr + " Fecha de Movimiento: " & F.NuloString(MovimientoItem.FechaMovimiento) _
                                + vbCr + " Documento: " & MovimientoItem.NumeroDocumento
                End If
                mCostoMovimientoDetalle = MovimientoItem.CostoUnitarioPromedio * MovimientoItem.Cantidad
                
            ' INVENTARIO INICIAL
            Case F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", mConexion))
                    mSQL = "SELECT alm_inventarioinicialdet.costo " _
                        + vbCr + "FROM (alm_inventarioinicial INNER JOIN (alm_ingresodet INNER JOIN alm_ingreso ON alm_ingresodet.id = alm_ingreso.id) ON alm_inventarioinicial.idinventarioinicial = alm_ingreso.iddocref) INNER JOIN alm_inventarioinicialdet ON alm_inventarioinicial.idinventarioinicial = alm_inventarioinicialdet.idinventarioinicial " _
                        + vbCr + "WHERE (((alm_inventarioinicialdet.iditem)=" & MovimientoItem.IdItem & ") AND ((alm_ingresodet.idmovdet)=" & MovimientoItem.IdMovimientoDet & ") AND ((alm_inventarioinicial.idestado)=" & F.NuloNumeric(F.KeyValue("EstadoAprobadoInventarioInicial", mConexion)) & ") AND ((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", mConexion)) & "))"
                    
                    Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
                    If mRecord.RecordCount = 0 Then
                        Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en Inventario Inicial. " _
                                            + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                            + vbCr + " Item: " & MovimientoItem.Item _
                                            + vbCr + " Movimiento: " & MovimientoItem.NumeroDocumento _
                                            + vbCr + " Fecha: " & MovimientoItem.FechaMovimiento
                    ElseIf F.NuloNumeric(mRecord("costo")) = 0 Then
                        Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en Inventario Inicial. " _
                                                + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                                + vbCr + " Item: " & MovimientoItem.Item _
                                                + vbCr + " Movimiento: " & MovimientoItem.NumeroDocumento _
                                                + vbCr + " Fecha: " & MovimientoItem.FechaMovimiento
                    End If
                    mCostoMovimientoDetalle = F.NuloNumeric(mRecord("costo")) * MovimientoItem.Cantidad
            
            ' AJUSTE DE INVENTARIO
            Case F.NuloNumeric(F.KeyValue("IdDocumentoAjusteInventario", mConexion))
                mSQL = "SELECT alm_tomainventariodet.preuni AS costo " _
                    + vbCr + "FROM (alm_ingreso INNER JOIN (alm_tomainventario INNER JOIN alm_tomainventariodet ON alm_tomainventario.idtomainventario = alm_tomainventariodet.idtomainventario) ON alm_ingreso.iddocref = alm_tomainventario.idtomainventario) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id " _
                    + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & MovimientoItem.IdMovimientoDet & ") AND ((alm_tomainventariodet.iditem)=" & MovimientoItem.IdItem & ") AND ((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoAjusteInventario", mConexion)) & ") AND ((alm_tomainventario.idestadoinventario)=" & F.NuloNumeric(F.KeyValue("EstadoAprobadoInventario", mConexion)) & "))"
                
                Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
                If mRecord.RecordCount = 0 Then
                    Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en Ajuste de Inventario. " _
                                        + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                        + vbCr + " Item: " & MovimientoItem.Item _
                                        + vbCr + " Movimiento: " & MovimientoItem.NumeroDocumento _
                                        + vbCr + " Fecha: " & MovimientoItem.FechaMovimiento
                ElseIf F.NuloNumeric(mRecord("costo")) = 0 Then
                    ' Se busca el ultimo costo del item
                    mSQL = "SELECT TOP 1 com_comprasdet.preuni AS costo " _
                        + vbCr + "FROM com_compras INNER JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom " _
                        + vbCr + "WHERE (((com_comprasdet.IdItem) = " & MovimientoItem.IdItem & ")) " _
                        + vbCr + "ORDER BY com_compras.fchdoc DESC"
                    
                    Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
                    If mRecord.RecordCount = 0 Then
                        Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en Ajuste de Inventario. No se encontraron movimientos para hallar el ultimo costo" _
                                            + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                            + vbCr + " Item: " & MovimientoItem.Item _
                                            + vbCr + " Movimiento: " & MovimientoItem.NumeroDocumento _
                                            + vbCr + " Fecha: " & MovimientoItem.FechaMovimiento
                    ElseIf F.NuloNumeric(mRecord("costo")) = 0 Then
                        Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en último registro de compras. " _
                                            + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                            + vbCr + " Item: " & MovimientoItem.Item _
                                            + vbCr + " Movimiento: " & MovimientoItem.NumeroDocumento _
                                            + vbCr + " Fecha: " & MovimientoItem.FechaMovimiento
                    End If
                    mCostoMovimientoDetalle = F.NuloNumeric(mRecord("costo")) * MovimientoItem.Cantidad
                Else
                    mCostoMovimientoDetalle = F.NuloNumeric(mRecord("costo")) * MovimientoItem.Cantidad
                End If
            
            'NOTAS DE CREDITO
            Case F.NuloNumeric(F.KeyValue("IdDocumentoFactura", mConexion))
                ' Validamos el costo promedio
                If MovimientoItem.CostoUnitarioPromedio = 0 Then
                    Err.Raise &HFFFFFF01, , "Costo Unitario Promedio igual a cero al intentar costear Factura. " _
                                + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                + vbCr + " Item: " & MovimientoItem.Item _
                                + vbCr + " Cantidad: " & F.NuloString(MovimientoItem.Cantidad) _
                                + vbCr + " Fecha de Movimiento: " & F.NuloString(MovimientoItem.FechaMovimiento) _
                                + vbCr + " Documento: " & MovimientoItem.NumeroDocumento
                End If
                mCostoMovimientoDetalle = MovimientoItem.CostoUnitarioPromedio * MovimientoItem.Cantidad
                
            Case Else
                mSQL = "SELECT Sum(IIf([com_compras].[idmon]=2,[com_comprasdet].[imptot]*[con_tc].[impven],[com_comprasdet].[imptot])) AS importetot, Sum(com_comprasdet.canpro) AS cantidadtot " _
                    + vbCr + "FROM ((com_compras INNER JOIN (com_comprasdet INNER JOIN alm_ingresodoc ON com_comprasdet.idcom = alm_ingresodoc.iddoc) ON com_compras.id = com_comprasdet.idcom) INNER JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN alm_ingresodet ON alm_ingresodoc.id = alm_ingresodet.id " _
                    + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & MovimientoItem.IdMovimientoDet & ") AND ((com_comprasdet.iditem)=" & MovimientoItem.IdItem & "))"
                
                Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
                If mRecord.RecordCount = 0 Then
                    Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en compras o no existe registro de compras. " _
                                            + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                            + vbCr + " Item: " & MovimientoItem.Item _
                                            + vbCr + " Movimiento: " & MovimientoItem.NumeroDocumento _
                                            + vbCr + " Fecha: " & MovimientoItem.FechaMovimiento
                ElseIf F.NuloNumeric(mRecord("cantidadtot")) = 0 Then
                    Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en compras o no existe registro de compras. " _
                                            + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                                            + vbCr + " Item: " & MovimientoItem.Item _
                                            + vbCr + " Movimiento: " & MovimientoItem.NumeroDocumento _
                                            + vbCr + " Fecha: " & MovimientoItem.FechaMovimiento
                Else
                    mCostoMovimientoDetalle = (F.NuloNumeric(mRecord("importetot")) / F.NuloNumeric(mRecord("cantidadtot"))) * MovimientoItem.Cantidad
                End If
        End Select
    ' SALIDAS
    Else
        ' Se calculan los costos unitarios hasta la fecha del movimiento
        If Not CosteaItem(LErrorCosto, IdAlmacenProceso, MovimientoItem.IdItem, MovimientoItem.IdAlmacen, FechaInicioProceso, MovimientoItem.FechaMovimiento, mConexion) Then
            Err.Raise &HFFFFFF01, , "Error luego de procesar el costeo del movimiento: " _
                        + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                        + vbCr + " Item: " & MovimientoItem.Item _
                        + vbCr + " Cantidad: " & F.NuloString(MovimientoItem.Cantidad) _
                        + vbCr + " Fecha de Movimiento: " & F.NuloString(MovimientoItem.FechaMovimiento) _
                        + vbCr + " Documento: " & MovimientoItem.NumeroDocumento
        End If
        ' Se vuelve a validar el costo del movimiento
        MovimientoItem.Fetch MovimientoItem.IdMovimientoDet
        If MovimientoItem.Costo = 0 Then
            Err.Raise &HFFFFFF01, , "Costo igual a cero despues de ejecutar costeo del movimiento: " _
                        + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                        + vbCr + " Item: " & MovimientoItem.Item _
                        + vbCr + " Cantidad: " & F.NuloString(MovimientoItem.Cantidad) _
                        + vbCr + " Fecha de Movimiento: " & F.NuloString(MovimientoItem.FechaMovimiento) _
                        + vbCr + " Documento: " & MovimientoItem.NumeroDocumento
        End If
        mCostoMovimientoDetalle = MovimientoItem.Costo
    
    End If
        
    CosteaMovimientoDetalle = mCostoMovimientoDetalle
    Exit Function

BloqueError:
    'Resume
    CosteaMovimientoDetalle = 0
    Err.Raise Err.Number, "[CosteaMovimientoDetalle] " & Err.Source, Err.Description
End Function

Public Function HallaDocumentoAnexado(Optional MovimientoItem As ContabilidadEntidad.EMovimientoItem = Nothing, _
                            Optional ParteProdItem As ContabilidadEntidad.EParteProdItem = Nothing, _
                            Optional mConexion As ADODB.Connection) As Long
    
    Dim mRecord As New ADODB.Recordset
    Dim mSQL As String
    Dim mIdDocumentoAnexado As Long
    Dim F As New SistemaLogica.Funciones
    Dim database As New SistemaData.EDataBase

On Error GoTo BloqueError
    mIdDocumentoAnexado = 0
    ' Se valida el tipo de documento de referencia
    If Not MovimientoItem Is Nothing Then
        Set mRecord = Nothing
        mSQL = "SELECT pro_producciondet.idproddet, alm_ingresodet.idmovdet, alm_ingresodet_1.idmovdet AS idmovdetanex, alm_ingresodet_1.iddocref " _
                + vbCr + "FROM (((pro_produccion INNER JOIN alm_ingreso ON pro_produccion.id = alm_ingreso.iddocref) INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_ingresodet AS alm_ingresodet_1 ON pro_producciondet.idproddet = alm_ingresodet_1.iddocref " _
                + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & MovimientoItem.IdMovimientoDet & ") AND ((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoParteProduccion", mConexion)) & ") AND ((pro_producciondet.iditem)=[alm_ingresodet].[iditem]) AND ((pro_producciondet.cantidad)=[alm_ingresodet].[cantidad]))"
        Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
        If mRecord.RecordCount = 0 Then
            Err.Raise &HFFFFFF01, , "El detalle del Movimiento: " & MovimientoItem.NumeroDocumento & ", no cuenta con un Parte asociado en Producción" _
                            + vbCr + " Codigo de Item: " & MovimientoItem.CodigoItem _
                            + vbCr + " Item: " & MovimientoItem.Item _
                            + vbCr + " Cantidad: " & F.NuloString(MovimientoItem.Cantidad)
        Else
            mRecord.Filter = adFilterNone
            mRecord.Filter = "idmovdetanex = " & MovimientoItem.IdMovimientoDet
            If mRecord.RecordCount = 0 Then
                mRecord.Filter = adFilterNone
                mRecord.Filter = "idmovdetanex = 0 Or idmovdetanex = null"
                mRecord.MoveFirst
                ' Se amarra al primer movimiento encontrado
                Set database.Connection = mConexion
                database.ClearParameter
                database.CommandText = "UPDATE alm_ingresodet SET alm_ingresodet.iddocref = " & F.NuloNumeric(mRecord("idproddet")) & " " _
                                + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & MovimientoItem.IdMovimientoDet & "))"
                database.Execute
            End If
        End If
        mIdDocumentoAnexado = F.NuloNumeric(mRecord("idproddet"))
    End If
    If Not ParteProdItem Is Nothing Then
        Set mRecord = Nothing
        mSQL = "SELECT pro_producciondet.idproddet, alm_ingresodet.idmovdet, alm_ingresodet.iddocref " _
            + vbCr + "FROM (pro_produccion INNER JOIN (alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) ON pro_produccion.id = alm_ingreso.iddocref) INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
            + vbCr + "WHERE (((pro_producciondet.idproddet)=" & ParteProdItem.IdParteProduccionDet & ") AND ((alm_ingresodet.iddocref) Is Null Or (alm_ingresodet.iddocref)=[pro_producciondet].[idproddet] Or (alm_ingresodet.iddocref)=0) AND ((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoParteProduccion", mConexion)) & ") AND ((pro_producciondet.iditem)=[alm_ingresodet].[iditem]) AND ((pro_producciondet.cantidad)=[alm_ingresodet].[cantidad]))"
        Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
        If mRecord.RecordCount = 0 Then
            Err.Raise &HFFFFFF01, , "El detalle del parte de produccion: " & ParteProdItem.NumeroDocumento & ", no cuenta con un movmiento asociado en Almacén" _
                            + vbCr + " Codigo de Item: " & ParteProdItem.CodigoItem _
                            + vbCr + " Item: " & ParteProdItem.Item _
                            + vbCr + " Cantidad: " & F.NuloString(ParteProdItem.CantidadProducida)
        Else
            ' Buscamos el registro anexado si existiese
            mRecord.Filter = adFilterNone
            mRecord.Filter = "iddocref=" & ParteProdItem.IdParteProduccionDet
            If mRecord.RecordCount = 0 Then
                mRecord.Filter = adFilterNone
                mRecord.Filter = "iddocref = 0 Or iddocref = null"
                mRecord.MoveFirst
                ' Se amarra al primer movimiento encontrado
                Set database.Connection = mConexion
                database.ClearParameter
                database.CommandText = "UPDATE alm_ingresodet SET alm_ingresodet.iddocref = " & ParteProdItem.IdParteProduccionDet & " " _
                                + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & F.NuloNumeric(mRecord("idmovdet")) & "))"
                database.Execute
            End If
        End If
        mIdDocumentoAnexado = F.NuloNumeric(mRecord("idmovdet"))
    End If
    HallaDocumentoAnexado = mIdDocumentoAnexado
    Set mRecord = Nothing
    'Set MovimientoItem = Nothing
    Set ParteProdItem = Nothing
    Exit Function

BloqueError:
    Set mRecord = Nothing
    'Set MovimientoItem = Nothing
    Set ParteProdItem = Nothing
    HallaDocumentoAnexado = 0
    Err.Raise Err.Number, "[HallaDocumentoAnexado] " & Err.Source, Err.Description
End Function


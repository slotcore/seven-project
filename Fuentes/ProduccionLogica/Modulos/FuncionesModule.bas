Attribute VB_Name = "FuncionesModule"
Option Explicit

Private PForm As New SistemaWindows.SistemaWindowsClass
Public mError As Boolean
    
''' <summary>
''' Genera los movimientos automaticos para guias y facturas de venta en un rango de fechas
''' </summary>
Function mGenerarMovimientosGuiasVentas(IdUsuario As Long, FechaInicio As Date, _
                                    FechaFin As Date, FechaInicioMovimientos As Date, _
                                    mConexion As ADODB.Connection) As Boolean
    Dim database As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones
    Dim RecordParent As New ADODB.Recordset
    Dim RecordChild As New ADODB.Recordset
    
On Error GoTo BloqueError
    Set database.Connection = mConexion
    ' Se encuentran las facturas y guias
    database.CommandText = "SELECT A.id, A.fechmov, A.numser, A.numdoc, A.idalm, A.tipo, A.tipmov " _
                + vbCr + "FROM " _
                + vbCr + "( " _
                + vbCr + "SELECT vta_guia.id, vta_guiadet.iditem, vta_guia.fecgiro AS fechmov, vta_guia.numser, vta_guia.numdoc, Iif(vta_guia.idalm Is Null, 2, vta_guia.idalm) As idalm, vta_guiadet.canpro, 'GR' AS tipo, 0 AS tipmov " _
                + vbCr + "FROM vta_guia LEFT JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui " _
                + vbCr + "WHERE (((vta_guia.fecgiro)>=CDate('" & FechaInicio & "') And (vta_guia.fecgiro)<=CDate('" & FechaFin & "') And (vta_guia.fecgiro)>=CDate('" & FechaInicioMovimientos & "'))) " _
                + vbCr + "Union All " _
                + vbCr + "SELECT vta_ventas.id, vta_ventasdet.iditem, vta_ventas.fchdoc AS fechmov, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.idalm, vta_ventasdet.canpro, 'V' AS tipo, 0 AS tipmov " _
                + vbCr + "FROM vta_ventas RIGHT JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
                + vbCr + "WHERE (((vta_ventas.fchdoc)>=CDate('" & FechaInicio & "') And (vta_ventas.fchdoc)<=CDate('" & FechaFin & "') And (vta_ventas.fchdoc)>=CDate('" & FechaInicioMovimientos & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref) Is Null Or (vta_ventas.iddocref)=0)) " _
                + vbCr + "UNION ALL " _
                + vbCr + "SELECT vta_ventas.id, vta_ventasdet.iditem, vta_ventas.fchdoc AS fechmov, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.idalm, vta_ventasdet.canpro, 'VNC' AS tipo, -1 AS tipmov " _
                + vbCr + "FROM vta_ventas RIGHT JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
                + vbCr + "WHERE (((vta_ventas.fchdoc)>=CDate('" & FechaInicio & "') And (vta_ventas.fchdoc)<=CDate('" & FechaFin & "') And (vta_ventas.fchdoc)>=CDate('" & FechaInicioMovimientos & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref)<>0) AND ((vta_ventas.idmotnotcre)=4))" _
                + vbCr + ") AS A " _
                + vbCr + "GROUP BY A.id, A.fechmov, A.numser, A.numdoc, A.idalm, A.tipo, A.tipmov"
    Set RecordParent = database.GetRecordset
    ' Se encuentran los detalles de las facturas y guias
    database.ClearParameter
    database.CommandText = "SELECT vta_guia.id, vta_guiadet.iditem, vta_guia.fecgiro AS fechmov, vta_guia.numser, vta_guia.numdoc, Iif(vta_guia.idalm Is Null, 2, vta_guia.idalm) As idalm, vta_guiadet.canpro, 'GR' AS tipo, 0 AS tipmov " _
                + vbCr + "FROM vta_guia LEFT JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui " _
                + vbCr + "WHERE (((vta_guia.fecgiro)>=CDate('" & FechaInicio & "') And (vta_guia.fecgiro)<=CDate('" & FechaFin & "') And (vta_guia.fecgiro)>=CDate('" & FechaInicioMovimientos & "'))) " _
                + vbCr + "Union All " _
                + vbCr + "SELECT vta_ventas.id, vta_ventasdet.iditem, vta_ventas.fchdoc AS fechmov, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.idalm, vta_ventasdet.canpro, 'V' AS tipo, 0 AS tipmov " _
                + vbCr + "FROM vta_ventas RIGHT JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
                + vbCr + "WHERE (((vta_ventas.fchdoc)>=CDate('" & FechaInicio & "') And (vta_ventas.fchdoc)<=CDate('" & FechaFin & "') And (vta_ventas.fchdoc)>=CDate('" & FechaInicioMovimientos & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref) Is Null Or (vta_ventas.iddocref)=0)) " _
                + vbCr + "Union All " _
                + vbCr + "SELECT vta_ventas.id, vta_ventasdet.iditem, vta_ventas.fchdoc AS fechmov, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.idalm, vta_ventasdet.canpro, 'VNC' AS tipo, -1 AS tipmov " _
                + vbCr + "FROM vta_ventas RIGHT JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
                + vbCr + "WHERE (((vta_ventas.fchdoc)>=CDate('" & FechaInicio & "') And (vta_ventas.fchdoc)<=CDate('" & FechaFin & "') And (vta_ventas.fchdoc)>=CDate('" & FechaInicioMovimientos & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref)<>0) AND ((vta_ventas.idmotnotcre)=4))"
    Set RecordChild = database.GetRecordset
    
    If RecordParent.RecordCount = 0 Then Err.Raise &HFFFFFF01, , "No se encontrar registros para la busqueda"
    RecordParent.MoveFirst
    mConexion.BeginTrans
    Dim mNumeroSerie As String
    Dim mNumeroDocumento As String
    mNumeroSerie = "0020"
    While Not RecordParent.EOF
        Dim Movimiento As New AlmacenEntidad.EMovimiento
        ' Cabecera
        Movimiento.IdTipoMovimiento = F.NuloNumeric(RecordParent("tipmov"))
        Movimiento.FechaMovimiento = CDate(RecordParent("fechmov"))
        Movimiento.NumeroSerie = F.NuloString(mNumeroSerie)
        Movimiento.NumeroDocumento = F.HallaNumeroDocumento("alm_ingreso", "'" & mNumeroSerie & "'", "numser", mConexion)
        Movimiento.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoMovimiento", mConexion))
        Movimiento.IdAlmacen = F.NuloNumeric(RecordParent("idalm"))
        Select Case F.NuloString(RecordParent("tipo"))
            Case "GR"
                Movimiento.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("IdDocumentoGuiaRemision", mConexion))
            Case "V"
                Movimiento.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("IdDocumentoFactura", mConexion))
            Case "VNC"
                Movimiento.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("IdDocumentoFactura", mConexion))
        End Select
        Movimiento.IdDocumentoReferencia = F.NuloNumeric(RecordParent("id"))
        Movimiento.DocumentoReferencia = F.NuloString(RecordParent("numser")) & " - " & F.NuloString(RecordParent("numdoc"))
        Movimiento.MesTrabajo = Month(CDate(RecordParent("fechmov")))
        Movimiento.AnhoTrabajo = Year(CDate(RecordParent("fechmov")))
        
        '*********************************************************
        ' Se elimina los movimientos relacionados
        Dim mDataBase As New SistemaData.EDataBase
        Dim mRecordAux As New ADODB.Recordset
        
        Set mDataBase.Connection = mConexion
        mDataBase.CommandText = "SELECT alm_ingreso.id AS idmov, alm_ingreso.numser, alm_ingreso.numdoc " _
                    + vbCr + "FROM alm_ingreso " _
                    + vbCr + "WHERE (((alm_ingreso.idtipdocref)=" & Movimiento.IdTipoDocumentoReferencia & ") AND ((alm_ingreso.iddocref)=" & Movimiento.IdDocumentoReferencia & "))"
        Set mRecordAux = mDataBase.GetRecordset
        If mRecordAux.RecordCount > 0 Then
            mRecordAux.MoveFirst
            While Not mRecordAux.EOF
                '--
                ' Se elimina el detalle
                mDataBase.ClearParameter
                mDataBase.CommandText = "DELETE FROM alm_ingresodet " _
                    + vbCr + "WHERE alm_ingresodet.id = " & F.NuloNumeric(mRecordAux("idmov")) & ""
                mDataBase.Execute
                ' Se elimina la cabecera
                mDataBase.ClearParameter
                mDataBase.CommandText = "DELETE FROM alm_ingreso " _
                    + vbCr + "WHERE alm_ingreso.id = " & F.NuloNumeric(mRecordAux("idmov")) & ""
                mDataBase.Execute
                '--
                mRecordAux.MoveNext
            Wend
        End If
        Set mRecordAux = Nothing
        Set mDataBase = Nothing
        '*********************************************************
        
        ' Detalle
        RecordChild.Filter = adFilterNone
        RecordChild.Filter = "id=" & F.NuloNumeric(RecordParent("id"))
        If RecordChild.RecordCount > 0 Then
            RecordChild.MoveFirst
            While Not RecordChild.EOF
                Dim MovimientoDet As New AlmacenEntidad.EMovimientoDet
                MovimientoDet.IdItem = F.NuloNumeric(RecordChild("iditem"))
                MovimientoDet.Cantidad = F.NuloNumeric(RecordChild("canpro"))
                MovimientoDet.CantidadTeorica = F.NuloNumeric(RecordChild("canpro"))
                ' Se agrega al padre
                Movimiento.LMovimientoDet.Add MovimientoDet
                Set MovimientoDet = Nothing
                RecordChild.MoveNext
            Wend
        End If
        ' Se graba el registro
        Set Movimiento.Conexion = mConexion
        Movimiento.Called = True
        If Not Movimiento.Save(IdUsuario, F.MachineName()) Then Err.Raise &HFFFFFF01, , F.ErrorDescriptionDLL(Err.LastDllError)
        Set Movimiento = Nothing
        
        RecordParent.MoveNext
    Wend
    mConexion.CommitTrans
    mGenerarMovimientosGuiasVentas = True
    Exit Function
    
BloqueError:
    mConexion.RollbackTrans
    Set RecordParent = Nothing
    Set RecordChild = Nothing
    Set Movimiento = Nothing
    mGenerarMovimientosGuiasVentas = False
    F.MostrarMensajeError Err.Description, "GenerarMovimientosGuiasVentas", Err.Source, Err.Number
End Function

''' <summary>
''' Halla el importe unitario promedio de partes de produccion en un rango de fechas
''' </summary>
Function mCostoPrimo(IdAlmacen As Long, FechaInicio As Date, FechaFin As Date, FechaInicioMovimientos As Date, _
                                    Optional mConexion As ADODB.Connection) As Double

    Dim F As New SistemaLogica.Funciones
    Dim LparteProdDet As New ProduccionEntidad.LEParteProdDet
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim mCostoParteDet As Double
    Dim mCostoUnitario As Double
    Dim mCostoTotal As Double
    Dim mContador As Long
    
    Dim mCostoUnitarioPromedio As Double
    Dim mCostoMovimiento As Double

On Error GoTo BloqueError

    ' Se carga el listado de partes de produccion detalle
    If Not mConexion Is Nothing Then Set LparteProdDet.Conexion = mConexion
    LparteProdDet.LoadChild = False
    LparteProdDet.Fetch , IdAlmacen, FechaInicio, FechaFin, FechaInicioMovimientos
    
    PForm.ShowProgress "Proc.Parte:", 0, LparteProdDet.Count
    mContador = 0
    ' Se recorre la lista para calcular su importe unitario
    Dim ParteProdDet As New ProduccionEntidad.EParteProdDet
    For Each ParteProdDet In LparteProdDet
        Dim mNumeroParte As String
        Dim mFechaParte As String
        Dim mIdMovimientoDetalle As Long
        Dim mIdAlmacen As Long
        
        mContador = mContador + 1
        DoEvents
        mNumeroParte = F.NuloString(F.BuscaCodigoTabla(ParteProdDet.IdParteProduccion, "id", "numser", "pro_produccion", "N", mConexion))
        mNumeroParte = mNumeroParte & "-" & F.NuloString(F.BuscaCodigoTabla(ParteProdDet.IdParteProduccion, "id", "numdoc", "pro_produccion", "N", mConexion))
        mFechaParte = F.NuloString(F.BuscaCodigoTabla(ParteProdDet.IdParteProduccion, "id", "fchdoc", "pro_produccion", "N", mConexion))
        PForm.SetProgress mFechaParte & " - " & mNumeroParte & " - " & ParteProdDet.Item, mContador
        ' Se halla el movimiento detalle relacionado
        mIdMovimientoDetalle = mHallaMovimientoDetalle(ParteProdDet.IdParteProduccionDet, mConexion)
                
        ' Validar si ya esta costeado
        mCostoMovimiento = F.NuloNumeric(F.BuscaCodigoTabla(mIdMovimientoDetalle, "idmovdet", "costoprimo", "con_librocostotemp", "N", mConexion))
        If mCostoMovimiento > 0 Then ' Si esta costeado
             mCostoParteDet = mCostoMovimiento
        Else ' Si no esta costeado
            ' Se halla el costo del movimiento generado
            mCostoParteDet = mCalcularCostoMovimiento(IdAlmacen, mIdMovimientoDetalle, mConexion)
            If mError Then Exit Function
        End If
    Next
    PForm.HideProgress

    mCostoPrimo = mCostoTotal
    Exit Function
BloqueError:
    Set LparteProdDet = Nothing
    PForm.HideProgress
    F.MostrarMensajeError Err.Description, "CostoPrimo", Err.Source, Err.Number
End Function

''' <summary>
''' Amarra el Parte Detalle a el Movimiento detalle correspondiente
''' </summary>
Function mHallaMovimientoDetalle(IdParteProduccionDet As Long, Optional mConexion As ADODB.Connection) As Long
    Dim mRecord As New ADODB.Recordset
    Dim database As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones

On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    
    database.CommandText = "SELECT pro_producciondet.idproddet, alm_ingresodet.idmovdet, alm_ingresodet.iddocref " _
        + vbCr + "FROM (pro_produccion INNER JOIN (alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) ON pro_produccion.id = alm_ingreso.iddocref) INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr + "WHERE (((pro_producciondet.idproddet)=" & IdParteProduccionDet & ") AND ((alm_ingresodet.iddocref) Is Null Or (alm_ingresodet.iddocref)=[pro_producciondet].[idproddet] Or (alm_ingresodet.iddocref)=0) AND ((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoParteProduccion", mConexion)) & ") AND ((pro_producciondet.iditem)=[alm_ingresodet].[iditem]) AND ((pro_producciondet.cantidad)=[alm_ingresodet].[cantidad]))"
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount = 0 Then
        Dim mNumDoc As String
        Dim mIdParteProduccion As Long
        Dim mIdItem As Long
        Dim mCodigoItem As String
        Dim mItem As String
        Dim mCantidad As Double
        
        mIdItem = F.BuscaCodigoTabla(IdParteProduccionDet, "idproddet", "iditem", "pro_producciondet", "N", mConexion)
        mCodigoItem = F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion)
        mItem = F.BuscaCodigoTabla(mIdItem, "id", "descripcion", "alm_inventario", "N", mConexion)
        mIdParteProduccion = F.BuscaCodigoTabla(IdParteProduccionDet, "idproddet", "idpro", "pro_producciondet", "N", mConexion)
        mCantidad = F.BuscaCodigoTabla(IdParteProduccionDet, "idproddet", "cantidad", "pro_producciondet", "N", mConexion)
        mNumDoc = F.BuscaCodigoTabla(mIdParteProduccion, "id", "numser", "pro_produccion", "N", mConexion)
        mNumDoc = mNumDoc & "-" & F.BuscaCodigoTabla(mIdParteProduccion, "id", "numdoc", "pro_produccion", "N", mConexion)
        Err.Raise &HFFFFFF01, , "El detalle del parte de produccion: " & mNumDoc & ", no cuenta con un movmiento asociado en Almacén" _
                        + vbCr + "Codigo de Item: " & mCodigoItem _
                        + vbCr + "Item: " & mItem _
                        + vbCr + "Cantidad: " & F.NuloString(mCantidad)
    Else
        ' Buscamos el registro anexado si existiese
        mRecord.Filter = adFilterNone
        mRecord.Filter = "iddocref=" & IdParteProduccionDet
        If mRecord.RecordCount = 0 Then
            mRecord.Filter = adFilterNone
            mRecord.Filter = "iddocref = 0 Or iddocref = null"
            mRecord.MoveFirst
            ' Se amarra al primer movimiento encontrado
            database.ClearParameter
            database.CommandText = "UPDATE alm_ingresodet SET alm_ingresodet.iddocref = " & IdParteProduccionDet & " " _
                            + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & F.NuloNumeric(mRecord("idmovdet")) & "))"
            database.Execute
        End If
    End If
    mHallaMovimientoDetalle = F.NuloNumeric(mRecord("idmovdet"))
    Set database = Nothing
    Set mRecord = Nothing
    Exit Function

BloqueError:
    Set database = Nothing
    Set mRecord = Nothing
    mHallaMovimientoDetalle = 0
    Err.Raise Err.Number, "[HallaMovimientoDetalle] " & Err.Source, Err.Description
End Function

''' <summary>
''' Amarra el Parte Detalle a el Movimiento detalle correspondiente
''' </summary>
Function mHallaParteDetalle(IdMovimientoDetalle As Long, Optional mConexion As ADODB.Connection) As Long
    Dim mRecord As New ADODB.Recordset
    Dim database As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones

On Error GoTo BloqueError

    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    database.CommandText = "SELECT pro_producciondet.idproddet, alm_ingresodet.idmovdet, alm_ingresodet_1.idmovdet AS idmovdetanex, alm_ingresodet_1.iddocref " _
        + vbCr + "FROM (((pro_produccion INNER JOIN alm_ingreso ON pro_produccion.id = alm_ingreso.iddocref) INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_ingresodet AS alm_ingresodet_1 ON pro_producciondet.idproddet = alm_ingresodet_1.iddocref " _
        + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & IdMovimientoDetalle & ") AND ((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoParteProduccion", mConexion)) & ") AND ((pro_producciondet.iditem)=[alm_ingresodet].[iditem]) AND ((pro_producciondet.cantidad)=[alm_ingresodet].[cantidad]))"
        
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount = 0 Then
        Dim mNumDoc As String
        Dim mIdMovimiento As Long
        Dim mIdItem As Long
        Dim mCodigoItem As String
        Dim mItem As String
        Dim mCantidad As Double
        
        mIdItem = F.BuscaCodigoTabla(IdMovimientoDetalle, "idmovdet", "iditem", "alm_ingresodet", "N", mConexion)
        mCodigoItem = F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion)
        mItem = F.BuscaCodigoTabla(mIdItem, "id", "descripcion", "alm_inventario", "N", mConexion)
        mIdMovimiento = F.BuscaCodigoTabla(IdMovimientoDetalle, "idmovdet", "id", "alm_ingresodet", "N", mConexion)
        mCantidad = F.BuscaCodigoTabla(IdMovimientoDetalle, "idmovdet", "cantidad", "alm_ingresodet", "N", mConexion)
        mNumDoc = F.BuscaCodigoTabla(mIdMovimiento, "id", "numser", "alm_ingreso", "N", mConexion)
        mNumDoc = mNumDoc & "-" & F.BuscaCodigoTabla(mIdMovimiento, "id", "numdoc", "alm_ingreso", "N", mConexion)
        Err.Raise &HFFFFFF01, , "El detalle del Movimiento: " & mNumDoc & ", no cuenta con un Parte asociado en Producción" _
                        + vbCr + "Codigo de Item: " & mCodigoItem _
                        + vbCr + "Item: " & mItem _
                        + vbCr + "Cantidad: " & F.NuloString(mCantidad)
    Else
        mRecord.Filter = adFilterNone
        mRecord.Filter = "idmovdetanex = " & IdMovimientoDetalle
        If mRecord.RecordCount = 0 Then
            mRecord.Filter = adFilterNone
            mRecord.Filter = "idmovdetanex = 0 Or idmovdetanex = null"
            mRecord.MoveFirst
            ' Se amarra al primer movimiento encontrado
            database.ClearParameter
            database.CommandText = "UPDATE alm_ingresodet SET alm_ingresodet.iddocref = " & F.NuloNumeric(mRecord("idproddet")) & " " _
                            + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & IdMovimientoDetalle & "))"
            database.Execute
        End If
    End If
    mHallaParteDetalle = F.NuloNumeric(mRecord("idproddet"))
    Set database = Nothing
    Set mRecord = Nothing
    Exit Function

BloqueError:
    Set database = Nothing
    Set mRecord = Nothing
    mHallaParteDetalle = 0
    Err.Raise Err.Number, "[HallaParteDetalle] " & Err.Source, Err.Description
End Function

''' <summary>
''' Halla el costo del parte especificado
''' </summary>
Function mCalcularCostoParte(IdAlmacenProceso As Long, mIdParteProduccion As Long, _
                                Optional mConexion As ADODB.Connection) As Double
                                
    Dim mParteProd As New ProduccionEntidad.EParteProd
    Dim mLParteProdDet As New ProduccionEntidad.LEParteProdDet
    Dim mCostoTotalParteProd As Double
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set mLParteProdDet.Conexion = mConexion
    mLParteProdDet.LoadChild = False
    mLParteProdDet.Fetch mIdParteProduccion
    
    ' Detalles del Parte
    Dim ParteProdDet As EParteProdDet
    mCostoTotalParteProd = 0
    For Each ParteProdDet In mLParteProdDet
        mCostoTotalParteProd = mCostoTotalParteProd + mCalcularCostoParteDetalle(IdAlmacenProceso, ParteProdDet.IdParteProduccionDet, mConexion)
    Next
    mCalcularCostoParte = mCostoTotalParteProd
    Exit Function

BloqueError:
    Set mParteProd = Nothing
    Err.Raise Err.Number, "[mCalcularCostoParte] " & Err.Source, Err.Description
End Function

Function mCalcularCostoParteDetalle(IdAlmacenProceso As Long, IdParteProdDet As Long, Optional mConexion As ADODB.Connection) As Double
    Dim F As New SistemaLogica.Funciones
    Dim mParteProdDet As New ProduccionEntidad.EParteProdDet
    Dim mCostoTotalInsumo As Double
    Dim mCostoTotalMovimientos As Double
    Dim mContador As Long
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set mParteProdDet.Conexion = mConexion
    mParteProdDet.Fetch IdParteProdDet
    ' Insumos del Parte
    Dim ParteProdDetIns As New ProduccionEntidad.EParteProdDetIns
    mCostoTotalInsumo = 0
    
    If mParteProdDet.LParteProduccionDetIns.Count = 0 Then
        Dim mNumParteProd As String
        
        mNumParteProd = F.NuloString(F.BuscaCodigoTabla(mParteProdDet.IdParteProduccion, "id", "numser", "pro_produccion", "N", mConexion))
        mNumParteProd = mNumParteProd & "-" & F.NuloString(F.BuscaCodigoTabla(mParteProdDet.IdParteProduccion, "id", "numdoc", "pro_produccion", "N", mConexion))
        Err.Raise &HFFFFFF01, , "Los insumos del detalle del Parte de Produccion: " & mNumParteProd & " no cuentan con movimientos en el almacén " _
            + vbCr + "Fecha: " & F.NuloString(mParteProdDet.Fecha) _
            + vbCr + "Codigo de Item: " & mParteProdDet.CodigoItem _
            + vbCr + "Item: " & mParteProdDet.Item _
            + vbCr + "Cantidad: " & F.NuloString(mParteProdDet.CantidadProducida)
    End If
    
    mContador = 0
    For Each ParteProdDetIns In mParteProdDet.LParteProduccionDetIns
        mContador = mContador + 1
        DoEvents
        PForm.SetSubProgress "Proc.Parte Det.", ParteProdDetIns.CodigoItem & " - " & ParteProdDetIns.Item
        If mError Then
            Err.Raise vbObjectError + 1, Err.Source, Err.Description
        End If
        ' Movimientos del Insumo
        Dim ParteProdDetInsMov As New ProduccionEntidad.EParteProdDetInsMov
        mCostoTotalMovimientos = 0
        For Each ParteProdDetInsMov In ParteProdDetIns.LParteProduccionDetInsMov
            mCostoTotalMovimientos = mCostoTotalMovimientos + mCalcularCostoMovimiento(IdAlmacenProceso, ParteProdDetInsMov.IdMovimientoDetalle, mConexion)
        Next
        mCostoTotalInsumo = mCostoTotalInsumo + mCostoTotalMovimientos
    Next
    Set ParteProdDetIns = Nothing
    mCalcularCostoParteDetalle = mCostoTotalInsumo
        
    Exit Function

BloqueError:
    Set mParteProdDet = Nothing
    Err.Raise Err.Number, "[mCalcularCostoParteDetalle] " & Err.Source, Err.Description
End Function


''' <summary>
''' Valida si los Items asociados a un movimiento detalle tienen gastos distribuidos
''' </summary>
Function mMovimientoDetalleTieneGastos(IdMovimientoDetalle As Long, _
                                Optional mConexion As ADODB.Connection) As Boolean
    Dim F As New SistemaLogica.Funciones
    Dim mDataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set mDataBase.Connection = mConexion
    mDataBase.CommandText = "SELECT con_librocostotemp.idmovdet " _
        + vbCr + "FROM alm_ingresodet INNER JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet " _
        + vbCr + "WHERE (((alm_ingresodet.iditem) In ( " _
        + vbCr + "SELECT alm_ingresodet.iditem " _
        + vbCr + "FROM alm_ingresodet INNER JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet " _
        + vbCr + "WHERE (((alm_ingresodet.idmovdet) = " & IdMovimientoDetalle & ")) " _
        + vbCr + ")) AND ((con_librocostotemp.costomod)>0) AND ((con_librocostotemp.costocif)>0))"

    Set mRecord = mDataBase.GetRecordset
    If mRecord.RecordCount = 0 Then
        mMovimientoDetalleTieneGastos = False
    Else
        mMovimientoDetalleTieneGastos = True
    End If
    Exit Function

BloqueError:
    mMovimientoDetalleTieneGastos = False
    F.MostrarMensajeError Err.Description, "[MovimientoDetalleTieneGastos]", Err.Source, Err.Number
End Function

''' <summary>
''' Halla el importe unitario promedio del movimiento especificado
''' </summary>
Function mMovimientoDetalleCosteado(IdMovimientoDetalle As Long, _
                                ByRef CostoMovimiento As Double, _
                                ByRef CostoUnitarioPromedio As Double, _
                                Optional mConexion As ADODB.Connection, _
                                Optional Called As Boolean = False) As Boolean
    Dim F As New SistemaLogica.Funciones
    Dim mCostoMovimiento As Double
    Dim mCostoUnitarioPromedio As Double
    
On Error GoTo BloqueError
    ' Se verifica si es un movimiento de produccion
    ' Para validar el MOD y CIF
    Dim mMovimientoItem As New AlmacenEntidad.EMovimientoItem
    If Not mConexion Is Nothing Then Set mMovimientoItem.Conexion = mConexion
    mMovimientoItem.Fetch IdMovimientoDetalle
    ' Si es un parte de Produccion
    If mMovimientoDetalleTieneGastos(IdMovimientoDetalle, mConexion) Then
        If Called Then
            mCostoUnitarioPromedio = ((mMovimientoItem.CostoPrimo + mMovimientoItem.CostoMOD + mMovimientoItem.CostoCIF) / mMovimientoItem.Cantidad)
            If mCostoUnitarioPromedio > 0 Then
                CostoMovimiento = mMovimientoItem.CostoPrimo + mMovimientoItem.CostoMOD + mMovimientoItem.CostoCIF
                CostoUnitarioPromedio = mMovimientoItem.CostoUnitarioPromedio
                mMovimientoDetalleCosteado = True
            Else
                CostoMovimiento = 0
                CostoUnitarioPromedio = 0
                mMovimientoDetalleCosteado = False
            End If
        Else
            CostoMovimiento = 0
            CostoUnitarioPromedio = 0
            mMovimientoDetalleCosteado = False
        End If
        
    Else
        If mMovimientoItem.CostoPrimo > 0 Then
            CostoMovimiento = mMovimientoItem.CostoPrimo
            CostoUnitarioPromedio = mMovimientoItem.CostoUnitarioPromedio
            mMovimientoDetalleCosteado = True
        Else
            CostoMovimiento = 0
            CostoUnitarioPromedio = 0
            mMovimientoDetalleCosteado = False
        End If
    End If
    Exit Function
    
BloqueError:
    CostoMovimiento = 0
    CostoUnitarioPromedio = 0
    mMovimientoDetalleCosteado = False
    F.MostrarMensajeError Err.Description, "[CalcularCostoMovimientoSimple]", Err.Source, Err.Number
End Function

''' <summary>
''' Halla el importe unitario promedio del movimiento especificado
''' </summary>
Function mCalcularCostoMovimiento(IdAlmacenProceso As Long, IdMovimientoDetalle As Long, _
                                    Optional mConexion As ADODB.Connection) As Double
    Dim F As New SistemaLogica.Funciones
    Dim mIdAlmacen As Long
    Dim mMovDetalle As New AlmacenEntidad.EMovimientoDet
    Dim CostoUnitarioPromedio As Double
    Dim CostoPrimo As Double
    Dim mCostoMovimiento As Double
    Dim mCodigoItem As String
    Dim mIdItem As Long
    Dim mItem As String
    Dim mCantidad As Double
    Dim mNumeroMovimiento As String
    
On Error GoTo BloqueError
    mError = False
    If Not mConexion Is Nothing Then Set mMovDetalle.Conexion = mConexion
    
    ' Se obtiene el ojeto
    mMovDetalle.Fetch IdMovimientoDetalle
    ' Se halla el almacen del movimiento
    mIdAlmacen = F.NuloNumeric(F.BuscaCodigoTabla(mMovDetalle.IdMovimiento, "id", "idalm", "alm_ingreso", "N", mConexion))
    If mIdAlmacen = 0 Then
        mNumeroMovimiento = F.NuloString(F.BuscaCodigoTabla(mMovDetalle.IdMovimiento, "id", "numser", "alm_ingreso", "N", mConexion))
        mNumeroMovimiento = mNumeroMovimiento & "-" & F.NuloString(F.BuscaCodigoTabla(mMovDetalle.IdMovimiento, "id", "numdoc", "alm_ingreso", "N", mConexion))
           
        Err.Raise &HFFFFFF01, , "El movimiento no cuenta con almacen de referencia: " _
                        + vbCr + "Codigo de Item: " & mCodigoItem _
                        + vbCr + "Item: " & mMovDetalle.Item _
                        + vbCr + "Cantidad: " & F.NuloString(mMovDetalle.Cantidad) _
                        + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovDetalle.FechaMovimiento) _
                        + vbCr + "Documento: " & mNumeroMovimiento
    End If
    ' Validar si ya esta costeado
    If mMovimientoDetalleCosteado(IdMovimientoDetalle, mCostoMovimiento, CostoUnitarioPromedio, mConexion) Then ' Si esta costeado
         mCalcularCostoMovimiento = mCostoMovimiento
    Else ' Si no esta costeado
        ' Se calculan los costos unitarios hasta la fecha del movimiento
        mCalcularCostoUnitarioItem IdAlmacenProceso, mMovDetalle.IdItem, mIdAlmacen, , mMovDetalle.FechaMovimiento, mConexion
        ' Se vuelve a validar el costo del movimiento
        If Not mMovimientoDetalleCosteado(IdMovimientoDetalle, mCostoMovimiento, CostoUnitarioPromedio, mConexion, True) Then
            
            mNumeroMovimiento = F.NuloString(F.BuscaCodigoTabla(mMovDetalle.IdMovimiento, "id", "numser", "alm_ingreso", "N", mConexion))
            mNumeroMovimiento = mNumeroMovimiento & "-" & F.NuloString(F.BuscaCodigoTabla(mMovDetalle.IdMovimiento, "id", "numdoc", "alm_ingreso", "N", mConexion))
            
            mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mMovDetalle.IdItem, "id", "codpro", "alm_inventario", "N", mConexion))
            Err.Raise &HFFFFFF01, , "Error desconocido al intentar costear el movimiento: " _
                        + vbCr + "Codigo de Item: " & mCodigoItem _
                        + vbCr + "Item: " & mMovDetalle.Item _
                        + vbCr + "Cantidad: " & F.NuloString(mMovDetalle.Cantidad) _
                        + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovDetalle.FechaMovimiento) _
                        + vbCr + "Documento: " & mNumeroMovimiento
        End If
        mCalcularCostoMovimiento = mCostoMovimiento
    End If
    Exit Function

BloqueError:
    mCalcularCostoMovimiento = 0
    Set mMovDetalle = Nothing
    mError = True
    F.MostrarMensajeError Err.Description, "CalcularCostoMovimiento", Err.Source, Err.Number
End Function

''' <summary>
''' Halla el importe unitario promedio del item especificado y el de sus movimientos
''' </summary>
Function mCalcularCostoUnitarioItem(mIdAlmacenProceso As Long, _
                            mIdItem As Long, _
                            mIdAlmacen As Long, _
                            Optional FechaInicio As Date, _
                            Optional FechaFin As Date, _
                            Optional mConexion As ADODB.Connection) As Double
                            
    Dim F As New SistemaLogica.Funciones
    Dim mLMovimientoItem As New AlmacenEntidad.LEMovimientoItem
    Dim mCostoTotal As Double
    Dim mCantidadAcumulada As Double
    Dim mCantidadAcumuladaEntrada As Double
    Dim mCantidadAcumuladaSalida As Double
    Dim mCostoAcumulado As Double
    Dim mCostoUnitarioPromedio As Double
    Dim mCostoUnitarioMovimiento As Double
    Dim mCostoMovimiento As Double
    Dim mCostoUnitarioPromedioAux As Double
    Dim mCostoMovimientoAux As Double
    Dim mCodigoItem As String
    Dim mContador As Long
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set mLMovimientoItem.Conexion = mConexion
        
    mLMovimientoItem.Fetch mIdItem, mIdAlmacen, FechaInicio, FechaFin
    If mLMovimientoItem.Count <= 0 Then
        Dim mItem As String
        mItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "descripcion", "alm_inventario", "N", mConexion))
        mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
        Err.Raise &HFFFFFF01, , "Inconsistencia en cantidad, el Item no cuenta con movimientos a la fecha." _
                            + vbCr + "Codigo de Item: " & mCodigoItem _
                            + vbCr + "Item: " & mItem _
                            + vbCr + "Fecha: " & F.NuloString(FechaFin)
    End If
    mCostoTotal = 0
    
    mContador = 0
    Dim mMovimientoItem As AlmacenEntidad.EMovimientoItem
    For Each mMovimientoItem In mLMovimientoItem
        DoEvents
        mContador = mContador + 1
        PForm.SetSubProgress "Proc. Mov.", mMovimientoItem.FechaMovimiento & " - " & mMovimientoItem.Item
                
        ' INGRESOS
        If mMovimientoItem.TipoMovimiento = "I" Then
            '*****************************
            mCantidadAcumulada = mCantidadAcumulada + mMovimientoItem.Cantidad
            mCantidadAcumuladaEntrada = mCantidadAcumuladaEntrada + mMovimientoItem.Cantidad
            ' Validar si ya esta costeado
             If Not mMovimientoDetalleCosteado(mMovimientoItem.IdMovimientoDet, mCostoMovimiento, mCostoUnitarioPromedioAux, mConexion, True) Then ' Si no esta costeado
                ' Se valida el tipo de documento de referencia
                Select Case mMovimientoItem.IdTipoDocumentoReferenciaPadre
                    ' Partes de Produccion
                    Case F.NuloNumeric(F.KeyValue("ParteProduccion", mConexion))
                        If mMovimientoItem.IdDocumentoReferencia = 0 Then
                            mMovimientoItem.IdDocumentoReferencia = mHallaParteDetalle(mMovimientoItem.IdMovimientoDet, mConexion)
                        End If
                        mCostoMovimiento = mCalcularCostoParteDetalle(mIdAlmacenProceso, mMovimientoItem.IdDocumentoReferencia, mConexion)
                        ' Se agregar los gastos distribuidos
                        mCostoMovimiento = mCostoMovimiento + mMovimientoItem.CostoMOD + mMovimientoItem.CostoCIF
                        mCostoUnitarioMovimiento = mCostoMovimiento / mMovimientoItem.Cantidad
                    
                    ' Solicitud de Materiales
                    Case F.NuloNumeric(F.KeyValue("SolictudMateriales", mConexion))
                        mCostoMovimiento = mCostoUnitarioPromedio * mMovimientoItem.Cantidad
                        mCostoUnitarioMovimiento = mCostoUnitarioPromedio
                    
                    ' Inventarios Iniciales
                    Case F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", mConexion))
                        mCostoUnitarioMovimiento = mHallaCostoUnitarioSimple(mMovimientoItem.IdMovimientoDet, mMovimientoItem.IdItem, "II", mConexion)
                        mCostoMovimiento = mCostoUnitarioMovimiento * mMovimientoItem.Cantidad
                    
                    ' Ajustes de Inventario
                    Case F.NuloNumeric(F.KeyValue("IdDocumentoAjusteInventario", mConexion))
                        mCostoUnitarioMovimiento = mHallaCostoUnitarioSimple(mMovimientoItem.IdMovimientoDet, mMovimientoItem.IdItem, "AJ", mConexion)
                        mCostoMovimiento = mCostoUnitarioMovimiento * mMovimientoItem.Cantidad
                        
                    'Notas de Credito
                    Case F.NuloNumeric(F.KeyValue("IdDocumentoFactura", mConexion))
                        ' Validamos el costo promedio
                        If mCostoUnitarioPromedio = 0 Then
                            mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                            Err.Raise &HFFFFFF01, , "Costo Unitario Promedio igual a cero al intentar costear Factura. " _
                                        + vbCr + "Codigo de Item: " & mCodigoItem _
                                        + vbCr + "Item: " & mMovimientoItem.Item _
                                        + vbCr + "Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                                        + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                        + vbCr + "Documento: " & mMovimientoItem.Documento
                        End If
                        mCostoMovimiento = mCostoUnitarioPromedio * mMovimientoItem.Cantidad
                        mCostoUnitarioMovimiento = mCostoUnitarioPromedio
                        
                    Case Else
                        mCostoUnitarioMovimiento = mHallaCostoUnitarioSimple(mMovimientoItem.IdMovimientoDet, mMovimientoItem.IdItem, "AI", mConexion)
                        mCostoMovimiento = mCostoUnitarioMovimiento * mMovimientoItem.Cantidad
                        
                End Select
                
                ' Validamos el costo unitario del movimiento
                If mCostoUnitarioMovimiento <= 0 Then
                    mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                    Err.Raise &HFFFFFF01, , "Costo Unitario de movimiento igual a cero. " _
                                + vbCr + "Codigo de Item: " & mCodigoItem _
                                + vbCr + "Item: " & mMovimientoItem.Item _
                                + vbCr + "Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                                + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                + vbCr + "Documento: " & mMovimientoItem.Documento
                End If
                
                mCostoAcumulado = mCostoAcumulado + mCostoMovimiento
                If mCantidadAcumulada > 0 Then mCostoUnitarioPromedio = mCostoAcumulado / mCantidadAcumulada
                
                ' Validamos el costo unitario promedio
                If mCostoUnitarioPromedio = 0 Then
                    mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                    Err.Raise &HFFFFFF01, , "Costo Unitario Promedio igual a cero. " _
                                + vbCr + "Codigo de Item: " & mCodigoItem _
                                + vbCr + "Item: " & mMovimientoItem.Item _
                                + vbCr + "Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                                + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                + vbCr + "Documento: " & mMovimientoItem.Documento
                End If
                ' Grabamos el costo del movimiento
                If Not mGrabaCostoMovTemp(mIdAlmacenProceso, mMovimientoItem.IdMovimientoDet, mMovimientoItem.Cantidad, _
                            mCostoUnitarioMovimiento, mCostoUnitarioPromedio, mCostoMovimiento - mMovimientoItem.CostoMOD - mMovimientoItem.CostoCIF, _
                            mMovimientoItem.CostoMOD, mMovimientoItem.CostoCIF, mMovimientoItem.FechaMovimiento, _
                            mMovimientoItem.TipoMovimiento, mConexion) Then
                            
                    mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                    Err.Raise &HFFFFFF01, , "Error interno al intentar grabar el costo del movimiento. " _
                                            + vbCr + "Codigo de Item: " & mCodigoItem _
                                            + vbCr + "Item: " & mMovimientoItem.Item _
                                            + vbCr + "Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                                            + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                            + vbCr + "Documento: " & mMovimientoItem.Documento
                End If
                
            Else
                ' Buscamos el importe del movimiento
                mCostoAcumulado = mCostoAcumulado + mCostoMovimiento
                If mCantidadAcumulada > 0 Then
                    mCostoUnitarioPromedio = mCostoAcumulado / mCantidadAcumulada
                    mCostoUnitarioMovimiento = mCostoMovimiento / mMovimientoItem.Cantidad
                End If
                
                ' Se valida que este bien costeado
                If mCostoUnitarioPromedio = 0 Then
                    mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                    Err.Raise &HFFFFFF01, , "Costo Unitario Promedio igual a cero. " _
                                            + vbCr + "Codigo de Item: " & mCodigoItem _
                                            + vbCr + "Item: " & mMovimientoItem.Item _
                                            + vbCr + "Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                                            + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                            + vbCr + "Documento: " & mMovimientoItem.Documento
                ElseIf Abs(mCostoUnitarioPromedio - mMovimientoItem.CostoUnitarioPromedio) > 0.01 Then ' Si esta mal costeado
                    ' Costeamos el movimiento
                    If Not mGrabaCostoMovTemp(mIdAlmacenProceso, mMovimientoItem.IdMovimientoDet, mMovimientoItem.Cantidad, _
                                mCostoUnitarioMovimiento, mCostoUnitarioPromedio, mCostoMovimiento - mMovimientoItem.CostoMOD - mMovimientoItem.CostoCIF, _
                                mMovimientoItem.CostoMOD, mMovimientoItem.CostoCIF, mMovimientoItem.FechaMovimiento, _
                                mMovimientoItem.TipoMovimiento, mConexion) Then
                                
                        mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                        Err.Raise &HFFFFFF01, , "Error interno al intentar grabar el costo del movimiento. " _
                                                + vbCr + "Codigo de Item: " & mCodigoItem _
                                                + vbCr + "Item: " & mMovimientoItem.Item _
                                                + vbCr + "Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                                                + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                                + vbCr + "Documento: " & mMovimientoItem.Documento
                    End If
                End If
            End If
        
        Else ' SALIDAS
            mCostoMovimiento = mCostoUnitarioPromedio * mMovimientoItem.Cantidad
            mCantidadAcumulada = mCantidadAcumulada - mMovimientoItem.Cantidad
            mCantidadAcumuladaSalida = mCantidadAcumuladaSalida + mMovimientoItem.Cantidad
            mCostoUnitarioMovimiento = mCostoUnitarioPromedio
                                
            If CDbl(Format(mCantidadAcumulada, "0.00000")) < 0 Then
                mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                Err.Raise &HFFFFFF01, , "Cantidad acumulada menor a cero. " _
                                        + vbCr + "Codigo de Item: " & mCodigoItem _
                                        + vbCr + "Item: " & mMovimientoItem.Item _
                                        + vbCr + "Cantidad: " & F.NuloString(mCantidadAcumulada) _
                                        + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                        + vbCr + "Documento: " & mMovimientoItem.Documento
            End If
            
            ' Si no esta costeado
            If Not mMovimientoDetalleCosteado(mMovimientoItem.IdMovimientoDet, mCostoMovimientoAux, mCostoUnitarioPromedioAux, mConexion, True) Then
                ' Costeamos el movimiento
                If Not mGrabaCostoMovTemp(mIdAlmacenProceso, mMovimientoItem.IdMovimientoDet, mMovimientoItem.Cantidad, _
                            mCostoUnitarioMovimiento, mCostoUnitarioPromedio, mCostoMovimiento, _
                            0, 0, mMovimientoItem.FechaMovimiento, _
                            mMovimientoItem.TipoMovimiento, mConexion) Then
                            
                    mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                    Err.Raise &HFFFFFF01, , "Error interno al intentar grabar el costo del movimiento. " _
                                            + vbCr + "Codigo de Item: " & mCodigoItem _
                                            + vbCr + "Item: " & mMovimientoItem.Item _
                                            + vbCr + "Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                                            + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                            + vbCr + "Documento: " & mMovimientoItem.Documento
                End If
            Else
                mCostoUnitarioMovimiento = mCostoMovimientoAux / mMovimientoItem.Cantidad
                ' Si esta mal costeado
                If mCostoUnitarioPromedio <> mCostoUnitarioPromedioAux Then
                    ' Costeamos el movimiento
                    If Not mGrabaCostoMovTemp(mIdAlmacenProceso, mMovimientoItem.IdMovimientoDet, mMovimientoItem.Cantidad, _
                                mCostoUnitarioMovimiento, mCostoUnitarioPromedio, mCostoMovimiento, _
                                0, 0, mMovimientoItem.FechaMovimiento, _
                                mMovimientoItem.TipoMovimiento, mConexion) Then
                                
                        mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                        Err.Raise &HFFFFFF01, , "Error interno al intentar grabar el costo del movimiento. " _
                                                + vbCr + "Codigo de Item: " & mCodigoItem _
                                                + vbCr + "Item: " & mMovimientoItem.Item _
                                                + vbCr + "Cantidad: " & F.NuloString(mMovimientoItem.Cantidad) _
                                                + vbCr + "Fecha de Movimiento: " & F.NuloString(mMovimientoItem.FechaMovimiento) _
                                                + vbCr + "Documento: " & mMovimientoItem.Documento
                    End If
                End If
            End If
            
            ' IMPORTE ACUMULADO
            mCostoAcumulado = mCostoAcumulado - mCostoMovimiento
        End If
    Next
    mCalcularCostoUnitarioItem = mCostoUnitarioPromedio
    Exit Function

BloqueError:
    Set mLMovimientoItem = Nothing
    Err.Raise Err.Number, "[mCalcularCostoUnitarioItem] " & Err.Source, Err.Description
End Function

''' <summary>
''' Halla el importe unitario de compra de un item
''' </summary>
Function mHallaCostoUnitarioSimple(IdDocumento As Long, IdItem As Long, DondeBuscar As String, Optional mConexion As ADODB.Connection) As Double
    Dim xRst As New ADODB.Recordset
    Dim nSQL As String
    Dim mItem As String
    Dim mIdMov As Long
    Dim mNumDoc As String
    Dim mFechMov As String
    Dim mDataBase As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set mDataBase.Connection = mConexion
    
    mItem = F.NuloString(Busca_Codigo(IdItem, "id", "descripcion", "alm_inventario", "N", mConexion))
    If DondeBuscar = "AI" Then
        nSQL = "SELECT Sum(IIf([com_compras].[idmon]=2,[com_comprasdet].[imptot]*[con_tc].[impcom],[com_comprasdet].[imptot])) AS importetot, Sum(com_comprasdet.canpro) AS cantidadtot " _
            + vbCr + "FROM ((com_compras INNER JOIN (com_comprasdet INNER JOIN alm_ingresodoc ON com_comprasdet.idcom = alm_ingresodoc.iddoc) ON com_compras.id = com_comprasdet.idcom) INNER JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN alm_ingresodet ON alm_ingresodoc.id = alm_ingresodet.id " _
            + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & IdDocumento & ") AND ((com_comprasdet.iditem)=" & IdItem & "))"
        
        Set xRst = Nothing
        mDataBase.CommandText = nSQL
        Set xRst = mDataBase.GetRecordset
        If xRst.RecordCount = 0 Then
            mIdMov = F.NuloNumeric(Busca_Codigo(IdDocumento, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
            mNumDoc = F.NuloString(Busca_Codigo(mIdMov, "id", "numser", "alm_ingreso", "N", mConexion))
            mNumDoc = mNumDoc & "-" & F.NuloString(Busca_Codigo(mIdMov, "id", "numdoc", "alm_ingreso", "N", mConexion))
            mFechMov = F.NuloString(Busca_Codigo(mIdMov, "id", "fchdoc", "alm_ingreso", "N", mConexion))
            Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en compras o no existe registro de compras. " _
                                    + vbCr + "Item: " & mItem _
                                    + vbCr + "Movimiento: " & mNumDoc _
                                    + vbCr + "Fecha: " & mFechMov
        End If
        
        If NulosN(xRst("cantidadtot")) = 0 Then
            mIdMov = F.NuloNumeric(Busca_Codigo(IdDocumento, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
            mNumDoc = F.NuloString(Busca_Codigo(mIdMov, "id", "numser", "alm_ingreso", "N", mConexion))
            mNumDoc = mNumDoc & "-" & F.NuloString(Busca_Codigo(mIdMov, "id", "numdoc", "alm_ingreso", "N", mConexion))
            mFechMov = F.NuloString(Busca_Codigo(mIdMov, "id", "fchdoc", "alm_ingreso", "N", mConexion))
            Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en compras o no existe registro de compras. " _
                                    + vbCr + "Item: " & mItem _
                                    + vbCr + "Movimiento: " & mNumDoc _
                                    + vbCr + "Fecha: " & mFechMov
        Else
            mHallaCostoUnitarioSimple = NulosN(xRst("importetot")) / NulosN(xRst("cantidadtot"))
        End If
        Exit Function
        
    ElseIf DondeBuscar = "II" Then
        mIdMov = F.NuloNumeric(Busca_Codigo(IdDocumento, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
        
        nSQL = "SELECT alm_inventarioinicialdet.costo " _
            + vbCr + "FROM (alm_inventarioinicial INNER JOIN (alm_ingresodet INNER JOIN alm_ingreso ON alm_ingresodet.id = alm_ingreso.id) ON alm_inventarioinicial.idinventarioinicial = alm_ingreso.iddocref) INNER JOIN alm_inventarioinicialdet ON alm_inventarioinicial.idinventarioinicial = alm_inventarioinicialdet.idinventarioinicial " _
            + vbCr + "WHERE (((alm_inventarioinicialdet.iditem)=" & IdItem & ") AND ((alm_ingresodet.idmovdet)=" & IdDocumento & ") AND ((alm_inventarioinicial.idestado)=" & F.NuloNumeric(F.KeyValue("EstadoAprobadoInventarioInicial", mConexion)) & ") AND ((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", mConexion)) & "))"
        
        Set xRst = Nothing
        mDataBase.CommandText = nSQL
        Set xRst = mDataBase.GetRecordset
        If xRst.RecordCount = 0 Then
            mIdMov = F.NuloNumeric(Busca_Codigo(IdDocumento, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
            mNumDoc = F.NuloString(Busca_Codigo(mIdMov, "id", "numser", "alm_ingreso", "N", mConexion))
            mNumDoc = mNumDoc & "-" & F.NuloString(Busca_Codigo(mIdMov, "id", "numdoc", "alm_ingreso", "N", mConexion))
            mFechMov = F.NuloString(Busca_Codigo(mIdMov, "id", "fchdoc", "alm_ingreso", "N", mConexion))
            Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en Inventario Inicial. " _
                                    + vbCr + "Item: " & mItem _
                                    + vbCr + "Movimiento: " & mNumDoc _
                                    + vbCr + "Fecha: " & mFechMov
        End If
        
        If NulosN(xRst("costo")) = 0 Then
            mIdMov = F.NuloNumeric(Busca_Codigo(IdDocumento, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
            mNumDoc = F.NuloString(Busca_Codigo(mIdMov, "id", "numser", "alm_ingreso", "N", mConexion))
            mNumDoc = mNumDoc & "-" & F.NuloString(Busca_Codigo(mIdMov, "id", "numdoc", "alm_ingreso", "N", mConexion))
            mFechMov = F.NuloString(Busca_Codigo(mIdMov, "id", "fchdoc", "alm_ingreso", "N", mConexion))
            Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en Inventario Inicial. " _
                                    + vbCr + "Item: " & mItem _
                                    + vbCr + "Movimiento: " & mNumDoc _
                                    + vbCr + "Fecha: " & mFechMov
        Else
            mHallaCostoUnitarioSimple = NulosN(xRst("costo"))
        End If
        Exit Function
        
    ElseIf DondeBuscar = "AJ" Then
        mIdMov = F.NuloNumeric(Busca_Codigo(IdDocumento, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
        
        nSQL = "SELECT alm_tomainventariodet.preuni AS costo " _
            + vbCr + "FROM (alm_ingreso INNER JOIN (alm_tomainventario INNER JOIN alm_tomainventariodet ON alm_tomainventario.idtomainventario = alm_tomainventariodet.idtomainventario) ON alm_ingreso.iddocref = alm_tomainventario.idtomainventario) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id " _
            + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & IdDocumento & ") AND ((alm_tomainventariodet.iditem)=" & IdItem & ") AND ((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoAjusteInventario", mConexion)) & ") AND ((alm_tomainventario.idestadoinventario)=" & F.NuloNumeric(F.KeyValue("EstadoAprobadoInventario", mConexion)) & "))"
        
        Set xRst = Nothing
        mDataBase.CommandText = nSQL
        Set xRst = mDataBase.GetRecordset
        If xRst.RecordCount = 0 Then
            mIdMov = F.NuloNumeric(Busca_Codigo(IdDocumento, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
            mNumDoc = F.NuloString(Busca_Codigo(mIdMov, "id", "numser", "alm_ingreso", "N", mConexion))
            mNumDoc = mNumDoc & "-" & F.NuloString(Busca_Codigo(mIdMov, "id", "numdoc", "alm_ingreso", "N", mConexion))
            mFechMov = F.NuloString(Busca_Codigo(mIdMov, "id", "fchdoc", "alm_ingreso", "N", mConexion))
            Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en Ajuste de Inventario. " _
                                    + vbCr + "Item: " & mItem _
                                    + vbCr + "Movimiento: " & mNumDoc _
                                    + vbCr + "Fecha: " & mFechMov
        End If
        
        If NulosN(xRst("costo")) = 0 Then
            ' Se busca el ultimo costo del item
            nSQL = "SELECT TOP 1 com_comprasdet.preuni AS costo " _
                + vbCr + "FROM com_compras INNER JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom " _
                + vbCr + "WHERE (((com_comprasdet.IdItem) = " & IdItem & ")) " _
                + vbCr + "ORDER BY com_compras.fchdoc DESC"
            
            Set xRst = Nothing
            mDataBase.ClearParameter
            mDataBase.CommandText = nSQL
            Set xRst = mDataBase.GetRecordset
            If xRst.RecordCount = 0 Then
                mIdMov = F.NuloNumeric(Busca_Codigo(IdDocumento, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
                mNumDoc = F.NuloString(Busca_Codigo(mIdMov, "id", "numser", "alm_ingreso", "N", mConexion))
                mNumDoc = mNumDoc & "-" & F.NuloString(Busca_Codigo(mIdMov, "id", "numdoc", "alm_ingreso", "N", mConexion))
                mFechMov = F.NuloString(Busca_Codigo(mIdMov, "id", "fchdoc", "alm_ingreso", "N", mConexion))
                Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en Ajuste de Inventario. No se encontraron movimientos para hallar el ultimo costo" _
                                        + vbCr + "Item: " & mItem _
                                        + vbCr + "Movimiento: " & mNumDoc _
                                        + vbCr + "Fecha: " & mFechMov
            ElseIf NulosN(xRst("costo")) = 0 Then
                mIdMov = F.NuloNumeric(Busca_Codigo(IdDocumento, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
                mNumDoc = F.NuloString(Busca_Codigo(mIdMov, "id", "numser", "alm_ingreso", "N", mConexion))
                mNumDoc = mNumDoc & "-" & F.NuloString(Busca_Codigo(mIdMov, "id", "numdoc", "alm_ingreso", "N", mConexion))
                mFechMov = F.NuloString(Busca_Codigo(mIdMov, "id", "fchdoc", "alm_ingreso", "N", mConexion))
                Err.Raise &HFFFFFF01, , "Costo de Item igual a cero en último registro de compras. " _
                                        + vbCr + "Item: " & mItem _
                                        + vbCr + "Movimiento: " & mNumDoc _
                                        + vbCr + "Fecha: " & mFechMov
            Else
                mHallaCostoUnitarioSimple = NulosN(xRst("costo"))
            End If
            
        Else
            mHallaCostoUnitarioSimple = NulosN(xRst("costo"))
        End If
        Exit Function
        
    ElseIf DondeBuscar = "GR" Then
        nSQL = "SELECT vta_guia.id, vta_ventasdet.iditem, Avg(vta_ventasdet.preuni) AS preuniprom " _
            + vbCr + " FROM vta_guia INNER JOIN vta_ventasdet ON vta_guia.iddocven = vta_ventasdet.idvta " _
            + vbCr + " GROUP BY vta_guia.id, vta_ventasdet.iditem " _
            + vbCr + " HAVING (((vta_guia.id)=" & IdDocumento & ") AND ((vta_ventasdet.iditem)=" & IdItem & ")); "
       
        Set xRst = Nothing
        RST_Busq xRst, nSQL, mConexion
        If xRst.State = 0 Then mHallaCostoUnitarioSimple = 0: Exit Function
        If xRst.RecordCount = 0 Then mHallaCostoUnitarioSimple = 0: Exit Function
        
        mHallaCostoUnitarioSimple = NulosN(xRst("preuniprom"))
        Exit Function
    Else
        mHallaCostoUnitarioSimple = 0
        Exit Function
    End If
    Exit Function

BloqueError:
    Set xRst = Nothing
    Err.Raise Err.Number, "[Preuni] " & Err.Source, Err.Description
End Function

''' <summary>
''' Graba el registro de kardex de un movimiento
''' </summary>
Function mGrabaCostoMovTemp(IdAlmacenProceso As Long, _
                        IdMovimientoDet As Long, Cantidad As Double, _
                        CostoUnitario As Double, CostoUnitarioPromedio As Double, _
                        CostoPrimo As Double, CostoMOD As Double, _
                        CostoCIF As Double, FechaMovimiento As Date, _
                        TipoMovimiento As String, mConexion As ADODB.Connection) As Boolean
    ' Se busca el registro de kardex asociado
    Dim mCostoTemp As New ContabilidadEntidad.ELibroCostoTemp
    Dim F As New SistemaLogica.Funciones
    Dim mCodigoItem As String
    Dim mIdItem As Long
    Dim mItem As String
    Dim mIdMovimiento As Long
    Dim mNumDoc As String
            
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set mCostoTemp.Conexion = mConexion
    ' Se trae el registro si existiese
    mCostoTemp.Fetch IdMovimientoDet
    ' Se modifican los datos
    'mCostoTemp
    mCostoTemp.IdAlmacenProceso = IdAlmacenProceso
    mCostoTemp.IdMovimientoDetalle = IdMovimientoDet
    mCostoTemp.TipoMovimiento = TipoMovimiento
    mCostoTemp.FechaMovimiento = FechaMovimiento
    mCostoTemp.Cantidad = Cantidad
    If CostoUnitario > 0 And mCostoTemp.CostoUnitario <> CostoUnitario Then
        mCostoTemp.CostoUnitario = CostoUnitario
    End If
    If CostoUnitarioPromedio > 0 And mCostoTemp.CostoUnitarioPromedio <> CostoUnitarioPromedio Then
        mCostoTemp.CostoUnitarioPromedio = CostoUnitarioPromedio
    End If
    mCostoTemp.CostoPrimo = CostoPrimo
    mCostoTemp.CostoMOD = CostoMOD
    mCostoTemp.CostoCIF = CostoCIF
    ' Se graba el registro
    If Not mConexion Is Nothing Then Set mCostoTemp.Conexion = mConexion
    If Not mCostoTemp.Save(0, "") Then
        mIdMovimiento = F.NuloString(F.BuscaCodigoTabla(IdMovimientoDet, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
        mNumDoc = F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numser", "alm_ingreso", "N", mConexion))
        mNumDoc = mNumDoc & "-" & F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numdoc", "alm_ingreso", "N", mConexion))
        mIdItem = F.NuloString(F.BuscaCodigoTabla(IdMovimientoDet, "idmovdet", "iditem", "alm_ingresodet", "N", mConexion))
        mCodigoItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "codpro", "alm_inventario", "N", mConexion))
        mItem = F.NuloString(F.BuscaCodigoTabla(mIdItem, "id", "descripcion", "alm_inventario", "N", mConexion))
        Err.Raise &HFFFFFF01, , "Error al intentar grabar el kardex. " _
                                    + vbCr + "Codigo de Item: " & mCodigoItem _
                                    + vbCr + "Item: " & mItem _
                                    + vbCr + "Movimiento: " & mNumDoc _
                                    + vbCr + "Fecha: " & FechaMovimiento
    End If
    mGrabaCostoMovTemp = True
    Exit Function

BloqueError:
    Set mCostoTemp = Nothing
    mGrabaCostoMovTemp = False
    F.MostrarMensajeError Err.Description, "[GrabaCostoMovTemp]", Err.Source, Err.Number
End Function

''' <summary>
''' Graba el registro de kardex de un movimiento
''' </summary>
Function mGrabaKardex(IdItem As Long, Cantidad As Double, _
                        CostoUnitario As Double, CostoUnitarioPromedio As Double, _
                        IdAlmacen As Long, FechaMovimiento As Date, _
                        IdMovimientoDet As Long, TipoMovimiento As String, _
                        mConexion As ADODB.Connection) As Boolean
    ' Se busca el registro de kardex asociado
    Dim Kardx As New AlmacenEntidad.EKardex
    Dim F As New SistemaLogica.Funciones
    Dim mCodigoItem As String
    Dim mItem As String
    Dim mIdMovimiento As Long
    Dim mNumDoc As String
            
    Kardx.LoadChild = False
    If Not mConexion Is Nothing Then Set Kardx.Conexion = mConexion
    Kardx.Fetch 0, IdItem
    If Kardx.IdKardex = 0 Then ' Si es un kardex nuevo
        Kardx.IdItem = IdItem
        Kardx.Cantidad = Cantidad
        Kardx.CostoUnitario = CostoUnitario
        Kardx.CostoUnitarioPromedio = CostoUnitarioPromedio
        ' Se crea al hijo
        Dim KardxDet As New AlmacenEntidad.EKardexDet
        KardxDet.IdAlmacen = IdAlmacen
        KardxDet.UltimaFecha = FechaMovimiento
        KardxDet.Cantidad = Cantidad
        KardxDet.CostoUnitario = CostoUnitario
        KardxDet.CostoUnitarioPromedio = CostoUnitarioPromedio
        ' Se creal al nieto
        Dim KardxDetMov As New AlmacenEntidad.EKardexDetMov
        KardxDetMov.IdMovimientoDetalle = IdMovimientoDet
        KardxDetMov.TipoMovimiento = TipoMovimiento
        KardxDetMov.FechaMovimiento = FechaMovimiento
        KardxDetMov.Cantidad = Cantidad
        KardxDetMov.CostoUnitario = CostoUnitario
        KardxDetMov.CostoUnitarioPromedio = CostoUnitarioPromedio
        ' Se agregan a los padres
        KardxDet.LKardexDetalleMov.Add KardxDetMov
        Kardx.LKardexDetalle.Add KardxDet
        ' Se registra el kardex
        If Not mConexion Is Nothing Then Set Kardx.Conexion = mConexion
        If Not Kardx.Save(0, "") Then
            
            mIdMovimiento = F.NuloString(F.BuscaCodigoTabla(IdMovimientoDet, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
            mNumDoc = F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numser", "alm_ingreso", "N", mConexion))
            mNumDoc = mNumDoc & "-" & F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numdoc", "alm_ingreso", "N", mConexion))
            mCodigoItem = F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "codpro", "alm_inventario", "N", mConexion))
            mItem = F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "descripcion", "alm_inventario", "N", mConexion))
            Err.Raise &HFFFFFF01, , "Error al intentar grabar el kardex. " _
                                        + vbCr + "Codigo de Item: " & mCodigoItem _
                                        + vbCr + "Item: " & mItem _
                                        + vbCr + "Movimiento: " & mNumDoc _
                                        + vbCr + "Fecha: " & FechaMovimiento
        End If
    Else ' Si ya existe un kardex para esa mercaderia
        ' Se busca un kardex detalle para ese almacen
        Dim mKardxDet As New AlmacenEntidad.EKardexDet
        mKardxDet.LoadChild = False
        If Not mConexion Is Nothing Then Set mKardxDet.Conexion = mConexion
        mKardxDet.Fetch 0, Kardx.IdKardex, IdAlmacen
        If mKardxDet.IdKardexDetalle = 0 Then ' Si es un registro nuevo
            mKardxDet.IdKardex = Kardx.IdKardex
            mKardxDet.IdAlmacen = IdAlmacen
            mKardxDet.UltimaFecha = FechaMovimiento
            mKardxDet.Cantidad = Cantidad
            mKardxDet.CostoUnitario = CostoUnitario
            mKardxDet.CostoUnitarioPromedio = CostoUnitarioPromedio
            ' Se crea al al nieto
            Dim nKardxDetMov As New AlmacenEntidad.EKardexDetMov
            nKardxDetMov.IdMovimientoDetalle = IdMovimientoDet
            nKardxDetMov.TipoMovimiento = TipoMovimiento
            nKardxDetMov.FechaMovimiento = FechaMovimiento
            nKardxDetMov.Cantidad = Cantidad
            nKardxDetMov.CostoUnitario = CostoUnitario
            nKardxDetMov.CostoUnitarioPromedio = CostoUnitarioPromedio
            nKardxDetMov.UpdateParent = True
            ' Se agrega al apdre
            mKardxDet.LKardexDetalleMov.Add nKardxDetMov
            ' Se registra el kardex detalle
            If Not mConexion Is Nothing Then Set mKardxDet.Conexion = mConexion
            If Not mKardxDet.Save(0, "") Then
                mIdMovimiento = F.NuloString(F.BuscaCodigoTabla(IdMovimientoDet, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
                mNumDoc = F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numser", "alm_ingreso", "N", mConexion))
                mNumDoc = mNumDoc & "-" & F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numdoc", "alm_ingreso", "N", mConexion))
                mCodigoItem = F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                mItem = F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "descripcion", "alm_inventario", "N", mConexion))
                Err.Raise &HFFFFFF01, , "Error al intentar grabar el kardex. " _
                                            + vbCr + "Codigo de Item: " & mCodigoItem _
                                            + vbCr + "Item: " & mItem _
                                            + vbCr + "Movimiento: " & mNumDoc _
                                            + vbCr + "Fecha: " & FechaMovimiento
            End If
        Else ' si es un registro existente
            ' Se busca un kardex movimiento
            Dim mKardxDetMov As New AlmacenEntidad.EKardexDetMov
            If Not mConexion Is Nothing Then Set mKardxDetMov.Conexion = mConexion
            mKardxDetMov.Fetch 0, mKardxDet.IdKardexDetalle, IdMovimientoDet
            If mKardxDetMov.IdKardexDetalleMov = 0 Then ' Si es un nuevo registro
                mKardxDetMov.IdKardexDetalle = mKardxDet.IdKardexDetalle
                mKardxDetMov.IdMovimientoDetalle = IdMovimientoDet
                mKardxDetMov.TipoMovimiento = TipoMovimiento
                mKardxDetMov.FechaMovimiento = FechaMovimiento
                mKardxDetMov.Cantidad = Cantidad
                mKardxDetMov.CostoUnitario = CostoUnitario
                mKardxDetMov.CostoUnitarioPromedio = CostoUnitarioPromedio
                ' Se registra e movimiento de kardex
                If Not mConexion Is Nothing Then Set mKardxDetMov.Conexion = mConexion
                mKardxDetMov.UpdateParent = True
                If Not mKardxDetMov.Save(0, "") Then
                    mIdMovimiento = F.NuloString(F.BuscaCodigoTabla(IdMovimientoDet, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
                    mNumDoc = F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numser", "alm_ingreso", "N", mConexion))
                    mNumDoc = mNumDoc & "-" & F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numdoc", "alm_ingreso", "N", mConexion))
                    mCodigoItem = F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                    mItem = F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "descripcion", "alm_inventario", "N", mConexion))
                    Err.Raise &HFFFFFF01, , "Error al intentar grabar el kardex. " _
                                                + vbCr + "Codigo de Item: " & mCodigoItem _
                                                + vbCr + "Item: " & mItem _
                                                + vbCr + "Movimiento: " & mNumDoc _
                                                + vbCr + "Fecha: " & FechaMovimiento
                End If
            Else ' Si es un registro existente
                mKardxDetMov.TipoMovimiento = TipoMovimiento
                mKardxDetMov.FechaMovimiento = FechaMovimiento
                mKardxDetMov.Cantidad = Cantidad
                mKardxDetMov.CostoUnitario = CostoUnitario
                mKardxDetMov.CostoUnitarioPromedio = CostoUnitarioPromedio
                ' Se registra e movimiento de kardex
                If Not mConexion Is Nothing Then Set mKardxDetMov.Conexion = mConexion
                mKardxDetMov.UpdateParent = True
                If Not mKardxDetMov.Save(0, "") Then
                    mIdMovimiento = F.NuloString(F.BuscaCodigoTabla(IdMovimientoDet, "idmovdet", "id", "alm_ingresodet", "N", mConexion))
                    mNumDoc = F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numser", "alm_ingreso", "N", mConexion))
                    mNumDoc = mNumDoc & "-" & F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numdoc", "alm_ingreso", "N", mConexion))
                    mCodigoItem = F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "codpro", "alm_inventario", "N", mConexion))
                    mItem = F.NuloString(F.BuscaCodigoTabla(IdItem, "id", "descripcion", "alm_inventario", "N", mConexion))
                    Err.Raise &HFFFFFF01, , "Error al intentar grabar el kardex. " _
                                                + vbCr + "Codigo de Item: " & mCodigoItem _
                                                + vbCr + "Item: " & mItem _
                                                + vbCr + "Movimiento: " & mNumDoc _
                                                + vbCr + "Fecha: " & FechaMovimiento
                End If
            End If
        End If
    End If
End Function

Function mCalcularCostoSalidas(IdItem As Long, IdAlmacen As Long, _
                                    FechaInicio As Date, FechaFin As Date, _
                                    mConexion As ADODB.Connection) As Boolean
    Dim mCSQL As String
    Dim mCostoUnitario As Double
    Dim mCostoUnitarioPromedio As Double
    Dim mRecord As New ADODB.Recordset
    Dim mSaldoTotal As Double
    Dim mSaldoInicial As Double
    Dim mCostoUnitarioInicial As Double
    Dim mUltimoCostoUnitario As Double
    Dim mImporteTotal As Double
    Dim mSaldoSalidas As Double
    Dim mSaldoEntradas As Double
    Dim F As New SistemaLogica.Funciones
        
On Error GoTo ERROR_
    '--Generar la consulta SQL para obtener el detalle de movimientos del kardex
    mCSQL = F.KardexMovimientoSQL(IdItem, IdAlmacen, FechaInicio, FechaFin, mConexion)
    RST_Busq mRecord, mCSQL, mConexion
    mRecord.Sort = "fchdoc, tipmov, numdoc"
       
    mSaldoInicial = 0
    mCostoUnitarioInicial = 0
    
    mUltimoCostoUnitario = mCostoUnitarioInicial
    
    mSaldoTotal = mSaldoInicial
    mImporteTotal = mSaldoTotal * mCostoUnitarioInicial
    mSaldoEntradas = mSaldoEntradas + mSaldoInicial
        
    If mRecord.RecordCount = 0 Then
        mCalcularCostoSalidas = False
        Exit Function
    End If
    
    mRecord.MoveFirst
    While Not mRecord.EOF
        '***************
        ' INGRESOS
        '***************
        If F.NuloNumeric(mRecord("tipmov")) = -1 Then
        
            mSaldoTotal = mSaldoTotal + F.NuloNumeric(mRecord("canpro"))
            mSaldoEntradas = mSaldoEntradas + F.NuloNumeric(mRecord("canpro"))
            
            If F.NuloString(mRecord("descdoc")) = "DA" Then
                mCostoUnitario = mHallaCostoUnitarioSimple(F.NuloNumeric(mRecord("idmovdet")), _
                            F.NuloNumeric(mRecord("iditem")), "AJ", mConexion)
            Else
                mCostoUnitario = F.NuloNumeric(mRecord("costo")) / F.NuloNumeric(mRecord("canpro"))
            End If
            mImporteTotal = mImporteTotal + (F.NuloNumeric(mRecord("canpro")) * mCostoUnitario)
            If mSaldoTotal = 0 Then
                mCostoUnitarioPromedio = 0
            Else
                mCostoUnitarioPromedio = mImporteTotal / mSaldoTotal
            End If
            mUltimoCostoUnitario = mCostoUnitarioPromedio
            If F.NuloString(mRecord("descdoc")) = "DA" Then
                If Not mGrabaCostoMovTemp(IdAlmacen, F.NuloNumeric(NulosN(mRecord("idmovdet"))), _
                            F.NuloNumeric(mRecord("canpro")), mCostoUnitario, _
                            mCostoUnitarioPromedio, F.NuloNumeric(mRecord("canpro")) * mCostoUnitario, _
                            0, 0, mRecord("fchdoc"), "I", mConexion) Then
                        
                    Err.Raise &HFFFFFF01, , "Error interno al intentar grabar el costo del movimiento. " _
                                            + vbCr + "Numero de Produccion: " & F.NuloString(mRecord("numser")) & "-" & F.NuloString(mRecord("numdoc")) _
                                            + vbCr + "Item: " & F.NuloString(mRecord("codpro")) & " - " & F.NuloString(mRecord("descripcion")) _
                                            + vbCr + "Cantidad: " & F.NuloNumeric(mRecord("canpro")) _
                                            + vbCr + "Fecha de Movimiento: " & F.NuloString(mRecord("fchdoc"))
                End If
            End If
        
        '***************
        ' SALIDAS
        '***************
        Else
            If mSaldoTotal = 0 Then
                mCostoUnitarioPromedio = 0
            Else
                mCostoUnitarioPromedio = mImporteTotal / mSaldoTotal
            End If
            
            mSaldoTotal = mSaldoTotal - F.NuloNumeric(mRecord("canpro"))
            mSaldoSalidas = mSaldoSalidas + F.NuloNumeric(mRecord("canpro"))
            mImporteTotal = mImporteTotal - (F.NuloNumeric(mRecord("canpro")) * mCostoUnitarioPromedio)
            
            If Not mGrabaCostoMovTemp(IdAlmacen, F.NuloNumeric(NulosN(mRecord("idmovdet"))), _
                        F.NuloNumeric(mRecord("canpro")), mCostoUnitarioPromedio, _
                        mCostoUnitarioPromedio, F.NuloNumeric(mRecord("canpro")) * mCostoUnitarioPromedio, _
                        0, 0, mRecord("fchdoc"), "S", mConexion) Then
                    
                Err.Raise &HFFFFFF01, , "Error interno al intentar grabar el costo del movimiento. " _
                                        + vbCr + "Numero de Produccion: " & F.NuloString(mRecord("numser")) & "-" & F.NuloString(mRecord("numdoc")) _
                                        + vbCr + "Item: " & F.NuloString(mRecord("codpro")) & " - " & F.NuloString(mRecord("descripcion")) _
                                        + vbCr + "Cantidad: " & F.NuloNumeric(mRecord("canpro")) _
                                        + vbCr + "Fecha de Movimiento: " & F.NuloString(mRecord("fchdoc"))
            End If
        End If
        
        mRecord.MoveNext
    Wend
    
    mCalcularCostoSalidas = True
    Exit Function
    
ERROR_:
    mCalcularCostoSalidas = False
    F.MostrarMensajeError Err.Description, "[CalcularCostoSalidas]", Err.Source, Err.Number
End Function

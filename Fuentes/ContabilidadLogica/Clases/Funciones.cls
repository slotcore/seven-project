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

Function LimpiaTablaCostos(IdAlmacenProceso As Long, _
                                FechaInicio As Date, _
                                FechaFin As Date, _
                                mConexion As ADODB.Connection) As Boolean
    Dim mDataBase As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones

On Error GoTo BloqueError
    mConexion.BeginTrans
    Set mDataBase.Connection = mConexion
    ' Se actualiza cantidades de Kardex
    mDataBase.ClearParameter
    mDataBase.CommandText = "DELETE FROM con_librocostotemp " _
        + vbCr + "WHERE (((con_librocostotemp.idalmproc = " & IdAlmacenProceso & ")) " _
            & "AND ((con_librocostotemp.fecha)>=CDate('" & FechaInicio & "') AND (con_librocostotemp.fecha)<=CDate('" & FechaFin & "')))"
    mDataBase.Execute
    mConexion.CommitTrans
    LimpiaTablaCostos = True
    Exit Function
    
BloqueError:
    mConexion.RollbackTrans
    LimpiaTablaCostos = False
    F.MostrarMensajeError Err.Description, "LimpiaTablaCostos", Err.Source, Err.Number
End Function

Function EjecutaCosteoMateriales(ByRef LErrorCosto As ContabilidadEntidad.LEErrorCosto, _
                                SaltaError As Boolean, _
                                IdAlmacenProceso As Long, _
                                FechaInicio As Date, _
                                FechaFin As Date, _
                                mConexion As ADODB.Connection) As Boolean
    Dim mRecord As New ADODB.Recordset
    Dim F As New SistemaLogica.Funciones
    Dim FWin As New SistemaWindows.SistemaWindowsClass
    Dim mSQL As String
    Dim mContador As Long

On Error GoTo BloqueError
    
    ' Movimientos no Costeados
    mSQL = "SELECT alm_ingresodet.iditem, alm_ingreso.idalm, alm_inventario.codpro, alm_inventario.descripcion " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((alm_inventario.tippro) In (1,2,4)) AND ((alm_ingreso.fching)>=CDate('" & FechaInicio & "') And (alm_ingreso.fching)<=CDate('" & FechaFin & "')) AND (([con_librocostotemp].[costounitariopromedio]*[alm_ingresodet].[cantidad]) Is Null) AND ((alm_ingresodet.cantidad)>0)) " _
        + vbCr + "GROUP BY alm_ingresodet.iditem, alm_ingreso.idalm, alm_inventario.codpro, alm_inventario.descripcion"
    
    Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
    If mRecord.RecordCount > 0 Then
        mContador = 0
        FWin.ShowProgress "Procesando Insumos", 0, mRecord.RecordCount
        mRecord.MoveFirst
        While Not mRecord.EOF
            ' Se costea el item
            mContador = mContador + 1
            DoEvents
            FWin.SetProgress F.NuloString(mRecord("codpro")) & " - " & F.NuloString(mRecord("descripcion")), mContador
            If Not CosteaItem(LErrorCosto, IdAlmacenProceso, F.NuloNumeric(mRecord("iditem")), _
                                    F.NuloNumeric(mRecord("idalm")), FechaInicio, _
                                    FechaFin, mConexion) Then
                If Not SaltaError Then
                    Err.Raise &HFFFFFF01, , LErrorCosto(LErrorCosto.Count).DetalleError
                End If
            End If
            mRecord.MoveNext
        Wend
    End If

    FWin.HideProgress
    EjecutaCosteoMateriales = True
    Exit Function
    
BloqueError:
    FWin.HideProgress
    EjecutaCosteoMateriales = False
    F.MostrarMensajeError Err.Description, "LimpiaTablaCostos", Err.Source, Err.Number
End Function

Function EjecutaCosteoProductosIntermedios(ByRef LErrorCosto As ContabilidadEntidad.LEErrorCosto, _
                                        SaltaError As Boolean, _
                                        IdAlmacenProceso As Long, _
                                        FechaInicio As Date, _
                                        FechaFin As Date, _
                                        mConexion As ADODB.Connection) As Boolean
    Dim mRecord As New ADODB.Recordset
    Dim F As New SistemaLogica.Funciones
    Dim FWin As New SistemaWindows.SistemaWindowsClass
    Dim mSQL As String
    Dim mContador As Long

On Error GoTo BloqueError
    
    ' Movimientos no Costeados
    mSQL = "SELECT alm_ingresodet.iditem, alm_ingreso.idalm, alm_inventario.codpro, alm_inventario.descripcion " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((alm_inventario.tippro) In (8)) AND ((alm_ingreso.fching)>=CDate('" & FechaInicio & "') And (alm_ingreso.fching)<=CDate('" & FechaFin & "')) AND (([con_librocostotemp].[costounitariopromedio]*[alm_ingresodet].[cantidad]) Is Null) AND ((alm_ingresodet.cantidad)>0)) " _
        + vbCr + "GROUP BY alm_ingresodet.iditem, alm_ingreso.idalm, alm_inventario.codpro, alm_inventario.descripcion"
    
    Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
    If mRecord.RecordCount > 0 Then
        mContador = 0
        FWin.ShowProgress "Procesando Productos Intermedios", 0, mRecord.RecordCount
        mRecord.MoveFirst
        While Not mRecord.EOF
            ' Se costea el item
            mContador = mContador + 1
            DoEvents
            FWin.SetProgress F.NuloString(mRecord("codpro")) & " - " & F.NuloString(mRecord("descripcion")), mContador
            If Not CosteaItem(LErrorCosto, IdAlmacenProceso, F.NuloNumeric(mRecord("iditem")), F.NuloNumeric(mRecord("idalm")), FechaInicio, FechaFin, mConexion) Then
                If Not SaltaError Then
                    Err.Raise &HFFFFFF01, , LErrorCosto(LErrorCosto.Count).DetalleError
                End If
            End If
            mRecord.MoveNext
        Wend
    End If

    FWin.HideProgress
    EjecutaCosteoProductosIntermedios = True
    Exit Function
    
BloqueError:
    FWin.HideProgress
    EjecutaCosteoProductosIntermedios = False
    F.MostrarMensajeError Err.Description, "EjecutaCosteoProductosIntermedios", Err.Source, Err.Number
End Function

Function EjecutaCosteoProductosTerminados(ByRef LErrorCosto As ContabilidadEntidad.LEErrorCosto, _
                                        SaltaError As Boolean, _
                                        IdAlmacenProceso As Long, _
                                        FechaInicio As Date, _
                                        FechaFin As Date, _
                                        mConexion As ADODB.Connection) As Boolean
    Dim mRecord As New ADODB.Recordset
    Dim F As New SistemaLogica.Funciones
    Dim FWin As New SistemaWindows.SistemaWindowsClass
    Dim mSQL As String
    Dim mContador As Long

On Error GoTo BloqueError
    
    ' Movimientos no Costeados
    mSQL = "SELECT alm_ingresodet.iditem, alm_ingreso.idalm, alm_inventario.codpro, alm_inventario.descripcion " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((alm_inventario.tippro) In (3)) AND ((alm_ingreso.fching)>=CDate('" & FechaInicio & "') And (alm_ingreso.fching)<=CDate('" & FechaFin & "')) AND (([con_librocostotemp].[costounitariopromedio]*[alm_ingresodet].[cantidad]) Is Null) AND ((alm_ingresodet.cantidad)>0)) " _
        + vbCr + "GROUP BY alm_ingresodet.iditem, alm_ingreso.idalm, alm_inventario.codpro, alm_inventario.descripcion"
    
    Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
    If mRecord.RecordCount > 0 Then
        mContador = 0
        FWin.ShowProgress "Procesando Productos Terminados", 0, mRecord.RecordCount
        mRecord.MoveFirst
        While Not mRecord.EOF
            ' Se costea el item
            mContador = mContador + 1
            DoEvents
            FWin.SetProgress F.NuloString(mRecord("codpro")) & " - " & F.NuloString(mRecord("descripcion")), mContador
            If Not CosteaItem(LErrorCosto, IdAlmacenProceso, F.NuloNumeric(mRecord("iditem")), F.NuloNumeric(mRecord("idalm")), FechaInicio, FechaFin, mConexion) Then
                If Not SaltaError Then
                    Err.Raise &HFFFFFF01, , LErrorCosto(LErrorCosto.Count).DetalleError
                End If
            End If
            mRecord.MoveNext
        Wend
    End If

    FWin.HideProgress
    
    EjecutaCosteoProductosTerminados = True
    Exit Function
    
BloqueError:
    FWin.HideProgress
    EjecutaCosteoProductosTerminados = False
    F.MostrarMensajeError Err.Description, "EjecutaCosteoProductosTerminados", Err.Source, Err.Number
End Function



''' <summary>
''' Halla el importe unitario promedio de partes de produccion en un rango de fechas
''' </summary>
Public Function EjecutaReCosteoIndirectos(ByRef LErrorCosto As ContabilidadEntidad.LEErrorCosto, _
                                        SaltaError As Boolean, _
                                        IdAlmacenProceso As Long, _
                                        FechaInicio As Date, _
                                        FechaFin As Date, _
                                        Optional mConexion As ADODB.Connection) As Boolean

    Dim mRecord As New ADODB.Recordset
    Dim F As New SistemaLogica.Funciones
    Dim FWin As New SistemaWindows.SistemaWindowsClass
    Dim mSQL As String
    Dim mContador As Long

On Error GoTo BloqueError
    
    ' Movimientos no Costeados
    mSQL = "SELECT alm_ingresodet.iditem, alm_ingreso.idalm, alm_inventario.codpro, alm_inventario.descripcion " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((con_librocostotemp.idalmproc) = " & IdAlmacenProceso & ") AND ((alm_ingreso.fching)>=CDate('" & FechaInicio & "') And (alm_ingreso.fching)<=CDate('" & FechaFin & "')) AND ((con_librocostotemp.costomod)>0) AND ((con_librocostotemp.costocif)>0) AND ((alm_ingresodet.cantidad)>0)) " _
        + vbCr + "GROUP BY alm_ingresodet.iditem, alm_ingreso.idalm, alm_inventario.codpro, alm_inventario.descripcion"
        
    Set mRecord = F.GeneraRstSQL(mSQL, mConexion)
    If mRecord.RecordCount > 0 Then
        mContador = 0
        FWin.ShowProgress "Re-Procesando Productos con Gastos Indirectos", 0, mRecord.RecordCount
        mRecord.MoveFirst
        While Not mRecord.EOF
            ' Se costea el item
            mContador = mContador + 1
            DoEvents
            FWin.SetProgress F.NuloString(mRecord("codpro")) & " - " & F.NuloString(mRecord("descripcion")), mContador
            If Not CosteaItem(LErrorCosto, IdAlmacenProceso, F.NuloNumeric(mRecord("iditem")), F.NuloNumeric(mRecord("idalm")), FechaInicio, FechaFin, mConexion) Then
                If Not SaltaError Then
                    Err.Raise &HFFFFFF01, , LErrorCosto(LErrorCosto.Count).DetalleError
                End If
            End If
            mRecord.MoveNext
        Wend
    End If

    FWin.HideProgress
    EjecutaReCosteoIndirectos = True
    Exit Function
    
BloqueError:
    FWin.HideProgress
    EjecutaReCosteoIndirectos = False
    F.MostrarMensajeError Err.Description, "EjecutaReCosteoIndirectos", Err.Source, Err.Number
End Function

Public Function GrabaGastoIndirecto(IdAlmacenProceso As Long, _
                                    IdMovimientoDetalle As Long, _
                                    CostoMOD As Double, _
                                    CostoCIF As Double, _
                                    Optional mConexion As ADODB.Connection = Nothing) As Boolean
    
    Dim F As New SistemaLogica.Funciones
    Dim mMovimientoItem As New ContabilidadEntidad.EMovimientoItem
    
On Error GoTo BloqueError

    If Not mConexion Is Nothing Then Set mMovimientoItem.Conexion = mConexion
    mMovimientoItem.Fetch IdMovimientoDetalle
    ' Costeamos el movimiento
    mMovimientoItem.CostoMOD = CostoMOD
    mMovimientoItem.CostoCIF = CostoCIF
    mMovimientoItem.CostoUnitario = (mMovimientoItem.CostoPrimo + mMovimientoItem.CostoMOD + mMovimientoItem.CostoCIF) / mMovimientoItem.Cantidad
    
    GrabaCostoMovimiento IdAlmacenProceso, mMovimientoItem, mConexion
    GrabaGastoIndirecto = True
    Exit Function

BloqueError:
    GrabaGastoIndirecto = False
    Set mMovimientoItem = Nothing
    F.MostrarMensajeError Err.Description, "GrabaGastoIndirecto", Err.Source, Err.Number
End Function

Public Function MovimientosSinValorizar(FechaInicio As Date, FechaFin As Date, _
                                            Cnn As ADODB.Connection) As ADODB.Recordset
    Dim mSQL As String
    Dim F As New SistemaLogica.Funciones
    
    mSQL = "SELECT IIf([alm_ingreso].[tipmov]=-1,'I','S') AS [Tip Mov], alm_ingreso.fchdoc AS [Fecha Mov], [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc] AS [Numero Doc], alm_inventario.codpro AS [Codigo Item], alm_inventario.descripcion AS Item, alm_ingresodet.Cantidad " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((alm_ingreso.fchdoc)>=CDate('" & FechaInicio & "') And (alm_ingreso.fchdoc)<=CDate('" & FechaFin & "')) AND ((alm_ingresodet.iditem)<>2424) AND ((alm_ingresodet.Cantidad)>0) AND ((con_librocostotemp.costoprimo) Is Null) AND ((con_librocostotemp.costomod) Is Null) AND ((con_librocostotemp.costocif) Is Null)) " _
        + vbCr + "ORDER BY IIf([alm_ingreso].[tipmov]=-1,'I','S'), alm_ingreso.fchdoc, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc]"
    Set MovimientosSinValorizar = F.GeneraRstSQL(mSQL, Cnn)
End Function

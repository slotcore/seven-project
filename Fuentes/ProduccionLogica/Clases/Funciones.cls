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

Public Property Get GetError() As Boolean
    GetError = mError
End Property
Public Property Let GetError(ByVal NewValue As Boolean)
    mError = NewValue
End Property

''' <summary>
''' Halla el importe unitario promedio de partes de produccion en un rango de fechas
''' </summary>
Function CostoPrimo(IdAlmacen As Long, FechaInicio As Date, FechaFin As Date, FechaInicioMovimientos As Date, _
                                    Optional mConexion As ADODB.Connection) As Double
    CostoPrimo = mCostoPrimo(IdAlmacen, FechaInicio, FechaFin, FechaInicioMovimientos, mConexion)
End Function


''' <summary>
''' Halla el importe unitario promedio del parte especificado
''' </summary>
Function CalcularCostoParte(mIdAlmacenProceso As Long, mIdParteProduccion As Long, _
                                Optional mConexion As ADODB.Connection) As Double
    CalcularCostoParte = mCalcularCostoParte(mIdAlmacenProceso, mIdParteProduccion, mConexion)
End Function

''' <summary>
''' Halla el importe unitario promedio del movimiento especificado
''' </summary>
Function CalcularCostoMovimiento(IdAlmacenProceso As Long, IdMovimientoDetalle As Long, Optional mConexion As ADODB.Connection) As Double
    CalcularCostoMovimiento = mCalcularCostoMovimiento(IdAlmacenProceso, IdMovimientoDetalle, mConexion)
End Function

''' <summary>
''' Halla el importe unitario promedio del item especificado y el de sus movimientos
''' </summary>
Function CalcularCostoUnitarioItem(mIdAlmacenProceso As Long, mIdItemInsumo As Long, _
                            mIdAlmacen As Long, _
                            Optional FechaInicio As Date, _
                            Optional FechaFin As Date, _
                            Optional mConexion As ADODB.Connection) As Double
      CalcularCostoUnitarioItem = mCalcularCostoUnitarioItem(mIdAlmacenProceso, mIdItemInsumo, mIdAlmacen, FechaInicio, FechaFin, mConexion)
End Function

''' <summary>
''' Genera los movimientos automaticos para guias y facturas de venta en un rango de fechas
''' </summary>
Function GenerarMovimientosGuiasVentas(IdUsuario As Long, FechaInicio As Date, _
                                    FechaFin As Date, FechaInicioMovimientos As Date, _
                                    mConexion As ADODB.Connection) As Boolean
    GenerarMovimientosGuiasVentas = mGenerarMovimientosGuiasVentas(IdUsuario, FechaInicio, _
                                    FechaFin, FechaInicioMovimientos, _
                                    mConexion)
End Function

Function GrabaCostoMovTemp(IdAlmacenProceso As Long, IdMovimientoDet As Long, Cantidad As Double, _
                        CostoUnitario As Double, CostoUnitarioPromedio As Double, _
                        CostoPrimo As Double, CostoMOD As Double, _
                        CostoCIF As Double, FechaMovimiento As Date, _
                        TipoMovimiento As String, mConexion As ADODB.Connection) As Boolean
                        
    GrabaCostoMovTemp = mGrabaCostoMovTemp(IdAlmacenProceso, IdMovimientoDet, Cantidad, CostoUnitario, _
                            CostoUnitarioPromedio, CostoPrimo, CostoMOD, _
                            CostoCIF, FechaMovimiento, TipoMovimiento, mConexion)
End Function

Function HallaMovimientoDetalle(IdParteProduccionDet As Long, Optional mConexion As ADODB.Connection) As Long
    HallaMovimientoDetalle = mHallaMovimientoDetalle(IdParteProduccionDet, mConexion)
End Function

Function CalcularCostoSalidas(IdItem As Long, IdAlmacen As Long, _
                                    FechaInicio As Date, FechaFin As Date, _
                                    mConexion As ADODB.Connection) As Boolean
    CalcularCostoSalidas = mCalcularCostoSalidas(IdItem, IdAlmacen, _
                                    FechaInicio, FechaFin, mConexion)
End Function

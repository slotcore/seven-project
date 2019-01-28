VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub ConfiguraImpresion(Con As ADODB.Connection)
    Set xCon = Con
    FrmConfPlant.Show
End Sub

Sub InventarioCobranza(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmImpDeu.Show
End Sub

Sub InventarioPagos(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmImpDeuPag.Show
End Sub

Sub CargarVentas(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmImpDocVenta.Show
End Sub

Sub CargarVentasEstudio(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    frmImpDocVentaEst.Show
End Sub

Sub CargarCompras(Con As ADODB.Connection)
    'importacion de compras para savar
    Set xCon = Con
    CargaDatosEmpresa
    FrmImpDocCom2.Show
End Sub

Sub CargarCompras2(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmImpDocCom2.Show
End Sub

Sub CargarClientes(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmImpCliente.Show
End Sub

Sub CargarProveedores(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmImpProveedor.Show
End Sub

Sub ConfiguraCodigoUnico(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmSetupCodigo.Show vbModal
End Sub

Sub PersonalTesoreria(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmControlPersTes.Show
End Sub

Sub PersonalProduccion(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmManControlPers.Show
End Sub

Sub PersonalCompras(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmControlPersonalCompras.Show
End Sub

Sub CargarPlandeCuentas(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmManPlanCtasImportar.Show
End Sub

Sub TransferenciaOperaciones(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmTransOpe.Show
End Sub

Sub CorregirAsiento(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmAsientoCorregir.Show
End Sub


Sub ConsistenciaDatos(Con As ADODB.Connection)
    '--011208
    Set xCon = Con
    CargaDatosEmpresa
    FrmConsistenciaDatos.Show
End Sub

Sub BDEvaluar(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmBDEvaluar1.Show
End Sub

Private Sub Class_Initialize()
    NomSIS = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

Function IngRapidoItems(Con As ADODB.Connection) As Integer
    Set xCon = Con
    FrmIngRapItems.Show vbModal
    If FrmIngRapItems.xIdProducto = 0 Then
        IngRapidoItems = 0
    Else
        IngRapidoItems = FrmIngRapItems.xIdProducto
    End If
    Unload FrmIngRapItems
End Function



Sub AjustexDifCambio(Con As ADODB.Connection)
    '===================================================================================================
    'Creado : 11/02/10 Por: Johan Castro
    'Propósito: Identificar aquellos documentos provicionados que fueron cancelados luego ajustarlo
    '
    'Entradas:  xcon=Conexion activa a la base de datos
    '
    'Resultados: Generar el asiento de ajuste por dif de cambio
    
    '===================================================================================================
    Set xCon = Con
    CargaDatosEmpresa
    FrmAjusteDifCambio.Show

End Sub

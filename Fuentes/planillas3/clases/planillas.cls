VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "planillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL

Sub Comision(Con As ADODB.Connection)
'    Set xConPri = Con
    Set xCon = Con
    CargaDatos
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmComision.Show
End Sub


Sub ExportarSunat(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    FrmExportarSunat.Show
End Sub

Sub ManConcepto(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManConcepto.Show
End Sub

Sub AsignarSueldo(Con As ADODB.Connection, RutaEmpresa As String, mes As Integer)
    Set xCon = Con
    'Set xConPri = Con
    xMes = mes
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmAsignarSueldo.Show
End Sub

Sub ManRegimenPension(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManRegPensionario.Show
End Sub

Sub BoletaPago(Con As ADODB.Connection, RutaEmpresa As String, mes As Integer)
    Set xCon = Con
    'Set xConPri = Con
    xMes = mes
    xIdUsuario = IdUsuario
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    FrmBoleta1.Show
End Sub

Sub ConsPlanilla(Con As ADODB.Connection, RutaEmpresa As String, mes As Integer)
    Set xCon = Con
    'Set xConPri = Con
    xMes = mes
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    FrmConsPlanilla.Show
End Sub

Sub ImprimirBoletas(Con As ADODB.Connection, RutaEmpresa As String, mes As Integer)
    Set xCon = Con
    'Set xConPri = Con
    xMes = mes
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    FrmPrintBoletaxPeriodo.Show
End Sub


Sub ManConceptoOtros(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    FrmManConceptoVarios.Show
End Sub

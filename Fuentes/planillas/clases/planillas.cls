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


Sub ManNomina(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmNomina2.Show
End Sub

Sub ConsRotacion(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmRepRotacion.Show
End Sub

Sub ExportarNomina(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    FrmExportarSunat.Show
End Sub

Sub ConsAsistencia(Con As ADODB.Connection, RutaEmpresa As String, mes As Integer)
    Set xCon = Con
    xMes = mes
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    FrmRepAsistencia.Show
End Sub

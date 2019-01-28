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


Sub ManTipoPermiso(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManPermiso.Show
End Sub

Sub ManTipoLicencia(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManLicencia.Show
End Sub

Sub ManHorario(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManHorario2.Show
End Sub

Sub Permiso(Con As ADODB.Connection, RutaEmpresa As String, mes As Integer)
    Set xCon = Con
    'Set xConPri = Con
    xMes = mes
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmPermiso.Show
End Sub

Sub Licencia(Con As ADODB.Connection, RutaEmpresa As String, mes As Integer)
    Set xCon = Con
    'Set xConPri = Con
    xMes = mes
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmLicencia.Show
End Sub

Sub ManDiasFestivos(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManDiaFestivo.Show
End Sub

Sub Vacaciones(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmVacaciones.Show
End Sub

Sub Asistencia(Con As ADODB.Connection, RutaEmpresa As String, mes As Integer)
    Set xCon = Con
    'Set xConPri = Con
    xMes = mes
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmAsistencia.Show
End Sub

Sub ConsAsistencia(Con As ADODB.Connection, RutaEmpresa As String, mes As Integer)
    Set xCon = Con
    'Set xConPri = Con
    xMes = mes
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    FrmConsAsistencia.Show
End Sub

Sub ResumenHoras(Con As ADODB.Connection, RutaEmpresa As String)
    Set xCon = Con
    'Set xConPri = Con
    AP_RUTDATTRA = RutaEmpresa
    CargaDatos
    'Set xCon = AbrirConPlanilla(AP_RUTABD)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmResumenHoras.Show
End Sub

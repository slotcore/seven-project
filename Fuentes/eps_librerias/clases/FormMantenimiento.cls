VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Tabla As String
Public CampoOrdenado As String
Public CampoClave As String
Public TituloFormulario As String
Public DefinirCampoClave As Boolean
Public PermiteActualiza As Boolean
Public CadSQLVista As String

Sub MantTablas(Campos() As String, CamposBusqueda() As String, DatosVincula() As String, DatosVista() As String, Conex As ADODB.Connection)
    Set xConeccion = Conex
    xVincula = DatosVincula
    xCamposVista = DatosVista
    xCadSQL = CadSQLVista
    xTabla = Tabla
    xCampoOrdenado = CampoOrdenado
    xCampos = Campos
    xCamposBusqueda = CamposBusqueda
    xTituloForm = TituloFormulario
    xCampoClave = CampoClave
    xDefCampoClave = DefinirCampoClave
    xPermiteActualiza = PermiteActualiza
    FrmMantenimiento.Show vbModal
End Sub

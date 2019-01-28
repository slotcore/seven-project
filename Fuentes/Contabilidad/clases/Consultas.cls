VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub HojaTrabajo(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmHojaTrabajo1.Show
End Sub

Sub Mayor(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmConsultaMayor3.Show
End Sub

Sub VerDiario(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmConsultaDiario.Show
End Sub

Sub VerRegCompras(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmRegComVen1.Show
End Sub

Sub VerRegVentas(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmRegVentasOtros1.Show
End Sub

Sub VerRegRent4ta(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmRegHonorario1.Show
End Sub

Sub ReporteRetenciones(Con As ADODB.Connection)
    Set xCon = Con
    FrmRepRetenciones.Show
End Sub

Sub MostrarStockResumen(Con As ADODB.Connection, MostrarValor As Boolean)
    Set xCon = Con
    MostrarValorizado = MostrarValor
    CargaDatosEmpresa
    FrmResuMov.Show
End Sub

Sub MostrarKardexValorizado(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmResuMov2.Show
End Sub

Sub ConsultaCostoMovimiento(Con As ADODB.Connection)
    Set xCon = Con
    FrmConsultaCostoMov.Show
End Sub

Sub ConsultaCostoParte(Con As ADODB.Connection)
    Set xCon = Con
    FrmConsultaCostoParte.Show
End Sub

Sub AnalisisCostoProduccion(Con As ADODB.Connection)
    Set xCon = Con
    FrmRepCostoProduccion.Show
End Sub

Sub ConsultaKardexValResum(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmConKardexResuVal.Show
End Sub

Sub InformeKardexVal(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmInfoKardexValorizado.Show
End Sub

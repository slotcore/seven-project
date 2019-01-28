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


'Sub MostrarStockResumen(Con As ADODB.Connection, MostrarValor As Boolean)
'    Set xCon = Con
'    MostrarValorizado = MostrarValor
'    CargaDatosEmpresa
'    FrmResuMov.Show
'End Sub

'--------------------------------------
' MODIFICADO 02/07/2012 - JOSE CHACON
'--------------------------------------
Sub MostrarKardexValorizado(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmResuMov2.Show
End Sub


'********************

'********************
Sub MostrarCostosVenta(Con As ADODB.Connection)
Set xCon = Con
CargaDatosEmpresa
FrmVerCostoVenta.Show
End Sub


Sub MostrarCostosVentaRes(Con As ADODB.Connection)
Set xCon = Con
CargaDatosEmpresa
FrmCostoVentaRes.Show
End Sub


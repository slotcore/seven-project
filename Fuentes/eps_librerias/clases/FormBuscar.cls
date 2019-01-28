VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SQLCad As String                    'Sentencia SQL
Public Coneccion As New ADODB.Connection   'coneccion ADO
Public Titulo As String                    'titulo del formulario de busqueda
Public Criterio As String
Public Ordenado As String                  'Campo por el que se ordena la busqueda
Public CampoBusca As String                'especifica el campo en el que se va a efectuar la busqueda
Public RstConsulta As New ADODB.Recordset

Public Enum FormasBusqueda
    Principio = 1
    CualquierParte = 2
End Enum

Public FormaBusca As FormasBusqueda

Function BuscarReg(Campos() As String) As ADODB.Recordset
    Set xConeccion = Coneccion
    xTitulo = Titulo
    xCriterio = Criterio
    xCampos = Campos
    
    xFormaBusca = FormaBusca
    xOrdenado = Ordenado
    xCampoBusca = CampoBusca
    
    
    If F_NulosC(SQLCad) <> "" Then
        EjecutaSQL = True
        xSQLCad = SQLCad
    Else
        If RstConsulta.State <> 0 Then
            EjecutaSQL = False
            Set xRstConsulta = RstConsulta
        End If
    End If
        
    FrmBuscar.Show vbModal
        
    If Cancelado = True Then
        Set BuscarReg = Nothing
    Else
        Set BuscarReg = FrmBuscar.RstBusca
        Set FrmBuscar.RstBusca = Nothing
        Unload FrmBuscar
    End If
End Function


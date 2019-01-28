VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SQLCad As String                         'Sentencia SQL
Public Coneccion As New ADODB.Connection        'coneccion ADO
Public Titulo As String                         'titulo del formulario de busqueda

'Function Seleccionar(Campos() As String) As ADODB.Recordset
'    Set xConeccion = Coneccion
'    xTitulo = Titulo
'    xSQLCad = SQLCad
'    xCampos = Campos
'
'    FrmSelecciona.Show vbModal
'
'    If FrmSelecciona.Acepto = True Then
'        Set Seleccionar = FrmSelecciona.Rst
'    Else
'        Set FrmSelecciona.Rst = Nothing
'        Set Seleccionar = FrmSelecciona.Rst
'    End If
'    Unload FrmSelecciona
'End Function

Function Seleccionar(Campos() As String) As ADODB.Recordset

    Set xConeccion = Coneccion
    xTitulo = Titulo
    xSQLCad = SQLCad
    xCampos = Campos
    
    FrmSelecciona1.Show vbModal
    
    If FrmSelecciona1.Acepto = True Then
        Set Seleccionar = FrmSelecciona1.Rst
    Else
        Set FrmSelecciona1.Rst = Nothing
        Set Seleccionar = FrmSelecciona1.Rst
    End If
    Unload FrmSelecciona1
End Function

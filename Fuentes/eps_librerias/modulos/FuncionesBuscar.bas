Attribute VB_Name = "declaraciones"
Option Explicit

Public xConeccion As New ADODB.Connection
Public xTitulo As String
Public Cancelado  As Boolean
Public xCriterio As String
Public xCampos() As String
Public xSQLCad As String
Public xOrdenado As String
Public xCampoBusca As String
Public xFormaBusca As Integer
Public xRstConsulta As New ADODB.Recordset
Public EjecutaSQL As Boolean 'especifica si se ejecutara un consulta sql o un recordset

'variables para la impresion
Public ArrayPrin() As String
Public Prin_Cabecera1 As String
Public Prin_Cabecera2 As String
Public Prin_Fecha As String
Public Prin_Titulo1 As String
Public Prin_Titulo2 As String
Public xTamañoFuente  As Integer
Public RstPrin As New ADODB.Recordset
Public Prin_TamañoCabecera As Integer
Public Prin_FuenteCabecera As String
Public Prin_TamañoHoja As Integer
Public Prin_OrientacionHoja As Integer
Public Prin_TextoConsiderar As String
Public Prin_TextoConsiderarAncho As Integer
Public xFg As Object

Public xOpciones() As String

Public xVS1 As VSPrinter7LibCtl.VSPrinter

Function BuscaCampoLista(Texto As String, IndiceBusca As String, IndiceDevuelve As String, Campos() As String) As String
    Dim A As Integer
    
    For A = LBound(Campos) To UBound(Campos)
        If xCampos(A, IndiceBusca) = Texto Then
            BuscaCampoLista = Campos(A, IndiceDevuelve)
            Exit For
        End If
        If A = UBound(Campos) - 1 Then
            Exit For
        End If
    Next A
End Function

Public Function AddControl(Controls As CommandBarControls, ControlType As XTPControlType, _
                            Id As Long, Caption As String, Optional BeginGroup As Boolean = False, _
                            Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, _
                            Optional Category As String = "Controls") As CommandBarControl
    
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, Id, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category
    
    Set AddControl = Control
    
End Function


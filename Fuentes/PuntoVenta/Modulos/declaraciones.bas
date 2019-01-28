Attribute VB_Name = "declaraciones"
Option Explicit

Public xCon As ADODB.Connection
Public xTitulo As String
Public AnoTra As String
Public mIdEmpleado As Integer    '--este codigo puede ser del vendedor, cajero, supervisor (u otros cargos que pueda tener el personal)
Public mIdAlmacen As Integer     '--indica el almacen que inicia la sesion(tadas los operaciones se haran en funcion del almacen que selecciona)
Public ArrDocumento(2, 3) As String '--almacenara los valores del documento, se cargara al momenro de loguearse el usuario
Public Contabilizar As Boolean
    '(0,?) ticket               ||  (?,0) Tipo Doc.,     (?,1) Nº Serie
    '(1,?) factura              ||  (?,2) ID Plantilla   (?,3) Nombre Plantilla
    '(2,?) boleta de venta      ||
    '? indica fila o columna


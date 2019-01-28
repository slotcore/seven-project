Attribute VB_Name = "Declaraciones"
Option Explicit

Public xCadSQL As String
Public xTabla As String
Public xCampoOrdenado As String
Public xCampos() As String
Public xVincula() As String
Public xTituloForm As String
Public xConeccion As New ADODB.Connection
Public xCamposBusqueda() As String
Public xCampoClave As String
Public xCamposVista() As String
Public xDefCampoClave As Boolean      'VARIABLE QUE PERMITE SABER SI EL USUARIO DEFINE EL CAMPO CLAVE A LA
                                      'HORA DE AGREGAR UN NUEVO REGISTRO

Public xPermiteActualiza As Boolean   'ESPECIFICA SI SE PERMITE LA MODIFICACION DE DATOS
Public xRstOrigen As New ADODB.Recordset
Public xConTMP As New ADODB.Connection
Public xCon2 As New ADODB.Connection
Public xRutaData As String
Public xRutaFileTrabajo As String

Public xCon As New ADODB.Connection

Public xIdUsuario As Integer            'ALMACENA EL CODIGO DEL USUARIO
Public xIdMenu As Integer               'ALMACENA EL CODIGO DEL MENU


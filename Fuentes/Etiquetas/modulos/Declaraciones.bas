Attribute VB_Name = "Declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES.BAS
'* Tipo             : MODULO
'* Descripcion      : MODULO DONDE SE DECLARAN LAS VARIABLES PUBLICAS QUE SE UTILIZARAN EN LA CLASE
'*                    ASI COMO LA DEFINICION DE ALGUNAS FUNCIONES PROPIAS DE LA CLASE
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 28/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public xCon As New ADODB.Connection
Public xMes As Integer
Public xTitulo As String

Public NomEmp As String
Public NumRUC As String
Public AnoTra As String
Public Nomsis As String
Public xIdUsuario As Integer

Public CONTABILIZAR As Boolean

Global AP_RUTASY As String
Global AP_RUTABD As String
Global AP_RUTABM As String

Public xIdMenu As Integer
Public xConEtiqueta As New ADODB.Connection

Public Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    With Rst
        If .State <> 0 Then
            If TODO Then .Filter = adFilterNone
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
        End If
    End With
End Sub

Public Function GENERAR_SQL_ID_RST(Rst As ADODB.Recordset, nDesc As String, _
                            nCampo As String, Optional nTipoIn As String = "IN", _
                            Optional fEsNumero As Boolean = True) As String
    Dim nSQL As String
    Dim k&
    nSQL = ""
    
    If Rst.RecordCount = 0 Then Exit Function Else Rst.MoveFirst
    While Not Rst.EOF
        If Trim(CStr(Rst("" & nDesc & ""))) <> "" Then
            If fEsNumero = True Then
                nSQL = nSQL & NulosN(Rst("" & nDesc & "")) & ","
            Else
                nSQL = nSQL & "'" & NulosC(Rst("" & nDesc & "")) & "',"
            End If
        End If
        Rst.MoveNext
    Wend
    
    If nSQL <> "" Then nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") "
        
    GENERAR_SQL_ID_RST = nSQL
End Function

'*****************************************************************************************************
'* Nombre         : CargaDatosEmpresa()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : CARGA LOS DATOS DE LA EMPRESA ACTUAL Y LOS ALAMCENA EN LAS  VARIABLES PUBLICAS  YA
'*                  DEFINIDAS
'* Paranetros     :
'* Retorna        :
'*****************************************************************************************************
Sub CargaDatosEmpresa(xCon As ADODB.Connection)
    On Error Resume Next
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    
    Set Rst = Nothing
    Err.Clear
End Sub

Sub crearConexionEtiqueta()
    Dim xFun As New eps_librerias.FuncionesData
    
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")          ' LEEMOS LA RUTA DEL SISTEA
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")          ' LEEMOS LA RUTA DE LA BASE DE DATOS
    
    xTitulo = "Sistema Gestion Informatica"                                             ' CARGAMOS LOS TITULOS PARA LOS MSGBOX Y OTROS MENSAJES QUE EMITIERA EL SISTEMA
    
    xFun.F_BASEDATOS = Trim(AP_RUTABD) & "etiquetas.mdb"                                ' PASAMOS LA RUTA DE LA BASE DE DATOS PARA ABRIR LA CONECCION
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"                                       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABJO DE LA BASE DE DATOS
    xFun.F_PASSWORD = Eps_Pass                                                          ' PASAMOS EL PASWORD DE LA BASE DE DATOS
    xFun.F_USUARIO = Eps_User                                                           ' PASAMOS EL USUARIO DE LA BASE DE DATOS
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"                                        ' PASAMOS EL NOMBRE DEL PROVEEDORE DE DATOS PARA ADO 2.5
    
    Set xConEtiqueta = xFun.AbrirConeccion                                                      ' ABRIMOS LA CONECCION DE DATOS
    Set xFun = Nothing
End Sub


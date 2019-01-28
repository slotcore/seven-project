Attribute VB_Name = "Declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES.BAS
'* Tipo             : MODULO
'* Descripcion      : MODULO EN EL QUE SE DEFINEN LA PRINCIPALES VARIABLES A UTILIZARCE EN LA CLASE,
'*                    ADEMAS AQUI SE DEFINEN FUNCIONES QUE SERAN USADAS UNICAMENTE EN LA CLASE ACTUAL
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public xCon As New ADODB.Connection   ' CONECCION A LA BASE DE DATOS
Public xTitulo As String              ' TITULO PARA LA CLASE CUANDO SE MUESTRE UN MENSAJE
Public NomSIS As String               ' NOMBRE DEL SISTEMA
Public AnoTra As String               ' AÑO DE TRABAJO ACTUAL
Public CONTABILIZAR As Boolean        ' VARIABLE QUE INDICA SI SE HARAN PROCESOS CONTABLES  TRUE = SE HARA PROCESO CONTABLE ; FALSE = NO SE HARA PROCES CONTABLE
Public xMes As Integer                ' INDICA EL MES DE TRABAJO ACTUAL
Public xOrigen As Integer             ' ESPECIFICA DE DONDE ES INVOCADO EL FORMAULARIO ; 0 = MENU PRINCIPAL; 1 = ALGUNA LIBRERIA O FUNCION
Public xIdUsuario As Integer          ' ALAMACENA EL ID DEL USUARIO

Public xNumRUC As String              ' NUMERO DE RUC DE LA EMPRESA
Public xNomEmp As String              ' NOMBRE DE LA EMPRESA
Public xDirEmp As String              ' DIRECCION DE LA EMPRESA
Public xDisEmp As String              ' DISTRITO DE LA EMPRESA
Public xPagEmp As String              ' PAGINA WEB DE LA EMPRESA

Public AP_RUTASY As String            ' RUTA DEL SISTEMA
Public AP_RUTABD As String            ' RUTA DE LA BASE DE DATOS
Public AP_RUTABM As String            ' RUTA DE LOS ARCHIVOS GRAFICOS DEL SISTEMA
Public AP_AÑODAT As String            ' AÑO DE TRABAJO
Public AP_MESTRA As Integer           ' MES DE TRABAJO

Public xDeDonde As Integer            ' ESPECIFICA SI SE SINCRONIZA EL INGRESO DE DATOS CON LAS DEMAS BASES DE DATOS
Public IdCompraReg As Integer         ' alamcenara el id de la compra registrada esta funcion es para cuando se halla llamado el formulario desde otro formulario

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)


'*****************************************************************************************************
'* Nombre           : CargaDatosEmpresa
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA EMPRESA: Nombre de la empresa, Numero de Ruc, Año de trabajo
'*                    TAMBIEN CARGA LOS SIGUIENTES DATOS DEL SISTEMA: Nombre del sistema, Ruta de la
'*                    base de datos, Ruta del sistema, Ruta de los archivo de grafico, ADEMAS ESPECIFICA
'*                    SI EL SISTEMA HARA PROCESOS CONTABLES
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT mae_empresa.*, mae_distrito.descripcion AS descdis FROM mae_empresa LEFT JOIN mae_distrito ON mae_empresa.iddis = mae_distrito.id", xCon

    xNomEmp = Rst("nomemp")
    xNumRUC = Rst("numruc")
    xDirEmp = Rst("diremp")
    xDisEmp = Rst("descdis")
    'xPagEmp = Rst("pagweb")
    
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    Set Rst = Nothing
    NomSIS = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

'Function LLenarItemsRapido() As Integer
'    FrmIngRapItems.Show vbModal
'    LLenarItemsRapido = FrmIngRapItems.xIdNewItem
'    Set FrmIngRapItems = Nothing
'End Function

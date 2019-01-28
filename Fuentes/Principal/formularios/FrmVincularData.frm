VERSION 5.00
Begin VB.Form FrmVincularData 
   Caption         =   "Utilitarios - Actualizacion de Tablas Vinculadas"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   330
      TabIndex        =   0
      Top             =   180
      Width           =   3900
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Command1"
         Height          =   720
         Left            =   1965
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   300
         Width           =   810
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Command1"
         Height          =   720
         Left            =   1050
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ejecutar"
         Top             =   300
         Width           =   810
      End
   End
End
Attribute VB_Name = "FrmVincularData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : FRMVINCULARDATA
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO UTILIZADO PARA LA ACTUALIZACION DE LA RUTA DE LAS TABLAS VINCULADAS
'*                     DE LA EMPRESA ACTUAL
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 01/09/09
'* VERSION           : 1.0
'*****************************************************************************************************

Option Explicit

'*****************************************************************************************************
'* Nombre Modulo  : Vincular()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : VINCULA LAS TABLAS DE LAS BASE DE DATOS DE LA EMPRESA ACTUAL
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Vincular()
    'Modificado 24/01/11 Por Johan Castro
    '           Agregar linea de codigo para consultar lista de tablas vinculadas de tabla MSysObjects
    '           Modificar el modo de vincular en data.mdb
    '           Modificar el modo de vincular en planillas.mdb
    
    
    ' ASIGNAMOS EL ARCHIVO DE GRUPO DE TRABAJO, SI SE HACE DESPUES EL DAO NO DEJA ASIGNARLO POR ESO SE
    ' AL PRINCIPIO
    DBEngine.SystemDB = AP_RUTASY + "seven.mdw"
    
    Dim wrkODBC As DAO.Workspace                ' VARIABLE DAO PARA EL ESPACIO DE TRABAJO
    Dim BaseD As DAO.Database                   ' VARIABLE PARA ALMACENAR LA CONECCCION A LA BASE DE DATOS
    Dim Rt As DAO.Recordset
    Dim Td As DAO.TableDef
    Dim Tds As DAO.TableDefs                    ' VARAIBLE PARA ALMACENAR LA TABLA DE TRABAJO ACTUAL
    Dim xCla As String, xPas As String          ' VARIABLE PARA ALMACENAR LA CLAVE Y EL PASSOWR DE LA BASE DE DATOS
    
    Dim xFun As New eps_librerias.FuncionesData
    Dim xRutaData As String                        ' ALMACENA LA RUTA DE LA BASE DE DATOS
    Dim NuevaConeccion As String                ' ALMACENA LA RUTA DE LA BASE DE DATOS PARA REALIZAR LA VINCULACION
        
        
    Dim xConPlanilla As New ADODB.Connection
    Dim xRst As New ADODB.Recordset
    
    
    
    xRutaData = AP_RUTABD & AP_RUTDATTRA
    xCla = Eps_User
    xPas = Eps_Pass
    
    xFun.F_BASEDATOS = Mid(xRutaData, 1, Len(Trim(xRutaData)) - 8) & "planillas.mdb"                               ' PASAMOS LA RUTA DE LA BASE DE DATOS PARA ABRIR LA CONECCION
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"                                       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABJO DE LA BASE DE DATOS
    xFun.F_PASSWORD = xPas                                                              ' PASAMOS EL PASWORD DE LA BASE DE DATOS
    xFun.F_USUARIO = xCla                                                               ' PASAMOS EL USUARIO DE LA BASE DE DATOS
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xConPlanilla = xFun.AbrirConeccion                                                      ' ABRIMOS LA CONECCION DE DATOS
    Set xFun = Nothing

    Dim nSQLTablaVinculadas As String

    'Sentencia SQL para obtener listado de tablas que estan vinculadas, tabla MSysObjects esta oculta.
    nSQLTablaVinculadas = "SELECT MSysObjects.id , MSysObjects.Name AS tabla, MSysObjects.Database AS ruta FROM MSysObjects WHERE (((MSysObjects.Type)=6)); "

    RST_Busq xRst, nSQLTablaVinculadas, xCon
  
    'ACTUALIZAMOS data.mdb
    'ACTUALIZANDO LA RUTA DE LAS TABLAS VINCULADAS EN data1.mdb
    'ABRIMOS LA BD PARA PODER CAMBIAR LA RUTA DE LAS TABLAS VINCULADAS
    Set wrkODBC = CreateWorkspace("ODBCWorkspace", xCla, xPas, dbUseJet)
    Workspaces.Append wrkODBC
    Set BaseD = wrkODBC.OpenDatabase(NulosC(xRutaData), False, False)
    Set Tds = BaseD.TableDefs
   
    'ACTUALIZANDO LA RUTA CON LA BASE DE DATOS data.mdb(Normas Legales)
    NuevaConeccion = Mid(xRutaData, 1, Len(Trim(xRutaData)) - 8) & "planillas.mdb"
    
    If xRst.RecordCount <> 0 Then
        Do While Not xRst.EOF
            Tds(xRst("tabla")).Connect = ";DATABASE=" & NuevaConeccion
            Tds(xRst("tabla")).RefreshLink
            xRst.MoveNext
        Loop
    End If
    Set wrkODBC = Nothing
    Set xRst = Nothing
    
    'ACTUALIZAMOS planillas.mdb
    'ACTUALIZANDO LA RUTA DE LAS TABLAS VINCULADAS EN planillas.mdb
    'ABRIMOS LA BD PARA PODER CAMBIAR LA RUTA DE LAS TABLAS VINCULADAS
    Set wrkODBC = CreateWorkspace("ODBCWorkspace", xCla, xPas, dbUseJet)
    Workspaces.Append wrkODBC
    Set BaseD = wrkODBC.OpenDatabase(NulosC(NuevaConeccion), False, False)
    Set Tds = BaseD.TableDefs
   
    RST_Busq xRst, nSQLTablaVinculadas, xConPlanilla

    If xRst.RecordCount <> 0 Then
        Do While Not xRst.EOF
            Tds(xRst("tabla")).Connect = ";DATABASE=" & xRutaData
            Tds(xRst("tabla")).RefreshLink
            xRst.MoveNext
        Loop
    End If
   
   Set xRst = Nothing
   
'    Tds("con_centrocosto").Connect = ";DATABASE=" & xRutaData
'    Tds("con_centrocosto").RefreshLink
'
'    Tds("con_diario").Connect = ";DATABASE=" & xRutaData
'    Tds("con_diario").RefreshLink
'
'    Tds("con_meses").Connect = ";DATABASE=" & xRutaData
'    Tds("con_meses").RefreshLink
'
'    Tds("con_planctas").Connect = ";DATABASE=" & xRutaData
'    Tds("con_planctas").RefreshLink
'
'    Tds("con_tc").Connect = ";DATABASE=" & xRutaData
'    Tds("con_tc").RefreshLink
'
'    Tds("mae_area").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_area").RefreshLink
'
'    Tds("mae_cargo").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_cargo").RefreshLink
'
'    Tds("mae_departamento").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_departamento").RefreshLink
'
'    Tds("mae_distrito").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_distrito").RefreshLink
'
'    Tds("mae_dociden").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_dociden").RefreshLink
'
'    Tds("mae_documento").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_documento").RefreshLink
'
'    Tds("mae_documentocta").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_documentocta").RefreshLink
'
'    Tds("mae_empresa").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_empresa").RefreshLink
'
'    Tds("mae_libros").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_libros").RefreshLink
'
'    Tds("mae_moneda").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_moneda").RefreshLink
'
'    Tds("mae_provincia").Connect = ";DATABASE=" & xRutaData
'    Tds("mae_provincia").RefreshLink
'
'    Tds("var_cierre").Connect = ";DATABASE=" & xRutaData
'    Tds("var_cierre").RefreshLink
'
'    Tds("var_edicion").Connect = ";DATABASE=" & xRutaData
'    Tds("var_edicion").RefreshLink
    
    Set wrkODBC = Nothing
    xConPlanilla.Close
    MsgBox "El proceso se completo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    CmdSalir_Click
End Sub

Private Sub CmdOk_Click()
    Vincular
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO QUE SE EJECUTA AL CARGAR EL FORMULARIO
    Dim Ruta As String
    
    Ruta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
    
    Me.ScaleMode = 3
    CmdOk.Caption = ""
    CmdSalir.Caption = ""
    On Error Resume Next
    CmdOk.Picture = LeerIcono(Ruta + "toolbar\18.ico", T32x32, Me, Me.BackColor)
    Err.Clear
    CmdSalir.Picture = LeerIcono(Ruta + "toolbar\16.ico", T32x32, Me, Me.BackColor)
End Sub





Sub Vincular_xxx()
'Dado de baja el 01/24/11

    ' ASIGNAMOS EL ARCHIVO DE GRUPO DE TRABAJO, SI SE HACE DESPUES EL DAO NO DEJA ASIGNARLO POR ESO SE
    ' AL PRINCIPIO
    DBEngine.SystemDB = Trim(App.Path) + "\seven.mdw"

    Dim wrkODBC As DAO.Workspace                ' VARIABLE DAO PARA EL ESPACIO DE TRABAJO
    Dim BaseD As DAO.Database                   ' VARIABLE PARA ALMACENAR LA CONECCCION A LA BASE DE DATOS
    Dim Rt As DAO.Recordset
    Dim Td As DAO.TableDef
    Dim Tds As DAO.TableDefs                    ' VARAIBLE PARA ALMACENAR LA TABLA DE TRABAJO ACTUAL
    Dim xCla As String, xPas As String          ' VARIABLE PARA ALMACENAR LA CLAVE Y EL PASSOWR DE LA BASE DE DATOS
    Dim xFun As New eps_librerias.FuncionesData
    Dim xxRuta As String                        ' ALMACENA LA RUTA DE LA BASE DE DATOS
    Dim NuevaConeccion As String                ' ALMACENA LA RUTA DE LA BASE DE DATOS PARA REALIZAR LA VINCULACION
    
    Dim nSQLTablaVinculadas As String
    
    
    nSQLTablaVinculadas = "SELECT MSysObjects.Name AS tabla, MSysObjects.Database AS ruta FROM MSysObjects WHERE (((MSysObjects.Type)=6)); "

    xxRuta = AP_RUTABD & AP_RUTDATTRA
    
    xCla = Eps_User
    xPas = Eps_Pass
    
    'ACTUALIZAMOS data.mdb
    'ACTUALIZANDO LA RUTA DE LAS TABLAS VINCULADAS EN data1.mdb
    'ABRIMOS LA BD PARA PODER CAMBIAR LA RUTA DE LAS TABLAS VINCULADAS
    Set wrkODBC = CreateWorkspace("ODBCWorkspace", xCla, xPas, dbUseJet)
    Workspaces.Append wrkODBC
    Set BaseD = wrkODBC.OpenDatabase(NulosC(xxRuta), False, False)
    Set Tds = BaseD.TableDefs
   
    'ACTUALIZANDO LA RUTA CON LA BASE DE DATOS data.mdb(Normas Legales)
    NuevaConeccion = Mid(xxRuta, 1, Len(Trim(xxRuta)) - 8) & "planillas.mdb"
    
    Tds("pla_boleta").Connect = ";DATABASE=" & NuevaConeccion
    Tds("pla_boleta").RefreshLink
    
    Tds("pla_empleados").Connect = ";DATABASE=" & NuevaConeccion
    Tds("pla_empleados").RefreshLink
    Set wrkODBC = Nothing
    
    
    'ACTUALIZAMOS planillas.mdb
    'ACTUALIZANDO LA RUTA DE LAS TABLAS VINCULADAS EN planillas.mdb
    'ABRIMOS LA BD PARA PODER CAMBIAR LA RUTA DE LAS TABLAS VINCULADAS
    Set wrkODBC = CreateWorkspace("ODBCWorkspace", xCla, xPas, dbUseJet)
    Workspaces.Append wrkODBC
    Set BaseD = wrkODBC.OpenDatabase(NulosC(NuevaConeccion), False, False)
    Set Tds = BaseD.TableDefs
   
    Tds("con_centrocosto").Connect = ";DATABASE=" & xxRuta
    Tds("con_centrocosto").RefreshLink
    
    Tds("con_diario").Connect = ";DATABASE=" & xxRuta
    Tds("con_diario").RefreshLink
    
    Tds("con_meses").Connect = ";DATABASE=" & xxRuta
    Tds("con_meses").RefreshLink
    
    Tds("con_planctas").Connect = ";DATABASE=" & xxRuta
    Tds("con_planctas").RefreshLink
    
    Tds("con_tc").Connect = ";DATABASE=" & xxRuta
    Tds("con_tc").RefreshLink
    
    Tds("mae_area").Connect = ";DATABASE=" & xxRuta
    Tds("mae_area").RefreshLink
    
    Tds("mae_cargo").Connect = ";DATABASE=" & xxRuta
    Tds("mae_cargo").RefreshLink
    
    Tds("mae_departamento").Connect = ";DATABASE=" & xxRuta
    Tds("mae_departamento").RefreshLink
    
    Tds("mae_distrito").Connect = ";DATABASE=" & xxRuta
    Tds("mae_distrito").RefreshLink
    
    Tds("mae_dociden").Connect = ";DATABASE=" & xxRuta
    Tds("mae_dociden").RefreshLink
    
    Tds("mae_documento").Connect = ";DATABASE=" & xxRuta
    Tds("mae_documento").RefreshLink
    
    Tds("mae_documentocta").Connect = ";DATABASE=" & xxRuta
    Tds("mae_documentocta").RefreshLink
    
    Tds("mae_empresa").Connect = ";DATABASE=" & xxRuta
    Tds("mae_empresa").RefreshLink
    
    Tds("mae_libros").Connect = ";DATABASE=" & xxRuta
    Tds("mae_libros").RefreshLink
    
    Tds("mae_moneda").Connect = ";DATABASE=" & xxRuta
    Tds("mae_moneda").RefreshLink
    
    Tds("mae_provincia").Connect = ";DATABASE=" & xxRuta
    Tds("mae_provincia").RefreshLink
    
    Tds("var_cierre").Connect = ";DATABASE=" & xxRuta
    Tds("var_cierre").RefreshLink
    
    Tds("var_edicion").Connect = ";DATABASE=" & xxRuta
    Tds("var_edicion").RefreshLink
    
    Set wrkODBC = Nothing
    
    MsgBox "El proceso se completo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    CmdSalir_Click
End Sub


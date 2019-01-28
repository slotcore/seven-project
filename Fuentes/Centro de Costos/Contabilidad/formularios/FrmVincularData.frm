VERSION 5.00
Begin VB.Form FrmVincularData 
   Caption         =   "Utilitarios - Actualizacion de Tablas Vinculadas"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4590
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
Option Explicit

Sub Vincular()
    On Error GoTo LaCague
    Err.Clear
    
    DBEngine.SystemDB = Trim(App.Path) + "\seven.mdw"

    Dim wrkODBC As DAO.Workspace
    Dim BaseD As DAO.Database
    Dim Rt As DAO.Recordset
    Dim Td As DAO.TableDef
    Dim Tds As DAO.TableDefs
    Dim xCla As String, xPas As String
    
    
    Dim xFun As New eps_librerias.FuncionesData
    'Dim xRutaSIALP As String
    Dim xxRuta As String
    Dim NuevaConeccion As String
    
    xxRuta = AP_RUTABD & AP_RUTDATTRA
    
    xCla = Eps_User
    xPas = Eps_Pass
    
    'ACTUALIZAMOS data.mdb
    '**************************************************************
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
    '**************************************************************
    'ACTUALIZANDO LA RUTA DE LAS TABLAS VINCULADAS EN planillas.mdb
    'ABRIMOS LA BD PARA PODER CAMBIAR LA RUTA DE LAS TABLAS VINCULADAS
    Set wrkODBC = CreateWorkspace("ODBCWorkspace", xCla, xPas, dbUseJet)
    Workspaces.Append wrkODBC
    Set BaseD = wrkODBC.OpenDatabase(NulosC(NuevaConeccion), False, False)
    Set Tds = BaseD.TableDefs
   
    'ACTUALIZANDO LA RUTA CON LA BASE DE DATOS data.mdb(Normas Legales)
  
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
    
    '---ultimos
    Tds("var_cierre").Connect = ";DATABASE=" & xxRuta
    Tds("var_cierre").RefreshLink
    
    Tds("var_edicion").Connect = ";DATABASE=" & xxRuta
    Tds("var_edicion").RefreshLink
    
    
    
    
    Set wrkODBC = Nothing
    
    MsgBox "El proceso se completo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    CmdSalir_Click
    
    Exit Sub
LaCague:
    Set wrkODBC = Nothing
    MsgBox "Error." & Err.Description, vbCritical, xTitulo
    Err.Clear
End Sub

Private Sub CmdOk_Click()
    Vincular
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Ruta As String
    
    Ruta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
    
    Me.ScaleMode = 3
    CmdOk.Caption = ""
    CmdSalir.Caption = ""
    On Error Resume Next
    CmdOk.Picture = LeerIcono(Ruta + "toolbar\18.ico", T32x32, Me, Me.BackColor)
    CmdSalir.Picture = LeerIcono(Ruta + "toolbar\16.ico", T32x32, Me, Me.BackColor)
    Err.Clear

End Sub

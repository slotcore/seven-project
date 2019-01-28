VERSION 5.00
Begin VB.Form FrmMenuRapido 
   Caption         =   "Form4"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form4"
   ScaleHeight     =   4740
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command24 
      Caption         =   "Reporte de Rotacion"
      Height          =   405
      Left            =   6990
      TabIndex        =   24
      Top             =   3660
      Width           =   2295
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Orden de Produccion"
      Height          =   390
      Left            =   45
      TabIndex        =   23
      Top             =   2445
      Width           =   1650
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Analisis x Unidad de Negocio"
      Height          =   405
      Left            =   3765
      TabIndex        =   22
      Top             =   3855
      Width           =   2295
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Analisis x Orden de Despacho"
      Height          =   390
      Left            =   3765
      TabIndex        =   21
      Top             =   3450
      Width           =   2295
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Consulta de Items"
      Height          =   390
      Left            =   4620
      TabIndex        =   20
      Top             =   420
      Width           =   2295
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Compras"
      Height          =   405
      Left            =   7005
      TabIndex        =   19
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Egresos Caja y Bancos"
      Height          =   390
      Left            =   75
      TabIndex        =   18
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Aprobar Orden de Compra"
      Height          =   405
      Left            =   7005
      TabIndex        =   17
      Top             =   2055
      Width           =   2295
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Orden de Compra"
      Height          =   405
      Left            =   7005
      TabIndex        =   16
      Top             =   1650
      Width           =   2295
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Activar Cotizacion"
      Height          =   405
      Left            =   7005
      TabIndex        =   15
      Top             =   1230
      Width           =   2295
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Requerimiento"
      Height          =   390
      Left            =   6990
      TabIndex        =   14
      Top             =   15
      Width           =   2295
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Cotizar"
      Height          =   390
      Left            =   6990
      TabIndex        =   13
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Activar Requerimiento"
      Height          =   405
      Left            =   6990
      TabIndex        =   12
      Top             =   420
      Width           =   2295
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Hoja de Ruta"
      Height          =   390
      Left            =   75
      TabIndex        =   11
      Top             =   2040
      Width           =   1650
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Mantenimiento de Items"
      Height          =   390
      Left            =   4620
      TabIndex        =   10
      Top             =   15
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Programacion de Mantenimiento Preventivo"
      Height          =   390
      Left            =   2220
      TabIndex        =   9
      Top             =   1230
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Mantenimiento de Frecuencua de Mantenimiento"
      Height          =   390
      Left            =   2220
      TabIndex        =   8
      Top             =   825
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Mantenimiento de Tareas de Mantenimiento"
      Height          =   390
      Left            =   2220
      TabIndex        =   7
      Top             =   420
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Mantenimiento de Equipos"
      Height          =   390
      Left            =   2220
      TabIndex        =   6
      Top             =   15
      Width           =   2295
   End
   Begin VB.CommandButton Command55 
      Caption         =   "Linea de Tiempo"
      Height          =   390
      Left            =   75
      TabIndex        =   5
      Top             =   1650
      Width           =   1650
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Maestro de Recetas"
      Height          =   390
      Left            =   75
      TabIndex        =   4
      Top             =   1245
      Width           =   1650
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   390
      Left            =   2250
      TabIndex        =   3
      Top             =   2670
      Width           =   5880
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Maestro de Rendimientos"
      Height          =   390
      Left            =   75
      TabIndex        =   2
      Top             =   840
      Width           =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cronograma de Produccion"
      Height          =   390
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cronograma de Tareas"
      Height          =   390
      Left            =   75
      TabIndex        =   0
      Top             =   435
      Width           =   1650
   End
End
Attribute VB_Name = "FrmMenuRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.CronogramaTareas xCon
    Set xFun = Nothing
End Sub

Private Sub Command10_Click()
    Dim xFrm As New SGI2_almacen.Almacen
    xFrm.IdUsuario = xIdUsuario
    xFrm.MantItem xCon, 1
    Set xFrm = Nothing
End Sub

Private Sub Command11_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.HojaDeRuta xCon
    Set xFrm = Nothing
End Sub

Private Sub Command12_Click()
    Dim xFrm As New seven_compras2.compras
    xFrm.AprobarRequerimiento 1, xCon
    Set xFrm = Nothing
End Sub

Private Sub Command13_Click()
    Dim xFrm As New seven_compras2.compras
    xFrm.ManOrdenCotizacion xCon
    Set xFrm = Nothing
End Sub

Private Sub Command14_Click()
    Dim xFrm As New seven_compras2.compras
    xFrm.ManOrdenrequerimiento xCon
    Set xFrm = Nothing
End Sub

Private Sub Command15_Click()
    Dim xFrm As New seven_compras2.compras
    xFrm.AprobarCotizacion 1, xCon
    Set xFrm = Nothing
End Sub

Private Sub Command16_Click()
    Dim xFrm As New seven_compras2.compras
    xFrm.ManOrdenCompra 1, xCon
    Set xFrm = Nothing
End Sub

Private Sub Command17_Click()
    Dim xFrm As New seven_compras2.compras
    xFrm.AprobarOrdenCompra 1, xCon
    Set xFrm = Nothing
End Sub

Private Sub Command18_Click()
    
    Dim xFr As New sgi2_cajabancos.cajabancos
    xFr.EgresoCajaBanco2 xCon, 5
    Set xFr = Nothing
End Sub

Private Sub Command19_Click()
    Dim xFrm As New sgi2_compras.compras
    xFrm.RegCompras2 xCon, 1, 1
    Set xFrm = Nothing
End Sub

Private Sub Command2_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.CronogramaProduccion xCon
    Set xFun = Nothing
End Sub

Private Sub Command20_Click()
    Dim xFrm As New SGI2_almacen.Almacen
    xFrm.ConsultaItems xCon
    Set xFrm = Nothing
End Sub

Private Sub Command21_Click()
    Dim xFun As New sgi2_gestion.Gestion
    xFun.ConsultaDocReferencia xCon
    Set xFun = Nothing
End Sub

Private Sub Command22_Click()
    Dim xFun As New sgi2_gestion.Gestion
    xFun.ConsultaUnidadNegocio xCon
    Set xFun = Nothing
End Sub

Private Sub Command23_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.GenOrdenProduccion xCon
    Set xFun = Nothing
End Sub

Private Sub Command24_Click()
    Dim xFun As New sgi2_planillas.planillas
    xFun.ConsRotacion xCon, AP_RUTASY
    Set xFun = Nothing
End Sub

Private Sub Command3_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.Rendimiento xCon
    Set xFun = Nothing
End Sub

Private Sub Command4_Click()
    Unload Me
    Set xCon = Nothing
    End
End Sub


Private Sub Command5_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.MamRecetas xCon
    Set xFun = Nothing

End Sub

Private Sub Command55_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.LineaDeTiempo xCon
    Set xFun = Nothing
    
End Sub

Private Sub Command6_Click()
    Dim xFrm As New sgi2_mantenimiento.mantenimiento
    xFrm.MantEquipos xCon
    Set xFrm = Nothing
End Sub

Private Sub Command7_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT man_tareas.* From man_tareas ORDER BY man_tareas.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "id":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Codigo":         xCamposVista(1, 1) = "cod":            xCamposVista(1, 2) = "900":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Descripcion":    xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "6000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Id":             xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "cod":          xCampos(1, 2) = "C":    xCampos(1, 3) = "800"
    xCampos(2, 0) = "Descripcion":    xCampos(2, 1) = "descripcion":  xCampos(2, 2) = "C":    xCampos(2, 3) = "6000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "man_tareas"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Id"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Tareas de Mantenimiento"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub Command8_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT man_frecuencia.* From man_frecuencia ORDER BY man_frecuencia.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "id":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Codigo":         xCamposVista(1, 1) = "cod":            xCamposVista(1, 2) = "900":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Descripcion":    xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "6000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Id":             xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "cod":          xCampos(1, 2) = "C":    xCampos(1, 3) = "800"
    xCampos(2, 0) = "Descripcion":    xCampos(2, 1) = "descripcion":  xCampos(2, 2) = "C":    xCampos(2, 3) = "6000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "man_frecuencia"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Id"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Frecuencia de Mantenimiento"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub Command9_Click()
    Dim xFrm As New sgi2_mantenimiento.mantenimiento
    xFrm.ManPreventivo xCon
    Set xFrm = Nothing
End Sub

Private Sub Form_Load()
    If xNivelUsuario = 2 Then
        'Command1.Enabled = True
'        Command2.Enabled = True
'        Command5.Enabled = True
'        Command55.Enabled = True
    End If
End Sub

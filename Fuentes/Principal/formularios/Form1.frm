VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command44 
      Caption         =   "Ver Letras"
      Height          =   450
      Left            =   8385
      TabIndex        =   49
      Top             =   6570
      Width           =   1620
   End
   Begin VB.CommandButton Command43 
      Caption         =   "Command43"
      Height          =   360
      Left            =   6240
      TabIndex        =   48
      Top             =   6750
      Width           =   1920
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Produccion"
      Height          =   390
      Left            =   6405
      TabIndex        =   47
      Top             =   6270
      Width           =   1620
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Transferir Operaciones"
      Height          =   390
      Left            =   8385
      TabIndex        =   46
      Top             =   6165
      Width           =   1620
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Registrar Tareas"
      Height          =   390
      Left            =   3990
      TabIndex        =   44
      Top             =   6300
      Width           =   1620
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Registrar Producto Dia"
      Height          =   390
      Left            =   2250
      TabIndex        =   43
      Top             =   6225
      Width           =   1620
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Command37"
      Height          =   390
      Left            =   285
      TabIndex        =   42
      Top             =   6240
      Width           =   1620
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Ventas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   30
      TabIndex        =   13
      Top             =   4995
      Width           =   11025
      Begin VB.CommandButton Command27 
         Caption         =   "Actualizar Datos de Venta"
         Height          =   495
         Left            =   7020
         TabIndex        =   32
         Top             =   375
         Width           =   1425
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Actualizar Datos de Venta"
         Height          =   495
         Left            =   5190
         TabIndex        =   31
         Top             =   375
         Width           =   1800
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Choferes"
         Height          =   495
         Left            =   3960
         TabIndex        =   21
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Clientes"
         Height          =   495
         Left            =   2730
         TabIndex        =   16
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Ventas"
         Height          =   495
         Left            =   1500
         TabIndex        =   15
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Empaques"
         Height          =   495
         Left            =   270
         TabIndex        =   14
         Top             =   375
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Caja y Bancos ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   30
      TabIndex        =   9
      Top             =   15
      Width           =   11040
      Begin VB.CommandButton Command36 
         Caption         =   "Honorarios"
         Height          =   495
         Left            =   5205
         TabIndex        =   41
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command34 
         Caption         =   "Proveddores"
         Height          =   495
         Left            =   3975
         TabIndex        =   39
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Actualizar Precios"
         Height          =   495
         Left            =   2745
         TabIndex        =   30
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Mantenimiento de Items"
         Height          =   495
         Left            =   1515
         TabIndex        =   29
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Compras"
         Height          =   495
         Left            =   270
         TabIndex        =   10
         Top             =   375
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Planillas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   30
      TabIndex        =   5
      Top             =   3930
      Width           =   11025
      Begin VB.CommandButton Command1 
         Caption         =   "Planillas"
         Height          =   495
         Left            =   270
         TabIndex        =   8
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Nomina"
         Height          =   495
         Left            =   1500
         TabIndex        =   7
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Comision"
         Height          =   495
         Left            =   2730
         TabIndex        =   6
         Top             =   375
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Caja y Bancos ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   30
      TabIndex        =   3
      Top             =   2655
      Width           =   11040
      Begin VB.CommandButton Command32 
         Caption         =   "Maestro Origen de Ingresos"
         Height          =   465
         Left            =   5355
         TabIndex        =   37
         Top             =   600
         Width           =   1545
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Maestro Destino de Ingresos"
         Height          =   465
         Left            =   6915
         TabIndex        =   36
         Top             =   600
         Width           =   1545
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Caja y Bancos Ingreso"
         Height          =   495
         Left            =   1560
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Diferencia Tipo  de Cambio"
         Height          =   465
         Left            =   9360
         TabIndex        =   34
         Top             =   120
         Width           =   1545
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Maestro Destino de Egresos"
         Height          =   465
         Left            =   6930
         TabIndex        =   27
         Top             =   120
         Width           =   1545
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Maestro Origen de Egresos"
         Height          =   465
         Left            =   5370
         TabIndex        =   26
         Top             =   120
         Width           =   1545
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Libro Bancos"
         Height          =   495
         Left            =   4080
         TabIndex        =   25
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Analisis Cliente Proveedor"
         Height          =   495
         Left            =   2835
         TabIndex        =   19
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Caja y Bancos Egresos"
         Height          =   495
         Left            =   270
         TabIndex        =   4
         Top             =   375
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Contabilidad ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   30
      TabIndex        =   0
      Top             =   1110
      Width           =   11040
      Begin VB.CommandButton Command40 
         Caption         =   "Analisis de cuenta"
         Height          =   495
         Left            =   9270
         TabIndex        =   45
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton Command35 
         Caption         =   "Registro de Ventas"
         Height          =   495
         Left            =   5565
         TabIndex        =   40
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command33 
         Caption         =   "Libro Diario"
         Height          =   495
         Left            =   9270
         TabIndex        =   38
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Maestro de Documentos"
         Height          =   495
         Left            =   8055
         TabIndex        =   33
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Plan de Cuentas"
         Height          =   495
         Left            =   8055
         TabIndex        =   28
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Libro Mayor"
         Height          =   495
         Left            =   6825
         TabIndex        =   24
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Hoja de Trabajo"
         Height          =   495
         Left            =   6825
         TabIndex        =   23
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Registro de Compras"
         Height          =   495
         Left            =   4320
         TabIndex        =   22
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton Cmdaaa 
         Caption         =   "Kardex Unificado"
         Height          =   495
         Left            =   2760
         TabIndex        =   20
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Detraccion"
         Height          =   495
         Left            =   1515
         TabIndex        =   18
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Maestro de Destinos"
         Height          =   495
         Left            =   4320
         TabIndex        =   17
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Centro de Costos Unificado"
         Height          =   495
         Left            =   2760
         TabIndex        =   12
         Top             =   375
         Width           =   1545
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Centro de Costos"
         Height          =   495
         Left            =   1515
         TabIndex        =   11
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Percepcion"
         Height          =   495
         Left            =   270
         TabIndex        =   2
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Proviciones Diversas"
         Height          =   495
         Left            =   270
         TabIndex        =   1
         Top             =   375
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub pLimpiarBD(Cnn As ADODB.Connection)
    Err.Clear
    On Error GoTo ERROR
    If MsgBox("Seguro desea Limpiar La base de Datos", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    MsgBox "A Continuación se procederá a limpiar la Base de Datos" + vbCr + "Este proceso puede tardar...espere...", vbInformation, xTitulo
    
    Cnn.BeginTrans
    
    Cnn.Execute "DELETE FROM ges_planabadet; "
    Cnn.Execute "DELETE FROM ges_planabapropro; "
    Cnn.Execute "DELETE FROM ges_planaba; "
    Cnn.Execute "DELETE FROM ges_planventasdet; "
    Cnn.Execute "DELETE FROM ges_planventas; "
    Cnn.Execute "DELETE FROM ges_plaproddet2; "
    Cnn.Execute "DELETE FROM ges_plaproddet; "
    Cnn.Execute "DELETE FROM ges_plaprod; "
    Cnn.Execute "DELETE FROM ges_ventaproydet; "
    Cnn.Execute "DELETE FROM ges_ventaproydetori; "
    Cnn.Execute "DELETE FROM ges_ventaproy; "
    Cnn.Execute "DELETE FROM ges_rutahis; "
    
    Cnn.Execute "DELETE FROM pla_comisiondet; "
    Cnn.Execute "DELETE FROM pla_comision; "
    Cnn.Execute "DELETE FROM pla_planillas; "
    
    Cnn.Execute "DELETE FROM pro_estacionalidadestados; "
    Cnn.Execute "DELETE FROM pro_estacionalidad; "
    Cnn.Execute "DELETE FROM pro_ordensalidadet; "
    Cnn.Execute "DELETE FROM pro_ordensalida; "
    Cnn.Execute "DELETE FROM pro_producciondetins; "
    Cnn.Execute "DELETE FROM pro_producciondettar; "
    Cnn.Execute "DELETE FROM pro_producciondet; "
    Cnn.Execute "DELETE FROM pro_produccion; "
    Cnn.Execute "DELETE FROM pro_programadet; "
    Cnn.Execute "DELETE FROM pro_programa; "
    Cnn.Execute "DELETE FROM pro_recetains; "
    Cnn.Execute "DELETE FROM pro_recetatar; "
    Cnn.Execute "DELETE FROM pro_tareas; "
    Cnn.Execute "DELETE FROM pro_receta; "
    Cnn.Execute "DELETE FROM pro_emp; "
    
    Cnn.Execute "DELETE FROM pvt_cotizaciondet; "
    Cnn.Execute "DELETE FROM pvt_cotizacion; "
    Cnn.Execute "DELETE FROM pvt_desccorporativo; "
    Cnn.Execute "DELETE FROM pvt_descgeneral; "
    Cnn.Execute "DELETE FROM pvt_items; "
    Cnn.Execute "DELETE FROM pvt_emp; "
    
    Cnn.Execute "DELETE FROM vta_cotizaciondet; "
    Cnn.Execute "DELETE FROM vta_cotizacion; "
    Cnn.Execute "DELETE FROM vta_notascreabodet; "
    Cnn.Execute "DELETE FROM vta_notascreabo; "
    Cnn.Execute "DELETE FROM vta_guiadet; "
    Cnn.Execute "DELETE FROM vta_guia; "
    Cnn.Execute "DELETE FROM vta_pedidodet; "
    Cnn.Execute "DELETE FROM vta_pedido; "
    Cnn.Execute "DELETE FROM vta_ventasdet; "
    Cnn.Execute "DELETE FROM vta_ventas; "
    Cnn.Execute "DELETE FROM vta_puntoVenta; "
    Cnn.Execute "DELETE FROM vta_vendedores; "
    Cnn.Execute "DELETE FROM com_preciosdet; "
    Cnn.Execute "DELETE FROM com_precios; "
    Cnn.Execute "DELETE FROM com_ordencompradet; "
    Cnn.Execute "DELETE FROM com_ordencompra; "
    Cnn.Execute "DELETE FROM com_comprascosto; "
    Cnn.Execute "DELETE FROM com_comprasdet; "
    Cnn.Execute "DELETE FROM com_compras; "
    Cnn.Execute "DELETE FROM con_cajabancocon; "
    Cnn.Execute "DELETE FROM con_cajabancodet; "
    Cnn.Execute "DELETE FROM con_cajabancoorides; "
    Cnn.Execute "DELETE FROM con_cajabanco; "
    Cnn.Execute "DELETE FROM con_canjesdet; "
    Cnn.Execute "DELETE FROM con_canjes; "
    Cnn.Execute "DELETE FROM con_detraccion; "
    Cnn.Execute "DELETE FROM con_devolucionesdet; "
    Cnn.Execute "DELETE FROM con_devoluciones; "
    Cnn.Execute "DELETE FROM con_ctasrendir; "
    Cnn.Execute "DELETE FROM con_letradet; "
    Cnn.Execute "DELETE FROM con_letradoc; "
    Cnn.Execute "DELETE FROM con_letra; "
    Cnn.Execute "DELETE FROM con_ordenpagodet; "
    Cnn.Execute "DELETE FROM con_ordenpago; "
    Cnn.Execute "DELETE FROM con_percepciondet; "
    Cnn.Execute "DELETE FROM con_percepcion; "
    Cnn.Execute "DELETE FROM con_provicionesdet; "
    Cnn.Execute "DELETE FROM con_proviciones; "
    Cnn.Execute "DELETE FROM con_retenciondet; "
    Cnn.Execute "DELETE FROM con_retencion; "
    Cnn.Execute "DELETE FROM con_diario; "
    
    Cnn.Execute "DELETE FROM alm_ingresodet; "
    Cnn.Execute "DELETE FROM alm_ingreso; "
    Cnn.Execute "DELETE FROM alm_invencencos; "
    Cnn.Execute "DELETE FROM alm_inventarioalmacen; "
    Cnn.Execute "DELETE FROM alm_inventariofoto; "
    Cnn.Execute "DELETE FROM alm_inventario; "

    Cnn.CommitTrans
    MsgBox "La base de datos esta limpio", vbInformation, xTitulo
    Exit Sub

ERROR:
    Cnn.RollbackTrans
    MsgBox "No se pudo limpiar la base de datos por el siguiente motivo: " + vbCr + _
    Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
End Sub

Private Sub Cmdaaa_Click()
    Dim xFun As New sgi2_contabilidad2.Reportes
    xFun.KardexUnificado xCon
    Set xFun = Nothing
End Sub

Private Sub Command1_Click()
'    Dim xFrm As New sgi2_planillas.planillas
'    xFrm.ProcesarPlanilla xCon
'    Set xFrm = Nothing
End Sub


Private Sub Command10_Click()
    Dim xFun As New sgi2_ventas.ventas
    xFun.Empaques xCon
    Set xFun = Nothing
End Sub

Private Sub Command11_Click()
    Dim xFun As New sgi2_ventas.ventas
    xFun.ventas xCon, 4
    Set xFun = Nothing
End Sub

Private Sub Command12_Click()
    Dim xfm As New sgi2_ventas.ventas
    xfm.Clientes xCon, 1
    Set xfm = Nothing
End Sub

Private Sub Command13_Click()
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.ManDestinos 2, xCon
    Set xFun = Nothing
End Sub

Private Sub Command14_Click()
    Dim xFun As New sgi2_contabilidad.Mantenimiento
    xFun.ManDetraccion xCon, 1, DET_Compra
    Set xFun = Nothing
End Sub

Private Sub Command15_Click()
    'Dim xFun As New sgi2_cajabancos.cajabancos
    'xFun.ConsultaCtaCte xCon
    'Set xFun = Nothing
End Sub


Private Sub Command16_Click()
    Dim xform As New Eps_MantTablas.Mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_chofer.id, mae_chofer.idvehiculo, UCase(pla_empleados!ape)+', '+pla_empleados!nom AS apenom, mae_chofer.numbre, " _
        & " mae_vehiculo.marca, mae_vehiculo.numpla, mae_chofer.categoria, mae_chofer.idper FROM pla_empleados RIGHT JOIN (mae_vehiculo " _
        & " RIGHT JOIN mae_chofer ON mae_vehiculo.id = mae_chofer.idvehiculo) ON pla_empleados.id = mae_chofer.idper " _
        & " ORDER BY UCase(pla_empleados!ape)+', '+pla_empleados!nom"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Apellido y Nombres":   xCamposVista(0, 1) = "apenom":         xCamposVista(0, 2) = "5000":   xCamposVista(0, 3) = "C":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Nº Brevete":           xCamposVista(1, 1) = "numbre":         xCamposVista(1, 2) = "1200":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Marca":                xCamposVista(2, 1) = "marca":          xCamposVista(2, 2) = "1500":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Nº Placa":             xCamposVista(3, 1) = "numpla":         xCamposVista(3, 2) = "1200":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Nom. Apellidos":  xCampos(1, 1) = "idper":        xCampos(1, 2) = "N":    xCampos(1, 3) = "3000"
    xCampos(2, 0) = "Nº Brevete":      xCampos(2, 1) = "numbre":       xCampos(2, 2) = "C":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Categoria":       xCampos(3, 1) = "categoria":    xCampos(3, 2) = "C":    xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Vehiculo":        xCampos(4, 1) = "idvehiculo":   xCampos(4, 2) = "N":    xCampos(4, 3) = "1500"
    
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
    xVinculos(0, 0) = "pla_empleados":              xVinculos(0, 1) = "id":              xVinculos(0, 2) = "ape,nom,id":
    xVinculos(0, 3) = "Apellidos,Nombres,Codigo":   xVinculos(0, 4) = "2000,2000,1000":  xVinculos(0, 5) = "C,C,N":
    xVinculos(0, 6) = "idper":                      xVinculos(0, 7) = "nom":             xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "ape"

    xVinculos(1, 0) = "mae_vehiculo":               xVinculos(1, 1) = "id":              xVinculos(1, 2) = "marca,numpla,id":
    xVinculos(1, 3) = "Marca,Nº Placa,Codigo":      xVinculos(1, 4) = "2000,2000,1000":  xVinculos(1, 5) = "C,C,N":
    xVinculos(1, 6) = "idvehiculo":                 xVinculos(1, 7) = "numpla":          xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "marca"

    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "idper"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_chofer"
    xform.CampoOrdenado = "id"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Choferes"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub Command17_Click()
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.VerRegCompras xCon
    Set xFun = Nothing
End Sub

Private Sub Command18_Click()
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.HojaTrabajo xCon
    Set xFun = Nothing
End Sub

Private Sub Command19_Click()
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.Mayor xCon
    Set xFun = Nothing
End Sub

Private Sub Command2_Click()
'    Dim xFun As New sgi2_planillas.planillas
'    xFun.ManNomina xCon
'    Set xFun = Nothing
End Sub

Private Sub Command20_Click()
    Dim xfrm As New sgi2_cajabancos.cajabancos
    xfrm.Librobancos xCon
    Set xfrm = Nothing
End Sub

Private Sub Command21_Click()
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.ManOrigenes 2, xCon
    Set xFun = Nothing

'    Dim xform As New Eps_MantTablas.Mantenimiento
'    Dim xNivelUsuario As Integer
'    Dim xCampos(5, 4) As String
'    Dim xVinculos(3, 10) As String
'    Dim xCampoBusca(2) As String
'    Dim xCamposVista(5, 4) As String
'    Dim xConsulta As String
'
'    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
'    xConsulta = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, con_origen.id, con_origen.idmon, con_origen.descripcion, " _
'        & " con_origen.idcue, con_origen.tipmov, con_origen.idorigen, mae_moneda.descripcion AS descmon " _
'        & " FROM con_planctas INNER JOIN (con_origen LEFT JOIN mae_moneda ON con_origen.idmon = mae_moneda.id) ON con_planctas.id = con_origen.idcue " _
'        & " WHERE (((con_origen.tipmov)=2))"
'
'
'    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
'    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
'    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "3000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
'    xCamposVista(2, 0) = "Cuenta":             xCamposVista(2, 1) = "cuenta":         xCamposVista(2, 2) = "1200":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
'    xCamposVista(3, 0) = "Descripcion Cuenta": xCamposVista(3, 1) = "descuen":        xCamposVista(3, 2) = "3000":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "R"
'    xCamposVista(4, 0) = "Moneda":             xCamposVista(4, 1) = "descmon":        xCamposVista(4, 2) = "1200":   xCamposVista(4, 3) = "C":    xCamposVista(4, 4) = "I"
'
'    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
'    xCampos(0, 0) = "Codigo":             xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
'    xCampos(1, 0) = "Descripcion":        xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
'    xCampos(2, 0) = "Cuenta":             xCampos(2, 1) = "idcue":        xCampos(2, 2) = "N":    xCampos(2, 3) = "1100"
'    xCampos(3, 0) = "Tipo Movimiento":    xCampos(3, 1) = "tipmov":       xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
'    xCampos(4, 0) = "Moneda":             xCampos(4, 1) = "idmon":        xCampos(4, 2) = "N":    xCampos(4, 3) = "1100"
'    'xCampos(5, 0) = "Origen":             xCampos(5, 1) = "idorigen":     xCampos(5, 2) = "N":    xCampos(5, 3) = "1100"
'
'    '0 = nombre de la tabla
'    '1 = nombre del campo con el que se iniciara la busqueda
'    '2 = lista de campos
'    '3 = lista de rotulos para los campos
'    '4 = tamaño de los campos
'    '5 = tipo de los campos
'    '6 = nombre del campo con el que se vincula el array anterior
'    '7 = campo que devolvera la busqueda
'    '8 = tipo del campo que iniciara la busqueda
'    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
'    xVinculos(0, 0) = "con_planctas":       xVinculos(0, 1) = "id":          xVinculos(0, 2) = "cuenta,descripcion":
'    xVinculos(0, 3) = "Cuenta,Descripcion": xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "C,C":
'    xVinculos(0, 6) = "idcue":              xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
'    xVinculos(0, 9) = "cuenta"
'
'    xVinculos(1, 0) = "mae_tipomov":        xVinculos(1, 1) = "id":          xVinculos(1, 2) = "id,descripcion":
'    xVinculos(1, 3) = "Codigo,Descripcion": xVinculos(1, 4) = "1100,4000":   xVinculos(1, 5) = "N,C":
'    xVinculos(1, 6) = "tipmov":             xVinculos(1, 7) = "descripcion": xVinculos(1, 8) = "N":
'    xVinculos(1, 9) = "id"
'
'    xVinculos(2, 0) = "mae_moneda":         xVinculos(2, 1) = "id":          xVinculos(2, 2) = "id,descripcion":
'    xVinculos(2, 3) = "Codigo,Descripcion": xVinculos(2, 4) = "1100,4000":   xVinculos(2, 5) = "N,C":
'    xVinculos(2, 6) = "idmon":              xVinculos(2, 7) = "descripcion": xVinculos(2, 8) = "N":
'    xVinculos(2, 9) = "id"
'
''    xVinculos(3, 0) = "con_origenes":       xVinculos(3, 1) = "id":          xVinculos(3, 2) = "id,descripcion":
''    xVinculos(3, 3) = "Codigo,Descripcion": xVinculos(3, 4) = "1100,4000":   xVinculos(3, 5) = "N,C":
''    xVinculos(3, 6) = "idorigen":           xVinculos(3, 7) = "descripcion": xVinculos(3, 8) = "N":
''    xVinculos(3, 9) = "id"
'
'    'CAMPOS PARA EFECTUAR LA BUSQUEDA
'    xCampoBusca(0) = "descripcion"
'    xCampoBusca(1) = "id"
'
'    If xNivelUsuario = 0 Then
'        xform.PermiteActualiza = False
'    Else
'        xform.PermiteActualiza = True
'    End If
'    xform.CadSQLVista = xConsulta
'    xform.Tabla = "con_origen"
'    xform.CampoOrdenado = "descripcion"
'    xform.CampoClave = "Codigo"
'    xform.PermiteActualiza = True
'    xform.TituloFormulario = "Mantenimiento - Origen de Egresos"
'    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon

End Sub

Private Sub Command22_Click()
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.ManDestinos 2, xCon
    Set xFun = Nothing
End Sub

Private Sub Command23_Click()
    Dim xfrm As New sgi2_contabilidad.Mantenimiento
    xfrm.ManPlanCuentas xCon
    Set xfrm = Nothing
End Sub

Private Sub Command24_Click()
    Dim xFun As New SGI2_almacen.Almacen
    xFun.MantItem xCon, 1
    Set xFun = Nothing
End Sub

Private Sub Command25_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.ActualizaPrecio xCon
    Set xfrm = Nothing
End Sub

Private Sub Command27_Click()
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.ActualizarDatosVenta xCon
    Set xfrm = Nothing
End Sub

Private Sub Command28_Click()
    Dim xform As New Eps_MantTablas.Mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(6, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_documento.*, mae_impuestos.descripcion AS descimp, mae_impuestos.tasa" _
        & " FROM mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Cod. Sunat":           xCamposVista(1, 1) = "codsun":         xCamposVista(1, 2) = "1100":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "C"
    xCamposVista(2, 0) = "Descripcion":          xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "3000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Abreviatura":          xCamposVista(3, 1) = "abrev":          xCamposVista(3, 2) = "1100":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "C"
    xCamposVista(4, 0) = "Impuesto":             xCamposVista(4, 1) = "descimp":        xCamposVista(4, 2) = "3000":   xCamposVista(4, 3) = "C":    xCamposVista(4, 4) = "I"
    xCamposVista(5, 0) = "Tasa":                 xCamposVista(5, 1) = "tasa":           xCamposVista(5, 2) = "1000":   xCamposVista(5, 3) = "C":    xCamposVista(5, 4) = "D"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":        xCampos(2, 2) = "C":    xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Cod. Sunat":     xCampos(3, 1) = "codsun":       xCampos(3, 2) = "C":    xCampos(3, 3) = "1200"
    xCampos(4, 0) = "Impuesto":       xCampos(4, 1) = "idimp":        xCampos(4, 2) = "N":    xCampos(4, 3) = "1200"
    
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
    
    xVinculos(0, 0) = "mae_impuestos":       xVinculos(0, 1) = "id":            xVinculos(0, 2) = "id,descripcion":
    xVinculos(0, 3) = "Codigo,Descripcion":  xVinculos(0, 4) = "1000,5000":     xVinculos(0, 5) = "N,C":
    xVinculos(0, 6) = "idimp":               xVinculos(0, 7) = "descripcion":   xVinculos(0, 8) = "N"
    xVinculos(0, 9) = "id"
    
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_documento"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Documentos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub Command29_Click()
    Dim xfrm As New sgi2_cajabancos.cajabancos
    xfrm.DiferenciaCambio xCon, 1
End Sub

Private Sub Command3_Click()
'    Dim xFin  As New sgi2_planillas.planillas
'    xFin.Comision xCon
'    Set xFin = Nothing
End Sub

Private Sub Command30_Click()
    Dim xFr As New sgi2_cajabancos.cajabancos
    xFr.IngresoCajaBanco2 xCon, 1
    Set xFr = Nothing
End Sub

Private Sub Command31_Click()
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.ManDestinos 1, xCon
    Set xFun = Nothing
End Sub

Private Sub Command32_Click()
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.ManOrigenes 1, xCon
    Set xFun = Nothing
End Sub

Private Sub Command33_Click()
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.VerDiario xCon
    Set xFun = Nothing
End Sub

Private Sub Command34_Click()
    Dim xFun As New sgi2_compras.Compras
    xFun.ManProveedor xCon, 1
    Set xFun = Nothing
End Sub

Private Sub Command35_Click()
    Dim xFun As New sgi2_contabilidad.Consultas
    xFun.VerRegVentas xCon
    Set xFun = Nothing
End Sub

Private Sub Command36_Click()
    Dim xFun As New sgi2_compras.Compras
    xFun.RegHonorarios xCon, 2, 0
    Set xFun = Nothing
End Sub

Private Sub Command37_Click()
    Dim xfm As New Sgi2_planilla4.planillas
    xfm.IngresoRapidoEmpleados xCon
    Set xfm = Nothing
End Sub

Private Sub Command38_Click()
    Dim xFun As New Sgi2_planilla4.planillas
    xFun.RegProductosDia xCon
    Set xFun = Nothing
End Sub

Private Sub Command39_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.VerRegRent4ta xCon
    Set xfrm = Nothing
End Sub

Private Sub Command4_Click()
    Dim xFun As New sgi2_cajabancos.cajabancos
    xFun.EgresoCajaBanco2 xCon, 8
    'xFun.IngresoCajaBanco xCon, 1
    Set xFun = Nothing
End Sub

Private Sub Command40_Click()
'    Dim xF As New sgi2_contabilidad2.estadosfinancieros
'    xF.AnalisisCuenta xCon
'    Set xF = Nothing
End Sub

Private Sub Command41_Click()
'    Dim xFun As New sgi2_compras.Compras
'    xFun.Huevada xCon
'    Set xCon = Nothing
End Sub

Private Sub Command42_Click()
    Dim xFun As New sgi2_produccion.produccion
    xFun.MamRecetas xCon
    Set xFun = Nothing
End Sub

Private Sub Command43_Click()
    Dim xfrm As New Sgi2_Procesos.Procesos
    xfrm.BDEvaluar xCon
    Set xfrm = Nothing
End Sub

Private Sub Command44_Click()
    Dim xfrm As New sgi2_letras.letras
    xfrm.ManLetras 3, xCon
    Set xfrm = Nothing
End Sub

Private Sub Command5_Click()
    Dim xFun As New sgi2_contabilidad.Mantenimiento
    xFun.ManPercepcion xCon, 2
    Set xFun = Nothing
End Sub

Private Sub Command6_Click()
    Dim xFun As New sgi2_contabilidad.Mantenimiento
    xFun.ManProviciones2 xCon, 0
    Set xFun = Nothing
End Sub

Private Sub Command7_Click()
    Dim xFun As New sgi2_compras.Compras
    xFun.RegCompras2 xCon, 2, 0
    Set xFun = Nothing
End Sub

Private Sub Command8_Click()
    Dim xFun As New sgi2_contabilidad2.Reportes
    xFun.CentroCostos xCon
    Set xFun = Nothing
End Sub

Private Sub Command9_Click()
    Dim xFun As New sgi2_contabilidad2.Reportes
    xFun.CentroCostosUnificado xCon
    Set xFun = Nothing
End Sub


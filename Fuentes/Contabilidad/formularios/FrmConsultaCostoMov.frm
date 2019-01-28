VERSION 5.00
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaCostoMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Consulta de Costos de Movimientos"
   ClientHeight    =   1425
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "[ Opciones ]"
      Height          =   960
      Left            =   5020
      TabIndex        =   6
      Top             =   400
      Width           =   1815
      Begin VB.CheckBox OpcionCheck 
         Caption         =   "Detallado"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox OpcionCheck 
         Caption         =   "Resumido"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Seleccionar ]"
      Height          =   960
      Left            =   50
      TabIndex        =   1
      Top             =   400
      Width           =   4925
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1110
         TabIndex        =   4
         Top             =   400
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "23/03/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   3240
         TabIndex        =   5
         Top             =   400
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "23/03/2007"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   450
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   450
         Width           =   420
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":08D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":0A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":0DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":0F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":139A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":14B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":19F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":1F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":204E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":2162
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":25B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoMov.frx":2722
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Insertar Ítem"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu menu01 
         Caption         =   "Eliminar Ítem"
      End
   End
End
Attribute VB_Name = "FrmConsultaCostoMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FrmVerKardex.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA EL VINCAR DEL ITEM SELECCIONADO, ADEMAS PERMITE COSTEAS LAS SALIDAS
'*                    MEDIANTE EL METODO PROMEDIO PONDERADO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 23/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim rst As New ADODB.Recordset            ' RECORSET QUE ALAMCENARA LOS MOVIMIENTOS DEL ITEM
Dim SeEjecuto As Boolean                  ' VARIABLE QUE CONTROLARA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim StockIni As Double                    ' ALMACENA EL STOCK INICIAL DEL ITEM
Dim xPrecioIni As Double                  ' ALMACENA EL PRECIO INICIAL DEL ITEM
Dim MuestraRpt As Integer
Dim cSQL As String
Dim INDICE_ As Integer
Dim BAND_INTERRUMPIR As Boolean
Dim F As New SistemaLogica.Funciones

Private Sub pIniciarCampos()
    TxtFchIni.Valor = CDate("01/01/" & Year(Date))
    TxtFchFin.Valor = Date
    OpcionCheck(0).Value = 1
    BAND_INTERRUMPIR = False
    Blanquea
End Sub

Sub Blanquea()
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    pIniciarCampos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pCargarDatos
    
    If Button.Index = 5 Then
        Set rst = Nothing
        Unload Me
    End If
End Sub

Private Sub pCargarDatos()
    Dim mRecord As New ADODB.Recordset
    Dim FchInicio As Date
    Dim FchFinal As Date
    Dim mDataBase As New SistemaData.EDataBase
    Dim oExport As New SGI2_funciones.formularios
    Dim xCampos() As String
    
    If fValidarDatos() = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    Set mDataBase.Connection = xCon
    ' FILTROS DE BUSQUEDA
    BAND_INTERRUMPIR = False
    FchInicio = TxtFchIni.Valor
    FchFinal = TxtFchFin.Valor
    
    ' Resumido
    If OpcionCheck(0).Value = 1 Then
               
        cSQL = "SELECT CONSCOSTOMOV.fchmov As FechaMovimiento, CONSCOSTOMOV.tipmovcad As TipoMov, CONSCOSTOMOV.numdocconcat As NumeroDocumento, CONSCOSTOMOV.alm As Almacen, CONSCOSTOMOV.doc As TipoDocReferencia, CONSCOSTOMOV.numdocrefconcat As NumeroDocReferencia, Sum(CONSCOSTOMOV.costo) As CostoTotal" _
            + vbCr + "FROM ( " _
            + vbCr + F.SQL_MovDetallado(, , FchInicio, FchFinal, xCon) _
            + vbCr + ") As CONSCOSTOMOV " _
            + vbCr + "GROUP BY CONSCOSTOMOV.fchmov, CONSCOSTOMOV.tipmovcad, CONSCOSTOMOV.numdocconcat, CONSCOSTOMOV.alm, CONSCOSTOMOV.doc, CONSCOSTOMOV.numdocrefconcat"
        
'        cSQL = "SELECT alm_ingreso.fching AS FechaMovimiento, IIf([alm_ingreso]![tipmov]=-1,'ING.','SAL.') AS TipoMov, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS NumeroDocumento, alm_almacenes.descripcion AS Almacen, mae_documento.abrev AS TipoDocReferencia, " _
'                    & "IIf([alm_ingreso].[idtipdocref]=110,[pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc],IIf([alm_ingreso].[idtipdocref]=71,[alm_recepcion].[numser] & '-' & [alm_recepcion].[numdoc],IIf([alm_ingreso].[idtipdocref]=114,[alm_devolucion].[numser] & '-' & [alm_devolucion].[numdoc],IIf([alm_ingreso].[idtipdocref]=92,[com_ordencompra].[numser] & '-' & [com_ordencompra].[numdoc],IIf([alm_ingreso].[idtipdocref]=119,[alm_transferencia].[numser] & '-' & [alm_transferencia].[numdoc],IIf([alm_ingreso].[idtipdocref]=120,[pro_produccion].[numser] & '-' & [pro_produccion].[numdoc],IIf([alm_ingreso].[idtipdocref]=111,[alm_tomainventario].[numser] & '-' & [alm_tomainventario].[numdoc],IIf([alm_ingreso].[idtipdocref]=9,[vta_guia].[numser] & '-' & [vta_guia].[numdoc],IIf([alm_ingreso].[idtipdocref]=1,[vta_ventas].[numser] & '-' & [vta_ventas].[numdoc],''))))))))) AS NumeroDocReferencia, UCase([mae_estados].[descripcion]) AS Estado, " _
'                    & "Sum(con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS CostoTotal " _
'            + vbCr + "FROM (((((((((((((alm_ingreso LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id) LEFT JOIN mae_area ON alm_ingreso.idare = mae_area.id) LEFT JOIN mae_estados ON alm_ingreso.estado = mae_estados.id) LEFT JOIN pro_solicitudmat ON alm_ingreso.iddocref = pro_solicitudmat.id) LEFT JOIN alm_recepcion ON alm_ingreso.iddocref = alm_recepcion.id) LEFT JOIN alm_devolucion ON alm_ingreso.iddocref = alm_devolucion.id) LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id) LEFT JOIN com_ordencompra ON alm_ingreso.iddocref = com_ordencompra.id) LEFT JOIN alm_transferencia ON alm_ingreso.iddocref = alm_transferencia.idtransferencia) LEFT JOIN pro_produccion ON alm_ingreso.iddocref = pro_produccion.id) LEFT JOIN alm_tomainventario ON alm_ingreso.iddocref = alm_tomainventario.idtomainventario) LEFT JOIN vta_guia ON alm_ingreso.iddocref = vta_guia.id) LEFT JOIN vta_ventas ON alm_ingreso.iddocref = vta_ventas.id) " _
'                    & "LEFT JOIN (alm_ingresodet LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id " _
'            + vbCr + "WHERE (((alm_ingreso.fching) >= CDate('" & FchInicio & "')) And ((alm_ingreso.fching) <= CDate('" & FchFinal & "'))) " _
'            + vbCr + "GROUP BY alm_ingreso.fching, IIf([alm_ingreso]![tipmov]=-1,'ING.','SAL.'), [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc], alm_almacenes.descripcion, mae_documento.abrev, IIf([alm_ingreso].[idtipdocref]=110,[pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc],IIf([alm_ingreso].[idtipdocref]=71,[alm_recepcion].[numser] & '-' & [alm_recepcion].[numdoc],IIf([alm_ingreso].[idtipdocref]=114,[alm_devolucion].[numser] & '-' & [alm_devolucion].[numdoc],IIf([alm_ingreso].[idtipdocref]=92,[com_ordencompra].[numser] & '-' & [com_ordencompra].[numdoc],IIf([alm_ingreso].[idtipdocref]=119,[alm_transferencia].[numser] & '-' & [alm_transferencia].[numdoc],IIf([alm_ingreso].[idtipdocref]=120,[pro_produccion].[numser] & '-' & [pro_produccion].[numdoc],IIf([alm_ingreso].[idtipdocref]=111,[alm_tomainventario].[numser] & '-' & [alm_tomainventario].[numdoc],IIf([alm_ingreso].[idtipdocref]=9,[vta_guia].[numser] & '-' & [vta_guia].[numdoc], " _
'                    & "IIf([alm_ingreso].[idtipdocref]=1,[vta_ventas].[numser] & '-' & [vta_ventas].[numdoc],''))))))))), UCase([mae_estados].[descripcion]) " _
'            + vbCr + "ORDER BY alm_ingreso.fching DESC"
    ' Detallado
    Else
        cSQL = "SELECT CONSCOSTOMOV.fchmov As FechaMovimiento, CONSCOSTOMOV.tipmovcad As TipoMov, CONSCOSTOMOV.numdocconcat As NumeroDocumento, CONSCOSTOMOV.alm As Almacen, CONSCOSTOMOV.doc As TipoDocReferencia, CONSCOSTOMOV.numdocrefconcat As NumeroDocReferencia, CONSCOSTOMOV.coditem As CodigoItem, CONSCOSTOMOV.item As Item, CONSCOSTOMOV.cantidad As Cantidad, CONSCOSTOMOV.costounitariopromedio As CostoUniPromedio, CONSCOSTOMOV.costo As CostoTotal " _
            + vbCr + "FROM ( " _
            + vbCr + F.SQL_MovDetallado(, , FchInicio, FchFinal, xCon) _
            + vbCr + ") As CONSCOSTOMOV"
                    
'        cSQL = "SELECT alm_ingreso.fching AS FechaMovimiento, IIf([alm_ingreso]![tipmov]=-1,'ING.','SAL.') AS TipoMov, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS NumeroDocumento, alm_almacenes.descripcion AS Almacen, mae_documento.abrev AS TipoDocReferencia, IIf([alm_ingreso].[idtipdocref]=110,[pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc],IIf([alm_ingreso].[idtipdocref]=71,[alm_recepcion].[numser] & '-' & [alm_recepcion].[numdoc],IIf([alm_ingreso].[idtipdocref]=114,[alm_devolucion].[numser] & '-' & [alm_devolucion].[numdoc],IIf([alm_ingreso].[idtipdocref]=92,[com_ordencompra].[numser] & '-' & [com_ordencompra].[numdoc],IIf([alm_ingreso].[idtipdocref]=119,[alm_transferencia].[numser] & '-' & [alm_transferencia].[numdoc],IIf([alm_ingreso].[idtipdocref]=120,[pro_produccion].[numser] & '-' & [pro_produccion].[numdoc],IIf([alm_ingreso].[idtipdocref]=111, " _
'                & "[alm_tomainventario].[numser] & '-' & [alm_tomainventario].[numdoc],IIf([alm_ingreso].[idtipdocref]=9,[vta_guia].[numser] & '-' & [vta_guia].[numdoc],IIf([alm_ingreso].[idtipdocref]=1,[vta_ventas].[numser] & '-' & [vta_ventas].[numdoc],''))))))))) AS NumeroDocReferencia, UCase([mae_estados].[descripcion]) AS Estado, alm_inventario.codpro AS CodigoItem, alm_inventario.descripcion AS Item, alm_ingresodet.cantidad AS Cantidad, con_librocostotemp.costounitariopromedio AS CostoUniPromedio, (con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS CostoTotal " _
'            + vbCr + "FROM (((((((((((((alm_ingreso LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id) LEFT JOIN mae_area ON alm_ingreso.idare = mae_area.id) LEFT JOIN mae_estados ON alm_ingreso.estado = mae_estados.id) LEFT JOIN pro_solicitudmat ON alm_ingreso.iddocref = pro_solicitudmat.id) LEFT JOIN alm_recepcion ON alm_ingreso.iddocref = alm_recepcion.id) LEFT JOIN alm_devolucion ON alm_ingreso.iddocref = alm_devolucion.id) LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id) LEFT JOIN com_ordencompra ON alm_ingreso.iddocref = com_ordencompra.id) LEFT JOIN alm_transferencia ON alm_ingreso.iddocref = alm_transferencia.idtransferencia) LEFT JOIN pro_produccion ON alm_ingreso.iddocref = pro_produccion.id) LEFT JOIN alm_tomainventario ON alm_ingreso.iddocref = alm_tomainventario.idtomainventario) " _
'                & "LEFT JOIN vta_guia ON alm_ingreso.iddocref = vta_guia.id) LEFT JOIN vta_ventas ON alm_ingreso.iddocref = vta_ventas.id) LEFT JOIN ((alm_ingresodet LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id " _
'            + vbCr + "WHERE (((con_librocostotemp.costounitariopromedio) Is Not Null) AND ((alm_ingreso.fching) >= CDate('" & FchInicio & "')) And ((alm_ingreso.fching) <= CDate('" & FchFinal & "'))) " _
'            + vbCr + "ORDER BY alm_ingreso.fching DESC"
    End If
    
    mDataBase.CommandText = cSQL
    Set mRecord = mDataBase.GetRecordset
            
    If mRecord.RecordCount = 0 Then
        F.MostrarMensajeError "No se encontraron registros para la busqueda actual", "Costo de Movimientos"
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.MousePointer = vbDefault
    ' Se exporta a excel el recordset
    F.ExportarExcelRecordSet mRecord
    
    Set oExport = Nothing
    Set mRecord = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        F.MostrarMensajeError "La Fecha de Inicio no puede ser mayor a la de Fin", "Error"
        fValidarDatos = False
        Exit Function
    End If
    
    If OpcionCheck(0).Value = 1 And OpcionCheck(1).Value = 1 Then
        F.MostrarMensajeError "Debe seleccionar solo una de las opciones", "Error"
        fValidarDatos = False
        Exit Function
    End If
    fValidarDatos = True
End Function

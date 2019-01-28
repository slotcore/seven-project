VERSION 5.00
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaCostoParte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Consulta de Costos de Partes de Producción"
   ClientHeight    =   1425
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4440
      Top             =   120
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
            Picture         =   "FrmConsultaCostoParte.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":08D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":0A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":0DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":0F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":139A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":14B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":19F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":1F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":204E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":2162
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":25B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaCostoParte.frx":2722
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
      Width           =   5055
      _ExtentX        =   8916
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
Attribute VB_Name = "FrmConsultaCostoParte"
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

Dim SeEjecuto As Boolean                  ' VARIABLE QUE CONTROLARA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim cSQL As String
Dim F As New SistemaLogica.Funciones

Private Sub pIniciarCampos()
    TxtFchIni.Valor = CDate("01/01/" & Year(Date))
    TxtFchFin.Valor = Date
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
    FchInicio = TxtFchIni.Valor
    FchFinal = TxtFchFin.Valor
    
    ' Resumido
    
    cSQL = "SELECT pro_produccion.fchdoc AS FechaParte, [pro_produccion].[numser] & '-' & [pro_produccion].[numdoc] AS NumeroDocumentoParte, alm_inventario.codpro AS CodigoItemParte, alm_inventario.descripcion AS ItemParte, pro_producciondet.cantidad AS CantidadParte, alm_ingreso_1.fchdoc AS FechaMovimientoParte, [alm_ingreso_1].[numser] & '-' & [alm_ingreso_1].[numdoc] AS NumeroDocumentoMovimientoParte, alm_ingresodet_1.cantidad AS CantidadMovimientoParte, (con_librocostotemp_1.costoprimo + IIf(con_librocostotemp_1.costomod Is Null Or con_librocostotemp_1.costomod = 0, 0, con_librocostotemp_1.costomod) + IIf(con_librocostotemp_1.costocif Is Null Or con_librocostotemp_1.costocif = 0, 0, con_librocostotemp_1.costocif)) AS CostoParte, con_librocostotemp_1.costounitariopromedio AS CostoUnitarioPromedioParte, pro_ordenprod.fchpro AS FechaOP, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS NumeroDocumentoOP, " _
                & "pro_ordenprod.cantidad AS CantidadOP, pro_solicitudmat.fchdoc AS FechaSM, [pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc] AS NumeroDocumentoSM, alm_ingreso.fching AS FechaMovimiento, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc] AS NumeroDocumentoMovimiento, " _
                & "alm_inventario_1.codpro AS CodigoItemMovimiento, alm_inventario_1.descripcion AS ItemMovimiento, alm_ingresodet.cantidad AS CantidadMovimiento, (con_librocostotemp.costoprimo+ IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS CostoMovimiento, con_librocostotemp.costounitariopromedio AS CostoUnitarioPromedioMov " _
        + vbCr + "FROM alm_ingreso AS alm_ingreso_1 INNER JOIN (pro_produccion INNER JOIN (((((alm_ingresodet AS alm_ingresodet_1 INNER JOIN ((alm_ingreso INNER JOIN (pro_solicitudmat INNER JOIN (pro_producciondet INNER JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id) ON pro_solicitudmat.iddocref = pro_ordenprod.id) ON alm_ingreso.iddocref = pro_solicitudmat.id) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) ON alm_ingresodet_1.iddocref = pro_producciondet.idproddet) INNER JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) INNER JOIN alm_inventario AS alm_inventario_1 ON alm_ingresodet.iditem = alm_inventario_1.id) INNER JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) INNER JOIN con_librocostotemp AS con_librocostotemp_1 ON alm_ingresodet_1.idmovdet = con_librocostotemp_1.idmovdet) ON pro_produccion.id = pro_producciondet.idpro) ON alm_ingreso_1.id = alm_ingresodet_1.id " _
        + vbCr + "WHERE (((pro_produccion.fchdoc)>=CDate('" & FchInicio & "') And (pro_produccion.fchdoc)<=CDate('" & FchFinal & "')) AND ((pro_solicitudmat.idtipdocref)=115) AND ((alm_ingreso.idtipdocref)=110)) " _
        + vbCr + "ORDER BY pro_produccion.fchdoc, [pro_produccion].[numser] & '-' & [pro_produccion].[numdoc]"
    
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
    fValidarDatos = True
End Function
